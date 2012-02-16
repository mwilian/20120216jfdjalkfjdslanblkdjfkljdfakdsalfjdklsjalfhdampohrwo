using System;
using System.Collections.Generic;
using System.Text;
using System.Data;

namespace DTO
{
	/// <summary> 
	///Author: nnamthach@gmail.com 
	/// <summary>
	
    public class LIST_GADGETInfo
    {
		#region Local Variable
		public enum Field
		{
			ID,
			Description,
			QDCode,
			ReportTmp,
			AutoUpdate,
			IsScroll,
			Action,
			Argument,
			Image
		}
		private Int32 _ID;
		private String _Description;
		private String _QDCode;
		private String _ReportTmp;
		private Int32 _AutoUpdate;
		private Boolean _IsScroll;
		private String _Action;
		private String _Argument;
		private String _Image;
		
		public Int32 ID{	get{ return _ID;} set{_ID = value;} }
		public String Description{	get{ return _Description;} set{_Description = value;} }
		public String QDCode{	get{ return _QDCode;} set{_QDCode = value;} }
		public String ReportTmp{	get{ return _ReportTmp;} set{_ReportTmp = value;} }
		public Int32 AutoUpdate{	get{ return _AutoUpdate;} set{_AutoUpdate = value;} }
		public Boolean IsScroll{	get{ return _IsScroll;} set{_IsScroll = value;} }
		public String Action{	get{ return _Action;} set{_Action = value;} }
		public String Argument{	get{ return _Argument;} set{_Argument = value;} }
		public String Image{	get{ return _Image;} set{_Image = value;} }
		
        #endregion LocalVariable
        
        #region Constructor
		public LIST_GADGETInfo()
		{
			_ID = 0;
			_Description = "";
			_QDCode = "";
			_ReportTmp = "";
			_AutoUpdate = 0;
			_IsScroll = true;
			_Action = "";
			_Argument = "";
			_Image = "";
		}
		public LIST_GADGETInfo(
		Int32 ID,
		String Description,
		String QDCode,
		String ReportTmp,
		Int32 AutoUpdate,
		Boolean IsScroll,
		String Action,
		String Argument,
		String Image
		)
		{
			_ID = ID;
			_Description = Description;
			_QDCode = QDCode;
			_ReportTmp = ReportTmp;
			_AutoUpdate = AutoUpdate;
			_IsScroll = IsScroll;
			_Action = Action;
			_Argument = Argument;
			_Image = Image;
		}
		public LIST_GADGETInfo(DataRow dr)
		{
			if (dr != null)
			{
				_ID = dr[Field.ID.ToString()] == DBNull.Value?0:Convert.ToInt32(dr[Field.ID.ToString()]);
				_Description = dr[Field.Description.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.Description.ToString()]);
				_QDCode = dr[Field.QDCode.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.QDCode.ToString()]);
				_ReportTmp = dr[Field.ReportTmp.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.ReportTmp.ToString()]);
				_AutoUpdate = dr[Field.AutoUpdate.ToString()] == DBNull.Value?0:Convert.ToInt32(dr[Field.AutoUpdate.ToString()]);
				_IsScroll = dr[Field.IsScroll.ToString()] == DBNull.Value?true:Convert.ToBoolean(dr[Field.IsScroll.ToString()]);
				_Action = dr[Field.Action.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.Action.ToString()]);
				_Argument = dr[Field.Argument.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.Argument.ToString()]);
				_Image = dr[Field.Image.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.Image.ToString()]);
			}
		}
		public LIST_GADGETInfo(LIST_GADGETInfo objEntr)
		{			
			_ID = objEntr.ID;			
			_Description = objEntr.Description;			
			_QDCode = objEntr.QDCode;			
			_ReportTmp = objEntr.ReportTmp;			
			_AutoUpdate = objEntr.AutoUpdate;			
			_IsScroll = objEntr.IsScroll;			
			_Action = objEntr.Action;			
			_Argument = objEntr.Argument;			
			_Image = objEntr.Image;			
		}
        #endregion Constructor
        
        #region InitTable
		public static DataTable ToDataTable()
		{
			DataTable dt = new DataTable("LIST_GADGET");
			dt.Columns.AddRange(new DataColumn[] { 
				new DataColumn(Field.ID.ToString(), typeof(Int32)),
				new DataColumn(Field.Description.ToString(), typeof(String)),
				new DataColumn(Field.QDCode.ToString(), typeof(String)),
				new DataColumn(Field.ReportTmp.ToString(), typeof(String)),
				new DataColumn(Field.AutoUpdate.ToString(), typeof(Int32)),
				new DataColumn(Field.IsScroll.ToString(), typeof(Boolean)),
				new DataColumn(Field.Action.ToString(), typeof(String)),
				new DataColumn(Field.Argument.ToString(), typeof(String)),
				new DataColumn(Field.Image.ToString(), typeof(String))
			});
			return dt;
		}
		public DataRow ToDataRow(DataTable dt)
		{
			DataRow row = dt.NewRow();
			row[Field.ID.ToString()] = _ID;
			row[Field.Description.ToString()] = _Description;
			row[Field.QDCode.ToString()] = _QDCode;
			row[Field.ReportTmp.ToString()] = _ReportTmp;
			row[Field.AutoUpdate.ToString()] = _AutoUpdate;
			row[Field.IsScroll.ToString()] = _IsScroll;
			row[Field.Action.ToString()] = _Action;
			row[Field.Argument.ToString()] = _Argument;
			row[Field.Image.ToString()] = _Image;
			return row;
		}
        #endregion InitTable
    }
}
