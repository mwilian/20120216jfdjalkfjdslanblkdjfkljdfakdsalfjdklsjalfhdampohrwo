using System;
using System.Collections.Generic;
using System.Text;
using System.Data;

namespace DTO
{
	/// <summary> 
	///Author: nnamthach@gmail.com 
	/// <summary>
	
    public class LIST_TASKInfo
    {
		#region Local Variable
		public enum Field
		{
			DTB,
			Code,
			Description,
			Lookup,
			AttQD_ID,
			AttTmp,
			ValidRange,
			CntQD_ID,
			CntTmp,
			Emails,
			Server,
			Protocol,
			Port,
			UserID,
			Password,
			IsUse
		}
		private String _DTB;
		private String _Code;
		private String _Description;
		private String _Lookup;
		private String _AttQD_ID;
		private String _AttTmp;
		private String _ValidRange;
		private String _CntQD_ID;
		private String _CntTmp;
		private String _Emails;
		private String _Server;
		private String _Protocol;
		private String _Port;
		private String _UserID;
		private String _Password;
		private String _IsUse;
		
		public String DTB{	get{ return _DTB;} set{_DTB = value;} }
		public String Code{	get{ return _Code;} set{_Code = value;} }
		public String Description{	get{ return _Description;} set{_Description = value;} }
		public String Lookup{	get{ return _Lookup;} set{_Lookup = value;} }
		public String AttQD_ID{	get{ return _AttQD_ID;} set{_AttQD_ID = value;} }
		public String AttTmp{	get{ return _AttTmp;} set{_AttTmp = value;} }
		public String ValidRange{	get{ return _ValidRange;} set{_ValidRange = value;} }
		public String CntQD_ID{	get{ return _CntQD_ID;} set{_CntQD_ID = value;} }
		public String CntTmp{	get{ return _CntTmp;} set{_CntTmp = value;} }
		public String Emails{	get{ return _Emails;} set{_Emails = value;} }
		public String Server{	get{ return _Server;} set{_Server = value;} }
		public String Protocol{	get{ return _Protocol;} set{_Protocol = value;} }
		public String Port{	get{ return _Port;} set{_Port = value;} }
		public String UserID{	get{ return _UserID;} set{_UserID = value;} }
		public String Password{	get{ return _Password;} set{_Password = value;} }
		public String IsUse{	get{ return _IsUse;} set{_IsUse = value;} }
		
        #endregion LocalVariable
        
        #region Constructor
		public LIST_TASKInfo()
		{
			_DTB = "";
			_Code = "";
			_Description = "";
			_Lookup = "";
			_AttQD_ID = "";
			_AttTmp = "";
			_ValidRange = "";
			_CntQD_ID = "";
			_CntTmp = "";
			_Emails = "";
			_Server = "";
			_Protocol = "";
			_Port = "";
			_UserID = "";
			_Password = "";
			_IsUse = "";
		}
		public LIST_TASKInfo(
		String DTB,
		String Code,
		String Description,
		String Lookup,
		String AttQD_ID,
		String AttTmp,
		String ValidRange,
		String CntQD_ID,
		String CntTmp,
		String Emails,
		String Server,
		String Protocol,
		String Port,
		String UserID,
		String Password,
		String IsUse
		)
		{
			_DTB = DTB;
			_Code = Code;
			_Description = Description;
			_Lookup = Lookup;
			_AttQD_ID = AttQD_ID;
			_AttTmp = AttTmp;
			_ValidRange = ValidRange;
			_CntQD_ID = CntQD_ID;
			_CntTmp = CntTmp;
			_Emails = Emails;
			_Server = Server;
			_Protocol = Protocol;
			_Port = Port;
			_UserID = UserID;
			_Password = Password;
			_IsUse = IsUse;
		}
		public LIST_TASKInfo(DataRow dr)
		{
			if (dr != null)
			{
				_DTB = dr[Field.DTB.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.DTB.ToString()]);
				_Code = dr[Field.Code.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.Code.ToString()]);
				_Description = dr[Field.Description.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.Description.ToString()]);
				_Lookup = dr[Field.Lookup.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.Lookup.ToString()]);
				_AttQD_ID = dr[Field.AttQD_ID.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.AttQD_ID.ToString()]);
				_AttTmp = dr[Field.AttTmp.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.AttTmp.ToString()]);
				_ValidRange = dr[Field.ValidRange.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.ValidRange.ToString()]);
				_CntQD_ID = dr[Field.CntQD_ID.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.CntQD_ID.ToString()]);
				_CntTmp = dr[Field.CntTmp.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.CntTmp.ToString()]);
				_Emails = dr[Field.Emails.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.Emails.ToString()]);
				_Server = dr[Field.Server.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.Server.ToString()]);
				_Protocol = dr[Field.Protocol.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.Protocol.ToString()]);
				_Port = dr[Field.Port.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.Port.ToString()]);
				_UserID = dr[Field.UserID.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.UserID.ToString()]);
				_Password = dr[Field.Password.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.Password.ToString()]);
				_IsUse = dr[Field.IsUse.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.IsUse.ToString()]);
			}
		}
		public LIST_TASKInfo(LIST_TASKInfo objEntr)
		{			
			_DTB = objEntr.DTB;			
			_Code = objEntr.Code;			
			_Description = objEntr.Description;			
			_Lookup = objEntr.Lookup;			
			_AttQD_ID = objEntr.AttQD_ID;			
			_AttTmp = objEntr.AttTmp;			
			_ValidRange = objEntr.ValidRange;			
			_CntQD_ID = objEntr.CntQD_ID;			
			_CntTmp = objEntr.CntTmp;			
			_Emails = objEntr.Emails;			
			_Server = objEntr.Server;			
			_Protocol = objEntr.Protocol;			
			_Port = objEntr.Port;			
			_UserID = objEntr.UserID;			
			_Password = objEntr.Password;			
			_IsUse = objEntr.IsUse;			
		}
        #endregion Constructor
        
        #region InitTable
		public static DataTable ToDataTable()
		{
			DataTable dt = new DataTable("LIST_TASK");
			dt.Columns.AddRange(new DataColumn[] { 
				new DataColumn(Field.DTB.ToString(), typeof(String)),
				new DataColumn(Field.Code.ToString(), typeof(String)),
				new DataColumn(Field.Description.ToString(), typeof(String)),
				new DataColumn(Field.Lookup.ToString(), typeof(String)),
				new DataColumn(Field.AttQD_ID.ToString(), typeof(String)),
				new DataColumn(Field.AttTmp.ToString(), typeof(String)),
				new DataColumn(Field.ValidRange.ToString(), typeof(String)),
				new DataColumn(Field.CntQD_ID.ToString(), typeof(String)),
				new DataColumn(Field.CntTmp.ToString(), typeof(String)),
				new DataColumn(Field.Emails.ToString(), typeof(String)),
				new DataColumn(Field.Server.ToString(), typeof(String)),
				new DataColumn(Field.Protocol.ToString(), typeof(String)),
				new DataColumn(Field.Port.ToString(), typeof(String)),
				new DataColumn(Field.UserID.ToString(), typeof(String)),
				new DataColumn(Field.Password.ToString(), typeof(String)),
				new DataColumn(Field.IsUse.ToString(), typeof(String))
			});
			return dt;
		}
		public DataRow ToDataRow(DataTable dt)
		{
			DataRow row = dt.NewRow();
			row[Field.DTB.ToString()] = _DTB;
			row[Field.Code.ToString()] = _Code;
			row[Field.Description.ToString()] = _Description;
			row[Field.Lookup.ToString()] = _Lookup;
			row[Field.AttQD_ID.ToString()] = _AttQD_ID;
			row[Field.AttTmp.ToString()] = _AttTmp;
			row[Field.ValidRange.ToString()] = _ValidRange;
			row[Field.CntQD_ID.ToString()] = _CntQD_ID;
			row[Field.CntTmp.ToString()] = _CntTmp;
			row[Field.Emails.ToString()] = _Emails;
			row[Field.Server.ToString()] = _Server;
			row[Field.Protocol.ToString()] = _Protocol;
			row[Field.Port.ToString()] = _Port;
			row[Field.UserID.ToString()] = _UserID;
			row[Field.Password.ToString()] = _Password;
			row[Field.IsUse.ToString()] = _IsUse;
			return row;
		}
        #endregion InitTable
    }
}
