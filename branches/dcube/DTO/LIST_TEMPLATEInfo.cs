using System;
using System.Collections.Generic;
using System.Text;
using System.Data;

namespace DTO
{
    /// <summary> 
    ///Author: nnamthach@gmail.com 
    /// <summary>

    public class LIST_TEMPLATEInfo
    {
        #region Local Variable
        public enum Field
        {
            DTB,
            Code,
            Data,
            Length
        }
        private String _DTB;
        private String _Code;
        private Byte[] _Data;
        private Int32 _Length;

        public String DTB { get { return _DTB; } set { _DTB = value; } }
        public String Code { get { return _Code; } set { _Code = value; } }
        public Byte[] Data { get { return _Data; } set { _Data = value; } }
        public Int32 Length { get { return _Length; } set { _Length = value; } }

        #endregion LocalVariable

        #region Constructor
        public LIST_TEMPLATEInfo()
        {
            _DTB = "";
            _Code = "";
            _Data = null;
            _Length = 0;
        }
        public LIST_TEMPLATEInfo(
        String DTB,
        String Code,
        Byte[] Data,
            Int32 Length
        )
        {
            _DTB = DTB;
            _Code = Code;
            _Data = Data;
            _Length = Length;
        }
        public LIST_TEMPLATEInfo(DataRow dr)
        {
            if (dr != null)
            {
                _DTB = dr[Field.DTB.ToString()] == DBNull.Value ? "" : Convert.ToString(dr[Field.DTB.ToString()]);
                _Code = dr[Field.Code.ToString()] == DBNull.Value ? "" : Convert.ToString(dr[Field.Code.ToString()]);
                _Data = dr[Field.Data.ToString()] == DBNull.Value ? null : (byte[])(dr[Field.Data.ToString()]);
                _Length = dr[Field.Length.ToString()] == DBNull.Value ? 0 : (int)(dr[Field.Length.ToString()]);
            }
        }
        public LIST_TEMPLATEInfo(LIST_TEMPLATEInfo objEntr)
        {
            _DTB = objEntr.DTB;
            _Code = objEntr.Code;
            _Data = objEntr.Data;
            _Length = objEntr.Length;
        }
        #endregion Constructor

        #region InitTable
        public static DataTable ToDataTable()
        {
            DataTable dt = new DataTable("LIST_TEMPLATE");
            dt.Columns.AddRange(new DataColumn[] { 
				new DataColumn(Field.DTB.ToString(), typeof(String)),
				new DataColumn(Field.Code.ToString(), typeof(String)),
				new DataColumn(Field.Data.ToString(), typeof(Byte[])),
                new DataColumn(Field.Length.ToString(), typeof(int))
			});
            return dt;
        }
        public DataRow ToDataRow(DataTable dt)
        {
            DataRow row = dt.NewRow();
            row[Field.DTB.ToString()] = _DTB;
            row[Field.Code.ToString()] = _Code;
            row[Field.Data.ToString()] = _Data;
            row[Field.Length.ToString()] = _Length;
            return row;
        }
        #endregion InitTable
    }
}
