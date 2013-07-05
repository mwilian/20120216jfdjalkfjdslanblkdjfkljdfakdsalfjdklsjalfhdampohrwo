using System;
using System.Collections.Generic;
using System.Text;

namespace dCube
{
    public class objValidatedList : IConvertible
    {
        string _qd = "";
        string _message = "";
        string _field = "";

        public string QD
        {
            get { return _qd; }
            set { _qd = value; }
        }
        public string Field
        {
            get { return _field; }
            set { _field = value; }
        }
        public string Message
        {
            get { return _message; }
            set { _message = value; }
        }

        public objValidatedList(string qd, string fld, string message)
        {
            _qd = qd;
            _field = fld;
            _message = message;
        }
        public objValidatedList(string value)
        {
            InitFromString(value);
        }

        private void InitFromString(string value)
        {
            if (value != "")
            {
                string[] arr = value.Split('.');
                _qd = arr[0].Substring(1, arr[0].Length - 2);
                _field = arr[1].Substring(1, arr[1].Length - 2);
                _message = arr[2].Substring(1, arr[2].Length - 2);
            }
        }
        public objValidatedList(object value)
        {
            if (value is objValidatedList)
            {
                _qd = ((objValidatedList)value)._qd;
                _field = ((objValidatedList)value)._field;
                _message = ((objValidatedList)value)._message;
            }
            else if (value is string)
            {
                InitFromString(value.ToString());
            }

        }

        double GetDoubleValue()
        {
            return 0;
        }
        public TypeCode GetTypeCode()
        {
            return TypeCode.Object;
        }
        bool IConvertible.ToBoolean(IFormatProvider provider)
        {
            return false;
        }
        byte IConvertible.ToByte(IFormatProvider provider)
        {
            return Convert.ToByte(GetDoubleValue());
        }

        char IConvertible.ToChar(IFormatProvider provider)
        {
            return Convert.ToChar(GetDoubleValue());
        }

        DateTime IConvertible.ToDateTime(IFormatProvider provider)
        {
            return Convert.ToDateTime(GetDoubleValue());
        }

        decimal IConvertible.ToDecimal(IFormatProvider provider)
        {
            return Convert.ToDecimal(GetDoubleValue());
        }

        double IConvertible.ToDouble(IFormatProvider provider)
        {
            return GetDoubleValue();
        }

        short IConvertible.ToInt16(IFormatProvider provider)
        {
            return Convert.ToInt16(GetDoubleValue());
        }

        int IConvertible.ToInt32(IFormatProvider provider)
        {
            return Convert.ToInt32(GetDoubleValue());
        }

        long IConvertible.ToInt64(IFormatProvider provider)
        {
            return Convert.ToInt64(GetDoubleValue());
        }

        sbyte IConvertible.ToSByte(IFormatProvider provider)
        {
            return Convert.ToSByte(GetDoubleValue());
        }

        float IConvertible.ToSingle(IFormatProvider provider)
        {
            return Convert.ToSingle(GetDoubleValue());
        }

        string IConvertible.ToString(IFormatProvider provider)
        {
            return String.Format("[{0}].[{1}].[{2}]", _qd, _field, _message);
        }

        object IConvertible.ToType(Type conversionType, IFormatProvider provider)
        {
            return Convert.ChangeType(GetDoubleValue(), conversionType);
        }

        ushort IConvertible.ToUInt16(IFormatProvider provider)
        {
            return Convert.ToUInt16(GetDoubleValue());
        }

        uint IConvertible.ToUInt32(IFormatProvider provider)
        {
            return Convert.ToUInt32(GetDoubleValue());
        }

        ulong IConvertible.ToUInt64(IFormatProvider provider)
        {
            return Convert.ToUInt64(GetDoubleValue());
        }
    }
}
