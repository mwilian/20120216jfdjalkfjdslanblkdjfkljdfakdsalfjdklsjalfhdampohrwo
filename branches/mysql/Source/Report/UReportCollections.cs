#region Using directives

using System;
using System.Globalization;
using FlexCel.Core;
using System.Collections.Generic;

using System.Text;
using System.Data;
using System.Data.Common;
using System.Collections;

#endregion

namespace FlexCel.Report
{

	internal class TConfigFormat
	{
		internal int XF;
		internal TFlxApplyFormat ApplyFmt;
		internal bool ExteriorBorders;

		internal TConfigFormat(int aXF, TFlxApplyFormat aApplyFmt, bool aExteriorBorders)
		{
			XF = aXF;
			ApplyFmt = aApplyFmt;
			ExteriorBorders = aExteriorBorders;
		}
	}

#if(FRAMEWORK20)
    internal sealed class TFormatList : Dictionary<string, TConfigFormat>
    {
        public TFormatList(): base(StringComparer.InvariantCultureIgnoreCase)
        {
        }

        public TConfigFormat GetValue(string Key)
        {
			TConfigFormat Result;
	        if (!TryGetValue(Key, out Result))
				FlxMessages.ThrowException(FlxErr.ErrInvalidFormat, Key);
            return Result;

        }
    }

    internal sealed class TValueList : Dictionary<string, object>
    {
        public TValueList(): base(StringComparer.InvariantCultureIgnoreCase)
        {
        }
    }

    internal sealed class TUserFunctionList : Dictionary<string, TFlexCelUserFunction>
    {
        public TUserFunctionList(): base(StringComparer.InvariantCultureIgnoreCase)
        {
        }
    }

    internal sealed class TExpressionList : Dictionary<string, TExpression>
    {
        public TExpressionList(): base(StringComparer.InvariantCultureIgnoreCase)
        {
        }
    }

    internal sealed class TUsedRefs : Dictionary<string, string>
    {
        public TUsedRefs(): base(StringComparer.InvariantCultureIgnoreCase)
        {
        }
    }

#else
    internal class TFormatList: Hashtable
    {
        public TFormatList(): base(FormatComparer.HashProvider, FormatComparer.HashComparer)
        {
        }

        public TConfigFormat GetValue(string Key)
        {
            TConfigFormat obj=this[Key] as TConfigFormat;
            if (obj==null) FlxMessages.ThrowException(FlxErr.ErrInvalidFormat, Key);
            return obj;
        }
    }

    internal class TValueList: Hashtable
    {
        public TValueList(): base(FormatComparer.HashProvider, FormatComparer.HashComparer)
        {
        }

    }

    internal class TUserFunctionList: Hashtable
    {
        public TUserFunctionList(): base(FormatComparer.HashProvider, FormatComparer.HashComparer)
        {
        }
    }

    internal class TExpressionList: Hashtable
    {
        public TExpressionList(): base(FormatComparer.HashProvider, FormatComparer.HashComparer)
        {
        }

		public TExpression this[string key]
		{
			get
			{
				return (TExpression) base[key];
			}
		}
    }

    internal class TUsedRefs: Hashtable
    {
        public TUsedRefs(): base(FormatComparer.HashProvider, FormatComparer.HashComparer)
        {
        }
    }

#endif

	internal class TStackData
	{
		internal TUsedRefs UsedRefs;
		internal TValueList ExpParams;

		internal TStackData(TUsedRefs aUsedRefs, TValueList aExpParams)
		{
			UsedRefs = aUsedRefs;
			ExpParams = aExpParams;
		}
	}



	#region Adapter
	internal class TAdapterData
	{
		internal IDbDataAdapter Adapter;
		internal CultureInfo Locale;
		internal bool CaseSensitive;

		internal TAdapterData(IDbDataAdapter aAdapter, CultureInfo aLocale, bool aCaseSensitive)
		{
			Adapter = aAdapter;
			Locale = aLocale;
			CaseSensitive = aCaseSensitive;
		}
	}

    internal class TDataAdapterList : IEnumerable<string>
    {
        #region Privates
#if (FRAMEWORK20)
        Dictionary<string, TAdapterData> FList = null;
#else
		Hashtable FList=null;
#endif
        #endregion

        internal TDataAdapterList()
        {
            FList = new Dictionary<string, TAdapterData>(StringComparer.InvariantCultureIgnoreCase);
        }

        internal void Clear()
        {
            FList.Clear();
        }

        internal void Add(string dtName, TAdapterData dt)
        {
            FList[dtName] = dt;
        }

        internal TAdapterData this[string key]
        {
            get
            {
                TAdapterData Result = null;
                if (FList.TryGetValue(key, out Result))
                    return Result;
                return null;
            }
        }

        internal int Count
        {
            get
            {
                return FList.Count;
            }
        }

        #region IEnumerable<string> Members

        IEnumerator<string> IEnumerable<string>.GetEnumerator()
        {
            return FList.Keys.GetEnumerator();
        }

        #endregion

        #region IEnumerable Members

        IEnumerator IEnumerable.GetEnumerator()
        {
            return FList.Keys.GetEnumerator();
        }

        #endregion
    }
	#endregion

	#region SQLParameters
    internal class TSqlParameterList : IEnumerable, IEnumerable<string>
    {
        #region Privates
#if (FRAMEWORK20)
        Dictionary<string, IDbDataParameter> FList = null;
#else
		Hashtable FList=null;
#endif
        #endregion

        internal TSqlParameterList()
        {
            FList = new Dictionary<string, IDbDataParameter>(StringComparer.InvariantCultureIgnoreCase);
        }

        internal void Clear()
        {
            FList.Clear();
        }

        internal void Add(string dtName, IDbDataParameter dt)
        {
            FList[dtName] = dt;
        }

        internal IDbDataParameter this[string key]
        {
            get
            {
                IDbDataParameter Result = null;
                if (!FList.TryGetValue(key, out Result))
                    FlxMessages.ThrowException(FlxErr.ErrSqlParameterNotFound, key);
                return Result;
            }
        }

        internal int Count
        {
            get
            {
                return FList.Count;
            }
        }

        #region IEnumerable<string> Members

        public IEnumerator<string> GetEnumerator()
        {
            return FList.Keys.GetEnumerator();
        }

        #endregion


        #region IEnumerable Members

        IEnumerator IEnumerable.GetEnumerator()
        {
            return FList.Keys.GetEnumerator();
        }

        #endregion
    }
	#endregion

    internal class TAddedRowColList  
    {
        SortedList<int, int> FList;

        public TAddedRowColList(int start)
        {
            FList = new SortedList<int,int>();
            FList.Add(start, 0);
        }

        public void Add(int StartCell, int EndCell, int Count)
        {
            if (Count == 0) return;
            if (!FList.ContainsKey(EndCell + 1)) //we need to add a new value at ixend
            {
                FList.Add(EndCell + 1, 0);
                int ixLast = FList.IndexOfKey(EndCell + 1);
                if (ixLast > 0) FList[EndCell + 1] = FList.Values[ixLast - 1];
            }

            int StartCellCount = 0;
            bool StartCellExists = FList.TryGetValue(StartCell, out StartCellCount);
            if (StartCellExists) FList[StartCell] += Count; else FList[StartCell] = Count;


            int ix = FList.IndexOfKey(StartCell); //always exists, as we added it above.
            for (int i = ix + 1; i < FList.Count; i++)
            {
                if (FList.Keys[i] > EndCell) break;
                FList[FList.Keys[i]] += Count;
            }
        }

        public int Max(int EndCell)
        {
            if (FList.Count == 0) return 0;
            int Result = FList.Values[0];
            for (int i = 1; i < FList.Count; i++)
            {
                if (FList.Values[i] > Result && FList.Keys[i] < EndCell) Result = FList.Values[i];  //FList.Key[i] will never be less than the start of the range, since ranges don't intersect. It might be more than the end though, +1.
            }

            return Result;
        }

        public IList<int> Cells { get { return FList.Keys; } }
        public IList<int> InsertedCount { get { return FList.Values; } }

        public int Count { get { return FList.Count; } }
    }

}
