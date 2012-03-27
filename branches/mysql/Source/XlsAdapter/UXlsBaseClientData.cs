using System;
using System.IO;
using FlexCel.Core;

namespace FlexCel.XlsAdapter
{
    internal enum TClientType { Null, TTXO, TMsObj}

	/// <summary>
	/// Abstract class holding either a TextBox (TXO) or an ole Object (OBJ) on its 2 descendants.
	/// It gives a common interface to handle embedded objects.
	/// </summary>
	internal abstract class TBaseClientData 
	{
		protected virtual int GetId() {return 0;}
		protected virtual void SetId(int Value){}
		
		internal TBaseRecord RemainingData;

		internal int Id{ get {return GetId();} set {SetId(value);}}
		internal virtual void ArrangeId(ref int MaxId){}

		internal abstract void Clear();
		protected abstract TBaseClientData DoCopyTo(TSheetInfo SheetInfo);
		internal static TBaseClientData Clone(TBaseClientData Self, TSheetInfo SheetInfo) 
		{
			if (Self!=null) return Self.DoCopyTo(SheetInfo); else return null;
		}
		
		internal abstract void LoadFromStream(TBaseRecordLoader RecordLoader, TWorkbookGlobals Globals, TBaseRecord First);
		internal abstract void SaveToStream(IDataStream DataStream, TSaveData SaveData);
		internal abstract long TotalSize();

		internal virtual void ArrangeCopyRange(int RowOfs, int ColOfs, TSheetInfo SheetInfo){}
		internal abstract void ArrangeInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo);
		internal abstract void ArrangeMoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo);
		internal abstract void ArrangeCopySheet(TSheetInfo SheetInfo);
		internal abstract void UpdateDeletedRanges(TDeletedRanges DeletedRanges);

        internal virtual bool HasExternRefs()
        {
            return false;
        }

		internal virtual TClientType ObjRecord()
		{
			return TClientType.Null;
		}

		internal virtual TFlxChart Chart()
		{
			return null;
		}

	}

}
