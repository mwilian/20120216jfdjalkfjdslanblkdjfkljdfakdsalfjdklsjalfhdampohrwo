using System;
using FlexCel.Core;
using System.Collections.Generic;

namespace FlexCel.XlsAdapter
{
    /// <summary>
    /// Horizontal Page break.
    /// </summary>
    internal class TBiff8HPageBreakRecord: TxBaseRecord
    {
        private const int SizeOfBreak=6;
        internal TBiff8HPageBreakRecord(int aId, byte[] aData): base(aId, aData){}

        internal int Count
        {
            get
            {
                return GetWord(0);
            }
        }
     
        internal THPageBreak CreatePageBreak(int Index)
        {
            return new THPageBreak(GetWord(2+Index*SizeOfBreak), GetWord(4+Index*SizeOfBreak), GetWord(6+Index*SizeOfBreak), true, false);
        }
    
        internal int Row(int Index) 
        {
            return GetWord(2+Index*SizeOfBreak);
        }

		internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
		{
            if (Loader.CustomView != null) Loader.CustomView.HPageBreaks.AddRecord(this);
            else ws.SheetGlobals.HPageBreaks.AddRecord(this);
		}

    }

    /// <summary>
    /// Vertical Page break.
    /// </summary>
    internal class TBiff8VPageBreakRecord: TxBaseRecord
    {
        private const int SizeOfBreak=6;
        internal TBiff8VPageBreakRecord(int aId, byte[] aData): base(aId, aData){}

        internal int Count
        {
            get
            {
                return GetWord(0);
            }
        }
     
        internal TVPageBreak CreatePageBreak(int Index)
        {
            return new TVPageBreak(GetWord(2+Index*SizeOfBreak), GetWord(4+Index*SizeOfBreak), GetWord(6+Index*SizeOfBreak), true, false);
        }    

        internal int Col(int Index) 
        {
            return GetWord(2+Index*SizeOfBreak);
        }

		internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
		{
			ws.SheetGlobals.VPageBreaks.AddBiff8Record(this);		
		}

    }

   
    internal abstract class TPageBreak
    {
        internal bool GoesAfter; //for compatibility with FlexCelReports, where <#page breaks> add a page break *after* the tag.

        internal TPageBreak(bool aGoesAfter)
        {
            GoesAfter = aGoesAfter;
        }

        internal abstract byte[] Biff8Data();
        internal static int Biff8Length{get {return 6;}}    
        protected abstract int GetSearchId();
        protected abstract void SetSearchId(int value);
        internal int SearchId{get {return GetSearchId();} set{ SetSearchId(value);}}

        internal abstract int Id { get; }
        internal abstract int Min { get; }
        internal abstract int Max { get; }
        }

    internal class THPageBreak: TPageBreak, IComparable
    {
        internal int Row;
        internal int Col1;
        internal int Col2;

        internal THPageBreak(int aRow, int aCol1, int aCol2, bool Expand, bool aGoesAfter): base(aGoesAfter)
        {
            Row=aRow;
            Col1=aCol1;
            if (Expand)
            {
                Col2= Biff8Utils.ExpandBiff8Col(aCol2);
            }
            else Col2 = aCol2;
        }

        internal THPageBreak CopyTo()
        {
            return new THPageBreak(Row, Col1, Col2, false, GoesAfter);
        }

        internal override byte[] Biff8Data()
        {
            byte[] Result= new byte[6];
            Biff8Utils.CheckRow(Row);
            Biff8Utils.CheckCol(Col1);
            int Col3 = Col2 <= FlxConsts.Max_Columns97_2003? Col2: FlxConsts.Max_Columns97_2003;  //Col can be stored up to 16384
            int C2 = Biff8Utils.CheckAndContractBiff8Col(Col3);
            BitOps.SetWord(Result, 0, Row);
            BitOps.SetWord(Result, 2, Col1);
            BitOps.SetWord(Result, 4, C2);
            return Result;
        }

        protected override int GetSearchId()
        {
            return Row;
        }

        protected override void SetSearchId(int value)
        {
            Row=value;
        }

        internal override int Id
        {
            get { return Row; }
        }

        internal override int Min
        {
            get { return Col1; }
        }

        internal override int Max
        {
            get { return Col2; }
        }

        #region IComparable Members

        public int CompareTo(object obj)
        {
            THPageBreak p2=(THPageBreak)obj;
            if (Row<p2.Row) return -1; else if (Row>p2.Row) return 1; else
            return 0;
        }

        #endregion
    }

    internal class TVPageBreak: TPageBreak, IComparable
    {
        internal int Col;
        internal int Row1;
        internal int Row2;

        internal TVPageBreak(int aCol, int aRow1, int aRow2, bool Expand, bool aGoesAfter): base(aGoesAfter)
        {
            Col=aCol;
            Row1 = aRow1;
            if (Expand)
            {
                Row2 = Biff8Utils.ExpandBiff8Row(aRow2);
            }
            else Row2 = aRow2;
        }

        internal TVPageBreak CopyTo()
        {
            return new TVPageBreak(Col, Row1, Row2, false, GoesAfter);

        }
        internal override byte[] Biff8Data()
        {
            byte[] Result= new byte[6];
            Biff8Utils.CheckCol(Col);
            Biff8Utils.CheckRow(Row1);
            int R2 = Biff8Utils.CheckAndContractBiff8Row(Row2);
            BitOps.SetWord(Result, 0, Col);
            BitOps.SetWord(Result, 2, Row1);
            BitOps.SetWord(Result, 4, R2);
            return Result;
        }

        internal override int Id
        {
            get { return Col; }
        }

        internal override int Min
        {
            get { return Row1; }
        }

        internal override int Max
        {
            get { return Row2; }
        }

        #region IComparable Members

        public int CompareTo(object obj)
        {
            TVPageBreak p2=(TVPageBreak)obj;
            if (Col<p2.Col) return -1; else if (Col>p2.Col) return 1; else
                return 0;
        }

        protected override int GetSearchId()
        {
            return Col;
        }

        protected override void SetSearchId(int value)
        {
            Col=value;
        }

        #endregion
    }

    interface IPageBreakList
    {
        void AddBreak(int RowCol, int MinColRow, int MaxColRow, bool aGoesAfter);
        int RealCount();
        TPageBreak GetItem(int i);
    }

    /// <summary>
    /// A base class for storing page breaks.
    /// </summary>
    internal abstract class TPageBreakList<T>: IPageBreakList where T: TPageBreak
    {
        protected int MaxPageBreaks;

        protected int RecordId=0;
        protected bool FSorted;
        protected T SearchRecord=null;

        protected List<T> FList;
        internal TPageBreakList(int aMaxPageBreaks)
        {
            MaxPageBreaks = aMaxPageBreaks;
            FSorted = false;
            FList = new List<T>();
        }

        protected void SetThis(T value, int index)
        {
            FList[index]=value;
            FSorted=false;  //When we add the list gets unsorted
        }

        internal T this[int index] 
        {
            get {return FList[index];} 
            set {SetThis(value, index);}
        }

        public TPageBreak GetItem(int index)
        {
            return FList[index];
        }

        public int RealCount()
        {
            return FList.Count; 
        }

        internal int VirtualCount
        {
            get { return FList.Count < MaxPageBreaks ? FList.Count : MaxPageBreaks; }
        }

        internal void Clear()
        {
            FList.Clear();
        }

        internal void Sort()
        {
            FList.Sort();
            FSorted=true;
        }
        
        internal bool Find(int aSearchId, ref int Index)
        {
            if (!FSorted) Sort();
            SearchRecord.SearchId=aSearchId;
            Index=FList.BinarySearch(0, FList.Count, SearchRecord, null);  //Only BinarySearch compatible with CF.
            bool Result=Index>=0;
            if (Index<0) Index=~Index;
            return Result;
        }

        private void SaveToStreamExt(IDataStream DataStream, TSaveData SaveData, int FirstRecord, int RecordCount)
        {
            if (RecordCount > 0)
            {
                Sort(); //just in case...
                int MyRecordCount = RecordCount;

                DataStream.WriteHeader((UInt16)RecordId, (UInt16)(2 + RecordCount * TPageBreak.Biff8Length));
                DataStream.Write16((UInt16)MyRecordCount);
                for (int i = FirstRecord; i < FirstRecord + RecordCount; i++)
                {
                    DataStream.Write(FList[i].Biff8Data(), TPageBreak.Biff8Length);
                }
            }
        }

        protected virtual void CalcIncludedRangeRecords(TXlsCellRange CellRange, ref int FirstRecord, ref int RecordCount)
        {
            Sort(); //just in case
            FirstRecord = -1;
        }

        internal static long TotalSizeExt(int RecordCount)
        {
            if (RecordCount == 0) return 0;
            else return XlsConsts.SizeOfTRecordHeader + 2 + RecordCount * TPageBreak.Biff8Length;
        }

        internal void SaveToStream(IDataStream DataStream, TSaveData SaveData, TXlsCellRange CellRange)
        {
            if (CellRange == null) SaveAllToStream(DataStream, SaveData); else SaveRangeToStream(DataStream, SaveData, CellRange);
        }

        internal void SaveAllToStream(IDataStream DataStream, TSaveData SaveData)
        {
			if (RealCount() > MaxPageBreaks)
			{
				if (SaveData.ThrowExceptionOnTooManyPageBreaks)
					XlsMessages.ThrowException(XlsErr.ErrTooManyPageBreaks);
				
				if (FlexCelTrace.Enabled) FlexCelTrace.Write(new TXlsTooManyPageBreaksError(XlsMessages.GetString(XlsErr.ErrTooManyPageBreaks), DataStream.FileName));
			}

            SaveToStreamExt(DataStream, SaveData, 0, VirtualCount);
        }

        internal void SaveRangeToStream(IDataStream DataStream, TSaveData SaveData, TXlsCellRange CellRange)
        {
            int FirstRecord=0;
            int RecordCount=0;
            CalcIncludedRangeRecords(CellRange, ref FirstRecord, ref RecordCount);

			if (RecordCount > MaxPageBreaks)
			{
				if (SaveData.ThrowExceptionOnTooManyPageBreaks)
					XlsMessages.ThrowException(XlsErr.ErrTooManyPageBreaks);
				else
					RecordCount = MaxPageBreaks;
				if (FlexCelTrace.Enabled) FlexCelTrace.Write(new TXlsTooManyPageBreaksError(XlsMessages.GetString(XlsErr.ErrTooManyPageBreaks), DataStream.FileName));
			}

            SaveToStreamExt(DataStream, SaveData, FirstRecord, RecordCount);
        }

        internal long TotalSize(TXlsCellRange CellRange)
        {
            if (CellRange == null) return TotalSizeAll();
            return TotalRangeSize(CellRange);
        }
        internal long TotalSizeAll()
        {
            return TotalSizeExt(VirtualCount);
        }

        internal long TotalRangeSize(TXlsCellRange CellRange)
        {
            int FirstRecord=0;
            int RecordCount=0;
            CalcIncludedRangeRecords(CellRange, ref FirstRecord, ref RecordCount);

            if (RecordCount > MaxPageBreaks)
                RecordCount = MaxPageBreaks;

            return TotalSizeExt(RecordCount);
        }

        public abstract void AddBreak(int RowCol, int MinColRow, int MaxColRow, bool aGoesAfter);
    }

    /// <summary>
    /// List of Horizontal Page breaks.
    /// </summary>
    internal class THPageBreakList: TPageBreakList<THPageBreak>
    {
        internal THPageBreakList() : base(XlsConsts.MaxHPageBreaks)
        {
            RecordId=(int)xlr.HORIZONTALPAGEBREAKS;
            SearchRecord=new THPageBreak(0,0,0, false, false);
        }

        protected override void CalcIncludedRangeRecords(TXlsCellRange CellRange, ref int FirstRecord, ref int RecordCount)
        {
            base.CalcIncludedRangeRecords (CellRange, ref FirstRecord, ref RecordCount);
            int LastRecord=-1;
            for (int i=0; i< VirtualCount;i++)
            {
                if ((FirstRecord<0) && (FList[i].Row>=CellRange.Top)) FirstRecord=i;
                if (FList[i].Row<=CellRange.Bottom) LastRecord=i;
            }
            if ((FirstRecord>=0) && (LastRecord>=0) && (FirstRecord<=LastRecord))
                RecordCount=LastRecord-FirstRecord+1;
            else
            {
                FirstRecord=0;
                RecordCount=0;
            }
        }

        internal void AddRecord(TBiff8HPageBreakRecord aRecord)
        {
            int Index = -1;
            for (int i = 0; i < aRecord.Count; i++)
                if (!Find(aRecord.Row(i), ref Index)) FList.Insert(Index, aRecord.CreatePageBreak(i));
        }
        
        internal void AddBreak(int aRow, bool aGoesAfter)
        {
            int Index=-1;
            if (!Find(aRow, ref Index)) FList.Insert(Index, new THPageBreak(aRow, 0, FlxConsts.Max_Columns, false, aGoesAfter));
        }

        public override void AddBreak(int RowCol, int MinColRow, int MaxColRow, bool aGoesAfter)
        {
            FList.Add(new THPageBreak(RowCol, MinColRow, MaxColRow, false, aGoesAfter));
        }

        internal void DeleteBreak(int aRow)
        {
            int Index=-1;
            if (Find(aRow, ref Index)) FList.RemoveAt(Index);
        }

        internal void CopyFrom(THPageBreakList aBreakList)
        {
            if (aBreakList == null) return;
            if (aBreakList.FList == FList) XlsMessages.ThrowException(XlsErr.ErrInternal); //Should be different objects
            for (int i = 0; i < aBreakList.RealCount(); i++)
                FList.Add(aBreakList[i].CopyTo());
        }

        internal void InsertRows(int DestRow, int aCount)
        {
            int Index = -1;
            Find(DestRow, ref Index);
            if (Index < RealCount() && Index >= 0 && FList[Index].GoesAfter && FList[Index].Row == DestRow) Index++;
            for (int i = Index; i < RealCount(); i++)
            {
                BitOps.IncWord(ref FList[i].Row, aCount, FlxConsts.Max_Rows, XlsErr.ErrTooManyRows);
            }
        }

        internal void DeleteRows(int DestRow, int aCount)
        {
            int Index = -1;
            Find(DestRow, ref Index);
            if (Index < RealCount() && Index >= 0 && FList[Index].GoesAfter && FList[Index].Row == DestRow) Index++;

            for (int i = RealCount() - 1; i >= Index; i--)
            {
                if (FList[i].Row < DestRow + aCount) FList.RemoveAt(i); else FList[i].Row -= aCount;
            }
        }

        internal bool HasPageBreak(int Row)
        {
            int Index=-1;
            return Find(Row, ref Index);
        }

    }

    /// <summary>
    /// List of Vertical Page breaks.
    /// </summary>
    internal class TVPageBreakList: TPageBreakList<TVPageBreak>
    {
        internal TVPageBreakList() : base(XlsConsts.MaxVPageBreaks)
        {
            RecordId=(int)xlr.VERTICALPAGEBREAKS;
            SearchRecord = new TVPageBreak(0, 0, 0, false, false);
        }

        protected override void CalcIncludedRangeRecords(TXlsCellRange CellRange, ref int FirstRecord, ref int RecordCount)
        {
            base.CalcIncludedRangeRecords (CellRange, ref FirstRecord, ref RecordCount);
            int LastRecord=-1;
            for (int i=0; i< VirtualCount;i++)
            {
                if ((FirstRecord < 0) && (FList[i].Col >= CellRange.Left)) FirstRecord = i;
                if (FList[i].Col <= CellRange.Right) LastRecord = i;
            }
            if ((FirstRecord>=0) && (LastRecord>=0) && (FirstRecord<=LastRecord))
                RecordCount=LastRecord-FirstRecord+1;
            else
            {
                FirstRecord=0;
                RecordCount=0;
            }
        }

        internal void AddBiff8Record(TBiff8VPageBreakRecord aRecord)
        {
            int Index=-1;
            for (int i=0; i< aRecord.Count;i++)
                if (!Find(aRecord.Col(i), ref Index)) FList.Insert(Index, aRecord.CreatePageBreak(i));
        }
        
        internal void AddBreak(int aCol, bool aGoesAfter)
        {
            int Index=-1;
            if (!Find(aCol, ref Index)) FList.Insert(Index, new TVPageBreak(aCol,0,FlxConsts.Max_Rows, false, aGoesAfter));
        }

        public override void AddBreak(int RowCol, int MinColRow, int MaxColRow, bool aGoesAfter)
        {
            FList.Add(new TVPageBreak(RowCol, MinColRow, MaxColRow, false, aGoesAfter));
        }

        internal void DeleteBreak(int aCol)
        {
            int Index=-1;
            if (Find(aCol, ref Index)) FList.RemoveAt(Index);
        }

        internal void CopyFrom(TVPageBreakList aBreakList)
        {
            if (aBreakList == null) return;
            if (aBreakList.FList == FList) XlsMessages.ThrowException(XlsErr.ErrInternal); //Should be different objects

            for (int i = 0; i < aBreakList.RealCount(); i++)
                FList.Add(aBreakList[i].CopyTo());

        }

        internal void InsertCols(int DestCol, int aCount)
        {
            int Index = -1;
            Find(DestCol, ref Index);
            if (Index < RealCount() && Index >= 0 && FList[Index].GoesAfter && FList[Index].Col == DestCol) Index++;
            for (int i = Index; i < RealCount(); i++)
            {
                BitOps.IncWord(ref FList[i].Col, aCount, FlxConsts.Max_Columns, XlsErr.ErrTooManyColumns);
            }
        }

        internal void DeleteCols(int DestCol, int aCount)
        {
            int Index = -1;
            Find(DestCol, ref Index);
            if (Index < RealCount() && Index >= 0 && FList[Index].GoesAfter && FList[Index].Col == DestCol) Index++;

            for (int i = RealCount() - 1; i >= Index; i--)
            {
                if (FList[i].Col < DestCol + aCount) FList.RemoveAt(i); else FList[i].Col -= aCount;
            }
        }

        internal bool HasPageBreak(int Col)
        {
            int Index=-1;
            return Find(Col, ref Index);
        }

    }
}

