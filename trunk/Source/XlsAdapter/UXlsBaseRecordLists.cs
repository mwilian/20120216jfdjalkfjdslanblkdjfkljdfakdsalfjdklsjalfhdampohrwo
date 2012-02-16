using System;
using System.Collections.Generic;
using System.IO;
using FlexCel.Core;
using System.Globalization;

namespace FlexCel.XlsAdapter
{
    internal interface ISaveBiff8
    {
        void SaveToStream(IDataStream DataStream, TSaveData SaveData, int Row);
        void SaveToPxl(TPxlStream PxlStream, int Row, TPxlSaveData SaveData);
    }

    /// <summary>
    /// Base list of records. All FlexCel record lists derive from here.
    /// </summary>
    internal class TBaseRecordList<T>: ISaveBiff8 where T:TBaseRecord
    {
        protected bool FSorted;

        protected List<T> FList;
        internal TBaseRecordList()
        {
            FSorted=true; //when empty, it is sorted.
            FList = new List<T>();
        }

        #region Generics
        internal void Add(T a)
        {
            FList.Add(a);
            OnAdd(a, Count - 1);
            if (FSorted && AddUnSorts(a))
                FSorted = false;  //When we add the list gets unsorted. This is optimized so if we are loading an ordered list, we do not need to set Sorted=false.
        }

        internal void Insert (int index, T a)
        {
            FList.Insert(index, a);
            OnAdd(a, index);
            // We assume that when we insert, we respect the order, so we dont set Sorted=false
        }

        protected virtual void SetThis(T value, int index)
        {
            if (FList[index]!=null) OnDelete(FList[index], index);
            FList[index]=value;
            FSorted=false;  //When we add the list gets unsorted
            if (value!=null) OnAdd(value, index);
        }

        internal T this[int index] 
        {
            get {return FList[index];} 
            set {SetThis(value, index);}
        }

        internal virtual void Delete(int Index)
        {
            OnDelete(FList[Index], Index);
            FList.RemoveAt(Index);
        }
        #endregion

        internal virtual void OnAdd(T r, int index)
        {
        }

        internal virtual bool AddUnSorts(T r)
        {
            return true;
        }

        internal virtual void OnDelete(T r, int index)
        {
        }

        internal int Count
        {
            get {return FList.Count;}
        }

        public virtual void Clear()
        {
            for (int i=0;i<Count;i++) OnDelete(FList[i], i);
            FList.Clear();
        }

        internal virtual void Sort()
        {
            FList.Sort();
            FSorted=true;
        }
        
        internal bool Find(T aRecord, ref int Index)
        {
            if (!FSorted) Sort();
            Index=FList.BinarySearch(0, FList.Count, aRecord, null);  //Only BinarySearch compatible with CF.
            bool Result = Index >= 0;
            if (Index < 0) Index = ~Index;
            return Result;
        }

		protected virtual T CloneRecord(T br, TSheetInfo SheetInfo)
		{
			return (T)TBaseRecord.Clone(br, SheetInfo);
		}

        internal virtual void CopyFrom(TBaseRecordList<T> aBaseRecordList, TSheetInfo SheetInfo)
        {
            if (aBaseRecordList.FList == FList) XlsMessages.ThrowException(XlsErr.ErrInternal); //If they are, the for loop will grow ad infinitum.
            if (aBaseRecordList != null)
            {
                FList.Capacity += aBaseRecordList.Count;
                for (int i = 0; i < aBaseRecordList.Count; i++)
                {
                    T br = CloneRecord(aBaseRecordList[i], SheetInfo);
                    FList.Add(br);
                    OnAdd(br, i);
                }
            }
        }


        internal virtual long TotalSize
        {
            get
            {
                int Result = 0;
                for (int i = Count - 1; i >= 0; i--)
                {
					T it= (T)FList[i];
                    if (it != null) Result += it.TotalSize();
				}

                return Result;
            }
        }

        internal virtual long TotalSizeNoHeaders
        {
            get
            {
                int Result = 0;
                for (int i = Count - 1; i >= 0; i--)
                {
                    T it = (T)FList[i];
                    if (it != null) Result += it.TotalSizeNoHeaders();
                }

                return Result;
            }
        }

        public virtual void SaveToStream(IDataStream DataStream, TSaveData SaveData, int Row)
        {
            int aCount=FList.Count;
            for (int i = 0; i < aCount; i++)
            {
                T it = (T)FList[i];
                if (it != null) it.SaveToStream(DataStream, SaveData, Row);
            }
        }

        public virtual void SaveToPxl(TPxlStream PxlStream, int Row, TPxlSaveData SaveData)
        {
            int aCount = FList.Count;
            for (int i = 0; i < aCount; i++)
            {
                T it = (T)FList[i];
                if (it != null) it.SaveToPxl(PxlStream, Row, SaveData);
            }
        }

    } //TBaseRecordList

    //For comparing columns on a row.
    internal class ColumnComparer<T>: IComparer<T> where T: TBaseRowColRecord
    {
        #region IComparer Members

        public int Compare(T x, T y)
        {
            return x.Col.CompareTo(y.Col);
        }

        #endregion

    }


    /// <summary>
    /// Common ancestor for all list of records containing rows and columns.
    /// </summary>
    internal class TBaseRowColRecordList<T> : TBaseRecordList<T> where T : TBaseRowColRecord
    {
        protected static readonly ColumnComparer<T> ColumnComparerMethod = new ColumnComparer<T>();

        internal TBaseRowColRecordList()
        {
        }

        #region Generics
        internal new T this[int index]
        {
            get { return (T)FList[index]; }
            set { SetThis(value, index); }
        }

        protected override void SetThis(T value, int index)
        {
            T r = this[index];

            if (r != null)
            {
                OnDelete(r, index);
                r.Destroy();
            }
            FList[index] = value;
            FSorted = r != null; //When r is not null, we assume we are inserting at the right place.
            if (value != null) OnAdd(value, index);
        }

        internal override void Delete(int index)
        {
            T r = this[index];
            if (r != null)
            {
                r.Destroy();
            }
            OnDelete(r, index);
            FList.RemoveAt(index);
        }

        internal void DeleteAndNotDestroy(int index)
        {
            T r = FList[index];
            OnDelete(r, index);
            FList.RemoveAt(index);
        }

        internal void Destroy()
        {
            for (int k = 0; k < Count; k++)
            {
                T r = FList[k];
                if (r != null) r.Destroy();
            }
        }

        internal void ClearAndNotDestroy()
        {
            base.Clear();
        }

        public override void Clear()
        {
            Destroy();
            base.Clear();
        }
        #endregion

        internal override void Sort()
        {
            FList.Sort(ColumnComparerMethod);
            FSorted = true;
        }

        internal void SortOptional()
        {
            if (!FSorted) Sort();
        }

        /// <summary>
        /// There are 2 reasons why we can perform better than a binary search here:
        ///    1) blocks of cells are normally together, so it you know cell 1 is column 5, column 8 will probably be cell 3
        ///    2) this is a discrete search. There is no col=5.5, so if you know cell 1 is column 5 and you look for column 8, it can't be more far away than 3 positions.
        /// </summary>
        /// <returns></returns>
        internal bool GuessSearch(int aCol, ref int Guess)
        {
            int MinPos = 0;
            Guess = aCol;
            int MaxPos = Count - 1;
            while (true)
            {
                if (Guess > MaxPos) Guess = MaxPos;
                if (Guess < MinPos) Guess = MinPos;
                if (MinPos > MaxPos) return false;
                int c = FList[Guess].Col;
                if (c == aCol) return true;
                if (c > aCol)
                {
                    MaxPos = Guess - 1;
                    Guess += aCol - c;
                }
                else
                {
                    MinPos = Guess + 1;
                    Guess += aCol - c;
                }

            }
        }
        internal bool Find(int aCol, ref int Index)
        {
            if (!FSorted) Sort();
            /*SearchRecord.Col=aCol;
            Index=FList.BinarySearch(0, FList.Count, SearchRecord, ColumnComparerMethod);  //Only BinarySearch compatible with CF.
            bool Result=Index>=0;
            if (Index<0) Index=~Index;
            return Result;*/
            return GuessSearch(aCol, ref Index);
        }


        internal void ArrangeCopyRange(TXlsCellRange SourceRange, int Row, int RowOffset, int ColOffset, TSheetInfo SheetInfo)
        {
            int aCount = Count;
            for (int i = 0; i < aCount; i++)
            {
                T it = this[i];
                if (it != null) it.ArrangeCopyRange(SourceRange, Row, RowOffset, ColOffset, SheetInfo);
            }

        }

        internal void ArrangeInsertRangeCols(int Row, TXlsCellRange CellRange, int aColCount, TSheetInfo SheetInfo)
        {
            int LeftCol = CellRange.Left;
            for (int i = Count - 1; i >= 0; i--)
            {
                T it = this[i]; //Here it is everything except formulas.
                if (it.Col < LeftCol) return;
                if (it != null && !(it is TFormulaRecord))
                {
                    it.ArrangeInsertRange(Row, CellRange, 0, aColCount, SheetInfo);
                }
            }
        }


        internal virtual void SaveRangeToStream(IDataStream DataStream, TSaveData SaveData, int Row, TXlsCellRange CellRange)
        {
            if (!FSorted) Sort();
            if ((Row < CellRange.Top) || (Row > CellRange.Bottom)) return;
            int aCount = Count;
            for (int i = 0; i < aCount; i++)
            {
                T it = this[i];
                if (it != null)
                {
                    int c = it.Col;
                    if ((c >= CellRange.Left) && (c <= CellRange.Right))
                    {
                        it.SaveToStream(DataStream, SaveData, Row);
                    }

                }
            }

        }

        internal virtual long TotalRangeSize(TXlsCellRange CellRange)
        {
            long Result = 0;
            int aCount = Count;
            for (int i = 0; i < aCount; i++)
            {
                T it = this[i];
                if (it != null)
                {
                    //int r=Row;
                    int c = it.Col;
                    if  //((r>=CellRange.Top) && (r<=CellRange.Bottom)&& 
                        ((c >= CellRange.Left) && (c <= CellRange.Right)) Result += it.TotalSize();
                }
            }
            return Result;

        }

        internal virtual void UpdateCopy(TBaseRowColRecordList<T> CopyFrom, TSheetInfo SheetInfo, bool IsFullRows)
        {
        }

        internal virtual void UpdateRow(TBaseRowColRecordList<T> CopyFrom)
        {
        }

        internal virtual bool HasData()
        {
            return Count > 0;
        }

    }

	internal class TDeletedRanges
	{
		private bool[] Refs;
		private int[] Offs;
		internal TReferences References;
		internal TNameRecordList Names;
		internal int Length;

		internal bool Update; //We will contemplate 2 ways of working for this. When Update is true, ranges will be modified, else we will just build the list.

		internal TDeletedRanges(int NameCount, TReferences aReferences, TNameRecordList aNames)
		{
			Refs = new bool[NameCount];
			Offs = new int[NameCount];
			Length = NameCount;
			References = aReferences;
			Names = aNames;
			Update = false;
		}

		internal bool Referenced(int i)
		{
			return Refs[i];
		}

		internal void Reference(int i)
		{
			Refs[i] = true;
		}

		internal bool NeedsUpdate
		{
			get
			{
				if (Length <= 0) return false;

				int k = Length - 1;
				if (k < 0) return false; //no nodes are referenced.

				return Offs[k] != 0;
			}
		}

		internal void AddNameForDelete(int n)
		{
			for (int i = n; i < Length; i++) Offs[i]++;
		}

		internal int Offset(int n)
		{
			return Offs[n];
		}

		internal int Count
		{
			get
			{
				return Length;
			}
		}

	}

    /// <summary>
    /// List of NAME Records on Global Section
    /// </summary>
    internal class TNameRecordList: TBaseRecordList<TNameRecord>, INameRecordList //Records are TNameRecord
    {
        private TFutureStorage FutureStorage;

        #region Generics
        internal static void CheckInternalNames(int OptionFlags, TWorkbookGlobals Globals)
        {
            //If a name is added and it is internal, we can't trust the ordering and need to delete the country record.
            if ((OptionFlags & 0x20)!=0)
                Globals.DeleteCountry();
        }

        #endregion

        internal void ArrangeInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
        {
            for (int i = 0; i < Count; i++)
            {
                TNameRecord it = this[i];
                if (it != null) it.ArrangeInsertRange(CellRange, aRowCount, aColCount, SheetInfo);
            }
        }

        internal void ArrangeMoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
        {
            for (int i = 0; i < Count; i++)
            {
                TNameRecord it = this[i];
                if (it != null) it.ArrangeMoveRange(CellRange, NewRow, NewCol, SheetInfo);
            }
        }

        internal void InsertSheets(int CopyFrom, int BeforeSheet, int SheetCount, TSheetInfo SheetInfo, bool CopyFilterDatabase)
        {
            int aCount = SheetInfo.SourceNames.Count;
            int DestCount = Count; //Get this before adding the new sheets.
            for (int i = DestCount - 1; i >= 0; i--)
            {
                this[i].ArrangeInsertSheets(BeforeSheet, SheetCount);
            }

            
            for (int i = 0; i < aCount; i++)
            {
                if (
                    CopyFrom >= 0 
                    && NameMustBeCopiedToSheet(CopyFrom, SheetInfo, i) 
                    && (CopyFilterDatabase || SheetInfo.SourceNames[i].Name != TXlsNamedRange.GetInternalName(InternalNameRange.Filter_DataBase)) //filterdatabase will be copied when copying autofilters.
                    )
                {
                    for (int k = 0; k < SheetCount; k++)
                    {
                        SheetInfo.InsSheet = BeforeSheet + k;
                        TNameRecord name = (TNameRecord.Clone(SheetInfo.SourceNames[i], SheetInfo) as TNameRecord).ArrangeCopySheet(SheetInfo); //this could add its own names, so we need to recheck name wasn't added in next line.
                        if (!NameAlreadyExistsInLocal(SheetInfo.DestFormulaSheet, name.Name)) Add(name);
                        CheckInternalNames(SheetInfo.SourceNames[i].OptionFlags, SheetInfo.SourceGlobals);
                    }
                }
            }
        }

        private bool NameMustBeCopiedToSheet(int CopyFrom, TSheetInfo SheetInfo, int i)
        {
            return
                SheetInfo.SourceNames[i].RangeSheet == CopyFrom ||
                 (
                   (SheetInfo.SourceNames[i].RangeSheet == -1) && (SheetInfo.SourceNames[i].RefersToSheet(SheetInfo.SourceReferences) == SheetInfo.SourceFormulaSheet)
                   && !NameAlreadyExistsInLocal(SheetInfo.DestFormulaSheet, SheetInfo.SourceNames[i].Name)
                 );  
        }

		private bool NameAlreadyExistsInGlobal(int r, string name)
		{
			for (int i = Count -1; i >= 0; i--)
			{
				if ((i != r) && 
					this[i].RangeSheet == -1 && 
					String.Equals(this[i].Name, name, StringComparison.InvariantCultureIgnoreCase)) return true;
			}
			return false;
		}

		private bool NameAlreadyExistsInLocal(int sheet, string name)
		{
			for (int i = Count -1; i >= 0; i--)
			{
				if (
					this[i].RangeSheet == sheet && 
					String.Equals(this[i].Name, name, StringComparison.InvariantCultureIgnoreCase)) return true;
			}
			return false;
		}

		internal TNameRecord GetName(int sheet, string aName)
		{
			for (int i = Count -1; i >= 0; i--)
			{
				if (
					this[i].RangeSheet == sheet && 
					String.Equals(this[i].Name, aName, StringComparison.InvariantCultureIgnoreCase)) return this[i];
			}
			return null;
		}

        internal int GetNamePos(int sheet, string aName)
        {
            for (int i = Count - 1; i >= 0; i--)
            {
                if (
                    this[i].RangeSheet == sheet &&
                    String.Equals(this[i].Name, aName, StringComparison.InvariantCultureIgnoreCase)) return i;
            }
            return -1;
        }


		private void EnsureUniqueGlobalName(int r)
		{
			int cnt = 0;
			string OriginalName = this[r].Name;
			string Name = OriginalName;
			while (NameAlreadyExistsInGlobal(r, Name))
			{
				Name = "ref?" + cnt.ToString() + "_" + OriginalName;
				cnt++;
			}
			if (cnt == 0) return;

			this[r].Name = Name;
		}

		private void ClearName(int r)
		{
			this[r] = this[r].CopyWithoutFormulaData();
			//No need to Adaptsize, it is done when changing this[r]
		}

		internal static bool CanDeleteName(int i)
		{
			return true; //no extra checks needed.
		}


        internal void DeleteSheets(int SheetIndex, int SheetCount, TWorkbook Workbook)
        {
            TDeletedRanges DeletedRanges = Workbook.FindUnreferencedRanges(SheetIndex, SheetCount);

            for (int i = Count - 1; i >= 0; i--)
            {
                if ((this[i].RangeSheet >= SheetIndex) && (this[i].RangeSheet < SheetIndex + SheetCount))
                {
                    /* We cannot just delete the range, or formulas referring this range would crash (or refer to the wrong name).
                     * To actually delete here, we need to first find out whether this range is used.
                     * If it is not, go through all the formulas, charts, pivot tables, etc, and update the references to
                     * ranges less than this one to one less.
                    */

                    if (!DeletedRanges.Referenced(i) && CanDeleteName(i))  //don't delete internal names, or macro names.
                    {
                        FList.RemoveAt(i);
                        DeletedRanges.AddNameForDelete(i);
                    }
                    else
                    {
                        EnsureUniqueGlobalName(i);
                        this[i].RangeSheet = -1;
                    }
                }
                else
                    this[i].ArrangeInsertSheets(SheetIndex, -SheetCount);
            }

            if (DeletedRanges.NeedsUpdate)
            {
                Workbook.UpdateDeletedRanges(SheetIndex, SheetCount, DeletedRanges); //Update formulas, charts, etc.
            }
        }

		internal void UpdateDeletedRanges(int SheetIndex, int SheetCount, TDeletedRanges DeletedRanges)
		{
			for (int i=Count-1;i>=0;i--)
			{
				if ((this[i].RangeSheet<SheetIndex) || (this[i].RangeSheet>=SheetIndex+SheetCount)) 
				{
					this[i].UpdateDeletedRanges(DeletedRanges);
				}
			}
			
		}

        internal int AddName (TXlsNamedRange Range, TWorkbookGlobals Globals, TCellList CellList)
        {
            int aCount=Count;
            bool IsInternal;
            bool ValidName= TXlsNamedRange.IsValidRangeName(Range.Name, out IsInternal);
            if (IsInternal) Range.OptionFlags |= 0x020;

            for (int i = 0; i < aCount; i++)
            {
                //Documentation is wrong. We need the sheet index (1 based) not the externsheetindex
                /*int rSheet=-1;
                if (this[i].RangeSheet>=0)
                {
                    rSheet = GetSheet(this[i].RangeSheet);
                }
                */
                int rSheet = this[i].RangeSheet;
                if (
                    (rSheet == Range.NameSheetIndex) &&
                    (String.Equals(this[i].Name, Range.Name, StringComparison.CurrentCultureIgnoreCase))
                )
                {
                    this[i] = new TNameRecord(Range, Globals, CellList);  //We have to be careful not to change name ordering, or formulas would point to wrong ranges.
                    //If we found it, then it *is* a valid name
                    return i;
                }
            }
            if (!ValidName) XlsMessages.ThrowException(XlsErr.ErrInvalidNameForARange, Convert.ToString(Range.Name));
            Add(new TNameRecord(Range, Globals, CellList));
            CheckInternalNames(Range.OptionFlags, Globals);

            return Count - 1;
        }

        internal bool AddNameIfNotExists(TNameRecord NewName, out int NamePos)
        {
            int aCount = Count;
            for (int i = 0; i < aCount; i++)
            {
                int rSheet = this[i].RangeSheet;
                if (
                    (rSheet == NewName.RangeSheet) &&
                    (String.Equals(this[i].Name, NewName.Name, StringComparison.CurrentCultureIgnoreCase))
                )
                {
                   NamePos = i; //Name already exists, we won't replace it.
                   return false;
                }
            }
            Add(NewName);
            NamePos = Count - 1;
            return true;
        }

		internal void ReplaceName (int Index, TXlsNamedRange Range, TWorkbookGlobals Globals, TCellList CellList)
		{
            bool IsInternal;
            bool ValidName = TXlsNamedRange.IsValidRangeName(Range.Name, out IsInternal);
            if (!ValidName) XlsMessages.ThrowException(XlsErr.ErrInvalidNameForARange, Convert.ToString(Range.Name));
			if (Index <0 || Index >= Count) XlsMessages.ThrowException(XlsErr.ErrXlsIndexOutBounds, Index, "Index", 0, Count -1);
            if (IsInternal) Range.OptionFlags |= 0x020;

			this[Index]= new TNameRecord(Range, Globals, CellList);  //We have to be careful not to change name ordering, or formulas would point to wrong ranges.
			CheckInternalNames(Range.OptionFlags, Globals);
		}

		internal void DeleteName(int Index, TWorkbook Workbook)
		{
			if (Index <0 || Index >= Count) XlsMessages.ThrowException(XlsErr.ErrXlsIndexOutBounds, Index, "Index", 0, Count -1);

			TDeletedRanges DeletedRanges = Workbook.FindUnreferencedRanges(-1, 0);
			if (!DeletedRanges.Referenced(Index) && CanDeleteName(Index))  //don't delete internal names, or macro names.
			{
				FList.RemoveAt(Index);
				DeletedRanges.AddNameForDelete(Index);
				Workbook.UpdateDeletedRanges(-1, 0, DeletedRanges);
			}
			else ClearName(Index);
		}

		internal void CleanUnusedNames(TWorkbook Workbook)
		{
			TDeletedRanges DeletedRanges = Workbook.FindUnreferencedRanges(-1, 0);
			for (int i = DeletedRanges.Count -1; i >= 0; i--)
			{
				if (!DeletedRanges.Referenced(i) && !this[i].HasFormulaData && CanDeleteName(i))  //don't delete internal names, or macro names.
				{
					FList.RemoveAt(i);
					DeletedRanges.AddNameForDelete(i);
				}
			}
			
			if (DeletedRanges.NeedsUpdate) Workbook.UpdateDeletedRanges(-1, 0, DeletedRanges);
		}


        public int GetCount()
        {
            return Count;
        }

        public string GetName(int pos)
        {
            return this[pos].Name;
        }

		public int NameSheet(int pos)
		{
			return this[pos].RangeSheet;
		}

        internal int AddAddin(string functionName, bool AddErrorDataToFormula)
        {
            for (int i = 0; i < Count; i++)
            {
                TNameRecord name = this[i];
                if (name.IsAddin && String.Equals(name.Name, functionName, StringComparison.InvariantCultureIgnoreCase))
                {
                    return i;
                }
            }

            Add(TNameRecord.CreateAddin(functionName, AddErrorDataToFormula));
            return Count - 1;
        }

        internal void AddFutureStorage(TFutureStorageRecord R)
        {
            TFutureStorage.Add(ref FutureStorage, R);
        }

        internal void AddComment(TNameCmtRecord aComment)
        {
            if (aComment == null || Count == 0) return;
            //This is a little hairy. In theory, we could know the comment because it is after the last record.
            //But, some apps like old flexcel or Excel 2003 not aware of namecmt comments will write them all at the bottom.
            //So, what we need to do, and what Excel does, is to compare the names, and don't worry about the location.
            //But, as the same name might be defined in more than one sheet, we might lose comments here. That's what Excel 2007 does.
            //We are trying to be a little smarter here, by searching in order.

            for (int i = 0; i < Count; i++)
            {
                TNameRecord name = this[i];
                if (name.Comment == null && String.Equals(name.Name, aComment.Name, StringComparison.InvariantCultureIgnoreCase))
                {
                    name.Comment = aComment.Comment;
                    return;
                }
            }

        }
    }

    internal class TMiscRecordList : TBaseRecordList<TBaseRecord>
    {
    }

    internal class TBoundSheetRecordList: TBaseRecordList<TBoundSheetRecord>
    {		
        internal string GetSheetName(int index)
        {
            return this[index].SheetName;
        }
		
        internal void SetSheetName(int index, string Value)
        {
            this[index].SheetName=Value;
        }
		
        internal TXlsSheetVisible GetSheetVisible(int index)
        {
            //Wrong docs? the byte to hide a sheet is the low, not the high on grbit
            switch ((this[index].OptionFlags) & 0x3)
            {
                case 1: 
                    return TXlsSheetVisible.Hidden;
                case 2:
                    return TXlsSheetVisible.VeryHidden;
                default: return TXlsSheetVisible.Visible;
            }
        }

        internal void SetSheetVisible(int index, TXlsSheetVisible Value)
        {
            //Wrong docs? the byte to hide a sheet is the low, not the high on grbit

            byte p=this[index].OptionFlags1;
            p= (byte)( p & (0xFF-0x3)); //clear the 2 first bits;
            switch (Value)
            {
                case TXlsSheetVisible.Hidden: 
                    p=(byte)(p | 0x1);
                    break;
                case TXlsSheetVisible.VeryHidden: 
                    p=(byte)(p | 0x2);
                    break;
            }
            this[index].OptionFlags1=p;
        }

    }

    /// <summary>
    /// A list with cells.
    /// </summary>
    internal class TCellRecordList: TBaseRowColRecordList<TCellRecord>
    {
        private TFormulaCache FormulaCache;
        internal bool CanJoinRecords;

        internal TCellRecordList(TFormulaCache aFormulaCache, bool aCanJoinRecords): base()
        {
            FormulaCache=aFormulaCache;
            CanJoinRecords = aCanJoinRecords;
        }

        #region Generics
        internal new TCellRecord this[int index] 
        {
            get {return FList[index];} 
            set {SetThis(value, index);}
        }

        internal override void OnAdd(TCellRecord r, int index)
        {
            base.OnAdd(r, index);
            TFormulaRecord rf= (r as TFormulaRecord);
            if (rf!=null)
                FormulaCache.Add(rf);
        }

		public void TrimToSize()
		{
			FList.TrimExcess();
		}

        internal override bool AddUnSorts(TCellRecord r)
        {
            int LCount = FList.Count;
            return (r == null || LCount > 1 && r.Col < FList[LCount - 2].Col);
        }


        internal override void OnDelete(TCellRecord r, int index)
        {
            base.OnDelete(r, index);
            TFormulaRecord rf= (r as TFormulaRecord);
            if (rf!=null)
                FormulaCache.Delete(rf);
        }

        #endregion

        private void GoNext(ref int i, int aCount, ref TCellRecord it, ref TCellRecord NextRec)
        {
            it=NextRec;
            i++;
            if (i<aCount) NextRec=this[i]; else NextRec=null;
        }


        private int SaveAndCalcRange(IDataStream DataStream, TSaveData SaveData, int Row, TXlsCellRange CellRange)
        {
            if (!FSorted) Sort();
            int Result = 0;
            if ((Row < CellRange.Top) || (Row > CellRange.Bottom)) return Result;
            int aCount = Count;
            if (aCount <= 0) return Result;

            TCellRecord NextRec = this[0];
            TCellRecord it = null;
            int i = 0;
            while (i < aCount)
            {
                GoNext(ref i, aCount, ref it, ref NextRec);
                if (it != null)
                {
                    int c = it.Col;
                    if ((c >= CellRange.Left) && (c <= CellRange.Right))
                    {
                        //Search for MulRecords. We need 2 blanks together for this to work.
                        if (it.CanJoinNext(NextRec, CellRange.Right) && CanJoinRecords)
                        {
                            //Calc Total.
                            int i2 = i; TCellRecord it2 = it; TCellRecord NextRec2 = NextRec;
                            int JoinSize = it2.TotalSizeFirst();
                            GoNext(ref i2, aCount, ref it2, ref NextRec2);

                            while (it2.CanJoinNext(NextRec2, CellRange.Right))
                            {
                                JoinSize += it2.TotalSizeMid();
                                GoNext(ref i2, aCount, ref it2, ref NextRec2);
                            }
                            JoinSize += it2.TotalSizeLast();
                            Result += JoinSize;

                            if (DataStream != null)
                            {
                                it.SaveFirstMul(DataStream, SaveData, Row, JoinSize);
                                GoNext(ref i, aCount, ref it, ref NextRec);

                                while (it.CanJoinNext(NextRec, CellRange.Right))
                                {
                                    it.SaveMidMul(DataStream, SaveData);
                                    GoNext(ref i, aCount, ref it, ref NextRec);
                                }
                                it.SaveLastMul(DataStream, SaveData);
                            }
                            else
                            {
                                it = it2;
                                i = i2;
                                NextRec = NextRec2;
                            }
                        }
                        else
                        {
                            if (DataStream != null) it.SaveToStream(DataStream, SaveData, Row);
                            Result += it.TotalSize();
                        }
                    }
                }
            }
            return Result;
        }

        internal override void SaveRangeToStream(IDataStream DataStream, TSaveData SaveData, int Row, TXlsCellRange CellRange)
        {
            SaveAndCalcRange(DataStream, SaveData, Row, CellRange);
        }

        public override void SaveToPxl(TPxlStream PxlStream, int Row, TPxlSaveData SaveData)
        {
            base.SaveToPxl (PxlStream, Row, SaveData);
        }


        internal override long TotalRangeSize(TXlsCellRange CellRange)
        {
            return SaveAndCalcRange(null, new TSaveData(), CellRange.Top ,CellRange);
        }

        public override void SaveToStream(IDataStream DataStream, TSaveData SaveData, int Row)
        {
            SaveRangeToStream (DataStream, SaveData, Row, new TXlsCellRange(0,0,FlxConsts.Max_Rows, FlxConsts.Max_Columns));
        }

        internal override long TotalSize
        {
            get
            {
                return TotalRangeSize(new TXlsCellRange(0, 0, FlxConsts.Max_Rows, FlxConsts.Max_Columns));
            }
        }


    }

    internal class TCellAndRowRecordList : TCellRecordList
    {
        internal TRowRecord RowRecord;

        internal TCellAndRowRecordList(TFormulaCache aFormulaCache, bool aCanJoinRecords)
            : base(aFormulaCache, aCanJoinRecords)
        {
        }

        internal override void CopyFrom(TBaseRecordList<TCellRecord> aBaseRecordList, TSheetInfo SheetInfo)
        {
            base.CopyFrom(aBaseRecordList, SheetInfo);
            RowRecord = (TRowRecord)TRowRecord.Clone(((TCellAndRowRecordList)aBaseRecordList).RowRecord, SheetInfo);
        }

        internal override long TotalRangeSize(TXlsCellRange CellRange)
        {
            int R = 0; if (RowRecord != null) R = RowRecord.TotalSize();
            return R + base.TotalRangeSize(CellRange);
        }

        internal override void UpdateCopy(TBaseRowColRecordList<TCellRecord> CopyFrom, TSheetInfo SheetInfo, bool IsFullRows)
        {
            if (IsFullRows) RowRecord = (TRowRecord)TRowRecord.Clone(((TCellAndRowRecordList)CopyFrom).RowRecord, SheetInfo);
            else RowRecord = null;
        }

        internal override void UpdateRow(TBaseRowColRecordList<TCellRecord> CopyFrom)
        {
            TCellAndRowRecordList rr = (TCellAndRowRecordList)CopyFrom;
            RowRecord = rr.RowRecord;
            rr.RowRecord = null;
        }

        internal override bool HasData()
        {
            return base.HasData() || RowRecord != null;
        }

    }

    internal class TSharedFormula
    {
        internal UInt64 Key;
        internal TParsedTokenList Data;

        internal TSharedFormula()
        {
        }

        internal TSharedFormula(TNameRecordList Names, UInt64 aKey, byte[] aData, int start, int len)
        {
            Key = aKey;
            TFormulaConvertBiff8ToInternal p = new TFormulaConvertBiff8ToInternal();
            Data = p.ParseRPN(Names, -1, -1, aData, start, len, true); //no real need for relative since shared formulas can't be 3d, and we only need relative for the non-existing tokens ptgarea3dn and ptgref3dn.
        }
    }

    /// <summary>
    /// A list with Shared formulas.
    /// </summary>
    internal class TShrFmlaRecordList
    {
        private Dictionary<ulong, TSharedFormula> FList;

        internal TShrFmlaRecordList()
        {
            FList = new Dictionary<ulong, TSharedFormula>();
        }

        #region Generics
        internal bool TryGetValue(UInt64 key, out TSharedFormula value) 
        {
            return FList.TryGetValue(key, out value);
        }
        #endregion

        internal void Add(TSharedFormula aSharedFormula)
        {
            if (FList.ContainsKey(aSharedFormula.Key)) return; //a key might exist more than once.
            FList.Add(aSharedFormula.Key, aSharedFormula);
        }
    }


}

