using System;
using System.IO;
using FlexCel.Core;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;

namespace FlexCel.XlsAdapter
{
    /// <summary>
    /// A basic list for handling rows of records. It has a list of rows, and each one of them can have a list of columns.
    /// </summary>
    internal abstract class TBaseRowColList<K, T>
        where K : TBaseRowColRecord
        where T : TBaseRowColRecordList<K>
    {
        protected List<T> FList;

        internal TBaseRowColList()
        {
            FList = new List<T>();
        }

        #region Generics
        internal void Add (T a)
        {
            FList.Add(a);
        }

        internal void Insert (int index, T a)
        {
            FList.Insert(index, a);
            if (Count > FlxConsts.Max_Rows + 1) XlsMessages.ThrowException(XlsErr.ErrTooManyRows, Count, FlxConsts.Max_Rows + 1);
        }


        internal T this[int index] 
        {
            get {return FList[index];} 
        }

        internal void RemoveAt(int index)
        {
            T a = this[index];
            for (int i = a.Count - 1; i >= 0; i--)
                a.OnDelete(a[i], i);
            FList.RemoveAt(index);
        }

        #endregion

        internal int Count
        {
            get {return FList.Count;}
        }

        public virtual void Clear()
        {
            int aCount = Count;
            for (int i = 0; i < aCount; i++)
            {
                this[i].Destroy();
                //no need to call ondelete here as FormulaCache will be cleared by the children when calling clear.
            }
            FList.Clear();
        }

        protected abstract T CreateRecord();

        internal void AddRow(T aRecordList, int aRow)
        {
            for (int i = Count; i < aRow; i++)
                FList.Add(CreateRecord());
            FList.Add(aRecordList);
        }

        internal void CopyFrom(TBaseRowColList<K, T> aList, TSheetInfo SheetInfo)
        {
            if (aList.FList == FList) XlsMessages.ThrowException(XlsErr.ErrInternal); //Should be different objects
            for (int i = 0; i < aList.Count; i++)
            {
                T Tr = CreateRecord();
                Tr.CopyFrom(aList[i], SheetInfo);
                FList.Add(Tr);
            }
        }

        internal void SaveToStream(IDataStream DataStream, TSaveData SaveData, TXlsCellRange CellRange)
        {
            if (CellRange == null) SaveAllToStream(DataStream, SaveData); else SaveRangeToStream(DataStream, SaveData, CellRange);
        }

        private void SaveAllToStream(IDataStream DataStream, TSaveData SaveData)
        {
            int aCount=Count;
            for (int i=0; i< aCount;i++)
                this[i].SaveToStream(DataStream, SaveData, i);
        }

        private void SaveRangeToStream(IDataStream DataStream, TSaveData SaveData, TXlsCellRange CellRange)
        {
            int aCount=Count;
            for (int i=0; i< aCount;i++)
                this[i].SaveRangeToStream(DataStream, SaveData, i, CellRange);
        }

        internal long TotalSize(TXlsCellRange CellRange)
        {
            if (CellRange == null) return TotalSizeAll();
            return TotalRangeSize(CellRange);
        }

        private long TotalSizeAll()
        {
            long Result = 0;
            int aCount = Count;
            for (int i = 0; i < aCount; i++)
                Result += this[i].TotalSize;
            return Result;
        }

        private long TotalRangeSize(TXlsCellRange CellRange)
        {
            long Result=0;
            for (int i=CellRange.Top; i<= CellRange.Bottom;i++)
                Result+=this[i].TotalRangeSize(CellRange);
            return Result;
        }

        internal long TotalSizeNoHeaders()
        {
            long Result = 0;
            int aCount = Count;
            for (int i = 0; i < aCount; i++)
                Result += this[i].TotalSizeNoHeaders;
            return Result;
        }


        /// <summary>
        /// This method moves the parts of the rows that go to a different parent row when deleting ranges.
        /// </summary>
        /// <param name="DestRow"></param>
        /// <param name="Left"></param>
        /// <param name="Right"></param>
        /// <param name="RowCount"></param>
        private void FixDeletedRowOffsets(int DestRow, int Left, int Right, int RowCount)
        {
            int aCount = Count;
            for (int i = DestRow + RowCount; i < aCount; i++)
            {
                T R = this[i];
                for (int c = R.Count - 1; c >= 0; c--)
                {
                    K Rc = R[c];
                    if (Rc.Col >= Left && Rc.Col <= Right)
                    {
                        int Index = 0;

                        if (!this[i - RowCount].Find(Rc.Col, ref Index))
                        {
                            R.DeleteAndNotDestroy(c); //do it first, so it is deleted from cacheindex.
                            this[i - RowCount].Insert(Index, Rc);
                        }
                        else R.Delete(c);
                    }
                }
            }
        }

        private void FixInsertedRowOffsets(int DestRow, int Left, int Right, int RowCount)
        {
            if (Left <= 0 && Right >= FlxConsts.Max_Columns) return; //no need to fix anything.
            int aCount = Count;
            for (int i = DestRow + RowCount; i < aCount; i++)
            {
                T RSource = this[i];
                T RDest = this[i - RowCount];

                RDest.UpdateRow(RSource);
                for (int c = RSource.Count - 1; c >= 0; c--)
                {
                    K Rc = RSource[c];
                    if (Rc.Col < Left || Rc.Col > Right)
                    {
                        int Index = 0;

                        if (!RDest.Find(Rc.Col, ref Index))
                        {
                            RSource.DeleteAndNotDestroy(c); //do it first, so it is deleted from cacheindex.
                            RDest.Insert(Index, Rc);
                        }
                        else RSource.Delete(c);
                    }
                }
            }


            while (Count > 0 && !this[Count - 1].HasData()) RemoveAt(Count - 1); //sadly this can't be done in the loop above.
        }

        internal virtual void InsertAndCopyRange(TXlsCellRange SourceRange, TFlxInsertMode InsertMode, int DestRow, int DestCol, int aRowCount, int aColCount, TRangeCopyMode CopyMode, TSheetInfo SheetInfo)
        {

             //There are 2 ways to do the insertandcopy thing:
             //*   1) First insert, then copy. It seems like the best method, but it has the drawback that
             //*      when you break the source range in 2, (for example, by insertingandcopying from a1:c2 in b1, by rows)
             //*      you can't access the source range anymore. Note that Excel does not allow copying ranges on this way anyway.
             //* 
             //*   2) First copy (to a temp place) then insert. Has the problem that, f.i. if you have cell A3=$A$4,
             //*      and you insert above, the copied cell will refer to $A$4 too. When you later insert,
             //*      old cell A3 will be moved to A4=$A$5, but the new A3 will remain A3=$A$4. If we applied both operations
             //*      to new cell A3 (offset when copying, then offset when inserting) we can go to negative temp results
             //*      that will result on a #ref, even when the final result is valid.
             

            // Insert the cells. we have to look at all the formulas, not only those below destrow
            TXlsCellRange DestInsRange = SourceRange.OffsetForIns(DestRow, DestCol, InsertMode);
            ArrangeInsertRange(DestInsRange, aRowCount, aColCount, SheetInfo);

            if (aRowCount > 0)
            {
                //Copy the cells
                int MyDestRow = DestRow;
                int CopyOffs = 0;
                bool IsFullRows = SourceRange.Left <= 0 && SourceRange.Right >= FlxConsts.Max_Columns;

                for (int j = 1; j <= aRowCount; j++)
                {
                    for (int i = SourceRange.Top; i <= SourceRange.Bottom; i++)
                    {
                        T aRecordList = CreateRecord();

                        if (i + CopyOffs >= Count && MyDestRow >= Count) //optimization, would work anyway without this line.
                        {
                            int Remaining = SourceRange.Bottom - i + 1;
                            MyDestRow += Remaining;
                            if (SourceRange.Top >= DestRow) CopyOffs += Remaining;

                            break;
                        }

                        //Will only copy the cells if copyfrom < recordcount.
                        if (i + CopyOffs < Count && (CopyMode != TRangeCopyMode.None))
                        {
                            T Rl = this[i + CopyOffs];
                            aRecordList.UpdateCopy(Rl, SheetInfo, IsFullRows);

                            for (int a = 0; a < Rl.Count; a++)
                            {
                                if (((CopyMode != TRangeCopyMode.OnlyFormulas && CopyMode != TRangeCopyMode.OnlyFormulasAndNoObjects) || Rl[a].AllowCopyOnOnlyFormula)
                                    && (Rl[a].Col >= SourceRange.Left && Rl[a].Col <= SourceRange.Right))
                                {
                                    K Rec;
                                    if (CopyMode == TRangeCopyMode.Formats && Rl[a] is TCellRecord)
                                    {
                                        TCellRecord Cell = (TCellRecord)(TBaseRecord)Rl[a];
                                        Rec = (K)(TBaseRecord)new TBlankRecord(Cell.Col, Cell.XF);
                                    }
                                    else Rec = (K)TBaseRowColRecord.Clone(Rl[a], SheetInfo);

                                    int InsOffs = 0;
                                    if (SourceRange.Top >= DestRow && DestInsRange.HasCol(Rec.Col))
                                        InsOffs = aRowCount * SourceRange.RowCount;

                                    Rec.ArrangeCopyRange(SourceRange, i + CopyOffs, MyDestRow - (i + InsOffs), DestCol - SourceRange.Left, SheetInfo);
                                    aRecordList.Add(Rec);
                                }
                            }

                            if ((aRecordList != null) && (aRecordList.Count > 0))
                            {
                                //aRecordList.ArrangeCopyRange(i+CopyOffs, MyDestRow-(i+InsOffs),DestCol-SourceRange.Left);
                            }
                        }

                        if (aRecordList.HasData() || MyDestRow < Count)
                        {
                            for (int z = Count; z < MyDestRow; z++)
                                Add(CreateRecord());
                            Insert(MyDestRow, aRecordList);
                        }
                        aRecordList = null;

                        MyDestRow++;
                        if (SourceRange.Top >= DestRow) CopyOffs++;
                    }
                }

                if (aRowCount != 0 && !IsFullRows) FixInsertedRowOffsets(DestRow, DestCol, DestCol + SourceRange.ColCount - 1, aRowCount * SourceRange.RowCount);

            }

            if (aColCount > 0)
            {
                //Copy the cells
                int CopyOffs = 0;
                int MyDestCol = DestCol;
                if (SourceRange.Left >= DestCol) CopyOffs = aColCount * SourceRange.ColCount;

                int FixedCount = Count;
                int FinalR = Math.Min(FixedCount, SourceRange.Bottom + 1);

                for (int k = 1; k <= aColCount; k++)
                    for (int i = SourceRange.Left; i <= SourceRange.Right; i++)
                    {
                        for (int r = SourceRange.Top; r < FinalR; r++)
                        {
                            int Index = 0;
                            if ((CopyMode != TRangeCopyMode.None) &&
                                (this[r].Find(i + CopyOffs, ref Index))
                                &&
                                ((CopyMode != TRangeCopyMode.OnlyFormulas && CopyMode != TRangeCopyMode.OnlyFormulasAndNoObjects) || (this[r][Index].AllowCopyOnOnlyFormula))
                                )
                            {
                                K Rec;
                                if (CopyMode == TRangeCopyMode.Formats && this[r][Index] is TCellRecord)
                                {
                                    TCellRecord Cell = (TCellRecord)(TBaseRecord)this[r][Index];
                                    Rec = (K)(TBaseRecord)new TBlankRecord(Cell.Col, Cell.XF);
                                }
                                else Rec = (K)TBaseRecord.Clone(this[r][Index], SheetInfo);

                                Rec.ArrangeCopyRange(SourceRange, r, DestRow - SourceRange.Top, MyDestCol - (i + CopyOffs), SheetInfo);
                                int Index2 = 0;
                                if (r + DestRow - SourceRange.Top >= Count)
                                {
                                    for (int z = Count; z <= r + DestRow - SourceRange.Top; z++)
                                    {
                                        FList.Add(CreateRecord());
                                    }
                                }
                                this[r + DestRow - SourceRange.Top].Find(Rec.Col, ref Index2);
                                this[r + DestRow - SourceRange.Top].Insert(Index2, Rec);
                            }
                        }
                        MyDestCol++;
                    }
            }

        }

        internal void DeleteRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
        {
            if (aRowCount > 0)
            {
                int Max = CellRange.Top + aRowCount * CellRange.RowCount; if (Max > Count) Max = Count;

                if ((CellRange.Left <= 0) && (CellRange.Right >= FlxConsts.Max_Columns))
                {
                    BreakPages(CellRange.Top, aRowCount);
                    for (int i = Max - 1; i >= CellRange.Top; i--)
                    {
                        this[i].Destroy();
                        RemoveAt(i);
                    }
                }
                else
                    for (int i = Max - 1; i >= CellRange.Top; i--)
                    {
                        for (int c = this[i].Count - 1; c >= 0; c--)
                            if (CellRange.HasCol(this[i][c].Col))
                            {
                                this[i].Delete(c);
                            }

                    }

            }

			if (aColCount>0)
			{
				int Index=0;
				for (int r=CellRange.Top; r<=CellRange.Bottom;r++)
				{
					if (r>=Count) break;
					this[r].Find(CellRange.Left, ref Index);
					int c=Index;
					while (c<this[r].Count)
					{
						if (CellRange.HasCol(this[r][c].Col))
							this[r].Delete(c);
						else 
						{
							if (c>Index) break;
							c++;
						}
					}
				}
			}
            //Delete the cells. we have to look at all the formulas, not only those below a row
            ArrangeInsertRange(CellRange, -aRowCount, -aColCount, SheetInfo);

            if (aRowCount>0 && !((CellRange.Left<=0)&&(CellRange.Right>= FlxConsts.Max_Columns))) 
                FixDeletedRowOffsets(CellRange.Top, CellRange.Left, CellRange.Right, aRowCount*CellRange.RowCount);
        }

        internal virtual void BreakPages(int aRow, int aRowCount)
        {
        }

        internal void ClearRange(TXlsCellRange CellRange)
        {
            int Max=CellRange.Bottom+1 ; if (Max>Count) Max= Count;
            for (int i= Max-1; i>=CellRange.Top;i--)
            {
                for (int c=this[i].Count-1;c>=0;c--)
                    if (CellRange.HasCol(this[i][c].Col))
                    {
                        this[i].Delete(c);
                    }

            }       
        }

        internal void ClearFormats(TXlsCellRange CellRange)
        {
            int Max = CellRange.Bottom + 1; if (Max > Count) Max = Count;
            for (int i = Max - 1; i >= CellRange.Top; i--)
            {
                for (int c = this[i].Count - 1; c >= 0; c--)
                    if (CellRange.HasCol(this[i][c].Col))
                    {
                        T crList = this[i];
                        K cr = crList[c];
                        if (cr is TBlankRecord) crList.Delete(c);
                        else
                        {
                            TCellRecord cell = cr as TCellRecord;
                            if (cell != null) cell.XF = FlxConsts.DefaultFormatId;
                        }
                    }
            }
        }

		internal void MoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
		{
			//Find out the loop direction. This is very important to avoid overwriting existing cells. for example when moving from (1,1: 5,5) to 2,2
			bool GoUp = NewRow > CellRange.Top;
			//bool GoLeft = NewCol > CellRange.Left; //GoLeft is not important, since we copy a whole row (or a whole part of a row) each time.
			int row = GoUp? CellRange.Bottom: CellRange.Top;
			int xRow = GoUp? -1: 1;
			int drow = NewRow - CellRange.Top;
			TXlsCellRange DestRange = CellRange.Offset(NewRow, NewCol);

			if ((CellRange.Left<=0)&&(CellRange.Right>= FlxConsts.Max_Columns) && NewCol == 0) //optimize moving rows.
			{
				for (int z= CellRange.Bottom; z>=CellRange.Top;z--)
				{
					T rec = null;
					if (row < Count && row >= 0)
					{
						rec = this[row];
					}

					if (row + drow < Count && row + drow >= 0) 
					{
						this[row + drow].Clear(); //this will destroy and clear formula cache.
						if (rec == null) FList[row + drow] = CreateRecord(); else FList[row + drow] = rec;
					}
					else
					{
						if (rec != null) AddRow(rec, row+ drow);
					}

					if (rec != null) FList[row] = CreateRecord();  //do not destroy things, they are on other part now.
					row += xRow;
				}
			}

			else
				for (int z= CellRange.Bottom; z>=CellRange.Top;z--)
				{
					T rec = null;
					if (row < Count && row >= 0)
					{
						rec = this[row];
					}

					T destrec = null;
					if (row + drow < Count && row + drow >= 0) 
					{
						destrec = this[row + drow];
					}
					else
					{
						if (rec != null) 
						{
							destrec = CreateRecord();
							AddRow(destrec, row+ drow);
						}
					}

					if (destrec != null)  //If null, there is no data in source or destiny.
					{
						//Delete all cells on existing destination that are NOT on source range.
						for (int c=destrec.Count-1;c>=0;c--)
						{
							if (DestRange.HasCol(destrec[c].Col))
							{
								bool IsInSource = CellRange.HasCol(destrec[c].Col) && CellRange.HasRow(row + drow);
								if (!IsInSource) destrec.Delete(c);
							}
						}

						if (rec != null)
						{
							K[] TmpRec = new K[rec.Count]; //rec and destrec can be the same object. (when moving in the same row). So we might be inserting on the same list we are modifying, and this is why we have to cache toe row first.
							int TmpRecordIndex = 0;
                            for (int c = rec.Count - 1; c >= 0; c--)
                            {
                                if (CellRange.HasCol(rec[c].Col))
                                {
                                    TmpRec[TmpRecordIndex] = rec[c];
                                    TmpRec[TmpRecordIndex].Col += NewCol - CellRange.Left;
                                    TmpRecordIndex++;
                                    rec.DeleteAndNotDestroy(c); //do it first so it is removed from formulacache.
                                }
                            }
					
							for (int i = 0; i < TmpRecordIndex; i++)
							{
								int Index = -1;
								if (destrec.Find(TmpRec[i].Col, ref Index))
								{
									//destrec might not have been cleared if ranges overlap.
									destrec[Index].Destroy();
									destrec[Index] = TmpRec[i]; //this will remove the old record from fmlacache and add the new.

								}
								else
								{
									destrec.Insert(Index, TmpRec[i]); //this will re-insert it on formulacache.
								}
							}
						}
						
					}

					row += xRow;
				}
             

			//Move references of the cells. Formulas that pointed to cellrange should point to the new range, and those pointing to newrange should point to #ref!
			ArrangeMoveRange(CellRange, NewRow, NewCol, SheetInfo);
		}


		internal virtual void ArrangeMoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
		{
			//nothing here.
		}

        internal virtual void ArrangeInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
        {
            if ((SheetInfo.InsSheet<0) || (SheetInfo.SourceFormulaSheet != SheetInfo.InsSheet)) return; //formulas are not arranged here. So this only makes sense when on the same sheet.
            //Row has nothing to arrange. Only formulas, and they are arranged on TCellList
            if (aColCount!=0)
            {
                //for performance, formulas will not be arranged here.
                int aCount= Math.Min(Count - 1, CellRange.Bottom);
                for (int i= CellRange.Top; i<= aCount;i++)
                {
                    this[i].ArrangeInsertRangeCols(i, CellRange, aColCount, SheetInfo);
                }
            }

        }
    }

    interface ICellList
    {
        bool FoundArrayFormula(int RowArr, int ColArr, out TArrayRecord ArrData);

        bool FoundTableFormula(int TopRow, int LeftCol, out TTableRecord TableData);

        TParsedTokenList ArrayFormula(int Row, int Col);

        TTableRecord TableFormula(int Row, int Col);

        ExcelFile Workbook { get; }

        TWorkbookGlobals Globals { get; }
    }

    /// <summary>
    /// A list of Cells, agrupated by RowLists. Items are TCellAndRowRecordList, each record has a row and a list of TCellRecord
    /// </summary>
    internal class TCellList : TBaseRowColList<TCellRecord, TCellAndRowRecordList>, ICellList
    {
        TWorkbookGlobals FGlobals;
        TColInfoList FColInfoList;
        internal TFormulaCache FormulaCache;
        internal TSheetGlobals SheetGlobals;
        internal Dictionary<int, TCellCondFmt[]> EmptyRowCF;
        int ActiveRow;

        internal TCellList(TWorkbookGlobals aGlobals, TColInfoList aColInfoList, TSheetGlobals aSheetGlobals)
        {
            FGlobals = aGlobals;
            SheetGlobals = aSheetGlobals;
            FColInfoList = aColInfoList;
            FormulaCache = new TFormulaCache();
            ActiveRow = -1;
            EmptyRowCF = new Dictionary<int, TCellCondFmt[]>();
        }

        #region Generics
        internal void AddRecord(TCellRecord aRecord, int aRow, TVirtualReader VirtualReader)
		{
            AddRecord(aRecord, aRow, false, VirtualReader);
		}

        internal void AddRecord(TCellRecord aRecord, int aRow, bool IfExistsSetFormulaValue, TVirtualReader VirtualReader)
        {
            //Virtual Mode
            if (VirtualReader != null)
            {
                if (IfExistsSetFormulaValue && aRecord != null)
                {
                    TFormulaRecord fr = VirtualReader.GetArray(aRow, aRecord.Col);
                    if (fr != null)
                    {
                        fr.FormulaValue = aRecord.GetValue(VirtualReader.CellList);
                        fr.XF = aRecord.XF;
                        aRecord = fr;
                    }
                }

                if (aRecord == null) return;
                TFormulaRecord fm = aRecord as TFormulaRecord;
                if (fm != null)
                {
                    if (fm.ArrayRecord != null || fm.TableRecord != null)
                    {
                        VirtualReader.CellList.AddArray(aRow, aRecord.Col, fm); //we need those for future cells that will refer to this.
                    }
                    if (fm.TableRecord != null)
                    {
                        VirtualReader.CellList.AddTable(aRow, aRecord.Col, fm); //we need those for future cells that will refer to this.
                    }
                }

                VirtualReader.Read(Workbook.PartialSheetCount(), aRow + 1, aRecord.Col + 1, aRecord);
                return;
            }                

            for (int i = Count; i <= aRow; i++)
                FList.Add(CreateRecord());

            TCellRecord Cell = aRecord as TCellRecord;
            TCellAndRowRecordList RowCells = this[aRow];
            if (IfExistsSetFormulaValue && Cell != null)
            {
                int index = -1;
                if (RowCells.Find(Cell.Col, ref index))
                {
                    TFormulaRecord fr = RowCells[index] as TFormulaRecord;
                    if (fr != null)
                    {
                        fr.FormulaValue = Cell.GetValue(this);
                        fr.XF = Cell.XF;
                    }
                    else //some third parties might have repeated values, and Excel stays with the latest ones.
                    {
                        RowCells.Delete(index);
                        RowCells.Add(aRecord);
                    }
                }
                else
                {
                    RowCells.Add(aRecord);
                }
            }
            else
            {
                int index = -1;
                if (Cell != null && RowCells.Find(Cell.Col, ref index)) RowCells.Delete(index); 
                RowCells.Add(aRecord);
            }


            if (ActiveRow != aRow)
            {
                if (ActiveRow >= 0 && ActiveRow < Count)
                {
                    this[ActiveRow].TrimToSize(); //Helps to not store the whole thing.
                }
                ActiveRow = aRow;
            }
        }

        #endregion


        public ExcelFile Workbook
        {
            get
            {
                return FGlobals.Workbook;
            }
        }

        public TWorkbookGlobals Globals
        {
            get
            {
                return FGlobals;
            }
        }

        protected override TCellAndRowRecordList CreateRecord()
        {
            return new TCellAndRowRecordList(FormulaCache, true);
        }

        public override void Clear()
        {
            base.Clear ();
            FormulaCache.Clear();
        }

        internal void Recalc(ExcelFile aXls, int SheetIndexBase1)
        {
			TBaseParsedToken.Dates1904 = aXls.OptionsDates1904;
            FormulaCache.Recalc(this, aXls, SheetIndexBase1);
        }

        internal void CleanFlags()
        {
            FormulaCache.CleanFlags();
        }

        internal void ClearFormulaResult()
        {
            FormulaCache.ClearResults();
        }

        internal void ForceAutoRecalc()
        {
            FormulaCache.ForceAutoRecalc();
        }

		internal bool Dates1904
		{
			get
			{
				return FGlobals.Dates1904;
			}
		}

        internal override void BreakPages(int aRow, int aRowCount)
        {
            if (!HasRow(aRow - 1) || this[aRow - 1].RowRecord.KeepTogether <= 0) return;

            //When deleting a row that has 0 level, we need to know if the row above is not 0. If it is not, we might need to change it to 0 
            //in order to keep the group separated.
            //In fact, even if the row we are deleting is not 0, we might need to change the row above anyway if any of the rows in the group had a split.

            int Level = this[aRow - 1].RowRecord.KeepTogether;
            for (int row = aRow; row < aRow + aRowCount; row++)
            {
                if (!HasRow(row) || this[row].RowRecord.KeepTogether != Level)
                {
                    //The group must be split.
                    this[aRow - 1].RowRecord.KeepTogether = 0;
                    return;
                }
            }

            //No need to split it, all rows 

        }

		#region Auto Fit
#if (!COMPACTFRAMEWORK && !MONOTOUCH && !SILVERLIGHT)
        private static int GetCellToGrow(int First, int Last, TAutofitMerged AutofitMerged)
        {
            switch (AutofitMerged)
            {
                case TAutofitMerged.OnLastCell:
                    return Last;
                case TAutofitMerged.OnLastCellMinusOne:
                    if (Last - 1 > First) return Last - 1; else return First;
                case TAutofitMerged.OnLastCellMinusTwo:
                    if (Last - 2 > First) return Last - 2; else return First;
                case TAutofitMerged.OnLastCellMinusThree:
                    if (Last - 3 > First) return Last - 3; else return First;
                case TAutofitMerged.OnLastCellMinusFour:
                    if (Last - 4 > First) return Last - 4; else return First;
                case TAutofitMerged.OnFirstCell:
                    return First;
                case TAutofitMerged.OnSecondCell:
                    if (First + 1 < Last) return First + 1; else return Last;
                case TAutofitMerged.OnThirdCell:
                    if (First + 2 < Last) return First + 2; else return Last;
                case TAutofitMerged.OnFourthCell:
                    if (First + 3 < Last) return First + 3; else return Last;
                case TAutofitMerged.OnFifthCell:
                    if (First + 4 < Last) return First + 4; else return Last;
            }

            FlxMessages.ThrowException(FlxErr.ErrInternal);
            return First;//just to compile.
        }


		internal void RecalcRowHeights(ExcelFile Workbook, int Row1, int Row2, bool Forced, bool KeepAutofit, bool OnlyMarkedForAutofit,
            float aAdjustment, int aAdjustmentFixed, int aMinHeight, int aMaxHeight, TAutofitMerged AutofitMerged)
		{
			using (TRowHeightCalc RowCalc = new TRowHeightCalc(FGlobals))
			{

				//For autofitting all the workoobk:
				//Row2 should be = FRowRecordList.Count - 1;
				//Row1 should be 0.

				float RowMultDisplay = ExcelMetrics.RowMultDisplay(Workbook) * 100F / FlexCel.Render.FlexCelRender.DispMul;
				float ColMultDisplay = ExcelMetrics.ColMultDisplay(Workbook) * 100F / FlexCel.Render.FlexCelRender.DispMul;

				TMultipleCellAutofitList MergedCells = new TMultipleCellAutofitList();

				for (int i = Row1; i <= Row2; i++)
				{
					if (!HasRow(i)) continue;
					TRowRecord Row = FList[i].RowRecord;
                    if (Row == null) continue;
					if (!Forced && !Row.IsAutoHeight()) continue;

                    if(OnlyMarkedForAutofit && !Row.MarkForAutofit && !Row.HasMergedCell) continue;
			
					int MaxCellHeight = RowCalc.CalcCellHeight(i+1, -1, null, Row.XF, Workbook, RowMultDisplay, ColMultDisplay, MergedCells);

					if (i < Count)
					{
						TCellRecordList Columns = this[i];
						int cCount = Columns.Count;
						for (int c=0; c<cCount;c++)
						{
							object val = Workbook.GetStringFromCell(i+1, Columns[c].Col+1);
						
							TRichString rx = val as TRichString;
							if (rx !=null)
							{
								int CellHeight = RowCalc.CalcCellHeight(i+1, Columns[c].Col + 1, rx, Columns[c].XF, Workbook, RowMultDisplay, ColMultDisplay, MergedCells);
								if (CellHeight>MaxCellHeight) MaxCellHeight = CellHeight;
							}
							else
							{
								string sx = val as string;
								if (sx !=null)
								{
									int CellHeight = RowCalc.CalcCellHeight(i+1, Columns[c].Col + 1, new TRichString(sx), Columns[c].XF, Workbook, RowMultDisplay, ColMultDisplay, MergedCells);
									if (CellHeight>MaxCellHeight) MaxCellHeight = CellHeight;
								}
							}
						}
					}

                    if(!OnlyMarkedForAutofit || Row.MarkForAutofit) ResizeRow(Row, Row, MaxCellHeight, OnlyMarkedForAutofit, KeepAutofit, aAdjustment, aAdjustmentFixed, aMinHeight, aMaxHeight);
					
				}
			
				//Merged Cells
                if (AutofitMerged != TAutofitMerged.None)
                {
                    FixCellsWithMergedRows(Workbook, AutofitMerged, MergedCells, OnlyMarkedForAutofit, KeepAutofit, aAdjustment, aAdjustmentFixed, aMinHeight, aMaxHeight);
                }

			}
		}

        private static void ReadLocalAdjustmentInRow(bool OnlyMarkedForAutofit, TRowRecord Row, ref float Adjustment, ref int AdjustmentFixed, ref int MinHeight, ref int MaxHeight)
        {
            if (OnlyMarkedForAutofit)
            {
                float adj = Row.AutofitAdjustment;
                if (adj > 0) Adjustment = adj;

                int adj2 = Row.AutofitAdjustmentFixed;
                if (adj2 != 0) AdjustmentFixed = adj2;

                if (Row.MinHeight != 0) MinHeight = Row.MinHeight;
                if (Row.MaxHeight != 0) MaxHeight = Row.MaxHeight;
            }
        }

        private static void ResizeRow(TRowRecord FirstRow, TRowRecord Row, int NewCellHeight, bool OnlyMarkedForAutofit, bool KeepAutofit, float aAdjustment, int aAdjustmentFixed, int aMinHeight, int aMaxHeight)
        {
            float Adjustment = aAdjustment; int AdjustmentFixed = aAdjustmentFixed; int MinHeight = aMinHeight; int MaxHeight = aMaxHeight;
            ReadLocalAdjustmentInRow(OnlyMarkedForAutofit, FirstRow, ref Adjustment, ref AdjustmentFixed, ref MinHeight, ref MaxHeight); //Adjustments are in the first row of a merged cell.

            if (MinHeight < 0) MinHeight = Row.Height;
            if (MaxHeight < 0) MaxHeight = Row.Height;

            if (Adjustment != 1 && Adjustment >= 0) NewCellHeight = (int)Math.Round(NewCellHeight * Adjustment);
            NewCellHeight += AdjustmentFixed;

            if (MaxHeight > 0 && NewCellHeight > MaxHeight) NewCellHeight = MaxHeight;
            if (NewCellHeight < MinHeight) NewCellHeight = MinHeight;

            if (NewCellHeight > XlsConsts.MaxRowHeight) NewCellHeight = XlsConsts.MaxRowHeight;
            Row.Height = (UInt16)NewCellHeight;
            if (!KeepAutofit) Row.ManualHeight();
        }

        private void FixCellsWithMergedRows(ExcelFile Workbook, TAutofitMerged AutofitMerged, TMultipleCellAutofitList MergedCells, bool OnlyMarkedForAutofit, bool KeepAutofit, float aAdjustment, int aAdjustmentFixed, int aMinHeight, int aMaxHeight)
        {
            for (int i = 0; i < MergedCells.Count; i++)
            {
                TMultipleCellAutofit mc = (TMultipleCellAutofit)MergedCells[i];

                int ExtraRows = 0;
                for (int r = mc.Cell.Top; r <= mc.Cell.Bottom; r++)
                {
					if (Workbook.GetRowHidden(r) && Workbook.GetRowOutlineLevel(r) == 0) continue;
					ExtraRows += Workbook.GetRowHeight(r, false); //hidden rows with outlining should not be considered
                }

                if (mc.NeededSize > ExtraRows)
                {
                    if (AutofitMerged == TAutofitMerged.Balanced)
                    {
                        int delta = (mc.NeededSize - ExtraRows) / mc.Cell.RowCount;
                        for (int r = mc.Cell.Top; r <= mc.Cell.Bottom; r++)
                        {
                            int d = r == mc.Cell.Bottom ? (mc.NeededSize - ExtraRows) - delta * (mc.Cell.RowCount - 1) : delta; //In the last row compensate for rounding errors.
                            RefitOneRow(r, OnlyMarkedForAutofit, KeepAutofit, aAdjustment, aAdjustmentFixed, aMinHeight, aMaxHeight, mc, d);                            
                        }
                    }
                    else
                    {
                        int RowToGrow = GetCellToGrow(mc.Cell.Top, mc.Cell.Bottom, AutofitMerged);
                        RefitOneRow(RowToGrow, OnlyMarkedForAutofit, KeepAutofit, aAdjustment, aAdjustmentFixed, aMinHeight, aMaxHeight, mc, mc.NeededSize - ExtraRows);
                    }
                }
            }
        }


        private void RefitOneRow(int RowToGrow, bool OnlyMarkedForAutofit, bool KeepAutofit, float aAdjustment, int aAdjustmentFixed, int aMinHeight, int aMaxHeight, TMultipleCellAutofit mc, int DeltaHeight)
        {
            AddRow(mc.Cell.Top - 1);
            AddRow(RowToGrow - 1);
            TRowRecord FirstRow = FList[mc.Cell.Top - 1].RowRecord;
            TRowRecord Row = FList[RowToGrow - 1].RowRecord;
            int NewHeight = Row.Height + DeltaHeight;
            if (NewHeight < 0) NewHeight = 0;
            ResizeRow(FirstRow, Row, NewHeight, OnlyMarkedForAutofit, KeepAutofit, aAdjustment, aAdjustmentFixed, aMinHeight, aMaxHeight);
        }

		internal void RecalcColWidths(ExcelFile Workbook, int Col1, int Col2, bool IgnoreStrings, bool OnlyMarkedForAutofit, 
			float aAdjustment, int aAdjustmentFixed, int aMinWidth, int aMaxWidth, TAutofitMerged AutofitMerged)
		{
			using (TColWidthCalc ColCalc = new TColWidthCalc(FGlobals))
			{
				float RowMultDisplay = ExcelMetrics.RowMultDisplay(Workbook) * 100F / FlexCel.Render.FlexCelRender.DispMul;
				float ColMultDisplay = ExcelMetrics.ColMultDisplay(Workbook) * 100F / FlexCel.Render.FlexCelRender.DispMul;
				TMultipleCellAutofitList MergedCells = new TMultipleCellAutofitList();

				for (int c = Col1; c <= Col2; c++)
				{
					if (OnlyMarkedForAutofit && (FColInfoList[c] == null || (!FColInfoList[c].MarkforAutofit && !FColInfoList[c].HasMergedCell) )) continue;

					bool ColNeedsAutofit = FColInfoList[c] == null? true: FColInfoList[c].MarkforAutofit;
					AutofitColumn(Workbook, OnlyMarkedForAutofit, ColNeedsAutofit, c + 1, ColCalc, RowMultDisplay, ColMultDisplay, IgnoreStrings, aAdjustment, aAdjustmentFixed, aMinWidth, aMaxWidth, MergedCells);
				}

				//Merged Cells
				if (AutofitMerged != TAutofitMerged.None)
				{
					FixCellsWithMergedCols(Workbook, AutofitMerged, MergedCells, OnlyMarkedForAutofit, aAdjustment, aAdjustmentFixed, aMinWidth, aMaxWidth);
				}

			}
		}

        private void ReadLocalAdjustmentInCol(bool OnlyMarkedForAutofit, int c, ref float Adjustment, ref int AdjustmentFixed, ref int MinWidth, ref int MaxWidth)
        {
            if (OnlyMarkedForAutofit)
            {
                float adj = FColInfoList[c].AutofitAdjustment;
                if (adj > 0) Adjustment = adj;

                int adj2 = FColInfoList[c].AutofitAdjustmentFixed;
                if (adj2 != 0) AdjustmentFixed = adj2;

                int m = FColInfoList[c].MinWidth; if (m != 0) MinWidth = m;
                m = FColInfoList[c].MaxWidth; if (m != 0) MaxWidth = m;

            }
        }

		private void AutofitColumn(ExcelFile Workbook, bool OnlyMarkedForAutofit, bool ColNeedsAutofit, int Column, TColWidthCalc ColCalc, float RowMultDisplay, float ColMultDisplay, bool IgnoreStrings, 
			float aAdjustment, int aAdjustmentFixed, int aMinWidth, int aMaxWidth, TMultipleCellAutofitList MergedCells)
		{
			int MaxCellWidth = 0;
			for (int r = Workbook.RowCount; r> 0; r--)
			{
				TRichString val = Workbook.GetStringFromCell(r, Column);
				if (val == null || val.Value == null || val.Value.Length == 0) continue;
				if (IgnoreStrings)
				{
					object obj = Workbook.GetCellValue(r, Column);
					if (obj is String || obj is TRichString) continue;
				}

				int ColumnXF = Workbook.GetCellVisibleFormat(r, Column);

				int CellWidth = ColCalc.CalcCellWidth(r, Column, val, ColumnXF, Workbook, RowMultDisplay, ColMultDisplay, MergedCells);
				if (CellWidth>MaxCellWidth) MaxCellWidth = CellWidth;
			}

            if (!OnlyMarkedForAutofit || ColNeedsAutofit) ResizeCol(Workbook, Column, Column, MaxCellWidth, OnlyMarkedForAutofit, aAdjustment, aAdjustmentFixed, aMinWidth, aMaxWidth);
        }

        private void ResizeCol(ExcelFile Workbook, int FirstColumn, int Column, int NewCellWidth, bool OnlyMarkedForAutofit, float aAdjustment, int aAdjustmentFixed, int aMinWidth, int aMaxWidth)
        {
            float Adjustment = aAdjustment; int AdjustmentFixed = aAdjustmentFixed; int MinWidth = aMinWidth; int MaxWidth = aMaxWidth;
            ReadLocalAdjustmentInCol(OnlyMarkedForAutofit, FirstColumn - 1, ref Adjustment, ref AdjustmentFixed, ref MinWidth, ref MaxWidth);
            
            if (MinWidth < 0) MinWidth = Workbook.GetColWidth(Column);
            if (MaxWidth < 0) MaxWidth = Workbook.GetColWidth(Column);

            if (Adjustment != 1 && Adjustment >= 0) NewCellWidth = (int)Math.Round(NewCellWidth * Adjustment);
            NewCellWidth += AdjustmentFixed;

            if (MaxWidth > 0 && NewCellWidth > MaxWidth) NewCellWidth = MaxWidth;
            if (NewCellWidth < MinWidth) NewCellWidth = MinWidth;
            if (NewCellWidth > 0xFFFF) NewCellWidth = 0xFFFF;

            if (NewCellWidth > 0) Workbook.SetColWidthInternal(Column, NewCellWidth);
        }

        private void FixCellsWithMergedCols(ExcelFile Workbook, TAutofitMerged AutofitMerged, TMultipleCellAutofitList MergedCells, bool OnlyMarkedForAutofit, float aAdjustment, int aAdjustmentFixed, int aMinWidth, int aMaxWidth)
        {
            for (int i = 0; i < MergedCells.Count; i++)
            {
                TMultipleCellAutofit mc = (TMultipleCellAutofit)MergedCells[i];

                int ExtraCols = 0;
                for (int c = mc.Cell.Left; c <= mc.Cell.Right; c++)
                {
					if (Workbook.GetColHidden(c) && Workbook.GetColOutlineLevel(c) == 0) continue;
					ExtraCols += Workbook.GetColWidth(c, false);
                }

                if (mc.NeededSize > ExtraCols)
                {
                    if (AutofitMerged == TAutofitMerged.Balanced)
                    {
                        int delta = (mc.NeededSize - ExtraCols) / mc.Cell.ColCount;
                        for (int c = mc.Cell.Left; c <= mc.Cell.Right; c++)
                        {
                            int d = c == mc.Cell.Right ? (mc.NeededSize - ExtraCols) - delta * (mc.Cell.ColCount - 1) : delta; //In the last column compensate for rounding errors.
                            RefitOneCol(Workbook, c, OnlyMarkedForAutofit, aAdjustment, aAdjustmentFixed, aMinWidth, aMaxWidth, mc, d);
                        }
                    }
                    else
                    {
                        int ColToGrow = GetCellToGrow(mc.Cell.Left, mc.Cell.Right, AutofitMerged);
                        RefitOneCol(Workbook, ColToGrow, OnlyMarkedForAutofit, aAdjustment, aAdjustmentFixed, aMinWidth, aMaxWidth, mc, mc.NeededSize - ExtraCols);
                    }
                }
            }
        }


        private void RefitOneCol(ExcelFile Workbook, int ColToGrow, bool OnlyMarkedForAutofit, float aAdjustment, int aAdjustmentFixed, int aMinWidth, int aMaxWidth, TMultipleCellAutofit mc, int DeltaWidth)
        {
            int NewWidth = Workbook.GetColWidth(ColToGrow, false) + DeltaWidth;
            if (NewWidth < 0) NewWidth = 0;
            ResizeCol(Workbook, mc.Cell.Left, ColToGrow, NewWidth, OnlyMarkedForAutofit, aAdjustment, aAdjustmentFixed, aMinWidth, aMaxWidth);
        }

#endif		


		#endregion

        internal void SetValueWithArrays(int Row, int Col, object Value, int ValueXF)
        {
            TFormula fmla = Value as TFormula;
            if (fmla != null)
            {
                TCellAddress colInputCell, rowInputCell;
                if (IsWhatIf(fmla.Text, out rowInputCell, out colInputCell))
                {
                    if (fmla.Span.IsTopLeft)
                    {
                        AddWhatIfTable(new TXlsCellRange(Row, Col, Row + fmla.Span.RowSpan - 1, Col + fmla.Span.ColSpan - 1), rowInputCell, colInputCell, ValueXF);
                    }
                    else
                    {
                        SetFormat(Row, Col, ValueXF);
                    }

                    return;
                }

                if (IsArray(fmla.Text))
                {
                    if (fmla.Span.IsTopLeft)
                    {
                        AddArrayFmla(Row, Col, fmla, ValueXF);
                    }
                    else
                    {
                        SetFormat(Row, Col, ValueXF);
                    }
                    return;
                }
            }

            SetValue(Row, Col, Value, ValueXF);
        }

        private bool IsWhatIf(string text, out TCellAddress rowInputCell, out TCellAddress colInputCell)
        {
            rowInputCell = null;
            colInputCell = null;
            if (text == null) return false;
            string StartTable = TBaseFormulaParser.fts(TFormulaToken.fmOpenArray) + TBaseFormulaParser.fts(TFormulaToken.fmStartFormula)
                + TBaseFormulaParser.fts(TFormulaToken.fmTableText) + TBaseFormulaParser.fts(TFormulaToken.fmOpenParen);

            if (text.Length < StartTable.Length || !String.Equals(text.Substring(0, StartTable.Length), StartTable, StringComparison.InvariantCultureIgnoreCase))
                return false;

            string EndTable = TBaseFormulaParser.fts(TFormulaToken.fmCloseParen) + TBaseFormulaParser.fts(TFormulaToken.fmCloseArray);
            if (!text.EndsWith(EndTable))
                return false;

            string CellRef = text.Substring(StartTable.Length, text.Length - StartTable.Length - EndTable.Length);

            string[] BothRefs = CellRef.Split(TFormulaMessages.TokenChar(TFormulaToken.fmFunctionSep));
            if (BothRefs == null || BothRefs.Length != 2) return false;
 
            TReferenceStyle rs = Workbook.FormulaReferenceStyle;
            rowInputCell = GetInputCell(BothRefs[0], rs);
            colInputCell = GetInputCell(BothRefs[1], rs);
            return true;
        }

        private static TCellAddress GetInputCell(string text, TReferenceStyle referenceStyle)
        {
            if (text == null) return null;
            text = text.Trim();
            if (text.Length == 0) return null;

            TCellAddress Result = new TCellAddress();
            if (!Result.TrySetCellRef(text, referenceStyle, 0, 0)) return null;
            return Result;
        }

        private static bool IsArray(string p)
        {
            return p != null && p.Length > 1 && p[0] == TBaseFormulaParser.ft(TFormulaToken.fmOpenArray);
        }

        internal void SetValue(int Row, int Col, object Value, int ValueXF)
        {
            if ((Row < 0) || (Row > FlxConsts.Max_Rows)) XlsMessages.ThrowException(XlsErr.ErrInvalidRow, Row);
            if ((Col > FlxConsts.Max_Columns) || (Col < 0)) XlsMessages.ThrowException(XlsErr.ErrInvalidCol, Col);

            AddRow(Row);

            int DefaultXF = FlxConsts.DefaultFormatId;
            int Index = 0;
            if (FList[Row].RowRecord.IsFormatted()) DefaultXF = FList[Row].RowRecord.XF;
            else if (FColInfoList[Col] != null) DefaultXF = FColInfoList[Col].XF;

            bool Found = (Row < Count) && this[Row].Find(Col, ref Index);
            int XF = DefaultXF;

            if (ValueXF >= 0) XF = ValueXF;
            else if (Found) XF = this[Row][Index].XF;

            TCellRecord Cell = CreateCellRecord(Row, Col, ref Value, DefaultXF, XF);

            if (Cell == null)
            {
                if (Found) this[Row].Delete(Index); //This will decrease references on labelssts.
                return;
            }

            //Not needed now. we are not keeping MaxCol/MinCol in sync.
            //if (Col+1> FRowRecordList[Row].MaxCol) FRowRecordList[Row].MaxCol=(UInt16)(Col+1);
            //if (Col< FRowRecordList[Row].MinCol) FRowRecordList[Row].MinCol=(UInt16)Col;

            if (Row >= Count) AddRecord(Cell, Row, null);
            else
                if (Found) this[Row][Index] = Cell; else this[Row].Insert(Index, Cell);
        }

        private TCellRecord CreateCellRecord(int Row, int Col, ref object Value, int DefaultXF, int XF)
        {
            TFormula Fmla = null;
            TCellRecord Cell = null;
            if (Value is char[]) Value = new string((char[])Value);  //Delphi 8 bdp returns char[] for memo fields.

            switch (TExcelTypes.ObjectToCellType(Value))
            {
                case TCellType.Empty:
                    if (XF != DefaultXF) Cell = new TBlankRecord(Col, XF); //.CreateFromData
                    break;
                case TCellType.Number:
                    double RealValue = Convert.ToDouble(Value, CultureInfo.CurrentCulture);
                    if (TRKRecord.IsRK(RealValue)) Cell = new TRKRecord(Col, XF, RealValue); //.CreateFromData
                    else Cell = new TNumberRecord(Col, XF, RealValue); //.CreateFromData
                    break;

                case TCellType.DateTime:
                    double RealDate = FlxDateTime.ToOADate((DateTime)Value, Dates1904);
                    if (TRKRecord.IsRK(RealDate)) Cell = new TRKRecord(Col, XF, RealDate); //.CreateFromData
                    else Cell = new TNumberRecord(Col, XF, RealDate); //.CreateFromData
                    break;

                case TCellType.String:
                case TCellType.Unknown:
					TRichString rs = Value as TRichString;
                    string RealString = Value.ToString();
					if (RealString.Length == 0)
					{
						if (XF != DefaultXF) Cell = new TBlankRecord(Col, XF); //.CreateFromData
					}
					else 
					{
						if (rs != null)
							Cell = new TLabelSSTRecord(Col, XF, FGlobals.SST, FGlobals.Workbook, rs); //.CreateFromData
						else
							Cell = new TLabelSSTRecord(Col, XF, FGlobals.SST, FGlobals.Workbook, Value); //.CreateFromData
					}
                    break;

                case TCellType.Bool:
                    Cell = new TBoolErrRecord(Col, XF, Convert.ToBoolean(Value, CultureInfo.CurrentCulture)); //.CreateFromData
                    break;

                case TCellType.Error:
                    Cell = new TBoolErrRecord(Col, XF, (TFlxFormulaErrorValue)Value);
                    break;

                case TCellType.Formula:
                    Fmla = (TFormula)Value;

                    TParsedTokenList FmlaData;
                    TParsedTokenList FmlaArrayData = null;

                    //We will try to use the data. Functions can have 3 different results (value, reference and array) and parameter types, so it is safer to use the working one. Also this way we keep the optimized ifs.
                    if (Fmla.Data == null || TTokenManipulator.HasExternRefs(Fmla.Data))  //Use Text
                    {
                        TFormulaConvertTextToInternal Ps = new TFormulaConvertTextToInternal(FGlobals.Workbook, FGlobals.Workbook.ActiveSheet, true, Fmla.Text, true);
                        Ps.SetStartForRelativeRefs(Row, Col);
                        Ps.Parse();
                        FmlaData = Ps.GetTokens();
                        FmlaArrayData = null;
                    }
                    else
                    {
                        FmlaData = Fmla.Data.Clone(); //Only clone if needed.
                        if (Fmla.ArrayData != null) FmlaArrayData = Fmla.ArrayData.Clone();
                    }

                    int OptionFlags = 0; int ArrayOptionFlags = 0;
                    Cell = new TFormulaRecord((int)xlr.FORMULA, Row, Col, null, XF, FmlaData, FmlaArrayData, Fmla.Result, OptionFlags, Dates1904, ArrayOptionFlags, false);
                    break;

                //Celltype.error can't be directly entered. It should be a formula.
            } //case
            return Cell;
        }

		internal void DeleteEmptyRowRecords(int DefaultRowHeight, int DefaultFlags, TNoteList Notes)  //no need for mergedcells here, there should be a blank record in place.
		{
			bool DefaultIsZero = (DefaultFlags & 0x2) != 0;
			//Remove all empty Rows.
            for (int k = FList.Count - 1; k >= 0; k--)
            {
                if (this[k] == null || this[k].Count == 0) //no cells on the row.
                {
                    if (this[k].RowRecord != null && !this[k].RowRecord.IsModified(DefaultRowHeight, DefaultIsZero)) // row does not have useful data.
                    {
                        if (Notes.HasNotes(k)) continue;  //no comments on the row.
                        if (k == FList.Count - 1) FList.RemoveAt(k); else FList[k].RowRecord = null;
                    }
                }
                else
                {
                    if (!HasRow(k)) FList[k].RowRecord = CreateDefaultRow();
                }
            }
		}

        internal void GetValue(ExcelFile aXls, int Row, int Col, int ColIndex, ref object V, ref int XF)
        {
#if(!COMPACTFRAMEWORK)
#endif

            if ((Row < 0) || (Row > FlxConsts.Max_Rows)) XlsMessages.ThrowException(XlsErr.ErrInvalidRow, Row);
            if (ColIndex < 0)
            {
                if ((Col > FlxConsts.Max_Columns) || (Col < 0)) XlsMessages.ThrowException(XlsErr.ErrInvalidCol, Col);
            }
            else
                if ((ColIndex >= this[Row].Count)) XlsMessages.ThrowException(XlsErr.ErrInvalidCol, Col);

            if (Row >= Count) { V = null; XF = -1; return; }
            int Index = ColIndex;
            bool Found = true;
            if (Index < 0) Found = (this[Row].Find(Col, ref Index));
            if (Found)
            {
                XF = this[Row][Index].XF;
                V = this[Row][Index].GetValue(this);
            }
            else
            { V = null; XF = -1; return; }
        }

        internal object GetValueAndRecalc(int SheetIndex, int Row, int Col, ExcelFile aXls, TCalcState CalcState, TCalcStack CalcStack)
        {
            if ((Row < 0) || (Row > FlxConsts.Max_Rows)) XlsMessages.ThrowException(XlsErr.ErrInvalidRow, Row);
            if ((Col > FlxConsts.Max_Columns) || (Col < 0)) XlsMessages.ThrowException(XlsErr.ErrInvalidCol, Col);

			if (CalcState.WhatIfRow > 0) // so even if the cell is null we still get the value.
			{
				if (CalcState.TableRowCell.IsCell(Row + 1, Col + 1)) return CalcState.TableRowValue;
				if (CalcState.TableColCell.IsCell(Row + 1, Col + 1)) return CalcState.TableColValue;
			}

            if (Row >= Count) { return null; }
            int Index = -1;
            if (this[Row].Find(Col, ref Index))
            {
                return GetRecalculatedCell(SheetIndex, Row, aXls, CalcState, CalcStack, Index, true);
            }
            else
            { return null; }
        }

        private object GetRecalculatedCell(int SheetIndex, int Row, ExcelFile aXls, TCalcState CalcState, TCalcStack CalcStack, int Index, bool ReturnFmlaValues)
        {
            TCellRecord r = this[Row][Index];

			if (CalcState.WhatIfRow > 0)
			{
				if (CalcState.TableRowCell.IsCell(Row + 1, r.Col + 1)) return CalcState.TableRowValue;
				if (CalcState.TableColCell.IsCell(Row + 1, r.Col + 1)) return CalcState.TableColValue;
			}


            TFormulaRecord f = (r as TFormulaRecord);
            if (f != null)
            {
                if (aXls != null)
                {
                    aXls.SetUnsupportedFormulaCellAddress(new TCellAddress(aXls.GetSheetName(SheetIndex + 1), Row + 1, f.Col + 1, false, false));
                }

				f.SetWhatIf(CalcState.WhatIfRow, CalcState.WhatIfCol, CalcState.WhatIfSheet);

                f.Recalc(this, aXls, SheetIndex + 1, CalcState, CalcStack);
                
                if ((f.HasSubtotal || f.HasAggregate) && CalcState.InSubTotal) return null;
                
				if (ReturnFmlaValues) return f.WhatIfFormulaValue; else return f.GetValue(this);
            }
            return r.GetValue(this);
        }


        internal object CalcExpression(ExcelFile aXls, int SheetIndexBase1, string FormulaText)
        {
            if (FormulaText == null || FormulaText.Length == 0) return null;

            TFormulaConvertTextToInternal Ps = new TFormulaConvertTextToInternal(aXls, SheetIndexBase1, false, FormulaText, true);
            Ps.Parse();
            TParsedTokenList FmlaData = Ps.GetTokens();
            TFormulaRecord r = new TFormulaRecord((int)xlr.FORMULA, 1, 1, null, 0, FmlaData, null, null, 0, aXls.OptionsDates1904, 0, false);
            r.NotRecalculated(); r.NotRecalculating();
            r.Recalc(this, aXls, SheetIndexBase1, new TCalcState(), new TCalcStack());
            return r.WhatIfFormulaValue;
        }

		private void ClearCell(int Row, int Col)
		{
			int Index = -1;
			bool Found=(Row<Count) && this[Row].Find(Col, ref Index);
			if (Found) this[Row].Delete(Index); //This will decrease references on labelssts.
		}

		private TCellRecord GetCell(int Row, int Col)
		{
			int Index = -1;
			bool Found=(Row<Count) && this[Row].Find(Col, ref Index);
			if (Found) return this[Row][Index]; 
			return null;
		}

        internal void CopyFmt(int SourceSheetIndex, int DestSheetIndex, TCellList Source, int SourceRow, int SourceCol, int DestRow, int DestCol, TSheet SourceSheet, TSheet DestSheet)
        {
            TCellRecord SourceCell = Source.GetCell(SourceRow, SourceCol);
            TCellRecord DestCell = GetCell(DestRow, DestCol);
            if (SourceCell == null || SourceCell.XF == FlxConsts.DefaultFormatId)
            {
                if (DestCell == null) ClearCell(DestRow, DestCol);
                else
                {
                    DestCell.XF = FlxConsts.DefaultFormatId;
                }
            }

            if (DestCell == null)
            {
                DestCell = new TBlankRecord(DestCol, SourceCell.XF);
                AddRow(DestRow);
                AddRecord(DestCell, DestRow, null);
            }

            DestCell.XF = SourceCell.XF;
            if (Source.FGlobals != FGlobals)
            {
                FixCopyFormat(Source.FGlobals, FGlobals, SourceCell.XF, ref DestCell.XF);
            }

        }

        internal void CopyCell(int SourceSheetIndex, int DestSheetIndex, TCellList Source, int SourceRow, int SourceCol, int DestRow, int DestCol, TRangeCopyMode CopyMode, TSheet SourceSheet, TSheet DestSheet)
        {
            if (CopyMode == TRangeCopyMode.None) return;
            if (CopyMode == TRangeCopyMode.Formats)
            {
                CopyFmt(SourceSheetIndex, DestSheetIndex, Source, SourceRow, SourceCol, DestRow, DestCol, SourceSheet, DestSheet);
                return;
            }

            TCellRecord SourceCell = Source.GetCell(SourceRow, SourceCol);
            if (CopyMode != TRangeCopyMode.All && CopyMode != TRangeCopyMode.AllIncludingDontMoveAndSizeObjects)
            {
                if (!SourceCell.AllowCopyOnOnlyFormula) return;
            }

            if (SourceCell == null)
            {
                ClearCell(DestRow, DestCol);
                return;
            }

            TSheetInfo SheetInfo = new TSheetInfo(SourceSheetIndex, SourceSheetIndex, DestSheetIndex, Source.FGlobals, FGlobals, SourceSheet, DestSheet, false);  //SemiAbsoluteMode here doesn't matter, since we are copying just one cell.

            TCellRecord DestCell;

            ClearCell(DestRow, DestCol);
            TLabelSSTRecord sst = SourceCell as TLabelSSTRecord;
            if (sst != null)
            {
                DestCell = TLabelSSTRecord.CreateFromOtherString(SourceCol, SourceCell.XF, FGlobals.SST, FGlobals.Workbook, sst.GetValue(null));
            }
            else
            {
                DestCell = (TCellRecord)TCellRecord.Clone(SourceCell, SheetInfo);
            }

            if (Source.FGlobals != FGlobals)
            {
                FixCopyFormat(Source.FGlobals, FGlobals, SourceCell.XF, ref DestCell.XF);
            }

            if (DestCell.XF == FlxConsts.DefaultFormatId && DestCell is TBlankRecord)
            {
                ClearCell(DestRow, DestCol);
                return;
            }

            DestCell.ArrangeCopyRange(new TXlsCellRange(0, 0, -1, -1), DestRow, DestRow - SourceRow, DestCol - SourceCol, SheetInfo);

            AddRow(DestRow);
            AddRecord(DestCell, DestRow, null);
        }

        internal static void FixCopyFont(int i, TWorkbookGlobals SourceGlobals, TWorkbookGlobals DestGlobals, byte[] SourceData, byte[] DestData)
        {
            if (SourceGlobals == null) return;

            int FontIndex = BitOps.GetWord(SourceData, i);
            TFontRecord SourceFnt = SourceGlobals.Fonts.GetFontRecord(FontIndex);

            if (SourceGlobals == DestGlobals && SourceFnt.Reuse) return; //No need to fix it.

            int DestFontIndex;
            if (!SourceFnt.Reuse && SourceFnt.CopiedTo > 0)
                DestFontIndex = SourceFnt.CopiedTo;
            else
            {
                TFlxFont Font = SourceGlobals.Fonts.GetFont(FontIndex);
                DestFontIndex = DestGlobals.Fonts.AddFont(Font);
            }
            BitOps.SetWord(DestData, i, DestFontIndex);
        }

        internal static void FixCopyFormat(TWorkbookGlobals SourceGlobals, TWorkbookGlobals DestGlobals, int SourceCellXF, ref int DestCellXF)
        {
            if (SourceGlobals == null) return;
            if (SourceGlobals == DestGlobals) return; //No need to fix it.

            TFlxFormat SourceFmt = SourceGlobals.CellXF[SourceCellXF].FlxFormat(SourceGlobals.Styles, SourceGlobals.Fonts, SourceGlobals.Formats, SourceGlobals.Borders, SourceGlobals.Patterns);

            SourceFmt.FixColors(SourceGlobals.Workbook, DestGlobals.Workbook);

            //When copying from different files, we can't reuse the Style selection in the source workbook.
            //Say Normal format in file2 has font = times new roman, and normal 1 has arial. When copying a cell with arial to 2, font must not be linked to the style.
            //SourceFmt.LinkedStyle.AutomaticChoose = false; 


            if (SourceFmt.ParentStyle != null) //style must be set before the cell, so AutomaticChoose makes sense when entering the cell.
            {
                int stylefmt = DestGlobals.Styles.GetStyle(SourceFmt.ParentStyle);
                if (stylefmt < 0) //style doesn't exists in dest file, must be copied.
                {
                    TFlxFormat fmt = SourceGlobals.GetStyleFormat(SourceGlobals.Styles.GetStyle(SourceFmt.ParentStyle));
                    fmt.FixColors(SourceGlobals.Workbook, DestGlobals.Workbook);
                    DestGlobals.Styles.SetStyle(SourceFmt.ParentStyle, DestGlobals.AddStyleFormat(fmt, SourceFmt.ParentStyle));
                }
            }


            TXFRecord XF = new TXFRecord(SourceFmt, false, DestGlobals, false);
            if (DestGlobals.CellXF.FindFormat(XF, ref DestCellXF)) return;

            DestGlobals.CellXF.Add(XF);
            DestCellXF = DestGlobals.CellXF.Count - 1;
        }


        internal int GetXFFormat(int Row, int Col, int ColIndex)
        {
            if ((Row<0) || (Row>FlxConsts.Max_Rows)) XlsMessages.ThrowException(XlsErr.ErrInvalidRow,Row);
            if (ColIndex<0)
            {
                if ((Col>FlxConsts.Max_Columns) || (Col<0)) XlsMessages.ThrowException(XlsErr.ErrInvalidCol,Col);
            }
            else
                if ((ColIndex>= this[Row].Count)) XlsMessages.ThrowException(XlsErr.ErrInvalidCol,Col);

            if (Row>=Count) return -1;
            int Index=ColIndex;
            bool Found=true;
            if (Index<0)Found = (this[Row].Find(Col,ref Index));
            if (Found)
            {
                return this[Row][Index].XF; 
            } 
            else
                return -1;
        }
        

        internal void SetFormat(int Row, int Col, int XF)
        {
            if ((Row<0) || (Row>FlxConsts.Max_Rows)) XlsMessages.ThrowException(XlsErr.ErrInvalidRow,Row);
            if ((Col>FlxConsts.Max_Columns) || (Col<0)) XlsMessages.ThrowException(XlsErr.ErrInvalidCol,Col);

            int Index=-1;
            if ((Row < Count) && HasRow(Row) && this[Row].Find(Col, ref Index))
			{
				if (XF < 0) XF = FlxConsts.DefaultFormatId;
				this[Row][Index].XF=XF; 
			}
			else
			{
				if (XF < 0) return;
				SetValue(Row,Col,null,XF);
			}
        }

        public TParsedTokenList ArrayFormula(int Row, int Col)
        {
            if ((Row < 0) || (Row >= Count)) XlsMessages.ThrowException(XlsErr.ErrInvalidRow, Row);
            if ((Col > FlxConsts.Max_Columns) || (Col < 0)) XlsMessages.ThrowException(XlsErr.ErrInvalidCol, Col);
            int Index = -1;
            if (this[Row].Find(Col, ref Index) && (this[Row][Index] is TFormulaRecord))
            {
                TFormulaRecord Fmla = (TFormulaRecord)this[Row][Index];
                if (Fmla.ArrayRecord == null) XlsMessages.ThrowException(XlsErr.ErrBadFormula, Row, Col, 1);
                return Fmla.ArrayRecord.Data;
            }
            else
            {
                XlsMessages.ThrowException(XlsErr.ErrShrFmlaNotFound);
                return null;
            }
        }

        public bool FoundArrayFormula(int Row, int Col, out TArrayRecord ArrResult)
        {
            ArrResult = null;
            if ((Row < 0) || (Row >= Count)) return false;
            if ((Col > FlxConsts.Max_Columns) || (Col < 0)) return false;
            int Index = -1;
            if (this[Row].Find(Col, ref Index) && (this[Row][Index] is TFormulaRecord))
            {
                TFormulaRecord Fmla = (TFormulaRecord)this[Row][Index];
                if (Fmla.ArrayRecord == null) return false;
                ArrResult = Fmla.ArrayRecord;
                return true;
            }
            return false;
        }

        public TTableRecord TableFormula(int Row, int Col)
        {
            if ((Row<0) || (Row>=Count)) XlsMessages.ThrowException(XlsErr.ErrInvalidRow,Row);
            if ((Col>FlxConsts.Max_Columns) || (Col<0)) XlsMessages.ThrowException(XlsErr.ErrInvalidCol,Col);
            int Index=-1;
            if (this[Row].Find(Col, ref Index) && (this[Row][Index] is TFormulaRecord)) 
            {
                TFormulaRecord Fmla=(TFormulaRecord)this[Row][Index];
                if (Fmla.TableRecord==null) XlsMessages.ThrowException(XlsErr.ErrBadFormula,Row, Col,1);
                return Fmla.TableRecord;
            } 
            else
            {
                XlsMessages.ThrowException(XlsErr.ErrShrFmlaNotFound);
                return null;
            }
        }

        public bool FoundTableFormula(int Row, int Col, out TTableRecord TableResult)
        {
            TableResult = null;
            if ((Row < 0) || (Row >= Count)) return false;
            if ((Col > FlxConsts.Max_Columns) || (Col < 0)) return false;
            int Index = -1;
            if (this[Row].Find(Col, ref Index) && (this[Row][Index] is TFormulaRecord))
            {
                TFormulaRecord Fmla = (TFormulaRecord)this[Row][Index];
                if (Fmla.TableRecord == null) return false;
                TableResult = Fmla.TableRecord;
                return true;
            }
            return false;
        }

        internal void Sort()
        {
            int aCount=Count;
            for (int i=0; i< aCount;i++)
            {
                TCellRecordList it=this[i];
                it.SortOptional();
            }
        }

		internal override void ArrangeMoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
		{
			FormulaCache.ArrangeMoveRange(CellRange, NewRow, NewCol, SheetInfo);

			if ((CellRange.Top<=0) && (CellRange.Bottom>=FlxConsts.Max_Rows))
				FColInfoList.ArrangeMoveCols(CellRange, NewCol, SheetInfo);
		}


        internal override void ArrangeInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
        {
            //No need here. only on colcount > 0  base.ArrangeInsertRange (CellRange, aRowCount, aColCount, SheetInfo);

            if (aRowCount != 0)
            {
                FormulaCache.ArrangeInsertRangeRows(CellRange, aRowCount, SheetInfo);
            }

            if (aColCount != 0)
            {
                base.ArrangeInsertRange(CellRange, aRowCount, aColCount, SheetInfo);
                FormulaCache.ArrangeInsertRangeCols(CellRange, aColCount, SheetInfo);
            }

            if ((CellRange.Top <= 0) && (CellRange.Bottom >= FlxConsts.Max_Rows))
                FColInfoList.ArrangeInsertCols(CellRange, aColCount, SheetInfo);
        }

        internal override void InsertAndCopyRange(TXlsCellRange SourceRange, TFlxInsertMode InsertMode, int DestRow, int DestCol, int aRowCount, int aColCount, TRangeCopyMode CopyMode, TSheetInfo SheetInfo)
        {
            base.InsertAndCopyRange(SourceRange, InsertMode, DestRow, DestCol, aRowCount, aColCount, CopyMode, SheetInfo);

            if ((SourceRange.Top <= 0) && (SourceRange.Bottom >= FlxConsts.Max_Rows) && (aColCount > 0) && (CopyMode != TRangeCopyMode.None))
            {
                TXlsCellRange sr = SourceRange;
                if (DestCol <= sr.Left && sr.Top <= 0 && sr.Bottom >= FlxConsts.Max_Rows) sr = sr.Offset(sr.Top, sr.Left + sr.ColCount * aColCount);
                FColInfoList.CopyCols(sr, DestCol, aColCount, SheetInfo);
            }
        }

        internal void ArrangeInsertSheet(TSheetInfo SheetInfo)
        {
            FormulaCache.ArrangeInsertSheet(SheetInfo);
        }

		#region Named Ranges
		internal void UpdateDeletedRanges(TDeletedRanges DeletedRanges)
		{
			FormulaCache.UpdateDeletedRanges(DeletedRanges);
		}
		#endregion

        private void AddArrayFmla(int Row, int Col, TFormula fmla, int ValueXF)
        {
            TFormula MainFmla = new TFormula(null, null, new TParsedTokenList(new TBaseParsedToken[] { new TExp_Token(Row, Col) }), null, true, new TFormulaSpan());
            for (int r = Row + fmla.Span.RowSpan - 1; r >= Row; r--)
            {
                for (int c = Col + fmla.Span.ColSpan - 1; c >= Col; c--)
                {
                    SetValue(r, c, MainFmla, ValueXF);
                }
            }

            TFormulaRecord f = GetCell(Row, Col) as TFormulaRecord;
            Debug.Assert(f != null, "The cell with the formula must not be null, we just added it.");

            TFormulaConvertTextToInternal Ps = new TFormulaConvertTextToInternal(FGlobals.Workbook, FGlobals.Workbook.ActiveSheet, true, fmla.Text, true);
            Ps.SetStartForRelativeRefs(Row, Col);
            Ps.Parse();
            f.ArrayRecord = new TArrayRecord((int)xlr.ARRAY, new TXlsCellRange(Row, Col, Row + fmla.Span.RowSpan -1, Col + fmla.Span.ColSpan -1),
                Ps.GetTokens(), 0);
        }

        #region What-if
        internal void AddWhatIfTable(TXlsCellRange range, TCellAddress rowInputCell, TCellAddress colInputCell, int ValueXF)
        {
			if (range.Left > range.Right || range.Top > range.Bottom) return;
            TFormula fmla = new TFormula(null, null, new TParsedTokenList(new TBaseParsedToken[] { new TTableToken(range.Top, range.Left) }), null, true, new TFormulaSpan());
            for (int r = range.Bottom; r >= range.Top; r--)
            {
                for (int c = range.Right; c >= range.Left; c--)
                {
                    SetValue(r, c, fmla, ValueXF);
                }
            }

            TFormulaRecord f = GetCell(range.Top, range.Left) as TFormulaRecord;
			Debug.Assert(f != null, "The cell with the formula must not be null, we just added it.");

            int oflags, rir, cir, ric, cic;
            CalcTableFlags(rowInputCell, colInputCell, out oflags, out rir, out cir, out ric, out cic);

            f.TableRecord = new TTableRecord((int)xlr.TABLE, oflags, range.Top, range.Left, range.Bottom, range.Right, rir, cir, ric, cic);
        }

        private static void CalcTableFlags(TCellAddress rowInputCell, TCellAddress colInputCell, out int oflags, out int rir, out int cir, out int ric, out int cic)
        {
            oflags = 0;
            if (colInputCell != null)
            {
                if (rowInputCell != null)
                {
                    oflags |= 0x08;
                }
            }
            else
            {
                oflags |= 0x04;
            }



            if (rowInputCell == null && colInputCell == null)
            {
                oflags |= TTableRecord.FlagDeleted1 | TTableRecord.FlagDeleted2;
                rir = -1;
                cir = -1;
            }
            else
            {
                rir = rowInputCell != null ? rowInputCell.Row - 1: colInputCell.Row - 1;
                cir = rowInputCell != null ? rowInputCell.Col - 1  : colInputCell.Col - 1;
            }
            ric = rowInputCell == null || colInputCell == null ? 0 : colInputCell.Row - 1;
            cic = rowInputCell == null || colInputCell == null ? 0 : colInputCell.Col - 1;
        }

        internal TCellAddress[] GetTables()
        {
            return FormulaCache.GetTables();
        }

        internal TXlsCellRange GetTable(int Sheet, int Row, int Col, out TCellAddress rowInputCell, out TCellAddress colInputCell)
        {
            rowInputCell = null;
            colInputCell = null;
            int Index = -1;
            bool Found = (Row < Count) && this[Row].Find(Col, ref Index);
            if (!Found) return null;

            TFormulaRecord f = this[Row][Index] as TFormulaRecord;
            if (f == null || f.TableRecord == null) return null;

            TCellAddress FirstCell = f.TableRecord.IsDeleted1 ? null : new TCellAddress(f.TableRecord.RwInpRw + 1, f.TableRecord.ColInpRw + 1);
            if (f.TableRecord.Has2Entries)
            {
                rowInputCell = FirstCell;
                colInputCell = f.TableRecord.IsDeleted2 ? null : new TCellAddress(f.TableRecord.RwInpCol + 1, f.TableRecord.ColInpCol + 1);
            }
            else
            {
                if (f.TableRecord.CellInputIsRow)
                {
                    rowInputCell = FirstCell;
                }
                else
                {
                    colInputCell = FirstCell;
                }
            }
            return new TXlsCellRange(f.TableRecord.FirstRow + 1, f.TableRecord.FirstCol + 1, f.TableRecord.LastRow + 1, f.TableRecord.LastCol + 1);
        }
        #endregion

        #region Rows

        internal int DefRowHeight
        {
            get
            {
                if (SheetGlobals == null || SheetGlobals.DefRowHeight == null) return 0xFF;
                return SheetGlobals.DefRowHeight.Height;
            }
        }

        
        internal void AddRowRecord(int Row, TRowRecord aRecord)
        {
            if (Row < Count)
            {
                if (FList[Row].RowRecord == null) FList[Row].RowRecord = aRecord; else XlsMessages.ThrowException(XlsErr.ErrDupRow);
            }
            else
            {
                for (int i = Count; i < Row; i++) base.Add(CreateRecord());
                TCellAndRowRecordList R = CreateRecord();
                R.RowRecord = aRecord;
                base.Add(R);
            }
        }

        internal bool HasRow(int index)
        {
            return (index >= 0) && (index < Count) && (FList[index].RowRecord != null);
        }

        internal void AddRow(int index)
        {
            if (HasRow(index)) return;
            TRowRecord aRecord = new TRowRecord(DefRowHeight);
            AddRowRecord(index, aRecord);        
        }

        internal TRowRecord CreateDefaultRow()
        {
            return new TRowRecord(DefRowHeight);
        }


		internal TCellCondFmt[] GetRowCondFmt(int aRow)
		{
			if (HasRow(aRow))
			{
				return (FList[aRow]).RowRecord.CondFmt;
			}

            TCellCondFmt[] Result;
            if (EmptyRowCF.TryGetValue(aRow, out Result)) return Result;
            return null;
		}

		internal void SetRowCondFmt(int aRow, TCellCondFmt[] fmt)
		{
			if (HasRow(aRow))
			{
				FList[aRow].RowRecord.CondFmt = fmt;
				return;
			}

			EmptyRowCF[aRow] = fmt;
		}

        internal void CleanRowCF()
        {
            EmptyRowCF.Clear();
            for (int i = FList.Count - 1; i >= 0; i--)
            {
                if (FList[i].RowRecord != null) FList[i].RowRecord.CondFmt = null;
            }
        }

        internal int RowOptions(int aRow)
        {
            if (HasRow(aRow)) return this[aRow].RowRecord.GetOptions(); else return 0;
        }

        internal void SetRowOptions(int aRow, int aOptions)
        {
            AddRow(aRow);
            this[aRow].RowRecord.SetOptions(aOptions);
        }

        internal void CollapseRows(int aRow, int Level, TCollapseChildrenMode CollapseChildren, bool IsNode)
        {
            AddRow(aRow);
            this[aRow].RowRecord.Collapse(Level, CollapseChildren, IsNode);
        }

        internal int RowHeight(int aRow)
        {
            if (!HasRow(aRow)) return 0; else return this[aRow].RowRecord.Height;
        }

        internal void SetRowHeight(int aRow, int Height)
        {
            AddRow(aRow);
            this[aRow].RowRecord.Height = (UInt16)Height;
            this[aRow].RowRecord.ManualHeight();
        }

        internal void AutoRowHeight(int aRow, bool Value)
        {
            if (HasRow(aRow))
                if (Value) this[aRow].RowRecord.AutoHeight(); else this[aRow].RowRecord.ManualHeight();
        }


        internal bool IsAutoRowHeight(int aRow)
        {
            if (HasRow(aRow)) return this[aRow].RowRecord.IsAutoHeight(); else return true;
        }


        internal void CalcRowGuts(TGutsRecord Guts)
        {
            int MaxGutLevel = 0;
            for (int i = Count - 1; i >= 0; i--)
            {
                TRowRecord rr = FList[i].RowRecord;
                if (rr != null)
                {
                    int GutLevel = rr.GetRowOutlineLevel();
                    if (GutLevel > MaxGutLevel) MaxGutLevel = GutLevel;
                }
            }
            Guts.RowLevel = MaxGutLevel;
        }

        
        #endregion

    }

    /// <summary>
    /// A list with both Cell records and ROW records.
    /// </summary>
    internal class TCells
    {
        TCellList FCellList;

        internal TCells(TWorkbookGlobals aGlobals, TColInfoList aColInfoList, TSheetGlobals aSheetGlobals)
        {
            FCellList = new TCellList(aGlobals, aColInfoList, aSheetGlobals);
        }

        internal static void WriteDimensions(IDataStream DataStream, TSaveData SaveData, TXlsCellRange CellRange)
        {

            DataStream.WriteHeader((UInt16)xlr.DIMENSIONS, 14);

            //No need to do biff8 check here since it already is 32/16 bits. And dimensions can go outside the range.
            DataStream.Write32((UInt32)CellRange.Top);
            DataStream.Write32((UInt32)(CellRange.Bottom + 1)); //This adds an extra row. Dimensions do from firstrow to lastrow+1
            DataStream.Write16((UInt16)CellRange.Left);
            DataStream.Write16((UInt16)(CellRange.Right + 1));
            DataStream.Write16(0);
        }

        internal static int DimensionsSize { get { return 14 + XlsConsts.SizeOfTRecordHeader; } }

        private void CalcUsedRange(TXlsCellRange CellRange)
        {
            CellRange.Top = 0;
            int aRowCount = FCellList.Count;
            while ((CellRange.Top < aRowCount) && (!FCellList.HasRow(CellRange.Top))) CellRange.Top++;
            CellRange.Bottom = aRowCount - 1;
            CellRange.Left = -1;
            CellRange.Right = 0;
            for (int i = CellRange.Top; i < aRowCount; i++)
                if (FCellList.HasRow(i))
                {
                    TCellRecordList Crl = null;
                    if (i < FCellList.Count) Crl = FCellList[i];
                    if ((Crl != null) && (Crl.Count > 0))
                    {
                        if (CellRange.Left == -1 || Crl[0].Col < CellRange.Left) CellRange.Left = Crl[0].Col;
                        if (Crl[Crl.Count - 1].Col + 1 > CellRange.Right) CellRange.Right = Crl[Crl.Count - 1].Col + 1;
                    }
                }

            if (CellRange.Left == -1) CellRange.Left = 0; //There are no cells on the sheet.
            if (CellRange.Right > 0) CellRange.Right--; //MaxCol is the max col+1
        }

        public TXlsCellRange UsedRange()
        {
            TXlsCellRange Result = new TXlsCellRange();
            CalcUsedRange(Result);
            return Result;
        }

        internal void Clear()
        {
            if (FCellList != null) FCellList.Clear();
        }

        internal void CopyFrom(TCells aList, TSheetInfo SheetInfo)
        {
            FCellList.CopyFrom(aList.FCellList, SheetInfo);
        }

        internal void DeleteEmptyRowRecords(int DefaultRowHeight, int DefaultFlags, TNoteList Notes)
        {
            FCellList.DeleteEmptyRowRecords(DefaultRowHeight, DefaultFlags, Notes);
        }

        internal void FixRows()
        {
            int aCount = FCellList.Count;
            for (int i = 0; i < aCount; i++)
                if (!FCellList.HasRow(i) && (FCellList[i].Count > 0)) FCellList[i].RowRecord = FCellList.CreateDefaultRow();
        }

        internal void SaveToStream(IDataStream DataStream, TSaveData SaveData, TXlsCellRange CellRange)
        {
            if (CellRange == null)
            {
                CellRange = new TXlsCellRange();
                CalcUsedRange(CellRange);
            }
            SaveRangeToStream(DataStream, SaveData, CellRange);
        }

        private void SaveRangeToStream(IDataStream DataStream, TSaveData SaveData, TXlsCellRange CellRange)
        {
            WriteDimensions(DataStream, SaveData, CellRange);
            int i = CellRange.Top;
            int CellListCount = FCellList.Count;
            while (i <= CellRange.Bottom)
            {
                int k = 0; int Written = 0;
                while ((Written < 32) && (k + i <= CellRange.Bottom)) //Write 32 ROW records, then 32 rows... etc.
                {
                    if (FCellList.HasRow(k + i))
                    {
                        /*if (FCellList[k+i].RowRecord.Row!=k+i) 
                            XlsMessages.ThrowException(XlsErr.ErrInternal);*/

                        TRowRecord it = FCellList[k + i].RowRecord;
                        TCellRecordList Crl = null;
                        if (CellListCount > k + i) Crl = FCellList[k + i];
                        if ((Crl != null) && (Crl.Count > 0))
                        {
                            it.rMinCol = (UInt16)Crl[0].Col;
                            it.rMaxCol = (UInt16)(Crl[Crl.Count - 1].Col + 1);
                        }
                        else
                        {
                            it.rMinCol = (UInt16)CellRange.Left;
                            it.rMaxCol = (UInt16)CellRange.Left;
                        }

                        it.SaveRangeToStream(DataStream, SaveData, CellRange.Left, CellRange.Right, k + i);
                        //inc(Written);  //We want 32 records in total, counting blanks. that's why not here
                    }
                    Written++;
                    k++;
                }

                for (int j = i; j < k + i; j++)
                    if ((j <= CellRange.Bottom) && (j < FCellList.Count))
                        FCellList[j].SaveRangeToStream(DataStream, SaveData, j, CellRange);

                i += k;
            }
        }

        internal void SaveToPxl(TPxlStream PxlStream, TPxlSaveData SaveData)
        {
            int i = 0;
            int RowListCount = Count;
            while (i <= RowListCount)
            {
                int k = 0; int Written = 0;
                while ((Written < 32) && (k + i <= RowListCount)) //Write 32 ROW records, then 32 rows... etc.
                {
                    if (FCellList.HasRow(k + i))
                    {
                        /*if (FRowList[k+i].Row!=k+i) 
                            XlsMessages.ThrowException(XlsErr.ErrInternal);*/

                        TRowRecord it = FCellList[k + i].RowRecord;
                        it.SaveToPxl(PxlStream, k + i, SaveData);
                        //inc(Written);  //We want 32 records in total, counting blanks. that's why not here
                    }
                    Written++;
                    k++;
                }

                for (int j = i; j < k + i; j++)
                    if ((j <= RowListCount) && (j < FCellList.Count))
                        FCellList[j].SaveToPxl(PxlStream, j, SaveData);

                i += k;
            }
        }

        internal long TotalSize(TXlsCellRange CellRange)
        {
            if (CellRange == null) return TotalSizeAll();
            return TotalRangeSize(CellRange);
        }

        private long TotalSizeAll()
        {
            return DimensionsSize + FCellList.TotalSize(null);
        }

        private long TotalRangeSize(TXlsCellRange CellRange)
        {
            return DimensionsSize + FCellList.TotalSize(CellRange);
        }

        internal void InsertAndCopyRange(TXlsCellRange SourceRange, TFlxInsertMode InsertMode, int DestRow, int DestCol, int aRowCount, int aColCount, TRangeCopyMode CopyMode, TSheetInfo SheetInfo)
        {
            FCellList.InsertAndCopyRange(SourceRange, InsertMode, DestRow, DestCol, aRowCount, aColCount, CopyMode, SheetInfo);

            //We must ensure we have rows for all the new inserted cells at the bottom. Also, when we shift down, some cells might go to previously unused rows.
            FixRows(); //We might need it even if we are inserting columns. dest rows can be below max row.
        }


        internal void DeleteRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
        {
            FCellList.DeleteRange(CellRange, aRowCount, aColCount, SheetInfo);

            if (aRowCount > 0)
                FixRows();
        }

        internal void MoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
        {
            FCellList.MoveRange(CellRange, NewRow, NewCol, SheetInfo);

            FixRows();
        }

        internal void ClearRange(TXlsCellRange CellRange)
        {
            FCellList.ClearRange(CellRange);
        }

        internal void ClearFormats(TXlsCellRange CellRange)
        {
            FCellList.ClearFormats(CellRange);
        }

        internal void ArrangeInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
        {
            FCellList.ArrangeInsertRange(CellRange, aRowCount, aColCount, SheetInfo);
        }

        internal void ArrangeMoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
        {
            FCellList.ArrangeMoveRange(CellRange, NewRow, NewCol, SheetInfo);
        }

        internal void ArrangeInsertSheet(TSheetInfo SheetInfo)
        {
            FCellList.ArrangeInsertSheet(SheetInfo);
        }

        internal void AddRow(int Row, TRowRecord aRecord)
        {
            FCellList.AddRowRecord(Row, aRecord);
        }

        internal void MarkRowForAutofit(int Row, bool Autofit, float Adjustment, int AdjustmentFixed, int minHeight, int maxHeight, bool IsMerged)
        {
            if (FCellList.HasRow(Row))
            {
                if (IsMerged)
                {
                    FCellList[Row].RowRecord.HasMergedCell = Autofit;
                }
                else
                {
                    FCellList[Row].RowRecord.MarkForAutofit = Autofit;
                }

                FCellList[Row].RowRecord.AutofitAdjustment = Adjustment;
                FCellList[Row].RowRecord.AutofitAdjustmentFixed = AdjustmentFixed;
                FCellList[Row].RowRecord.MinHeight = minHeight;
                FCellList[Row].RowRecord.MaxHeight = maxHeight;
            }
        }

        internal void KeepRowsTogeher(int Row1, int Row2, int Level, bool ReplaceLowerLevels)
        {
            if (Level == 0 && Row2 > FCellList.Count) Row2 = FCellList.Count; //we don't want this looping in the whole 65536 rows if Row2 is the full range.
            for (int Row = Row1; Row < Row2; Row++) //Row2 is not included.
            {
                if (!FCellList.HasRow(Row))
                {
                    if (Level == 0) continue;
                    FCellList.AddRow(Row);
                }

                if (ReplaceLowerLevels || FCellList[Row].RowRecord.KeepTogether < Level) FCellList[Row].RowRecord.KeepTogether = Level;

            }
        }

        internal int GetKeepRowsTogeher(int Row)
        {
            if (FCellList.HasRow(Row))
            {
                return FCellList[Row].RowRecord.KeepTogether;
            }
            return 0;
        }

        internal bool HasKeepRowsTogether()
        {
            for (int Row = 0; Row < FCellList.Count; Row++)
            {
                if (FCellList.HasRow(Row) && FCellList[Row].RowRecord.KeepTogether != 0) return true;
            }
            return false;
        }


        internal void AddCell(TCellRecord aRecord, int aRow, TVirtualReader VirtualReader)
        {
            FCellList.AddRecord(aRecord, aRow, VirtualReader);
        }

        internal void AddCell(TCellRecord aRecord, int aRow, bool IfExistsSetFormulaValue, TVirtualReader VirtualReader)
        {
            FCellList.AddRecord(aRecord, aRow, IfExistsSetFormulaValue, VirtualReader);
        }

        internal void AddMultipleCells(TMultipleValueRecord aRecord, int Row, TVirtualReader VirtualReader)
        {
            while (!aRecord.Eof())
            {
                TCellRecord OneRec = aRecord.ExtractOneRecord();
                FCellList.AddRecord(OneRec, Row, VirtualReader);
            }
        }

        internal TCellList CellList { get { return FCellList; } }

        internal int ColCount
        {
            get
            {
                int Result = 0;
                int aCount = FCellList.Count;
                for (int i = 0; i < aCount; i++)
                {
                    TCellRecordList Crl = FCellList[i];
                    if ((Crl != null) && (Crl.Count > 0))
                        if (Crl[Crl.Count - 1].Col + 1 > Result)
                            Result = Crl[Crl.Count - 1].Col + 1; //1 based result

                }
                return Result;
            }
        }

        internal int Count
        {
            get
            {
                return FCellList.Count;
            }
        }

        internal void MergeFromPxlCells(TCells Source)
        {
            for (int r = 0; r < Source.FCellList.Count; r++)
            {
                TCellAndRowRecordList Row = Source.FCellList[r];
                FCellList.Add(Row);
            }
        }
    }

    /// <summary>
    /// A list of range entries, parsed from RangeRecords. Records are TRangeEntry
    /// </summary>
    internal class TRangeList<T> where T: TRangeEntry
    {
        protected List<T> FList;

		internal int Modified;  //0 means never set, 1 modified, -1 the same as last time.

        internal TRangeList()
        {
            FList=new List<T>();
        }
            
        #region Generics
        internal void Add (T a)
        {
			Modified = 1;
            FList.Add(a);
        }

        internal void Insert(int index, T a)
        {
            Modified = 1;
            FList.Insert(index, a);
        }

        protected void SetThis(T value, int index)
        {
            FList[index]=value;
			Modified = 1;
        }

        internal T this[int index]
        {
            get { return FList[index]; }
            set { SetThis(value, index); }
        }

        internal int Count
        {
            get {return FList.Count;}
        }

        internal void Clear()
        {
			Modified = 1;
            FList.Clear();
        }
        #endregion

        internal void CopyFrom(TRangeList<T> aRangeList, TSheetInfo SheetInfo)
        {
            if (aRangeList.FList == FList) XlsMessages.ThrowException(XlsErr.ErrInternal); //Should be different objects
            Modified = 1;
            for (int i = 0; i < aRangeList.Count; i++)
                Add((T)TRangeEntry.Clone(aRangeList[i], SheetInfo));
        }


        internal void SaveToStream(IDataStream DataStream, TSaveData SaveData, TXlsCellRange CellRange)
        {
            if (CellRange == null) SaveAllToStream(DataStream, SaveData); else SaveRangeToStream(DataStream, SaveData, CellRange);
        }

        private void SaveAllToStream(IDataStream DataStream, TSaveData SaveData)
        {
            int aCount=Count;
            for (int i=0; i< aCount;i++)
                this[i].SaveToStream(DataStream, SaveData);
        }

        private void SaveRangeToStream(IDataStream DataStream, TSaveData SaveData, TXlsCellRange CellRange)
        {
            int aCount=Count;
            for (int i=0; i< aCount;i++)
                this[i].SaveRangeToStream(DataStream, SaveData, CellRange);
        }

        internal long TotalSize(TXlsCellRange CellRange)
        {
            if (CellRange == null) return TotalSizeAll();
            return TotalRangeSize(CellRange);
        }

        private long TotalSizeAll()
        {
            long Result = 0;
            int aCount = Count;
            for (int i = 0; i < aCount; i++)
                Result += this[i].TotalSize();
            return Result;
        }

        private long TotalRangeSize(TXlsCellRange CellRange)
        {
            long Result = 0;
            int aCount = Count;
            for (int i = 0; i < aCount; i++)
                Result += this[i].TotalRangeSize(CellRange);
            return Result;
        }

        internal void InsertAndCopyRange(TXlsCellRange SourceRange, int DestRow, int DestCol, int aRowCount, int aColCount, TRangeCopyMode CopyMode, TFlxInsertMode InsertMode, TSheetInfo SheetInfo)
        {
			Modified = 1;
			int aCount=Count;
            for (int i=0; i< aCount;i++)
                this[i].InsertAndCopyRange(SourceRange, DestRow, DestCol, aRowCount, aColCount, CopyMode, InsertMode, SheetInfo);

        }

		internal void DeleteRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
		{
			Modified = 1;
			int aCount=Count;
			for (int i=0; i< aCount;i++)
				this[i].DeleteRange(CellRange, aRowCount, aColCount, SheetInfo);
		}
		
		internal void MoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
		{
			Modified = 1;
			int aCount=Count;
			for (int i=0; i< aCount;i++)
				this[i].MoveRange(CellRange, NewRow, NewCol, SheetInfo);
		}

        internal void UpdateDeletedRanges(TDeletedRanges DeletedRanges)
        {
            int aCount = Count;
            for (int i = 0; i < aCount; i++)
                this[i].UpdateDeletedRanges(DeletedRanges);            
        }
    }

    internal class TChartCellList : TBaseRowColList<TCellRecord, TCellRecordList>
    {
        protected override TCellRecordList CreateRecord()
        {
            return new TCellRecordList(null, false);
        }

        internal void AddRecord(TCellRecord aRecord, int aRow)
        {
            for (int i = Count; i <= aRow; i++)
                FList.Add(CreateRecord());
            this[aRow].Add(aRecord);
        }

        internal void DeleteSeries(int index)
        {
            for (int i = Count - 1; i >= 0; i--)
            {
                TCellRecordList rl = this[i];
                int k = -1;
                if (rl.Find(index, ref k))
                {
                    rl.Delete(k);
                }

                for (int m = rl.Count - 1; m >= k; m--)
                {
                    rl[m].Col--;
                }
            }
        }
    }

}
