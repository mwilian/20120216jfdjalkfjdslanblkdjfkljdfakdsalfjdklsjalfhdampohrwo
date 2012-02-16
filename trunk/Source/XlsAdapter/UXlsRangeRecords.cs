using System;
using FlexCel.Core;
using System.Collections.Generic;

namespace FlexCel.XlsAdapter
{
    /// <summary>
    /// A simple Excel range we will use to store on a list.
    /// </summary>
    internal class TExcelRange: ICloneable
    {
        internal int R1;
        internal int R2;
        internal int C1;
        internal int C2;

		internal bool Reuse;  //must be initialized each time.

		internal TExcelRange()
		{
		}

		internal TExcelRange(int Row1, int Col1, int Row2, int Col2)
		{
			R1 = Row1;
			C1 = Col1;
			R2 = Row2;
			C2 = Col2;
		}

        internal static int Length{get {return 8;}}
        internal void LoadFromBiff8 (byte[] Data, bool Expand)
        {
            R1 = BitOps.GetWord(Data, 0);
            R2 = BitOps.GetWord(Data, 2);
            C1 = BitOps.GetWord(Data, 4);
            C2 = BitOps.GetWord(Data, 6);

            if (Expand)
            {
                if (R1 < FlxConsts.Max_Rows97_2003) R2 = Biff8Utils.ExpandBiff8Row(R2);
                if (C1 < FlxConsts.Max_Columns97_2003) C2 = Biff8Utils.ExpandBiff8Col(C2);
            }
        }

        internal byte[] Data(bool Expand)
        {
            int Rr1 = R1; int Rr2 = R2; int Cr1 = C1; int Cr2 = C2;
            if (Expand)
            {
                //here we really don't care if we lose CFs or DVs. If there are no cells in that range, we can safely save anyway.
                if (Rr1 > FlxConsts.Max_Rows97_2003) Rr1 = FlxConsts.Max_Rows97_2003;
                if (Rr2 > FlxConsts.Max_Rows97_2003) Rr2 = FlxConsts.Max_Rows97_2003;
                if (Cr1 > FlxConsts.Max_Columns97_2003) Cr1 = FlxConsts.Max_Columns97_2003;
                if (Cr2 > FlxConsts.Max_Columns97_2003) Cr2 = FlxConsts.Max_Columns97_2003;

            }

            Biff8Utils.CheckRow(Rr1);
            Biff8Utils.CheckRow(Rr2);
            Biff8Utils.CheckCol(Cr1);
            Biff8Utils.CheckCol(Cr2);
            
            byte[] Result = new byte[TExcelRange.Length];
            BitOps.SetWord(Result,0,Rr1);
            BitOps.SetWord(Result,2,Rr2);
            BitOps.SetWord(Result,4,Cr1);
            BitOps.SetWord(Result,6,Cr2);
            return Result;
        }

        internal bool IsOneCell
        {
            get
            {return (R1>=R2)&&(C1>=C2);}
        }

        #region ICloneable Members

        public object Clone()
        {
            return MemberwiseClone();
        }

        #endregion
    }

    /// <summary>
    /// A list of ranges. We use this to store merged cells or conditional formats.
    /// </summary>
    internal class TRangeValuesList   //Items are TExcelRange
    {
        private int FOtherDataLen;

        //Note that for merged cells we could go to 1026, but Excel 2007 doesn't save that either, so we will play safe. (Excel 2003 will save up to 1026 entries)
        //For CF while docs say 1026 (and that is the maximum we could fit in a register), office 2010 will crash with anything larger than half of it. And 513 is what Excel itself saves anyway.
        //For DV, docs say maximum is 432 but Excel doesn't seem to complain with more.
        internal int MaxRangesPerRecord;

        internal bool Allow1Cell=true;
        internal bool ExpandRanges;
        protected List<TExcelRange> FList;
        internal TRangeValuesList(int aMaxRangesPerRecord, int aOtherDataLen, bool aAllow1Cell, bool aExpandRanges)
        {
            MaxRangesPerRecord = aMaxRangesPerRecord;
            FOtherDataLen=aOtherDataLen;
            FList = new List<TExcelRange>();
            Allow1Cell=aAllow1Cell;
            ExpandRanges = aExpandRanges;
        }

		internal int OtherDataLen{get {return FOtherDataLen;}set{FOtherDataLen = value;}}

        #region Generics
        internal void AddForced (TExcelRange a)
        {
            FList.Add(a);
        }

		internal void AddAndMerge(TExcelRange a)
		{
			if (Allow1Cell) //this makes no sense in merged ranges.
			{
				for (int i = Count - 1; i >= 0; i--)
				{
					TExcelRange x = (TExcelRange)FList[i];
					if (x.C1 == a.C1 && x.C2 == a.C2)
					{
						if (x.R2 >= a.R1 && x.R1 <= a.R2)
						{
							x.R1 = Math.Min(x.R1, a.R1);
							x.R2 = Math.Max(x.R2, a.R2);
							return;
						}
					}
					else
					{
						if (x.R1 == a.R1 && x.R2 == a.R2)
						{
							if (x.C2 >= a.C1 && x.C1 <= a.C2)
							{
								x.C1 = Math.Min(x.C1, a.C1);
								x.C2 = Math.Max(x.C2, a.C2);
								return;
							}
						}
					}
				}
			}

			FList.Add(a);
		}


        internal void Insert (int index, TExcelRange a)
        {
            FList.Insert(index, a);
        }

        internal TExcelRange this[int index] 
        {
            get {return FList[index];} 
            set {FList[index] = value;}
        }

        internal int Count
        {
            get {return FList.Count;}
        }

        internal virtual void Clear()
        {
            FList.Clear();
        }

        internal void Delete(int Index)
        {
            FList.RemoveAt(Index);
        }
        #endregion


        internal void LoadFromBiff8(TxBaseRecord aRecord, int aPos)
        {
            int MyPos = aPos;
            TxBaseRecord MyRecord = aRecord;
            byte[]n=new byte[2];
            BitOps.ReadMem(ref MyRecord, ref MyPos, n);
            byte[] aData= new byte[TExcelRange.Length];
            for (int i=0; i< BitConverter.ToUInt16(n,0);i++)
            {
                TExcelRange ExcelRange= new TExcelRange();
                BitOps.ReadMem(ref MyRecord, ref MyPos, aData);
                ExcelRange.LoadFromBiff8(aData, ExpandRanges);
                AddForced(ExcelRange);
            }
        }

		#region Split into Continue Records
        //original methods used to split the record using Continue
        //but excel doesn't like continued range records, so we don't use them

        //these new methods are to split the record repeating it.  (That's why the "R" at the end)
        internal void SaveToStreamR(IDataStream DataStream, TSaveData SaveData, int Line)
        {
            int OneRecCount = MaxRangesPerRecord;
            UInt16 MyCount=0;
            if ((Line+1)*OneRecCount >Count) MyCount=(UInt16)(Count-Line*OneRecCount); else MyCount=(UInt16)OneRecCount;
            DataStream.Write16(MyCount);
            for (int i=Line*OneRecCount; i<Line*OneRecCount+MyCount; i++) DataStream.Write(this[i].Data(ExpandRanges),TExcelRange.Length);
        }

        private int NextInRange(TXlsCellRange Range, int k)
        {
            for (int i=k+1; i< Count;i++)
                if ((Range.Top<= this[i].R1) &&
                    (Range.Bottom>= this[i].R2) &&
                    (Range.Left<= this[i].C1) &&
                    (Range.Right>= this[i].C2)  )
                                                                                              
                    return i;
        
            return -1;
        }

        internal void SaveRangeToStreamR(IDataStream DataStream, TSaveData SaveData, int Line, int aCount, TXlsCellRange Range)
        {
            int OneRecCount = MaxRangesPerRecord;
            UInt16 MyCount=0;
            if ((Line+1)*OneRecCount >aCount) MyCount=(UInt16)(aCount-Line*OneRecCount); else MyCount=(UInt16)OneRecCount;
            DataStream.Write16(MyCount);
            int k=NextInRange(Range, -1);
            for (int i=Line*OneRecCount; i<Line*OneRecCount+MyCount;i++)
            {
                DataStream.Write(this[k].Data(ExpandRanges),TExcelRange.Length);
                k=NextInRange(Range, k);
            }
        }

        internal long TotalSizeR(int aCount)
        {
            return (XlsConsts.SizeOfTRecordHeader+ 2+ FOtherDataLen)* RepeatCountR(aCount)    //Base data
                + TExcelRange.Length*aCount;                               // Registers
        }

        internal int RepeatCountR(int aCount)
        {
            int OneRecCount = MaxRangesPerRecord;
            if (aCount>0) return ((aCount-1) / OneRecCount) +1; else return 1;
        }

        internal int RecordSizeR(int Line, int aCount)
        {
            int OneRecCount = MaxRangesPerRecord;
            int MyCount=0;
            if ((Line+1)*OneRecCount >aCount) MyCount=aCount-Line*OneRecCount; else MyCount=OneRecCount;
            return 2+ FOtherDataLen+MyCount*TExcelRange.Length;
        }

        internal int CountRangeRecords(TXlsCellRange Range)
        {
            int Result=0;
            for (int i=0; i< Count;i++)
                if (
                    (Range.Top<= this[i].R1 ) &&
                    (Range.Bottom>= this[i].R2 ) &&
                    (Range.Left<= this[i].C1 ) &&
                    (Range.Right>= this[i].C2 )) 

                    Result++;
            return Result;
        }

        #endregion

        internal void CopyFrom(TRangeValuesList aBaseRecordList)
        {
            if (aBaseRecordList.FList==FList) XlsMessages.ThrowException(XlsErr.ErrInternal); //Should be different objects

            if (aBaseRecordList!=null)
            {
                FList.Capacity+=aBaseRecordList.Count;
                for (int i=0;i<aBaseRecordList.Count;i++) 
                {
                    TExcelRange br=(TExcelRange)(aBaseRecordList[i].Clone());
                    FList.Add(br);
                }
            }
        }

        private static bool AddRows(TExcelRange R, TXlsCellRange CellRange, int aRowCount)
        {
            if (R.R1>= CellRange.Top) R.R1=BitOps.GetIncMaxMin(R.R1, aRowCount*CellRange.RowCount, FlxConsts.Max_Rows, CellRange.Top);
            if (R.R2>= CellRange.Top) R.R2=BitOps.GetIncMaxMin(R.R2, aRowCount*CellRange.RowCount, FlxConsts.Max_Rows, R.R1-1);
            return R.R2>=R.R1;
        }

		private void SplitRangesH(TExcelRange R, int BreakColumn)
		{
			TExcelRange NewRange=(TExcelRange)R.Clone();
			NewRange.C1=BreakColumn+1;
			R.C2=BreakColumn;
			AddForced(NewRange);
		}

		private void SplitRangesH2(TExcelRange R, int BreakColumn)
		{
			TExcelRange NewRange=(TExcelRange)R.Clone();
			NewRange.C2=BreakColumn-1;
			R.C1=BreakColumn;
			AddForced(NewRange);
		}

        private static bool AddCols(TExcelRange R, TXlsCellRange CellRange, int aColCount)
        {
            if (R.C1>= CellRange.Left) R.C1=BitOps.GetIncMaxMin(R.C1, aColCount*CellRange.ColCount, FlxConsts.Max_Columns, CellRange.Left);
            if (R.C2>= CellRange.Left) R.C2=BitOps.GetIncMaxMin(R.C2, aColCount*CellRange.ColCount, FlxConsts.Max_Columns, R.C1-1);
            return R.C2>=R.C1;
        }

		private void SplitRangesV(TExcelRange R, int BreakRow)
		{
			TExcelRange NewRange=(TExcelRange)R.Clone();
			NewRange.R1=BreakRow+1;
			R.R2=BreakRow;
			AddForced(NewRange);
		}
		
		private void SplitRangesV2(TExcelRange R, int BreakRow)
		{
			TExcelRange NewRange=(TExcelRange)R.Clone();
			NewRange.R2=BreakRow-1;
			R.R1=BreakRow;
			AddForced(NewRange);
		}

        private void CheckOneCell(int i)
        {
            if (this[i].IsOneCell) Delete(i); 
        }

		private void IsolateBlock(int i, TXlsCellRange CellRange)
		{
			this[i].Reuse = true;
			if (this[i].C2 > CellRange.Right) 
			{
				SplitRangesH(this[i], CellRange.Right);
				if (!Allow1Cell) {CheckOneCell(Count-1);}
			}

			if (this[i].R2 > CellRange.Bottom) 
			{
				SplitRangesV(this[i], CellRange.Bottom);
				if (!Allow1Cell) {CheckOneCell(Count-1);}
			}

			if (this[i].C1 < CellRange.Left) 
			{
				SplitRangesH2(this[i], CellRange.Left);
				if (!Allow1Cell) {CheckOneCell(Count-1);}
			}

			if (this[i].R1 < CellRange.Top) 
			{
				SplitRangesV2(this[i], CellRange.Top);
				if (!Allow1Cell) {CheckOneCell(Count-1);}
			}

		}

		private void SplitBlocks(TXlsCellRange CellRange, int NewRow, int NewCol, bool FirstPass)
		{
			for (int i=Count -1; i>=0;i--)
			{
				if (FirstPass) 
				{
					this[i].Reuse = false;
				}
				else
				{
					if (this[i].Reuse) continue; //avoid deleting ranges that were moved.
				}
                
				if (this[i].R2<CellRange.Top || this[i].R1 > CellRange.Bottom || this[i].C2 < CellRange.Left || this[i].C1 > CellRange.Right) continue;
				IsolateBlock(i, CellRange);

				if (FirstPass)
				{
					//should never overflow.
					this[i].R1+= NewRow - CellRange.Top;
					this[i].R2+= NewRow - CellRange.Top;
					this[i].C1+= NewCol - CellRange.Left;
					this[i].C2+= NewCol - CellRange.Left;
				}
				else
				{
					Delete(i);
					continue;
				}

				if (!Allow1Cell) {CheckOneCell(i);}   			
			}
		}

		internal void ArrangeMoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
		{
			if (SheetInfo.SourceFormulaSheet != SheetInfo.InsSheet) return;

			SplitBlocks(CellRange, NewRow, NewCol, true);
			SplitBlocks(CellRange.Offset(NewRow, NewCol), 0, 0, false);
		}

        internal void ArrangeInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
        {
            for (int i=Count -1; i>=0;i--)
            {
                if (aRowCount!=0)
                {
                    if (this[i].R2<CellRange.Top) continue;
                    if (CellRange.HasCol(this[i].C1) && CellRange.HasCol(this[i].C2))
                    {
                        if (!AddRows(this[i], CellRange, aRowCount))
                            Delete(i);
                        else
                            if (!Allow1Cell) {CheckOneCell(i);}   
                    }
                    else
                        if (CellRange.HasCol(this[i].C1))  //Range has been split in 2. C1 is in, C2 is out
                    {
                        SplitRangesH(this[i], CellRange.Right);
                        if (!Allow1Cell) CheckOneCell(Count-1);
                        if (!AddRows(this[i], CellRange, aRowCount))
                            Delete(i);
                        else
                            if (!Allow1Cell) CheckOneCell(i);
                    }
                    else
                        if (CellRange.HasCol(this[i].C2))  //Range has been split in 2. C1 is out, C2 is in.
                    {
                        SplitRangesH(this[i], CellRange.Left-1);
                        if (!Allow1Cell) CheckOneCell(i);
                        if (!AddRows(this[Count-1], CellRange, aRowCount))
                            Delete(Count-1);
                        else
                            if (!Allow1Cell) {CheckOneCell(Count-1);}
                    }
                    else
                        if ((this[i].C1<CellRange.Left)&& (this[i].C2>CellRange.Right)) //Range has been split in 3. C1 is out, C2 is out, but some cells on the middle are in.
                    {
                        SplitRangesH(this[i], CellRange.Right);
                        if (!Allow1Cell) {CheckOneCell(Count-1);}
                        SplitRangesH(this[i], CellRange.Left-1);
                        if (!Allow1Cell) CheckOneCell(i);
                        if (!AddRows(this[Count-1], CellRange, aRowCount))
                            Delete(Count-1);
                        else
                            if (!Allow1Cell) {CheckOneCell(Count-1);}
                    }
                }

                /////////////// Vertical inserts
                if (aColCount!=0)
                {
                    if (i<Count) //We might have deleted one cell.
                    {
                        if (this[i].C2<CellRange.Left) continue;
                        if (CellRange.HasRow(this[i].R1) && CellRange.HasRow(this[i].R2))
                        {
                            if (!AddCols(this[i], CellRange, aColCount))
                                Delete(i);
                            else
                                if (!Allow1Cell) {CheckOneCell(i);}
                        }
                        else
                            if (CellRange.HasRow(this[i].R1))  //Range has been split in 2. R1 is in, R2 is out
                        {
                            SplitRangesV(this[i], CellRange.Bottom);
                            if (!Allow1Cell) {CheckOneCell(Count-1);}
                            if (!AddCols(this[i], CellRange, aColCount))
                                Delete(i);
                            else
                                if (!Allow1Cell) {CheckOneCell(i);}
                        }
                        else
                            if (CellRange.HasRow(this[i].R2))  //Range has been split in 2. R1 is out, R2 is in.
                        {
                            SplitRangesV(this[i], CellRange.Top-1);
                            if (!Allow1Cell) {CheckOneCell(i);}
                            if (!AddCols(this[Count-1], CellRange, aColCount))
                                Delete(Count-1);
                            else
                                if (!Allow1Cell) {CheckOneCell(Count-1);}
                        }
                        else
                            if ((this[i].R1<CellRange.Top)&& (this[i].R2>CellRange.Bottom)) //Range has been split in 3. R1 is out, R2 is out, but some cells on the middle are in.
                        {
                            SplitRangesV(this[i], CellRange.Bottom);
                            if (!Allow1Cell) {CheckOneCell(Count-1);}
                            SplitRangesV(this[i], CellRange.Top-1);
                            if (!Allow1Cell) {CheckOneCell(i);}
                            if (!AddCols(this[Count-1], CellRange, aColCount))
                                Delete(Count-1);
                            else
                                if (!Allow1Cell) {CheckOneCell(Count-1);}
                        }
                    }
                }
            }
        }

		internal void ClearRange(TXlsCellRange CellRange)
		{
			for (int i=Count -1; i>=0;i--)
			{
				if (this[i].R2<CellRange.Top || this[i].R1 > CellRange.Bottom || this[i].C2 < CellRange.Left || this[i].C1 > CellRange.Right) continue;
				IsolateBlock(i, CellRange);

				Delete(i);
			}
		}


        #region Copying
        private void CopyOneRange(TExcelRange R, TXlsCellRange CellRange, int DestRow, int DestCol, int zr, int zc)
        {
            TExcelRange NewRange= (TExcelRange) R.Clone();
                       
            BitOps.IncWord(ref NewRange.C1, DestCol - CellRange.Left + CellRange.ColCount*zc, FlxConsts.Max_Columns, XlsErr.ErrTooManyColumns);
            BitOps.IncWord(ref NewRange.C2, DestCol - CellRange.Left + CellRange.ColCount*zc, FlxConsts.Max_Columns, XlsErr.ErrTooManyColumns);
            BitOps.IncWord(ref NewRange.R1, DestRow - CellRange.Top +CellRange.RowCount*zr, FlxConsts.Max_Rows, XlsErr.ErrTooManyRows);
            BitOps.IncWord(ref NewRange.R2, DestRow - CellRange.Top +CellRange.RowCount*zr, FlxConsts.Max_Rows, XlsErr.ErrTooManyRows);
                                             
            AddAndMerge(NewRange);
        }

        private void CopyIntersectRange(TExcelRange R, TXlsCellRange NewCellRange, int DestRow, int DestCol, int aRowCount, int aColCount, ref int MinR1, ref int MaxR2, ref int MinC1, ref int MaxC2)
        {
            if (aRowCount>0)
            {
                int Lc=NewCellRange.RowCount* aRowCount;
                if ((R.R1<=NewCellRange.Top) && (R.R2>=NewCellRange.Bottom)) // Just copy one big range
                {
                    TExcelRange NewRange=(TExcelRange)R.Clone(); //Automatically clones R.UseCols
                    NewRange.R1=DestRow;
                    NewRange.R2=DestRow+Lc-1;
                    NewRange.C1=Math.Max(DestCol+NewRange.C1-NewCellRange.Left, DestCol);
                    NewRange.C2=Math.Min(DestCol+NewRange.C2-NewCellRange.Left, DestCol+NewCellRange.ColCount-1);
                    AddAndMerge(NewRange);
                    if (NewRange.C1< MinC1) MinC1=NewRange.C1;
                    if (NewRange.C2> MaxC2) MaxC2=NewRange.C2;
                    if (NewRange.R1< MinR1) MinR1=NewRange.R1;
                    if (NewRange.R2> MaxR2) MaxR2=NewRange.R2;
                }        
                else // We have to copy one small range for each aCount
                {
                    for (int k=0; k< aRowCount ;k++)
                    {
                        TExcelRange NewRange=(TExcelRange)R.Clone();
                        NewRange.R1=DestRow+NewCellRange.RowCount*k;
                        if (R.R1>NewCellRange.Top) NewRange.R1+= R.R1-NewCellRange.Top;
                        NewRange.R2=DestRow+NewCellRange.RowCount*(k+1)-1;
                        if (R.R2<NewCellRange.Bottom) NewRange.R2-= NewCellRange.Bottom-R.R2;
                        NewRange.C1=Math.Max(DestCol+NewRange.C1-NewCellRange.Left, DestCol);
                        NewRange.C2=Math.Min(DestCol+NewRange.C2-NewCellRange.Left, DestCol+NewCellRange.ColCount-1);

                        AddAndMerge(NewRange);
                        if (NewRange.C1< MinC1) MinC1=NewRange.C1;
                        if (NewRange.C2> MaxC2) MaxC2=NewRange.C2;
                        if (NewRange.R1< MinR1) MinR1=NewRange.R1;
                        if (NewRange.R2> MaxR2) MaxR2=NewRange.R2;
                    }        
                }
            }

            if (aColCount>0)
            {
                int Lc=NewCellRange.ColCount* aColCount;
                if ((R.C1<=NewCellRange.Left) && (R.C2>=NewCellRange.Right)) // Just copy one big range
                {
                    TExcelRange NewRange=(TExcelRange)R.Clone(); //Automatically clones R.UseCols
                    NewRange.C1=DestCol;
                    NewRange.C2=DestCol+Lc-1;
                    NewRange.R1=Math.Max(DestRow+NewRange.R1-NewCellRange.Top, DestRow);
                    NewRange.R2=Math.Min(DestRow+NewRange.R2-NewCellRange.Top, DestRow+NewCellRange.RowCount-1);

                    AddAndMerge(NewRange);
                    if (NewRange.C1< MinC1) MinC1=NewRange.C1;
                    if (NewRange.C2> MaxC2) MaxC2=NewRange.C2;
                    if (NewRange.R1< MinR1) MinR1=NewRange.R1;
                    if (NewRange.R2> MaxR2) MaxR2=NewRange.R2;
                }        
                else // We have to copy one small range for each aCount
                {
                    for (int k=0; k< aColCount ;k++)
                    {
                        TExcelRange NewRange=(TExcelRange)R.Clone();
                        NewRange.C1=DestCol+NewCellRange.ColCount*k;
                        if (R.C1>NewCellRange.Left) NewRange.C1+= R.C1-NewCellRange.Left;
                        NewRange.C2=DestCol+NewCellRange.ColCount*(k+1)-1;
                        if (R.C2<NewCellRange.Right) NewRange.C2-= NewCellRange.Right-R.C2;
                        NewRange.R1=Math.Max(DestRow+NewRange.R1-NewCellRange.Top, DestRow);
                        NewRange.R2=Math.Min(DestRow+NewRange.R2-NewCellRange.Top, DestRow+NewCellRange.RowCount-1);

                        AddAndMerge(NewRange);
                        if (NewRange.C1< MinC1) MinC1=NewRange.C1;
                        if (NewRange.C2> MaxC2) MaxC2=NewRange.C2;
                        if (NewRange.R1< MinR1) MinR1=NewRange.R1;
                        if (NewRange.R2> MaxR2) MaxR2=NewRange.R2;
                    }        
                }
            }
        }

        //Formats are copied if the range intersects with the original. (Merged cells need all the range to be inside the original)
        internal void CopyRangeInclusive(TXlsCellRange SourceRange, int DestRow, int DestCol, int aRowCount, int aColCount, ref int MinR1, ref int MaxR2, ref int MinC1, ref int MaxC2)
        {
            if (aRowCount>0)  //We are inserting rows
            {
                int Lc= SourceRange.RowCount * aRowCount;

                //Adapt for the already inserted range. If the source block is below the destblock, it has already been shifted lc rows down.
                int NewFirstRow=SourceRange.Top; int NewLastRow=SourceRange.Bottom;
                if (SourceRange.Top>=DestRow) NewFirstRow+=Lc;
                if (SourceRange.Bottom>=DestRow) NewLastRow+=Lc;

				int aCount = Count; //Count might change, so this must be fixed!
                for (int i=0; i< aCount;i++)
                {
                    TExcelRange R=this[i];
                    if ((R.R1<= NewLastRow) &&
                        (R.R2>= NewFirstRow) &&
                        (R.C1<= SourceRange.Right) &&
                        (R.C2>= SourceRange.Left)
                        ) //Leave out blocks that do not intersect SourceRange.
                    {
                        //First Case, Block copied is above the original

                        if (SourceRange.Top>=DestRow) 
                            if ((R.R1<DestRow + Lc) && (SourceRange.Left==DestCol)) {} //nothing, range is automatically expanded. Remember copy range has already been shifted down., that's why we add Lc. In fact, there can be no cells on [DestRow...DestRow+Lc] since this is an inserted range.
                            else if ((R.R1==DestRow + Lc) && ( R.R2 >=NewLastRow)&& (SourceRange.Left==DestCol)) //expand the range to include inserted rows
                            {
                                R.R1-= Lc;
                                if (R.R1< MinR1) MinR1=R.R1;
                            }
                            else CopyIntersectRange(R, new TXlsCellRange(NewFirstRow, SourceRange.Left, NewLastRow, SourceRange.Right), DestRow, DestCol, aRowCount, aColCount, ref MinR1, ref MaxR2, ref MinC1, ref MaxC2); //We have to Copy the intersecting range, and clip the results

                            //Second Case, Block copied is below the original
                        else
                            if ((R.R2>DestRow-1)&& (SourceRange.Left==DestCol)) {} //nothing, range is automatically expanded
                        else if ((R.R2==DestRow -1) && (R.R1<=NewFirstRow) && (SourceRange.Left==DestCol)) //expand the range to include inserted rows
                        {
                            R.R2+=Lc;
                            if (R.R2> MaxR2) MaxR2=R.R2;
                        }
                        else CopyIntersectRange(R, new TXlsCellRange(NewFirstRow, SourceRange.Left, NewLastRow, SourceRange.Right), DestRow, DestCol, aRowCount, aColCount, ref MinR1, ref MaxR2, ref MinC1, ref MaxC2); //We have to Copy the intersecting range, and clip the results

                    }
                }
            }

            if (aColCount>0)  //We are inserting columns
            {
                int Lc= SourceRange.ColCount * aColCount;

                //Adapt for the already inserted range. If the source block is below the destblock, it has already been shifted lc rows down.
                int NewFirstCol=SourceRange.Left; int NewLastCol=SourceRange.Right;
                if (SourceRange.Left>=DestCol) NewFirstCol+=Lc;
                if (SourceRange.Right>=DestCol) NewLastCol+=Lc;

                int FixedCount=Count;
                for (int i=0; i< FixedCount;i++)
                {
                    TExcelRange R=this[i];
                    if ((R.C1<= NewLastCol) &&
                        (R.C2>= NewFirstCol) &&
                        (R.R1<= SourceRange.Bottom) &&
                        (R.R2>= SourceRange.Top)
                        ) //Leave out blocks that do not intersect SourceRange.
                    {
                        //First Case, Block copied is at the left from the original

                        if (SourceRange.Left>=DestCol) 
                            if ((R.C1<DestCol + Lc) && (SourceRange.Top==DestRow)) {} //nothing, range is automatically expanded. Remember copy range has already been shifted down., that's why we add Lc. In fact, there can be no cells on [DestRow...DestRow+Lc] since this is an inserted range.
                            else if ((R.C1==DestCol + Lc) && ( R.C2 >=NewLastCol)&& (SourceRange.Top==DestRow)) //expand the range to include inserted rows
                            {
                                R.C1-= Lc;
                                if (R.C1< MinC1) MinC1=R.C1;
                            }
                            else CopyIntersectRange(R, new TXlsCellRange(SourceRange.Top, NewFirstCol, SourceRange.Bottom, NewLastCol), DestRow, DestCol, aRowCount, aColCount, ref MinR1, ref MaxR2, ref MinC1, ref MaxC2); //We have to Copy the intersecting range, and clip the results

                            //Second Case, Block copied is at the right from the original
                        else
                            if ((R.C2>DestCol-1)&& (SourceRange.Top==DestRow)) {} //nothing, range is automatically expanded
                        else if ((R.C2==DestCol -1) && (R.C1<=NewFirstCol) && (SourceRange.Top==DestRow)) //expand the range to include inserted rows
                        {
                            R.C2+=Lc;
                            if (R.C2> MaxC2) MaxC2=R.C2;
                        }
                        else CopyIntersectRange(R, new TXlsCellRange(SourceRange.Top, NewFirstCol, SourceRange.Bottom, NewLastCol), DestRow, DestCol, aRowCount, aColCount, ref MinR1, ref MaxR2, ref MinC1, ref MaxC2); //We have to Copy the intersecting range, and clip the results
                    }
                }
            }

        }



        /// <summary>
        /// Used on merged cells. we won't copy partially contained merged cells, as we would with conditional formats.
        /// </summary>
        internal void CopyRangeExclusive(TXlsCellRange SourceRange, int DestRow, int DestCol, int aRowCount, int aColCount)
        {
            int Lc= SourceRange.RowCount * aRowCount;
            int Cc= SourceRange.ColCount * aColCount;

            //Adapt for the already inserted range. If the source block is below the destblock, it has already been shifted lc rows down.
            int NewFirstRow=SourceRange.Top; int NewLastRow=SourceRange.Bottom;
            if (SourceRange.Top>=DestRow) NewFirstRow+=Lc;
            if (SourceRange.Bottom>=DestRow) NewLastRow+=Lc;

            int NewFirstCol=SourceRange.Left; int NewLastCol=SourceRange.Right;
            if (SourceRange.Left>=DestCol) NewFirstCol+=Cc;
            if (SourceRange.Right>=DestCol) NewLastCol+=Cc;

            for (int i=0; i< Count;i++)
            {
                TExcelRange R=this[i];
                if ((R.R1>= NewFirstRow) &&
                    (R.R2<= NewLastRow) &&
                    (R.C1>=NewFirstCol) &&
                    (R.C2<=NewLastCol)  
                    ) //Range has to be COMPLETELY in to be copied.
                                              
                {
                    for (int k=0; k< aRowCount;k++)
                    {
                        int zr=k;
                        if (SourceRange.Top>=DestRow) zr=-(k+1);
                        CopyOneRange(R, SourceRange, DestRow, DestCol, zr, 0);
                    }
                 
                    for (int k=0; k< aColCount;k++)
                    {
                        int zc=k;
                        if (SourceRange.Left>=DestCol) zc=-(k+1);
                        CopyOneRange(R, SourceRange, DestRow, DestCol, 0, zc);
                    }
                }
            }
        }
    
        private struct TRect
        {
            internal int Left;
            internal int Top;
            internal int Right;
            internal int Bottom;
        }

        internal void PreAddNewRange(ref int R1, ref int C1, ref int R2, ref int C2)
        {
            //Check ranges are valid
            if ((R1<0) || (R2<R1) || (R2>FlxConsts.Max_Rows) ||
                (C1<0) || (C2<C1) || (C2>FlxConsts.Max_Columns)) return;

            if ((R1==R2)&&(C1==C2)) return;

            for (int i=Count-1;i>=0;i--)
            {
                TExcelRange R=this[i];
                TRect OutRect;
                OutRect.Left=Math.Max(R.C1, C1);
                OutRect.Top=Math.Max(R.R1, R1);
                OutRect.Right=Math.Min(R.C2, C2);
                OutRect.Bottom=Math.Min(R.R2, R2);
                if ((OutRect.Left<=OutRect.Right)&&(OutRect.Top<=OutRect.Bottom))//found
                {
                    R1=Math.Min(R.R1, R1);
                    R2=Math.Max(R.R2, R2);
                    C1=Math.Min(R.C1, C1);
                    C2=Math.Max(R.C2, C2);
                    FList.RemoveAt(i);
                }
            }
        }

        /// <summary>
        /// We always have to call PreAddNewRange to verify it doesn't exist
        /// </summary>
        internal void AddNewRange(int FirstRow, int FirstCol, int LastRow, int LastCol)
        {
            //Check ranges are valid
            if ((FirstRow<0) || (LastRow<FirstRow) || (LastRow>FlxConsts.Max_Rows) ||
                (FirstCol<0) || (LastCol<FirstCol) || (LastCol>FlxConsts.Max_Columns)) return;

            if ((FirstRow==LastRow)&&(FirstCol==LastCol)) return;

            TExcelRange NewRange= new TExcelRange();
            NewRange.R1=FirstRow;
            NewRange.R2=LastRow;
            NewRange.C1=FirstCol;
            NewRange.C2=LastCol;
            AddForced(NewRange);
        }

        #endregion
    }

    /// <summary>
    /// Base for merged cells and conditional formats. In abstract, to any kind of records that save list of ranges.
    /// </summary>
    internal abstract class TRangeEntry 
    {
        protected TRangeValuesList RangeValuesList;
                      
        internal TRangeEntry()
        {
            RangeValuesList=null; //It will be initialized by its children.
        }
                      
        protected virtual TRangeEntry DoCopyTo(TSheetInfo SheetInfo)
        {
            TRangeEntry b = (TRangeEntry)MemberwiseClone();
            b.RangeValuesList = new TRangeValuesList(RangeValuesList.MaxRangesPerRecord, RangeValuesList.OtherDataLen, RangeValuesList.Allow1Cell, RangeValuesList.ExpandRanges);
            b.RangeValuesList.CopyFrom(RangeValuesList);
            return b;
        }
        
        internal static TRangeEntry Clone(TRangeEntry Self, TSheetInfo SheetInfo) //this should be non-virtual. It allows you to obtain a clone, even if the object is null
        {
            if (Self==null) return null;   //for this to work, this can't be a virtual method
            else return Self.DoCopyTo(SheetInfo);
        }

        internal abstract void LoadFromStream(TBaseRecordLoader RecordLoader, TRangeRecord First);
        internal abstract void SaveToStream(IDataStream DataStream, TSaveData SaveData);
        internal abstract void SaveRangeToStream(IDataStream DataStream, TSaveData SaveData, TXlsCellRange CellRange);
        internal abstract long TotalSize();
        internal abstract long TotalRangeSize(TXlsCellRange CellRange);

        internal virtual void ArrangeInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
        {
            RangeValuesList.ArrangeInsertRange(CellRange, aRowCount, aColCount, SheetInfo);
        }

		internal virtual void ArrangeMoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
		{
			RangeValuesList.ArrangeMoveRange(CellRange, NewRow, NewCol, SheetInfo);
		}


        internal virtual void InsertAndCopyRange(TXlsCellRange SourceRange, int DestRow, int DestCol, int aRowCount, int aColCount, TRangeCopyMode CopyMode, TFlxInsertMode InsertMode, TSheetInfo SheetInfo)
		{
            if (InsertMode != TFlxInsertMode.NoneDown && InsertMode != TFlxInsertMode.NoneRight)
            {
                ArrangeInsertRange(SourceRange.OffsetForIns(DestRow, DestCol, InsertMode), aRowCount, aColCount, SheetInfo);
            }
		}

		internal virtual void DeleteRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
		{
			ArrangeInsertRange(CellRange, -aRowCount, -aColCount, SheetInfo);
		}
		
		internal virtual void MoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
		{
			ArrangeMoveRange(CellRange, NewRow, NewCol, SheetInfo);
		}

		internal void ClearRange(TXlsCellRange CellRange)
		{
			RangeValuesList.ClearRange(CellRange);
		}

        internal virtual void UpdateDeletedRanges(TDeletedRanges DeletedRanges)
        {
        }
	}


	#region Merged Cells
    /// <summary>
    /// Merged Cell list. Merged cells can't be continued. We have to write independent records.
    /// </summary>
    internal class TMergedCells: TRangeEntry
    {
        internal TMergedCells()
        {
            RangeValuesList= new TRangeValuesList(513, 0, false, false);
        }

        internal void Clear()
        {
            if (RangeValuesList!=null) RangeValuesList.Clear();
        }
        internal override void LoadFromStream(TBaseRecordLoader RecordLoader, TRangeRecord First)
        {
            Clear();
            RangeValuesList.LoadFromBiff8(First, 0);
        }

        internal override void SaveToStream(IDataStream DataStream, TSaveData SaveData)
        {
            if (RangeValuesList.Count==0) return; //don't save empty MergedCells
            int aCount = RangeValuesList.RepeatCountR(RangeValuesList.Count);
            for (int i=0; i< aCount;i++)
            {
                DataStream.WriteHeader((UInt16)xlr.CELLMERGING, (UInt16) RangeValuesList.RecordSizeR(i, RangeValuesList.Count));
                RangeValuesList.SaveToStreamR(DataStream, SaveData,  i);
            }
        }

        internal override void SaveRangeToStream(IDataStream DataStream, TSaveData SaveData, TXlsCellRange CellRange)
        {
            int Rc=RangeValuesList.CountRangeRecords(CellRange);
            if (Rc==0) return; //don't save empty MergedCells
            int aCount = RangeValuesList.RepeatCountR(RangeValuesList.Count);
            for (int i=0; i< aCount;i++)
            {
                DataStream.WriteHeader((UInt16)xlr.CELLMERGING, (UInt16)RangeValuesList.RecordSizeR(i,Rc) );
                RangeValuesList.SaveRangeToStreamR(DataStream, SaveData,  i, Rc, CellRange);
            }
        }

        internal override long TotalSize()
        {
            if (RangeValuesList.Count==0) return 0; else return RangeValuesList.TotalSizeR(RangeValuesList.Count);
        }

        internal override long TotalRangeSize(TXlsCellRange CellRange)
        {
            if (RangeValuesList.Count==0) return 0; else return RangeValuesList.TotalSizeR(RangeValuesList.CountRangeRecords(CellRange));
        }

        internal override void InsertAndCopyRange(TXlsCellRange SourceRange, int DestRow, int DestCol, int aRowCount, int aColCount, TRangeCopyMode CopyMode, TFlxInsertMode InsertMode, TSheetInfo SheetInfo)
        {
            base.InsertAndCopyRange(SourceRange, DestRow, DestCol, aRowCount, aColCount, CopyMode, InsertMode, SheetInfo);
            if (CopyMode!=TRangeCopyMode.None) RangeValuesList.CopyRangeExclusive(SourceRange, DestRow, DestCol, aRowCount, aColCount);
        }
        internal override void DeleteRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
        {
            base.DeleteRange(CellRange, aRowCount, aColCount, SheetInfo);
        }

		internal override void MoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
		{
			base.MoveRange (CellRange, NewRow, NewCol, SheetInfo);
		}

		internal override void ArrangeMoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
		{
			base.ArrangeMoveRange (CellRange, NewRow, NewCol, SheetInfo);
		}



        internal bool CheckCell(int aRow, int aCol, TXlsCellRange CellBounds)
        {
            int aCount=RangeValuesList.Count;
            for (int i=0; i< aCount;i++)
            {
                TExcelRange Cr= RangeValuesList[i];
                if ((Cr.R1<=aRow) &&
                    (Cr.R2>=aRow) &&
                    (Cr.C1<=aCol) &&
                    (Cr.C2>=aCol) )
                {
                    CellBounds.Left= Cr.C1;
                    CellBounds.Top= Cr.R1;
                    CellBounds.Right= Cr.C2;
                    CellBounds.Bottom= Cr.R2;
                    return true;
                }
            }
            return false;
        }

        internal int MergedCount()
        {
            return RangeValuesList.Count;
        }

        internal TXlsCellRange MergedCell(int i)
        {
            TXlsCellRange Result=new TXlsCellRange(RangeValuesList[i].R1,RangeValuesList[i].C1,RangeValuesList[i].R2, RangeValuesList[i].C2);
            return Result; 
        }

        internal TRangeValuesList GetRangeValueList()
        {
            return RangeValuesList;
        }

        internal void PreMerge(ref int R1, ref int C1, ref int R2, ref int C2)
        {
            RangeValuesList.PreAddNewRange(ref R1, ref C1, ref R2, ref C2);
        }

        /// <summary>
        /// Always call premergecell first...
        /// </summary>
        internal void MergeCells(int FirstRow, int FirstCol, int LastRow, int LastCol)
        {
            RangeValuesList.AddNewRange(FirstRow, FirstCol, LastRow, LastCol);
        }
        
        internal void UnMergeCells(int FirstRow, int FirstCol, int LastRow, int LastCol)
        {
            for (int i=RangeValuesList.Count-1; i>= 0;i--)
            {
                if ((RangeValuesList[i].R1==FirstRow) &&
                    (RangeValuesList[i].R2==LastRow) &&
                    (RangeValuesList[i].C1==FirstCol) &&
                    (RangeValuesList[i].C2==LastCol) )
                {
                    RangeValuesList.Delete(i);
                }
            }
        }
    }
	#endregion

	#region Data Validation

	internal class TDataValidationList   //could be a TRangeList, but has a headerdata too. Just to not add complexity to TRangeList, we will handle it separately.
	{
		private List<TDataValidation> FList;
		private byte[] HeaderData;

		internal TDataValidationList()
		{
			FList = new List<TDataValidation>();
		}

		internal void LoadDVal(TDValRecord r)
		{
            HeaderData = new byte[14];
			Array.Copy(r.Data, 0, HeaderData, 0, HeaderData.Length);
		}

		internal int Add(TDataValidation a)
		{
			FList.Add(a);
            return FList.Count - 1;
		}

		internal TDataValidation this[int index]
		{
			get
			{
				return FList[index];
			}
		}

		internal void ClearObjId()
		{
			if (HeaderData == null) return;
			BitOps.SetCardinal(HeaderData, HeaderData.Length - 4, 0xFFFFFFFF);
		}

		internal void CopyFrom(TDataValidationList aBaseRecordList, TDrawing aDrawing, TSheetInfo SheetInfo)
		{
			if (aBaseRecordList.FList==FList) XlsMessages.ThrowException(XlsErr.ErrInternal); //Should be different objects

			HeaderData = null;
			if (aBaseRecordList!=null)
			{
				if (aBaseRecordList.HeaderData != null) 
				{
					HeaderData = new byte[aBaseRecordList.HeaderData.Length];
					Array.Copy(aBaseRecordList.HeaderData, 0, HeaderData, 0, aBaseRecordList.HeaderData.Length);
					//If the dropdown is hidden, objid will be 0xffffffff and this will be copied.
					//If the dropdown is shown, we need to make sure we new obid corresponds with the new combo in the new list.
					
					/*uint ObjId = 0xFFFFFFFF; //aDrawing.AddNewDv().ObjId;
					byte[] ObjIdb = BitConverter.GetBytes(ObjId);
					Array.Copy(ObjIdb, 0, HeaderData, HeaderData.Length - 4, ObjIdb.Length);*/
				}

				FList.Capacity+=aBaseRecordList.FList.Count;
                for (int i = 0; i < aBaseRecordList.FList.Count; i++)
                {
                    TDataValidation br = (TDataValidation)(TDataValidation.Clone(aBaseRecordList[i], SheetInfo));
                    FList.Add(br);
                }
			}
		}


        internal long TotalSize(TXlsCellRange CellRange)
        {
            if (CellRange == null) return TotalSizeAll();
            return TotalRangeSize(CellRange);
        }

		private long TotalSizeAll()
		{
			ShrinkList();
			long Result = 0;
			foreach (TDataValidation dv in FList)
			{
				Result += dv.TotalSize();
			}
			if (Result == 0) return 0;
			return Result + HeaderData.Length + 4 + XlsConsts.SizeOfTRecordHeader;
		}

		private long TotalRangeSize(TXlsCellRange CellRange)
		{
			ShrinkList();
			long Result = 0;
			foreach (TDataValidation dv in FList)
			{
				Result += dv.TotalRangeSize(CellRange);
			}

			if (Result == 0) return 0;
			return Result + HeaderData.Length + 4 + XlsConsts.SizeOfTRecordHeader;
		}

		private void ShrinkList()
		{
			for (int i = FList.Count -1; i >= 0; i--)
			{
				if (this[i].IsEmpty()) FList.RemoveAt(i);
			}
		}

		private int GetCount()
		{
			return FList.Count; //it has been shrinked.
		}

		private int GetCount(TXlsCellRange CellRange)
		{
			int Result = 0;
			for (int i = 0; i < FList.Count; i++)
			{
				if (!this[i].IsEmpty(CellRange)) Result++;
			}

			return Result;
		}

        internal int GetXYWindow(int posi)
        {
                if (HeaderData == null) return 0;
                unchecked
                {
                    return (int) BitOps.GetCardinal(HeaderData, posi);
                }
        }

        internal void SetXYWindow(int posi, int value)
        {
            if (HeaderData == null) HeaderData = CreateNewHeaderData();
            if (value >= 0 && value <= 65535)
            {
                BitOps.SetCardinal(HeaderData, posi, value);
            }
            else FlxMessages.ThrowException(FlxErr.ErrInvalidValue, "xyWindow", value, 0, 65535);
        }

        internal int xWindow { get { return GetXYWindow(2); } set { SetXYWindow(2, value); } }
        internal int yWindow { get { return GetXYWindow(6); } set { SetXYWindow(6, value); } }

        internal bool DisablePrompts
        {
            get
            {
                if (HeaderData == null) return false;
                return (HeaderData[0] & 0x01) != 0;
            }
            set
            {
                if (HeaderData == null) HeaderData = CreateNewHeaderData();
                unchecked
                {
                    if (value) HeaderData[0] |= 1; else HeaderData[0] &= (byte)~1;
                }
            }
        }
            

        internal void SaveToStream(IDataStream DataStream, TSaveData SaveData, TXlsCellRange CellRange)
        {
            if (CellRange == null) SaveAllToStream(DataStream, SaveData); else SaveRangeToStream(DataStream, SaveData, CellRange);
        }

		private void SaveAllToStream(IDataStream DataStream, TSaveData SaveData)
		{
			int ListCount = GetCount();
			if (ListCount == 0) return;
			DataStream.WriteHeader((UInt16)xlr.DVAL, (UInt16) (HeaderData.Length + 4));
			DataStream.Write(HeaderData, HeaderData.Length);
			DataStream.Write32((UInt32)ListCount);

			foreach (TDataValidation dv in FList)
			{
				dv.SaveToStream(DataStream, SaveData);
			}
		}

		private void SaveRangeToStream(IDataStream DataStream, TSaveData SaveData, TXlsCellRange CellRange)
		{
			int ListCount = GetCount(CellRange);
			if (ListCount == 0) return;
			DataStream.WriteHeader((UInt16)xlr.DVAL, (UInt16) (HeaderData.Length + 4));
			DataStream.Write(HeaderData, HeaderData.Length);
			DataStream.Write32((UInt32)ListCount);

			foreach (TDataValidation dv in FList)
			{
				dv.SaveRangeToStream(DataStream, SaveData, CellRange);
			}
		}

		internal void Clear()
		{
			FList.Clear();
			HeaderData = null;
		}

		internal void ClearRange(TXlsCellRange CellRange, bool ClearHeader)
		{
			foreach (TDataValidation dv in FList)
			{
				dv.ClearRange(CellRange);
			}
			ShrinkList();
            if (FList.Count == 0 && ClearHeader) HeaderData = null;
		}

		private static byte[] CreateNewHeaderData()
		{
			return new byte[]{0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF};
		}

		internal void AddRange(TXlsCellRange CellRange, TDataValidationInfo DvInfo, TCellList CellList, bool ClearHeader, bool ReadingXlsx)
		{
			ClearRange(CellRange, ClearHeader); //Ensure no older dv stays.
			
			TDataValidation SearchDv = TDataValidation.CreateFromData(DvInfo, CellList, ReadingXlsx);
			TDataValidation NewDv = null;
			foreach (TDataValidation dv in FList)
			{
				if (dv.IsDvInfo(SearchDv))
				{
					NewDv = dv;
					break;
				}
			}
			if (NewDv == null) 
			{
				if (HeaderData == null) HeaderData = CreateNewHeaderData();
				NewDv = SearchDv;
				FList.Add(NewDv);
			}

			NewDv.AddRange(CellRange);
		}

		internal TDataValidationInfo GetDataValidation(int row, int col, TCellList CellList, bool WritingXlsx)
		{
			foreach (TDataValidation dv in FList)
			{
				if (dv.Contains(row, col)) return dv.GetDataValidationInfo(CellList, WritingXlsx);
			}
			return null;
		}

		internal int Count
		{
			get
			{
                return FList.Count;
			}
		}

		internal TDataValidationInfo GetDataValidation(int index, TCellList CellList, bool WritingXlsx)
		{
			return ((TDataValidation)FList[index]).GetDataValidationInfo(CellList, WritingXlsx);
		}

		internal TXlsCellRange[] GetDataValidationRange(int index, bool Inc1)
		{
			return ((TDataValidation)FList[index]).GetRanges(Inc1);
		}

		internal void InsertAndCopyRange(TXlsCellRange SourceRange, int DestRow, int DestCol, int aRowCount, int aColCount, TRangeCopyMode CopyMode, TFlxInsertMode InsertMode, TSheetInfo SheetInfo)
		{ 
			foreach (TDataValidation dv in FList)
			{
				dv.InsertAndCopyRange(SourceRange, DestRow, DestCol, aRowCount, aColCount, CopyMode, InsertMode, SheetInfo);
			}
		}

		internal void DeleteRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
		{
			foreach (TDataValidation dv in FList)
			{
				dv.DeleteRange(CellRange, aRowCount, aColCount, SheetInfo);
			}	
			ShrinkList();
		}
		
		internal void MoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
		{
			foreach (TDataValidation dv in FList)
			{
				dv.MoveRange(CellRange, NewRow, NewCol, SheetInfo);
			}		
		}

        internal void UpdateDeletedRanges(TDeletedRanges DeletedRanges)
        {
            foreach (TDataValidation dv in FList)
            {
                dv.UpdateDeletedRanges(DeletedRanges);
            }
        }
    }

	/// <summary>
	/// An internal representation of a Data Validation. 
	/// </summary>
	internal class TDataValidation : TRangeEntry
	{
		#region Privates
        private TDataValidationDataType FValidationType;
        private TDataValidationConditionType FCondition;
        private TParsedTokenList FirstFormula;
        private TParsedTokenList SecondFormula;
        private bool FIgnoreEmptyCells;
        private bool FInCellDropDown;
        private bool FExplicitList;

        private bool FShowErrorBox;
        private string FErrorBoxCaption;
        private string FErrorBoxText;

        private bool FShowInputBox;
        private string FInputBoxCaption;
        private string FInputBoxText;

        private TDataValidationIcon FErrorIcon;
        private TDataValidationImeMode FImeMode;

		int Reserved1, Reserved2; //Not needed, but to save exactly what we read.
		#endregion

        #region Constructors
        internal TDataValidation()
		{
			Init();
		}
        private void Init()
        {
            RangeValuesList = new TRangeValuesList(432, 0, true, true);
            Clear();
        }

        internal void Clear()
        {
            if (RangeValuesList != null) RangeValuesList.Clear();
            FValidationType = TDataValidationDataType.AnyValue;
            FCondition = TDataValidationConditionType.Between;
            FirstFormula = null;
            SecondFormula = null;
            FIgnoreEmptyCells = false;
            FInCellDropDown = false;
            FExplicitList = false;

            FShowErrorBox = false;
            FErrorBoxCaption = null;
            FErrorBoxText = null;

            FShowInputBox = false;
            FInputBoxCaption = null;
            FInputBoxText = null;
			FErrorIcon = TDataValidationIcon.Information;
            FImeMode = TDataValidationImeMode.NoControl;
        }

        #endregion

        #region Export to Biff8

        private static TExcelString GetDvString(string s, int MaxLen)
        {
            if (s == null || s.Length == 0) s = "\0";
            return new TExcelString(TStrLenLength.is16bits, s.Substring(0, Math.Min(s.Length, MaxLen)), null, false);
        }

        private static void WriteString(IDataStream DataStream, TExcelString s)
        {
			if (s == null) 
			{
				DataStream.Write16(0);
				return;
			}

            byte[] b = new byte[s.TotalSize()];
            s.CopyToPtr(b, 0);
            DataStream.Write(b, b.Length);
        }

		private static int GetSize(TExcelString s)
		{
			if (s == null) return 2;
			return s.TotalSize();
		}

        private void WriteBiff8Data(TNameRecordList Names, IDataStream DataStream, out int Len)
        {
            int op1 = 0;
            op1 |= (int)FValidationType & 0x0F;
            op1 |= ((int)FErrorIcon << 4) & 0x70;

            if (FExplicitList) op1 |= 0x80; //List is inside the formula.

            if (FIgnoreEmptyCells) op1 |= 0x100;
            if (!FInCellDropDown) op1 |= 0x200;

            int op2 = 0;
            unchecked
            {
                op1 |= ((int)FImeMode) << 11;
                //op2 will never be used.
            }


            if (FShowInputBox) op2 |= 0x4;
            if (FShowErrorBox) op2 |= 0x8;

            op2 |= ((int)FCondition << 4) & 0xF0;

            TExcelString InputCaption = GetDvString(FInputBoxCaption, FlxConsts.Max_DvInputTitleLen);
            TExcelString ErrorCaption = GetDvString(FErrorBoxCaption, FlxConsts.Max_DvErrorTitleLen);
            TExcelString InputText = GetDvString(FInputBoxText, FlxConsts.Max_DvInputTextLen);
            TExcelString ErrorText = GetDvString(FErrorBoxText, FlxConsts.Max_DvErrorTextLen);

            int FirstFormulaPos = 4 + GetSize(InputCaption) + GetSize(ErrorCaption) + GetSize(InputText) + GetSize(ErrorText) + 4;


            if (FirstFormula == null && FValidationType != TDataValidationDataType.AnyValue) XlsMessages.ThrowException(XlsErr.ErrDataValidationFmla1Null);
            if (SecondFormula == null)
            {
                if (FValidationType != TDataValidationDataType.AnyValue && FValidationType != TDataValidationDataType.List && FValidationType != TDataValidationDataType.Custom)
                {
                    if (FCondition == TDataValidationConditionType.Between || FCondition == TDataValidationConditionType.NotBetween)
                    {
                        XlsMessages.ThrowException(XlsErr.ErrDataValidationFmla2Null);
                    }
                }
            }

            int Fmla1Size = 0;

            int Fmla1LenWithoutArray = 0;
            int Fmla2LenWithoutArray = 0;

            byte[] ps1 = null;
            if (FirstFormula != null)
            {
                ps1 = TFormulaConvertInternalToBiff8.GetTokenData(Names, FirstFormula, TFormulaType.DataValidation, out Fmla1LenWithoutArray);
                Fmla1Size = ps1.Length;
            }

            int SecondFormulaPos = FirstFormulaPos + Fmla1Size + 4;

            int Fmla2Size = 0;
            byte[] ps2 = null;
            if (SecondFormula != null)
            {
                ps2 = TFormulaConvertInternalToBiff8.GetTokenData(Names, SecondFormula, TFormulaType.DataValidation, out Fmla2LenWithoutArray);
                Fmla2Size = ps2.Length;
            }

            Len = SecondFormulaPos + Fmla2Size;
            if (DataStream == null) return;

            DataStream.Write16((UInt16)op1);
            DataStream.Write16((UInt16)op2);

            WriteString(DataStream, InputCaption);
            WriteString(DataStream, ErrorCaption);
            WriteString(DataStream, InputText);
            WriteString(DataStream, ErrorText);

            DataStream.Write16((UInt16)Fmla1LenWithoutArray);
            DataStream.Write16((UInt16)Reserved1);
            if (ps1 != null) DataStream.Write(ps1, ps1.Length);

            DataStream.Write16((UInt16)Fmla2LenWithoutArray);
            DataStream.Write16((UInt16)Reserved2);
            if (ps2 != null) DataStream.Write(ps2, ps2.Length);
        }

        private int HeaderSize
        {
            get
            {
                int Len;
                WriteBiff8Data(null, null, out Len);
                return Len;
            }
        }
        #endregion

        #region Import From Biff8
        private static string GetMessage(byte[] Data, ref int aPos)
        {
            long asize = 0; string Result = null;
            StrOps.GetSimpleString(true, Data, aPos, false, 0, ref Result, ref asize);
            aPos += (int)asize;
            if (Result == "\0") Result = null;
            return Result;
        }

        private static TParsedTokenList GetFormula(TNameRecordList Names, byte[] Data, out int Reserved, ref int aPos)
        {
            TParsedTokenList Result = null;
			int Size = BitOps.GetWord(Data, aPos);
			Reserved = BitOps.GetWord(Data, aPos + 2);
            aPos += 4;
            if (Size > 0)
            {
                TFormulaConvertBiff8ToInternal Convert = new TFormulaConvertBiff8ToInternal();

                Result = Convert.ParseRPN(Names, -1, -1, Data, aPos, Size, true); //no real need for relative since shared formulas can't be 3d, and we only need relative for the non-existing tokens ptgarea3dn and ptgref3dn.
                aPos += Size;
            }

            return Result;
        }

        internal void ImportFromBiff8(TNameRecordList Names, byte[] Data, out int aPos)
        {
            int op1 = BitOps.GetWord(Data, 0);

            FValidationType = (TDataValidationDataType)(op1 & 0x0F);
            FErrorIcon = (TDataValidationIcon)((op1 & 0x70) >> 4);
            unchecked
            {
                FImeMode = (TDataValidationImeMode)(op1 >> 11); //op2 is not used.
            }

            FExplicitList = (op1 & 0x80) != 0; //List is inside the formula.

            FIgnoreEmptyCells = (op1 & 0x100) != 0;
            FInCellDropDown = (op1 & 0x200) == 0;

            int op2 = BitOps.GetWord(Data, 2);
            FShowInputBox = (op2 & 0x4) != 0;
            FShowErrorBox = (op2 & 0x8) != 0;

            FCondition = (TDataValidationConditionType)((op2 & 0xF0) >> 4);

            aPos = 4; //skip optionflags
            FInputBoxCaption = GetMessage(Data, ref aPos);
            FErrorBoxCaption = GetMessage(Data, ref aPos);
            FInputBoxText = GetMessage(Data, ref aPos);
            FErrorBoxText = GetMessage(Data, ref aPos);

            FirstFormula = GetFormula(Names, Data, out Reserved1, ref aPos);
            SecondFormula = GetFormula(Names, Data, out Reserved2, ref aPos);
        }
        #endregion

        internal static TDataValidation CreateFromData(TDataValidationInfo dv, TCellList CellList, bool ReadingXlsx)
        {
            TDataValidation Result = new TDataValidation();
            Result.FValidationType = dv.ValidationType;
            Result.FCondition = dv.Condition;
            Result.FIgnoreEmptyCells = dv.IgnoreEmptyCells;
            Result.FInCellDropDown = dv.InCellDropDown;
            Result.FExplicitList = dv.ExplicitList;

			Result.FErrorIcon = dv.ErrorIcon;

            Result.FImeMode = dv.ImeMode;

            Result.FShowErrorBox = dv.ShowErrorBox;
            Result.FErrorBoxCaption = dv.ErrorBoxCaption;
            Result.FErrorBoxText = dv.ErrorBoxText;

            Result.FShowInputBox = dv.ShowInputBox;
            Result.FInputBoxCaption = dv.InputBoxCaption;
            Result.FInputBoxText = dv.InputBoxText;

            if (! String.IsNullOrEmpty(dv.FirstFormula))
            {
                if (dv.FirstFormula.Length > 255) XlsMessages.ThrowException(XlsErr.ErrDataValidationFmla1TooLong, dv.FirstFormula);
                TFormulaConvertTextToInternal ps1 = new TFormulaConvertTextToInternal(CellList.Workbook, CellList.Workbook.ActiveSheet, true, dv.FirstFormula, true, true, false, null, TFmReturnType.Ref, false);
                if (ReadingXlsx) ps1.SetReadingXlsx();
                ps1.Parse();
                Result.FirstFormula = ps1.GetTokens();
            }

            if (!string.IsNullOrEmpty(dv.SecondFormula))
            {
                if (dv.SecondFormula.Length > 255) XlsMessages.ThrowException(XlsErr.ErrDataValidationFmla2TooLong, dv.SecondFormula);
                TFormulaConvertTextToInternal ps2 = new TFormulaConvertTextToInternal(CellList.Workbook, CellList.Workbook.ActiveSheet, true, dv.SecondFormula, true, true, false, null, TFmReturnType.Ref, false);
                if (ReadingXlsx) ps2.SetReadingXlsx();
                ps2.Parse();
                Result.SecondFormula = ps2.GetTokens();
            }


            Result.RangeValuesList.OtherDataLen = Result.HeaderSize;
            return Result;
        }

        internal bool IsDvInfo(TDataValidation dv)
        {
            //This is a little more complex than just comparing the data validations, because string case in formulas can make things different.
            //For example, if one fmla is "=m1" and the other is "=M1" they are actually the same, but they will be reported as different.
            //Doing a case insensitive compare does not help, because ="A" should be different from ="a".
            //If we did not fix it, adding "=m1" many times would end up in lots of different data validation blocks, because the converted one would be "=M1".
            //return dv == GetDataValidationInfo(CellList);

            //After changing this to store the parsedtokenlists instead of the rpns, now we can compare the datavalidationinfo.
            if (FValidationType != dv.FValidationType
                    || FCondition != dv.FCondition
                    || FIgnoreEmptyCells != dv.FIgnoreEmptyCells
                    || FInCellDropDown != dv.FInCellDropDown
                    || FExplicitList != dv.FExplicitList
				    || FErrorIcon != dv.FErrorIcon
                    || FImeMode != dv.FImeMode

                    || FShowErrorBox != dv.FShowErrorBox
                    || FErrorBoxCaption != dv.FErrorBoxCaption
                    || FErrorBoxText != dv.FErrorBoxText

                    || FShowInputBox != dv.FShowInputBox
                    || FInputBoxCaption != dv.FInputBoxCaption
                    || FInputBoxText != dv.FInputBoxText)
            {
                return false;
            }

            if (FirstFormula == null ^ dv.FirstFormula == null) return false;
			if (dv.FirstFormula != null) //dv first is also != null
			{
				if (!FirstFormula.SameTokens(dv.FirstFormula)) return false;
			}

			if (SecondFormula == null ^ dv.SecondFormula == null) return false;
			if (dv.SecondFormula != null)
			{
				if (!SecondFormula.SameTokens(dv.SecondFormula)) return false;
			}

            return true;
        }

		internal TDataValidationInfo GetDataValidationInfo(TCellList CellList, bool WritingXlsx)
		{
			TDataValidationInfo Result = new TDataValidationInfo();
            Result.ValidationType = FValidationType;
            Result.Condition = FCondition;
            Result.IgnoreEmptyCells = FIgnoreEmptyCells;
            Result.InCellDropDown = FInCellDropDown;
            Result.ExplicitList = FExplicitList;

            Result.ShowErrorBox = FShowErrorBox;
            Result.ErrorBoxCaption = FErrorBoxCaption;
            Result.ErrorBoxText = FErrorBoxText;

            Result.ShowInputBox = FShowInputBox;
            Result.InputBoxCaption = FInputBoxCaption;
            Result.InputBoxText = FInputBoxText;

			Result.ErrorIcon = FErrorIcon;
            Result.ImeMode = FImeMode;

            if (FirstFormula != null) Result.FirstFormula = TFormulaConvertInternalToText.AsString(FirstFormula, 0, 0, CellList, CellList.Globals, -1, WritingXlsx);
            if (SecondFormula != null) Result.SecondFormula = TFormulaConvertInternalToText.AsString(SecondFormula, 0, 0, CellList, CellList.Globals, -1, WritingXlsx);
			return Result;
		}

		protected override TRangeEntry DoCopyTo(TSheetInfo SheetInfo)
		{
			TDataValidation Result=(TDataValidation)base.DoCopyTo(SheetInfo);
            if (Result.FirstFormula != null) Result.FirstFormula = FirstFormula.Clone();
            if (Result.SecondFormula != null) Result.SecondFormula = SecondFormula.Clone();
			return Result;
		}


		internal override void LoadFromStream(TBaseRecordLoader RecordLoader, TRangeRecord First)
		{
			//continue is not allowed here.
			Clear();
			int aPos;
            ImportFromBiff8(RecordLoader.Names, First.Data, out aPos);
            RangeValuesList.OtherDataLen = HeaderSize;  //aPos would be dangerous to use here, since we might output something longer.
            
            RangeValuesList.LoadFromBiff8(First, aPos);
		}

		internal bool IsEmpty()
		{
			return RangeValuesList.Count == 0;
		}

		internal bool IsEmpty(TXlsCellRange CellRange)
		{
			return RangeValuesList.CountRangeRecords(CellRange) == 0;
        }

        #region SaveToStream

		internal override void SaveToStream(IDataStream DataStream, TSaveData SaveData)
		{
			if (RangeValuesList.Count==0) return; //Don't save empty Data Validations
			int aCount = RangeValuesList.RepeatCountR(RangeValuesList.Count);
			for (int i=0;i< aCount;i++)
			{
				DataStream.WriteHeader((UInt16)xlr.DV, (UInt16)RangeValuesList.RecordSizeR(i, RangeValuesList.Count)); //this includes the headers in headersize.
                int Len;
				WriteBiff8Data(SaveData.Globals.Names, DataStream, out Len);
				RangeValuesList.SaveToStreamR(DataStream, SaveData, i);
			}
		}
        
		internal override void SaveRangeToStream(IDataStream DataStream, TSaveData SaveData, TXlsCellRange CellRange)
		{
			int Rc=RangeValuesList.CountRangeRecords(CellRange);
			if (Rc==0) return; //Don't save empty CF's
			int aCount = RangeValuesList.RepeatCountR(Rc);
			for (int i=0; i< aCount;i++)
			{
				DataStream.WriteHeader((UInt16)xlr.DV, (UInt16)RangeValuesList.RecordSizeR(i, Rc));
                int Len;
				WriteBiff8Data(SaveData.Globals.Names, DataStream, out Len);
				RangeValuesList.SaveRangeToStreamR(DataStream, SaveData, i, Rc, CellRange);
			}
		}

		internal override long TotalSize()
		{
			if (RangeValuesList.Count==0) return 0;
			return RangeValuesList.TotalSizeR(RangeValuesList.Count);
		}

		internal override long TotalRangeSize(TXlsCellRange CellRange)
		{
			int i= RangeValuesList.CountRangeRecords(CellRange);
			if (RangeValuesList.Count==0) return 0; 
			else
				return RangeValuesList.TotalSizeR(i);
        }
        #endregion

        #region Insert And Copy
        private void ArrangeInsertFmlas(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
		{
			if (FirstFormula != null) TTokenManipulator.ArrangeInsertAndCopyRange(FirstFormula, CellRange, -1, -1, aRowCount, aColCount, 0, 0, SheetInfo, true, null);
			if (SecondFormula != null) TTokenManipulator.ArrangeInsertAndCopyRange(SecondFormula, CellRange, -1, -1, aRowCount, aColCount, 0, 0, SheetInfo, true, null);
		}
		
		private void ArrangeMoveFmlas(TXlsCellRange CellRange, int aNewRow, int aNewCol, TSheetInfo SheetInfo)
		{
            if (FirstFormula != null) TTokenManipulator.ArrangeMoveRange(FirstFormula, CellRange, -1, -1, aNewRow, aNewCol, SheetInfo, null);
            if (SecondFormula != null) TTokenManipulator.ArrangeMoveRange(SecondFormula, CellRange, -1, -1, aNewRow, aNewCol, SheetInfo, null);
		}

		internal override void ArrangeInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
		{
			base.ArrangeInsertRange(CellRange, aRowCount, aColCount, SheetInfo);
			ArrangeInsertFmlas(CellRange, aRowCount, aColCount, SheetInfo);
		}

        internal override void InsertAndCopyRange(TXlsCellRange SourceRange, int DestRow, int DestCol, int aRowCount, int aColCount, TRangeCopyMode CopyMode, TFlxInsertMode InsertMode, TSheetInfo SheetInfo)
		{
			base.InsertAndCopyRange(SourceRange, DestRow, DestCol, aRowCount, aColCount, CopyMode, InsertMode, SheetInfo);
			if (CopyMode!=TRangeCopyMode.None)
			{
				int R1 = 0; int R2 = 0; int C1 = 0; int C2 = 0;
				RangeValuesList.CopyRangeInclusive(SourceRange, DestRow, DestCol, aRowCount, aColCount, ref R1, ref R2, ref C1, ref C2);
			}
		}

		internal override void DeleteRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
		{
			base.DeleteRange(CellRange, aRowCount, aColCount, SheetInfo);
			ArrangeInsertFmlas(CellRange, -aRowCount, -aColCount, SheetInfo);
		}

		internal override void MoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
		{
			base.MoveRange (CellRange, NewRow, NewCol, SheetInfo);
			ArrangeMoveFmlas(CellRange, NewRow, NewCol, SheetInfo);
		}

		internal override void ArrangeMoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
		{
			base.ArrangeMoveRange (CellRange, NewRow, NewCol, SheetInfo);
        }
        #endregion

        internal override void UpdateDeletedRanges(TDeletedRanges DeletedRanges)
        {
            if (FirstFormula != null) TTokenManipulator.UpdateDeletedRanges(FirstFormula, DeletedRanges);
            if (SecondFormula != null) TTokenManipulator.UpdateDeletedRanges(SecondFormula, DeletedRanges);
        }

        internal void AddRange(TXlsCellRange CellRange)
		{
			RangeValuesList.AddAndMerge(new TExcelRange(CellRange.Top, CellRange.Left, CellRange.Bottom, CellRange.Right));
		}

		internal bool Contains(int aRow, int aCol)
		{
			for (int i= RangeValuesList.Count-1; i>=0; i--)
			{
				TExcelRange xr = RangeValuesList[i];
				if (xr.C1<=aCol && xr.C2>=aCol && xr.R1<=aRow && xr.R2>=aRow) return true;
			}
			return false;
		}

		internal TXlsCellRange[] GetRanges(bool Inc1)
		{
			TXlsCellRange[] Result = new TXlsCellRange[RangeValuesList.Count];
			for (int i= 0; i < Result.Length; i++)
			{
				TExcelRange xr = RangeValuesList[i];
				if (Inc1)
					Result[i] = new TXlsCellRange(xr.R1 + 1, xr.C1 + 1, xr.R2 + 1, xr.C2 + 1);
				else
					Result[i] = new TXlsCellRange(xr.R1, xr.C1, xr.R2, xr.C2);
			}
			return Result;
		}

    }

	#endregion
}
