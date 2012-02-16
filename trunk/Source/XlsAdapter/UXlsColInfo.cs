using System;
using System.IO;
using FlexCel.Core;
using System.Diagnostics;

namespace FlexCel.XlsAdapter
{
    /// <summary>
    /// Column information 
    /// </summary>
    internal class TColInfoRecord: TxBaseRecord
    {
        TBiff8XFMap XFMap;

        internal TColInfoRecord(int aId, byte[] aData, TBiff8XFMap aXFMap): base(aId, aData)
        {
            XFMap = aXFMap;
        }

        internal int FirstColumn { get { return GetWord(0); } }
        internal int LastColumn { get { return GetWord(2); } }
        internal int Width { get { return GetWord(4); } }
        internal int XF { get { if (XFMap == null) return GetWord(6); return XFMap.GetCellXF2007(GetWord(6)); } }
        internal int Options { get { return GetWord(8); } }
        internal int Reserved { get { return GetWord(10); } }

        internal static int Length {get {return 12;}}

		internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
		{
			ws.Columns.AddBiff8Record(this);
		}

    }

	/// <summary>
	/// Formatting Info on columns, on a friendly way.
	/// </summary>
	internal class TColInfo
	{
        internal int FWidth;
        internal int XF;
        internal int Options;
		internal bool MarkforAutofit;
		internal bool HasMergedCell;
        internal int KeepTogether;
		internal float AutofitAdjustment;
		internal int AutofitAdjustmentFixed;
		internal int MinWidth;
		internal int MaxWidth;

        public TColInfo(int aWidth, int aXF, int aOptions, bool KeepAutoWidth)
		{
            FWidth   = aWidth;
            XF      = aXF;
            Options = aOptions; 
			if (!KeepAutoWidth) Options |= 0x02;  //The | 02 is not documented, but if false and there is a standardwidth record below the column info, that value will be used (!)
			                              //It looks like 02 is used to mark the column as "has its own width" so it will not be resized when changing standard width.
		}

		public int Width 
		{
			get {return FWidth;} 
			set
			{
				FWidth = value;
				Options |= 0x02;  //If we change the width, we need to update this one.
			}
		}

        internal void SetColOutlineLevel(int Level)
        {
            Options= (Options & ~(0x7<<8)) | ((Level & 0x7)<<8);
        }

        internal int GetColOutlineLevel()
        {
            return (Options>>8) & 0x7;
        }
    
        internal bool IsEqual(TColInfo aColInfo)
        {
            if (aColInfo == null) return false;
            return // don't compare the column .... (Column = aColInfo.Column) and
                (FWidth  == aColInfo.Width)  &&
                (XF     == aColInfo.XF)     &&
                (Options== aColInfo.Options);
        }

        internal void ChangeStandardWidth(int OldStdValue, int NewStdValue, bool BeforeWas256)
		{
            if ((Options & 0x02) == 0)
            {
                if (!BeforeWas256) Options |= 0x02; //If we didn't had a stdwidth record, Options 0x02 wasn't used, so we need to set this value just in case.
                else FWidth = NewStdValue; //Columns with option 0x02 are the ones that do not have manual width.
            }
		}

   }

    /// <summary>
    /// List of ColInfos.
    /// </summary>
	internal class TColInfoList
	{
        private TColInfo DefaultColumn; //this is saved as 0x100 in biff8
		private TColInfo[][] FColumns;
        private const int Buckets = 257;

        internal TWorkbookGlobals FWorkbookGlobals;
        internal int DefColWidthChars;
        internal int DefColWidthChars256; //This is StandardWidth: It is used if and only if DefColWidhtChars is 8, despite what the docs say.
        internal bool AllowStandardWidth;
        internal bool IsDialog;


		internal TColInfoList(TWorkbookGlobals aWorkbookGlobals, bool aAllowStandardWidth)
		{
            FWorkbookGlobals = aWorkbookGlobals;
			FColumns = new TColInfo[(int)Math.Ceiling((FlxConsts.Max_Columns+2) / (double)Buckets)][];
            DefColWidthChars = 8; //Defcolwidthchars is a required record.
            DefColWidthChars256 = -1;
            AllowStandardWidth = aAllowStandardWidth;
		}

		#region Generics
        private TColInfo GetItem(int Column)
        {
            if (Column < 0 || Column > FlxConsts.Max_Columns + 1) return null; 
            if (FColumns[Column / Buckets] == null) return null;
            return FColumns[Column / Buckets][Column % Buckets];
        }

		internal void Add (int Column, TColInfo a)
		{
			if (Column<0 || Column>FlxConsts.Max_Columns+1) return;
            TColInfo[] Buck = FColumns[Column / Buckets];
            if (a == null)
            {
                if (Buck == null) return;
            }
            if (Buck == null) FColumns[Column / Buckets] = new TColInfo[Buckets];
            FColumns[Column / Buckets][Column % Buckets] = a;
        }
 
		internal TColInfo this[int index] 
		{
            get {return GetItem(index); } 
			set {Add(index, value);}
		}

		#endregion

        internal TFlxFont Get0Font()
        {
            if (FWorkbookGlobals.Fonts.Count > 0)
                return FWorkbookGlobals.Fonts[0].FlxFont();
            return null;
        }

        internal int DefColWidth
        {
            get
            {
                if (IsDialog) return 256; //dialogs don't change the col width
                if (UsesStdWidth256())
                {
                    return DefColWidthChars256;
                }

                if (DefColWidthChars >= 0) return ExcelMetrics.DefColWidthAdapt(DefColWidthChars);
                return ExcelMetrics.DefColWidthAdapt(8);
            }
            set
            {
                if (AllowStandardWidth)
                {
                    bool BeforeWas256 = UsesStdWidth256();
                    DefColWidthChars = 8; //If it is 8 it will be ignored.
                    ChangeStandardWidth(DefColWidthChars256, value, BeforeWas256);
                    DefColWidthChars256 = value;
                }
                else
                {
                    DefColWidthChars256 = -1;
                    DefColWidthChars = ExcelMetrics.InverseDefColWidthAdapt(value);
                }
            }
        }

        private bool UsesStdWidth256()
        {
            return AllowStandardWidth && DefColWidthChars256 >= 0 && (DefColWidthChars < 0 || DefColWidthChars == 8);
        }	

		internal void Clear()
		{
            DefColWidthChars = -1;
            DefColWidthChars256 = -1;
			for (int i= FColumns.Length-1; i>=0; i--)
				FColumns[i]=null;
		}

        internal void CopyFrom(TColInfoList aColInfoList)
        {
            DefColWidthChars = aColInfoList.DefColWidthChars;
            DefColWidthChars256 = aColInfoList.DefColWidthChars256;

            if (aColInfoList.FColumns == FColumns) XlsMessages.ThrowException(XlsErr.ErrInternal); //Should be different objects

            if(aColInfoList.DefaultColumn != null) DefaultColumn = new TColInfo(aColInfoList.DefaultColumn.Width, aColInfoList.DefaultColumn.XF, aColInfoList.DefaultColumn.Options, true);
            for (int i = FColumns.Length - 1; i >= 0; i--)
            {
                TColInfo[] a = aColInfoList.FColumns[i];
                if (a != null)
                {
                    FColumns[i] = new TColInfo[a.Length];
                    for (int k = a.Length - 1; k >= 0; k--)
                    {
                        if (a[k] != null) FColumns[i][k] = new TColInfo(a[k].Width, a[k].XF, a[k].Options, true);
                    }

                }
            }
        }

		internal void MarkColForAutofit(int Col, bool Autofit, float Adjustment, int AdjustmentFixed, int MinWidth, int MaxWidth, bool IsMerged)
		{
			if (this[Col] == null) Add(Col, new TColInfo(DefColWidth, FlxConsts.DefaultFormatId, 0, true));
            TColInfo ColInfo = this[Col];
			if (IsMerged)
			{
				ColInfo.HasMergedCell = Autofit;
			}
			else
			{
				ColInfo.MarkforAutofit = Autofit;
			}
            ColInfo.AutofitAdjustment = Adjustment;
            ColInfo.AutofitAdjustmentFixed = AdjustmentFixed;
            ColInfo.MinWidth = MinWidth;
            ColInfo.MaxWidth = MaxWidth;
		}

        internal void KeepColsTogether(int Col1, int Col2, int Level, bool replaceLowerLevels)
        {
            for (int Col = Col1; Col < Col2; Col++) //Col2 is not included.
            {
                if (this[Col] == null)
                {
                    if (Level == 0) continue;

                    this[Col] = new TColInfo(DefColWidth, FlxConsts.DefaultFormatId, 0, true);
                }
                if (replaceLowerLevels || this[Col].KeepTogether < Level) this[Col].KeepTogether = Level;
            }

        }

        internal int GetKeepColsTogether(int Col)
        {
            if (this[Col] == null) return 0;
            return this[Col].KeepTogether;
        }

		internal bool HasKeepColsTogether()
		{
			for (int Col = 0; Col < ColCount; Col++) 
			{
				if (this[Col] != null && this[Col].KeepTogether != 0) return true;
			}
			return false;
		}


        internal void AddBiff8Record(TColInfoRecord R)
        {
            //Excel can save a column 0x100 in a range, for example (1:0x100) and 0x100 will be used for the default column.
            //BUT, having a range 0x100:0x100 will crash Excel 2000 when closing, and have warning in Excel 2003!!
            //So, if the default column is not in the same format as the last column, it cannot be saved. Excel doesn't do it either,
            //and if you save a file with last column different from default col, close it and reopen, the default col will be gone.

            int LastColumn = R.LastColumn;
            if (LastColumn == FlxConsts.Max_Columns97_2003 + 1)
            {
                DefaultColumn = new TColInfo(R.Width, R.XF, R.Options, true);
                LastColumn--;
            }

            if (R.FirstColumn < 0 || LastColumn > FlxConsts.Max_Columns97_2003 || R.FirstColumn > LastColumn) return;

            for (int i = R.FirstColumn; i <= LastColumn; i++)
                Add(i, new TColInfo(R.Width, R.XF, R.Options, true));
        }

		private void SaveOneRecord(int i, int k, IDataStream DataStream, TSaveData SaveData)
		{
			DataStream.WriteHeader((UInt16)xlr.COLINFO, (UInt16) TColInfoRecord.Length);
			DataStream.Write16((UInt16)i);
			DataStream.Write16((UInt16)k);
			DataStream.Write16((UInt16)this[i].Width);
			DataStream.Write16(SaveData.GetBiff8FromCellXF(this[i].XF));
			DataStream.Write16((UInt16)(this[i].Options));  
			DataStream.Write16(0);
		}

		private void SaveOnePxlRecord(int i, int k, TPxlStream PxlStream, TPxlSaveData SaveData)
		{
			//We need to ensure this[i] is not bigger than Maxcolumns. this[k] can be=Maxcolumns+1.
			if (i>FlxConsts.Max_Columns97_2003) return;
			//really for pxl k should be Maxcolumns only
			if (k>FlxConsts.Max_Columns97_2003) k = FlxConsts.Max_Columns97_2003;

			PxlStream.WriteByte((byte)pxl.COLINFO);
			PxlStream.Write16((UInt16)i);
			PxlStream.Write16((UInt16)k);
			if (this[i] == null)
			{
				PxlStream.Write16((UInt16)0x900); //width
				PxlStream.Write16((UInt16)0); //xf
				PxlStream.WriteByte((byte)(0)); //options
			}
			else
			{
				PxlStream.Write16((UInt16)this[i].Width);
                PxlStream.Write16(SaveData.GetBiff8FromCellXF(this[i].XF));
				PxlStream.WriteByte((byte)(this[i].Options & 1));   
			}
		}

        /// <summary>
        /// Will return default as column 256.
        /// </summary>
        /// <returns></returns>
        private TColInfo GetBiff8Col(int k)
        {
            if (k == FlxConsts.Max_Columns97_2003 + 1) return DefaultColumn;
            return this[k];
        }

        internal void SaveDefColWidth(IDataStream DataStream, TSaveData SaveData)
        {
            TDefColWidthRecord.SaveRecord(DataStream, DefColWidthChars);
        }

		private long SaveRangeEx(IDataStream DataStream, TSaveData SaveData, int FirstCol, int LastCol)
		{
            long Result = 0;

            if (DataStream != null) TDefColWidthRecord.SaveRecord(DataStream, DefColWidthChars);
            Result += TDefColWidthRecord.StandardSize();

			int i=FirstCol; 
			int k=i;
			while (i<= LastCol)
			{
				if (this[i]==null) {i++; continue;}
				k=i+1;
				while (k<= LastCol && GetBiff8Col(k)!=null && GetBiff8Col(k).IsEqual(this[i])) k++;

				//We need to ensure this[i] is not bigger than Maxcolumns. this[k] can be=Maxcolumns+1.
				if (i<=FlxConsts.Max_Columns97_2003)
				{
					if (DataStream != null) SaveOneRecord(i, k-1, DataStream, SaveData);
					Result+= XlsConsts.SizeOfTRecordHeader+TColInfoRecord.Length;
				}
				i=k;
			}
			return Result;
		}

		internal void SaveToStream(IDataStream DataStream, TSaveData SaveData, TXlsCellRange CellRange)
		{
            int Start = CellRange == null ? 0 : CellRange.Left;
            int Finish = CellRange == null ? FlxConsts.Max_Columns97_2003 + 1 : CellRange.Right;
			SaveRangeEx(DataStream, SaveData, Start, Finish);
		}

		private static bool IsEqual(TColInfo a, TColInfo b)
		{
			if (a == null)
			{
				if (b == null) return true;
				return false;
			}

			//a != null
			if (b == null) return false;
			return a.IsEqual(b);
		}

		internal void SaveToPxl(TPxlStream PxlStream, TPxlSaveData SaveData)
		{
            TDefColWidthRecord.SaveToPxl(PxlStream, DefColWidthChars);
			//Pxl needs the full 255 columns defined, starting at 0 and ending af 255.
            int LastCol = FlxConsts.Max_Columns97_2003 + 1;
			int i=0; 
			int k=i;
			while (i<= LastCol)
			{
				k=i+1;
				while (k<= LastCol && IsEqual(this[k], this[i])) k++;
				if (PxlStream != null) SaveOnePxlRecord(i, k-1, PxlStream, SaveData);
				i=k;
			}

			if (k < LastCol  && (PxlStream != null))  
			{
				SaveOnePxlRecord(k, FlxConsts.Max_Columns97_2003, PxlStream, SaveData);
			}
		}

		internal long TotalSizeNoStdWidth(TXlsCellRange CellRange)
		{
            int Start = CellRange == null ? 0 : CellRange.Left;
            int Finish = CellRange == null ? FlxConsts.Max_Columns97_2003 + 1 : CellRange.Right;
            return SaveRangeEx(null, new TSaveData(), Start, Finish);
		}

        internal int StandardWidthSize()
        {
            int StdWidth = 0;
            if (AllowStandardWidth && DefColWidthChars256 >= 0) StdWidth = XlsConsts.SizeOfTRecordHeader + 2;
            return StdWidth;
        }

		private int LastColumn
		{
			get
			{
                for (int i = FColumns.Length - 1; i >= 0; i--)
                {
                    if (FColumns[i] != null)
                    {
                        for (int k = FColumns[i].Length - 1; k >= 0; k--)
                        {
							if (FColumns[i][k] != null) 
							{
								return (i * Buckets) + k;
							}
                        }
                    }
                }
				return -1;
			}
		}

		public int ColCount
		{
			get
			{
				int Result = LastColumn + 1;
				if (Result > FlxConsts.Max_Columns + 1 ) return FlxConsts.Max_Columns + 1;
				return Result;

			}
		}

        private void DeleteColumns(int First, int ColCount)
        {
			int Lc = LastColumn + 1;
            for (int i = First; i < Lc - ColCount; i++)
            {
                this[i] = this[i + ColCount];
            }

            for (int i = Math.Max(First, Lc - ColCount); i < Lc; i++) //delete entries that were not overriden by the goback
            {
                this[i] = null;
                /* We won't use defaultcolumn, as it would not allow to reduce the number of columns when deleting.
                if (DefaultColumn == null) this[i] = null;
                else
                {
                    //The new column will have the last column width and the other format from DefaultColumn.
                    //We will just use the default format width.
                    this[i] = new TColInfo(DefaultColumn.Width, DefaultColumn.XF, DefaultColumn.Options, true);
                }*/
            }
        }

        private void InsertColumns(int First, int ColCount)
        {
			int Lc = LastColumn + 1;
			for (int i = Lc - 1; i >= First; i--)
            {
                this[i + ColCount] = this[i];
            }

            for (int i = First; i < First + ColCount; i++)
            {
                this[i] = null;
            }
        }

		internal void ArrangeInsertCols(TXlsCellRange CellRange, int aColCount, TSheetInfo SheetInfo)
		{
			if (SheetInfo.SourceFormulaSheet != SheetInfo.InsSheet) return;
			int ColOffset=CellRange.ColCount*aColCount;

			if (aColCount<0) //Deleting columns. ColOffset is < 0.
			{
                DeleteColumns(CellRange.Left, -ColOffset);
            }
			else
			{
				//the check below might throw unwanted exceptions when all columns are formatted (even with fake formatting)
				//and disallow to insert columns on some sheets. (for example all pxl files have the full 255 columns formatted)
				//so we will allow to "lose" formatted columns if there is no data on them.
				//if (LastColumn+ColOffset>FlxConsts.Max_Columns) XlsMessages.ThrowException(XlsErr.ErrTooManyColumns, LastColumn + ColOffset + 1, FlxConsts.Max_Columns+1);
                
				if (CellRange.Left+ColOffset>FlxConsts.Max_Columns) XlsMessages.ThrowException(XlsErr.ErrTooManyColumns, CellRange.Left + ColOffset + 1, FlxConsts.Max_Columns+1);

                InsertColumns(CellRange.Left, ColOffset);
			}
		}

		internal void ArrangeMoveCols(TXlsCellRange CellRange, int NewCol, TSheetInfo SheetInfo)
		{
			if (SheetInfo.SourceFormulaSheet != SheetInfo.InsSheet) return;
			if (NewCol == CellRange.Left) return;

			for (int z = CellRange.Left; z <= CellRange.Right; z++)
			{
				int c = CellRange.Left > NewCol? z: CellRange.Left - z + CellRange.Right; //order here is important, to not overwrite existing cells.

				this[c + NewCol - CellRange.Left] = this[c];
				this[c] = null;
			}
		}

		internal void CopyCols(TXlsCellRange SourceRange, int DestCol, int aColCount, TSheetInfo SheetInfo)
		{
			//ArrangeInsertCols(SourceRange.Offset(SourceRange.Top, DestCol), aColCount, SheetInfo);  //This has already been called.
			
            for (int k=0;k<aColCount;k++)
				for (int i=SourceRange.Left; i<=SourceRange.Right;i++)
				{
					TColInfo C=this[i];
					int NewCol=DestCol+i-SourceRange.Left+k*SourceRange.ColCount;
					if (C!=null && NewCol>=0 && NewCol<=FlxConsts.Max_Columns+1) 
						this[NewCol]= new TColInfo(C.Width, C.XF, C.Options, true);
				}
		}

		internal void CalcGuts(TGutsRecord Guts)
		{
			int MaxGutLevel=0;
   			for (int i = 0; i < FColumns.Length; i++)
			{
				if (FColumns[i] != null) 
                {
                    for (int k = 0; k < FColumns[i].Length; k++)
			        {
                        TColInfo Col = FColumns[i][k];
			        	if (Col!=null) 
				        {
					        int GutLevel=Col.GetColOutlineLevel();
					        if (GutLevel>MaxGutLevel) MaxGutLevel=GutLevel;
				        }
			        }
                }
            }
		    Guts.ColLevel=MaxGutLevel;
		}

        internal void ChangeStandardWidth(int OldStdValue, int NewStdValue, bool BeforeWas256)
		{
			for (int i = 0; i < FColumns.Length; i++)
			{
				if (FColumns[i] != null) 
                {
                    for (int k = 0; k < FColumns[i].Length; k++)
			        {
                        TColInfo Col = FColumns[i][k];
                        if (Col != null) Col.ChangeStandardWidth(OldStdValue, NewStdValue, BeforeWas256);		 
                    }
                }
			}
		}

        internal void SaveStandardWidth(IDataStream DataStream)
        {
            if (!AllowStandardWidth || DefColWidthChars256 < 0) return;
            DataStream.WriteHeader((UInt16)xlr.STANDARDWIDTH, 2);
            DataStream.Write16((UInt16)DefColWidthChars256);
        }
    }
}

