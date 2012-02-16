using System;
using System.Text;
using System.Diagnostics;
using System.Collections.Generic;

using FlexCel.Core;

#if (MONOTOUCH)
using Color = MonoTouch.UIKit.UIColor;
#endif

#if (WPF)
using RectangleF = System.Windows.Rect;
using Rectangle = System.Windows.Rect;
using System.Windows.Media;
#else
using System.Drawing;
#endif

namespace FlexCel.XlsAdapter
{
	#region Base
	/// <summary>
	/// Base for all records on charts.
	/// </summary>
	internal abstract class TChartBaseRecord: TxBaseRecord  //This should descend from TBaseRecord, but we will leave it like this. If not, TxChartBaseRecord won't be able to descend from both TxBaseRecord and TChartBaseRecord.
	{
		protected TChartRecordList FChildren;

		internal TChartBaseRecord(int aId) : base (aId, null)
        {
        }

		internal TChartRecordList Children {get{return FChildren;}}

		internal void CreateChildren(TChartCache MasterCache)
		{
			FChildren = new TChartRecordList(MasterCache);
		}

		internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
		{
			base.SaveToStream(Workbook, SaveData, Row);
			if (FChildren != null)
			{
				int aCount = FChildren.Count;
				for (int i = 0; i < aCount; i++)
				{
					FChildren[i].SaveToStream(Workbook, SaveData, Row);
				}
			}
		}

		protected override TBaseRecord DoCopyTo(TSheetInfo SheetInfo)
		{
            if (FChildren != null)
            {
                XlsMessages.ThrowException(XlsErr.ErrInternal); //We should use CopyTo(AiCache, SeriesCache)
            }
            return base.DoCopyTo(SheetInfo);
			
		}

		internal virtual TBaseRecord CopyTo(TChartCache Cache, TSheetInfo SheetInfo)
		{
			TChartBaseRecord Result = (TChartBaseRecord) base.DoCopyTo(SheetInfo);
			if (FChildren != null)
			{
				Result.FChildren = new TChartRecordList(Cache);
				Result.FChildren.CopyFrom(FChildren, SheetInfo);
			}

			return Result;
		}

		internal override int TotalSize()
		{
			if (FChildren == null) return base.TotalSize();
			return base.TotalSize() + (int)FChildren.TotalSize;
		}

		internal override int TotalSizeNoHeaders()
		{
			if (FChildren == null) return base.TotalSizeNoHeaders();
			return base.TotalSizeNoHeaders() + (int) FChildren.TotalSizeNoHeaders;
		}

        internal T FindRec<T>() where T : TChartBaseRecord
        {
            T Result;
            if (FChildren == null) return null;
            if (FChildren.FindRec(out Result)) return Result;
            return null;
        }

		internal int FindRecPos(Type t)
		{
			if (FChildren== null) return -1;
			return FChildren.FindRec(t);
		}

		internal List<TBaseRecord> FindAllRec(Type t)
		{
			if (FChildren == null) return new List<TBaseRecord>();
			return FChildren.FindAllRec(t);
        }

        #region Utility Methods
        internal static Color ColorFromLong(int aColor)
		{
			unchecked
			{
				return ColorUtil.FromArgb(aColor & 0xFF, (aColor >> 8) & 0xFF, (aColor >> 16) & 0xFF);
			}
		}

		internal static void AdaptRTF(int Start, int Ofs, TWorkbookGlobals SourceGlobals, TWorkbookGlobals DestGlobals, byte[] SourceData, ref byte[] DestData)
		{
			if (SourceGlobals == null) return;
			Debug.Assert(SourceData.Length == DestData.Length);

			for (int i = Start + 2 ; i < SourceData.Length; i+=Ofs)
			{
				TCellList.FixCopyFont(i, SourceGlobals, DestGlobals, SourceData, DestData);
			}
		}

        #endregion
    }

    internal class TxChartBaseRecord: TChartBaseRecord
    {
        internal TxChartBaseRecord(int aId, byte[] aData)
            : base(aId)
        {
            Data = aData;
        }
    }
	#endregion

	#region Begin/End
	/// <summary>
	/// Starts a subrecord
	/// </summary>
	internal class TBeginRecord: TxChartBaseRecord
	{
		internal TBeginRecord(int aId, byte[] aData): base(aId, aData){}
	}

	/// <summary>
	/// Ends a subrecord
	/// </summary>
	internal class TEndRecord: TxChartBaseRecord
	{
		internal TEndRecord(int aId, byte[] aData): base(aId, aData){}
	}

	#endregion

	#region AI
	/// <summary>
	/// Linked series data or text
	/// </summary>
	internal class TChartAIRecord: TChartBaseRecord
	{
		private TFormulaBounds Bounds;
        private TParsedTokenList Tokens;
        internal int LinkId;
        internal int RefType;
        internal int FormatIndex;
        internal int OptionFlags;

        private TChartAIRecord(int aId)
            : base(aId)
        {
        }

		private TChartAIRecord(TNameRecordList Names, int aId, byte[] aData): this(aId)
        {
            bool HasSubtotal; bool HasAggregate;
            LinkId = aData[0];
            RefType = aData[1];
            OptionFlags = BitOps.GetWord(aData, 2);
			FormatIndex = BitOps.GetWord(aData, 4);
            Tokens = TTokenManipulator.CreateFromBiff8(Names, -1, -1, aData, 8, BitOps.GetWord(aData, 6), false, out HasSubtotal, out HasAggregate);
        }

        internal static TChartAIRecord CreateFromBiff8(TNameRecordList Names, int aId, byte[] aData)
        {
            return new TChartAIRecord(Names, aId, aData);
        }

        internal static TChartAIRecord CreateFromData(ExcelFile xls, byte LinkId, string value, TChartCache MasterCache)
        {
            TChartAIRecord Result = new TChartAIRecord((int)xlr.ChartAI);
            Result.Tokens = new TParsedTokenList(new TBaseParsedToken[0]);

            Result.LinkId = LinkId;
            Result.SetDefinition(xls, value, MasterCache);
            return Result;
        }

		internal void SetDefinition(ExcelFile xls, string value, TChartCache MasterCache)
		{
			if (FChildren != null)
			{
				for (int i = FChildren.Count - 1; i >=0; i--)
				{
					if (FChildren[i] is TChartSeriesTextRecord) FChildren.Delete(i);
				}
			}

			if (value == null)
			{
				RefType = 0;
			}
			else
				if (value.StartsWith(TBaseFormulaParser.fts(TFormulaToken.fmEQ)))
			{
				RefType = 2;
				SetFormulaDefinition(xls, value);
			}
			else
			{
				RefType = 1;
				if (FChildren == null) CreateChildren(MasterCache);
				FChildren.Add(TChartSeriesTextRecord.CreateFromData(value));
			}
		}
        
		internal override TBaseRecord CopyTo(TChartCache Cache, TSheetInfo SheetInfo)
		{
			TChartAIRecord Result = (TChartAIRecord) base.CopyTo (Cache, SheetInfo); //this copies the children.
			if (Tokens != null) Result.Tokens = Tokens.Clone();
			TWorkbookGlobals SourceGlobals = SheetInfo.SourceGlobals;
			TWorkbookGlobals DestGlobals = SheetInfo.DestGlobals;
			if (SourceGlobals != DestGlobals && SourceGlobals != null && DestGlobals != null) 
			{
				string fmt = SourceGlobals.Formats.Format(FormatIndex);
				Result.FormatIndex = DestGlobals.Formats.AddFormat(fmt);
			}

			Result.Bounds = null; //clear the cache.
			
			return Result;
        }

		#region SavetoStream
		internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
		{
            int FmlaLenWithoutArray;
			byte[] bData = TFormulaConvertInternalToBiff8.GetTokenData(SaveData.Globals.Names, Tokens, TFormulaType.Chart, out FmlaLenWithoutArray);
			Workbook.WriteHeader((UInt16) Id, (UInt16) (bData.Length + 8));
            Workbook.Write16((UInt16)(LinkId + (RefType << 8)));
			Workbook.Write16((UInt16) OptionFlags);
			Workbook.Write16((UInt16)FormatIndex);

			Workbook.Write16((UInt16)FmlaLenWithoutArray);
			Workbook.Write(bData, bData.Length);
			if (FChildren != null) FChildren.SaveToStream(Workbook, SaveData, Row);
		}

		internal override int TotalSizeNoHeaders()
		{
			int ChildrenSize = FChildren == null? 0: (int)FChildren.TotalSizeNoHeaders;
			return 8 + TTokenManipulator.TotalSizeWithArray(Tokens, TFormulaType.Chart) + ChildrenSize;
		}

		internal override int TotalSize()
		{
			int ChildrenSize = FChildren == null? 0: (int)FChildren.TotalSize;
			return 8 + TTokenManipulator.TotalSizeWithArray(Tokens, TFormulaType.Chart) + ChildrenSize  + XlsConsts.SizeOfTRecordHeader;
		}

		#endregion

        #region InsertAndCopy
        private void ArrangeTokensInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount, int CopyRowOffset, int CopyColOffset, TSheetInfo SheetInfo)
		{
			try
			{
				if (Bounds != null && CopyRowOffset == 0 && CopyColOffset == 0   //When copying we will always fix the formula.
					&& Bounds.OutBounds(CellRange, SheetInfo, aRowCount, aColCount)) return;
            
				if (Bounds == null) Bounds = new TFormulaBounds(); else Bounds.Clear();
                
				TTokenManipulator.ArrangeInsertAndCopyRange(Tokens, CellRange, -1, -1, aRowCount, aColCount, CopyRowOffset, CopyColOffset, SheetInfo, false, Bounds);
			}
			catch (ETokenException e)
			{
				XlsMessages.ThrowException(XlsErr.ErrBadChartFormula,e.Token);
			}
		}

		private void ArrangeTokensMoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
		{
			try
			{
				if (Bounds == null) Bounds = new TFormulaBounds(); else Bounds.Clear();
                
				TTokenManipulator.ArrangeMoveRange(Tokens, CellRange, -1, -1, NewRow, NewCol, SheetInfo, Bounds);
			}
			catch (ETokenException e)
			{
				XlsMessages.ThrowException(XlsErr.ErrBadChartFormula,e.Token);
			}
		}

		internal void ArrangeInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
		{
			if (Tokens.Count > 0) ArrangeTokensInsertRange(CellRange, aRowCount, aColCount, 0, 0, SheetInfo);
		}

		internal void ArrangeMoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
		{
			if (Tokens.Count > 0) ArrangeTokensMoveRange(CellRange, NewRow, NewCol, SheetInfo);
		}

		internal void ArrangeCopySheet(TSheetInfo SheetInfo)
		{
			if (Tokens.Count==0) return;
			try
			{
				TTokenManipulator.ArrangeInsertSheets(Tokens, SheetInfo);
			}
			catch (ETokenException e)
			{
				XlsMessages.ThrowException(XlsErr.ErrBadChartFormula,e.Token);
			}
		}

		//This shouldn't make sense... all ranges in charts are absolute. This is to support RelativeCharts
		//Also, when copying from one sheet to other, we need to fix the 3d references.
		internal void ArrangeCopyRange(int RowOffset, int ColOffset, TSheetInfo SheetInfo)
		{
			if (Tokens.Count>0) ArrangeTokensInsertRange(new TXlsCellRange(0,0,-1,-1), 0, 0, RowOffset, ColOffset, SheetInfo); //Sheet info doesn't have meaning on copy, except to create the Bounds cache.
		}

		internal bool HasExternRefs()
		{
			try
			{
				return TTokenManipulator.HasExternRefs(Tokens);
			}
			catch (ETokenException e)
			{
				XlsMessages.ThrowException(XlsErr.ErrBadChartFormula,e.Token);
			}
			return false;
        }
        #endregion

        #region Named Ranges
        internal void UpdateDeletedRanges(TDeletedRanges DeletedRanges)
		{
			try
			{
				TTokenManipulator.UpdateDeletedRanges(Tokens, DeletedRanges);
			}
			catch (ETokenException e)
			{
				XlsMessages.ThrowException(XlsErr.ErrBadChartFormula,e.Token);
			}
		}
		#endregion

		private static string[] GetOneRange(ExcelFile xls, TAddress[] OneRange)
		{
			int r0 = Math.Min(OneRange[0].Row, OneRange[1].Row);
			int r1 = Math.Max(OneRange[0].Row, OneRange[1].Row);
			int c0 = Math.Min(OneRange[0].Col, OneRange[1].Col);
			int c1 = Math.Max(OneRange[0].Col, OneRange[1].Col);
			int s0 = Math.Min(OneRange[0].Sheet, OneRange[1].Sheet);
			int s1 = Math.Max(OneRange[0].Sheet, OneRange[1].Sheet);
				
			string[] Result = new string [(r1 - r0 + 1) * (c1 - c0 +1) * (s1 - s0 +1)];

			int rpos = 0;
			for (int sheet = s0; sheet <= s1; sheet++)
			{
				for (int r = r0; r <= r1; r++)
				{
					for (int c = c0; c <= c1; c++)
					{
						Result[rpos] = xls.GetCellVisibleFormatDef(sheet, r, c).Format;
						rpos++;
					}

				}
			}

			return Result;
		}

		private static string[] GetCellFormats(ExcelFile xls, object Addresses)
		{
			TAddress OneCell = Addresses as TAddress;
			if (OneCell != null)
			{
				return new string[] {xls.GetCellVisibleFormatDef(OneCell.Sheet, OneCell.Row, OneCell.Col).Format};
			}

			TAddress[] OneRange = Addresses as TAddress[];
			if (OneRange != null && OneRange.Length == 2)
			{
				return GetOneRange(xls, OneRange);
			}

			TAddressList ManyRanges = Addresses as TAddressList;
			if (ManyRanges != null)
			{
				List<string> Result = new List<string>();
				for (int i =  ManyRanges.Count - 1; i >=0; i--)
				{
					string[] TmpResult = GetOneRange(xls, ManyRanges[i]);
					if (TmpResult != null && TmpResult.Length > 0)
					{
						for (int k = 0; k < TmpResult.Length; k++)
						{
							Result.Add(TmpResult[k]);
						}						
					}
				}
				if (Result.Count == 0) return null;
				return Result.ToArray();
			}


			return null;

		}

        internal object[] GetValues(int SeriesStored, XlsFile aXls, int SheetIndex, TChartRecordList Parent, bool GetFormats, out string[] Formats)
        {
            bool Dummy = true;
            return GetValues(SeriesStored, aXls, SheetIndex, Parent, GetFormats, out Formats, false, ref Dummy);
        }

		internal object[] GetValues(int SeriesStored, XlsFile aXls, int SheetIndex, TChartRecordList Parent, bool GetFormats, out string[] Formats, bool AllowSquare, ref bool SquareDown)
		{
			Formats = null;
			switch (SeriesStored)
			{
				case 1:
					TChartSeriesTextRecord SeriesText;
					if (Parent.FindRec(out SeriesText))
					{
						return new object[]{SeriesText.Text};
					}
					return null;
				case 2:
					if (Tokens.Count == 0)
					{
						return null;
					}
                    TWorkbookInfo wi = new TWorkbookInfo(aXls, SheetIndex, 0, 0, 0, 0, 0, 0, true);
					object ObjResult = TFormulaRecord.EvaluateFormula(TArrayAggregate.Instance, Tokens, wi, false);

                    object[,] ArrResult = ObjResult as Object[,];

					if (GetFormats || IsSquare(ArrResult))
					{
						object CellRange = TFormulaRecord.EvaluateFormula(null, Tokens, wi, true);
						Formats = GetCellFormats(aXls, CellRange);
					}


					if (ArrResult != null)
					{
                        int a = 0; int b = 1; 
                        if (ArrResult.GetLength(0) <= 1)
                        {
                            a = 1;
                            b = 0;
                        }
                        else if (ArrResult.GetLength(1) > 1) //square
                        {
                            if (!SquareDown) { a = 1; b = 0; }
                        }

                        if (!AllowSquare) SquareDown = a == 0;

                        int ColCount = ArrResult.GetLength(b);

						object[] Result = new object[ArrResult.GetLength(a)];
						for (int r = 0; r < ArrResult.GetLength(a); r++)
						{
							for (int c = 0; c < ArrResult.GetLength(b); c++)
							{
                                object o = a == 0? ArrResult[r,c]: ArrResult[c,r];
                                if (c == 0)
                                {
                                    Result[r] = o;
                                }
                                else
                                {
                                    int r1 = a == 0 ? r : c;
                                    int c1 = a == 0 ? c : r;
                                    Result[r] =   GetFormattedString(aXls, o, Formats, r1, c1, ColCount, a == 0)+
                                        (char)10 + (char)10 + GetFormattedString(aXls, Result[r], Formats, r1, 0, ColCount, a == 0);
                                        
                                }
							}
						}

                        if (GetFormats && ColCount > 1) Formats = new string[Result.Length]; //fix it once it has been used. If it is a square range, then individual cell formats don't make sense.
						return Result;
					}
					return new object[]{ObjResult};
			}
			return null;
		}

        private static bool IsSquare(object[,] ArrResult)
        {
            if (ArrResult == null) return false;
            bool rc = ArrResult.GetLength(0) > 1;
            bool cc = ArrResult.GetLength(1) > 1;
            return rc && cc;
        }

        private static string GetFormattedString(ExcelFile xls, object o, string[] Formats, int r, int c, int ColCount, bool direct)
        {
            string s1 = o as string;
            if (s1 != null) return s1;

            Color aColor = ColorUtil.Empty;
            int k = r * ColCount + c;
            string fmt = Formats != null && k < Formats.Length ? Formats[k] : string.Empty;
            return FlxConvert.ToString(TFlxNumberFormat.FormatValue(o, fmt, ref aColor, xls));
        }

		internal string GetDefinition(TCellList CellList)
		{
			switch (RefType)
			{
				case 1:
					TChartSeriesTextRecord SeriesText;
					if (FChildren != null && FChildren.FindRec(out SeriesText))
					{
						return SeriesText.Text;
					}
					return null;
				case 2:
					return TFormulaConvertInternalToText.AsString(Tokens, 0, 0, CellList); //No real problem here with relative references because charts always use absolute 3d ones, and can't include ranges. So, we can safely use (0,0) as (RowOfs, ColOfs)
			}
			return null;
		}
	
		internal void SetFormulaDefinition(ExcelFile xls, string Value)
		{
            TFormulaConvertTextToInternal Ps = new TFormulaConvertTextToInternal(xls, xls.ActiveSheet, true, Value, true, false, true, xls.SheetName, TFmReturnType.Ref, true);
			Ps.Parse();
			Tokens = Ps.GetTokens();
		}
	}

	#region Cache
	/// <summary>
	/// A class for grouping all the needed caches.
	/// </summary>
	internal class TChartCache
	{
		internal TChartAIRecordCache AI;
		internal TChartSeriesRecordCache Series;

		internal TChartCache(TChartAIRecordCache aAI, TChartSeriesRecordCache aSeries)
		{
			AI = aAI;
			Series = aSeries;
		}

		internal void ArrangeCopyRange(int RowOffset, int ColOffset, TSheetInfo SheetInfo)
		{
			AI.ArrangeCopyRange(RowOffset, ColOffset, SheetInfo);
		}

		internal void ArrangeInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
		{
			AI.ArrangeInsertRange(CellRange, aRowCount, aColCount, SheetInfo);
		}
		
		internal void ArrangeMoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
		{
			AI.ArrangeMoveRange(CellRange, NewRow, NewCol, SheetInfo);
		}

		internal void ArrangeCopySheet(TSheetInfo SheetInfo)
		{
			AI.ArrangeCopySheet(SheetInfo);
		}

		internal bool HasExternRefs()
		{
			return AI.HasExternRefs();
		}

		#region Named Ranges
		internal void UpdateDeletedRanges(TDeletedRanges DeletedRanges)
		{
			AI.UpdateDeletedRanges(DeletedRanges);
		}
		#endregion


	}

	/// <summary>
	/// Holds a Cache with AI records for fast finding them.
	/// </summary>
    internal class TChartAIRecordCache
    {
        protected List<TChartAIRecord> FList;

        internal TChartAIRecordCache()
        {
            FList = new List<TChartAIRecord>();
        }

        #region Generics
        internal void Add(TChartAIRecord a)
        {
            FList.Add(a);
        }

        internal void Remove(TChartAIRecord a)
        {
            FList.Remove(a);
        }

        protected void SetThis(TChartAIRecord value, int index)
        {
            FList[index] = value;
        }

        internal TChartAIRecord this[int index]
        {
            get { return FList[index]; }
            set { SetThis(value, index); }
        }

        #endregion

        internal int Count
        {
            get { return FList.Count; }
        }

        internal void ArrangeCopyRange(int RowOffset, int ColOffset, TSheetInfo SheetInfo)
        {
            for (int i = 0; i < Count; i++)
                FList[i].ArrangeCopyRange(RowOffset, ColOffset, SheetInfo);
        }

        internal void ArrangeInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
        {
            for (int i = 0; i < Count; i++)
                FList[i].ArrangeInsertRange(CellRange, aRowCount, aColCount, SheetInfo);
        }

        internal void ArrangeMoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
        {
            for (int i = 0; i < Count; i++)
                FList[i].ArrangeMoveRange(CellRange, NewRow, NewCol, SheetInfo);
        }

        internal void ArrangeCopySheet(TSheetInfo SheetInfo)
        {
            for (int i = 0; i < Count; i++)
                FList[i].ArrangeCopySheet(SheetInfo);
        }

        internal bool HasExternRefs()
        {
            for (int i = 0; i < Count; i++)
                if (FList[i].HasExternRefs()) return true;
            return false;
        }

        #region Named Ranges
        internal void UpdateDeletedRanges(TDeletedRanges DeletedRanges)
        {
            for (int i = 0; i < Count; i++)
                FList[i].UpdateDeletedRanges(DeletedRanges);
        }
        #endregion

    }


	/// <summary>
	/// Holds a Cache with Series records for fast finding them.
	/// </summary>
	internal class TChartSeriesRecordCache
	{
        protected List<TChartSeriesRecord> FList;

		internal TChartSeriesRecordCache()
		{
			FList = new List<TChartSeriesRecord>();
		}

		#region Generics
		internal void Add (TChartSeriesRecord a)
		{
			FList.Add(a);
		}

		internal void Remove (TChartSeriesRecord a)
		{
			FList.Remove(a);
		}

		protected void SetThis(TChartSeriesRecord value, int index)
		{
			FList[index]=value;
		}

		internal TChartSeriesRecord this[int index] 
		{
			get {return (TChartSeriesRecord) FList[index];} 
			set {SetThis(value, index);}
		}

		#endregion

		internal int Count
		{
			get {return FList.Count;}
		}

		#region Manipulation methods
		#endregion
	}


	#endregion
	/// <summary>
	/// Records inside a Chart.
	/// </summary>
	internal class TChartRecordList: TBaseRecordList<TBaseRecord>
	{
		private TChartCache Cache;
		
		internal TChartRecordList()
		{
			Cache = new TChartCache(new TChartAIRecordCache(), new TChartSeriesRecordCache());
		}

		internal TChartRecordList(TChartCache MasterCache)
		{
			Cache=MasterCache;
		}

		internal TChartCache GetCache()
		{
			return Cache;
		}


		internal override void OnAdd(TBaseRecord r, int index)
		{
			base.OnAdd(r, index);
			TChartAIRecord Ai=r as TChartAIRecord;
			if (Ai!=null)
			{
				Cache.AI.Add(Ai);
				return;
			}
		
			TChartSeriesRecord Sr=r as TChartSeriesRecord;
			if (Sr!=null)
			{
				Cache.Series.Add(Sr);
				return;
			}
		}

		internal override void OnDelete(TBaseRecord r, int index)
		{
			base.OnDelete (r, index);
			TChartAIRecord Ai=r as TChartAIRecord;
			if (Ai!=null)
			{
				Cache.AI.Remove(Ai);
				return;
			}
		
			TChartSeriesRecord Sr=r as TChartSeriesRecord;
			if (Sr!=null)
			{
				Cache.Series.Remove(Sr);
				return;
			}
		}

		internal void ArrangeCopyRange(int RowOffset, int ColOffset, TSheetInfo SheetInfo)
		{
			Cache.ArrangeCopyRange(RowOffset, ColOffset, SheetInfo);
		}

		internal void ArrangeInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
		{
			Cache.ArrangeInsertRange(CellRange, aRowCount, aColCount, SheetInfo);
		}

		internal void ArrangeMoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
		{
			Cache.ArrangeMoveRange(CellRange, NewRow, NewCol, SheetInfo);
		}

		internal void ArrangeCopySheet(TSheetInfo SheetInfo)
		{
			Cache.ArrangeCopySheet(SheetInfo);
		}

		internal bool HasExternRefs()
		{
			return Cache.HasExternRefs();
		}

		#region Named Ranges
		internal void UpdateDeletedRanges(TDeletedRanges DeletedRanges)
		{
			Cache.UpdateDeletedRanges(DeletedRanges);
		}
		#endregion

		protected override TBaseRecord CloneRecord(TBaseRecord br, TSheetInfo SheetInfo)
		{
			TChartBaseRecord cr = br as TChartBaseRecord;
			if (cr != null)
			{
				return cr.CopyTo(Cache, SheetInfo);
			}
			else
			{
				return TBaseRecord.Clone(br, SheetInfo);
			}

		}

        internal bool FindRec<T>(out T b) where T : TChartBaseRecord
        {
            for (int i = Count - 1; i >= 0; i--)
            {
                b = FList[i] as T;
                if (b != null) return true;

            }

            b = null;
            return false;
        }

		internal int FindRec(Type t)
		{
			for (int i = Count - 1; i >=0; i--)
			{
				if (t.IsInstanceOfType(FList[i]))
				{
					return i;
				}
			}
			return -1;
		}

		internal List<TBaseRecord> FindAllRec(Type t)
		{
			List<TBaseRecord> Result = new List<TBaseRecord>();
			for (int i = 0; i< Count; i++)
			{
				if (t.IsInstanceOfType(FList[i]))
				{
					Result.Add(FList[i]);
				}
			}
			return Result;
		}

	}
	#endregion

	#region More Chart Records
	internal class TChartFBIRecord: TxChartBaseRecord
	{
		internal TWorkbookGlobals Globals;

		internal TChartFBIRecord(int aId, byte[] aData): base(aId, aData)
		{
		}

		internal int Scale {get {return GetWord(6);}}
		internal int FontId {get {return GetWord(8);} set {SetWord(8, value);}}

		internal override TBaseRecord CopyTo(TChartCache Cache, TSheetInfo SheetInfo)
		{
			Debug.Assert(Globals != null);
			TChartFBIRecord Result = (TChartFBIRecord) base.CopyTo (Cache, SheetInfo);
			TWorkbookGlobals DestGlobals = SheetInfo.DestGlobals;
			if (DestGlobals == null) DestGlobals = Globals;
			Result.FixFontIdOnCopy(Globals, DestGlobals, SheetInfo);

			Globals = DestGlobals;
			return Result;
		}

		private void FixFontIdOnCopy(TWorkbookGlobals SourceGlobals, TWorkbookGlobals DestGlobals, TSheetInfo SheetInfo)
		{
			int fontIndex = FontId;
			if (fontIndex==4) fontIndex=0;  //font 4 does not exist
			if (fontIndex>4) fontIndex--;
			if ((fontIndex<0) || (fontIndex>= SourceGlobals.Fonts.Count)) fontIndex=0;

			TFontRecord NewFont = (TFontRecord)TFontRecord.Clone(SourceGlobals.Fonts[fontIndex],SheetInfo);
			NewFont.Reuse = false;
			FontId = DestGlobals.Fonts.AddNotReusableFont(NewFont);
			SourceGlobals.Fonts[fontIndex].CopiedTo = FontId;
		}

	}

	internal class TChartShtPropsRecord: TxChartBaseRecord
	{
		internal TChartShtPropsRecord(int aId, byte[] aData): base(aId, aData){}

		internal int PlotEmptyCells {get{return Data[2];}}

	}

	internal class TChartChartRecord: TxChartBaseRecord
	{
		internal TChartChartRecord(int aId, byte[] aData): base(aId, aData){}

		internal Int64 Left {get {return GetCardinal(0);} set{SetCardinal(0, value);}}
		internal Int64 Top {get {return GetCardinal(4);} set{SetCardinal(4, value);}}
		internal Int64 Width {get {return GetCardinal(8);} set{SetCardinal(8, value);}}
		internal Int64 Height {get {return GetCardinal(12);} set{SetCardinal(12, value);}}

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            TFlxChart ch = ws as TFlxChart;
            if (ch == null) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);

            ch.Chart.LoadFromStream(RecordLoader, this);
        }

	}

	internal class TChartObjectLinkRecord: TxChartBaseRecord
	{
		internal TChartObjectLinkRecord(int aId, byte[] aData): base(aId, aData){}

		internal int LinkType {get{return GetWord(0);}}
		internal int SeriesIndex {get{return GetWord(2);}}
		internal int DataPointIndex {
			get
			{
				unchecked
				{
					return (Int16) GetWord(4);
				}
			}
		}
	}

	internal class TChartALRunsRecord: TxChartBaseRecord
	{
		internal TChartALRunsRecord (int aId, byte[] aData): base(aId, aData){}

		internal TRTFRun[] Runs()
		{
			int Len = GetWord(0);
			if (Len <= 0) return null;
			TRTFRun[] Result = new TRTFRun[Len];

			int k = 2;
			for (int i = 0; i < Result.Length; i++)
			{
				Result[i].FirstChar = GetWord(k); k+=2;
				Result[i].FontIndex = GetWord(k); k+=2;
			}
			return Result;

		}

		internal override TBaseRecord CopyTo(TChartCache Cache, TSheetInfo SheetInfo)
		{
			TChartALRunsRecord Result = (TChartALRunsRecord) base.CopyTo (Cache, SheetInfo);
			TWorkbookGlobals SourceGlobals = SheetInfo.SourceGlobals;
			TWorkbookGlobals DestGlobals = SheetInfo.DestGlobals;
			AdaptRTF(2, 4, SourceGlobals, DestGlobals, Data, ref Result.Data);
			return Result;
		}

	}

	internal class TChartFrameRecord: TxChartBaseRecord
	{
		internal TChartFrameRecord(int aId, byte[] aData): base(aId, aData){}

		internal TChartFrameOptions GetFrameOptions()
		{
			return GetFrameOptions(Children);
		}
		
		internal static TChartFrameOptions GetFrameOptions(TChartRecordList Children)
		{
			TChartFrameOptions Result = new TChartFrameOptions(null, null, null);
			if (Children == null) return Result;

			for (int i = 0; i < Children.Count; i++)
			{
				TxChartBaseRecord CR = (Children[i] as TxChartBaseRecord);
				if (CR == null) continue;
				switch (CR.Id)
				{
					case (int) xlr.ChartLineformat:
						TChartLineFormatRecord Lf = CR as TChartLineFormatRecord;
						Result.LineOptions = Lf.GetLineFormat();
						break;
					case (int) xlr.ChartAreaformat:
						TChartAreaFormatRecord Af = CR as TChartAreaFormatRecord;
						Result.FillOptions = new ChartFillOptions(ColorFromLong(Af.ForeColor), ColorFromLong(Af.BackColor), (TChartPatternStyle)Af.Pattern);
						break;

					case (int) xlr.ChartGelframe:
						if (Result.ExtraOptions == null)
						{
							TChartGelFrameRecord Gf = CR as TChartGelFrameRecord;
							Result.ExtraOptions = Gf.ShapeProperties(Children, i + 1);
						}
						break;

				}
			}
			return Result;
		}
	}

	internal class TChartPlotGrowthRecord: TxChartBaseRecord
	{
		internal TChartPlotGrowthRecord(int aId, byte[] aData): base(aId, aData){}

		internal long XScaling {get{return GetCardinal(0);}}
		internal long YScaling {get{return GetCardinal(4);}}
	}

	internal class TChartTextRecord: TxChartBaseRecord
	{
		internal TChartTextRecord(int aId, byte[] aData): base(aId, aData){}

		internal TChartTextOptions GetTextBaseOptions()
		{
			TChartTextOptions Result = new TChartTextOptions();
			switch (Data[0])
			{
				case 2: Result.HAlign = THFlxAlignment.center;break;
				case 3: Result.HAlign = THFlxAlignment.right;break;
				case 4: Result.HAlign = THFlxAlignment.justify;break;
				default: Result.HAlign = THFlxAlignment.left;break;
			}

			switch (Data[1])
			{
				case 2: Result.VAlign = TVFlxAlignment.center;break;
				case 3: Result.VAlign = TVFlxAlignment.bottom;break;
				case 4: Result.VAlign = TVFlxAlignment.justify;break;
				default: Result.VAlign = TVFlxAlignment.top;break;
			}

			Result.BackgoundMode = (TBackgroundMode) Data[2];

			TChartPatternStyle Pattern = TChartPatternStyle.Automatic;
			Result.TextColor = new ChartFillOptions(ColorFromLong((Int32)GetCardinal(4)), ColorUtil.Empty, Pattern);

            TChartPosRecord ChPos = (TChartPosRecord)FindRec<TChartPosRecord>();
            if (ChPos != null)
            {
                Result.Position = new TChartLabelPosition(ChPos.TopLeftMode, ChPos.BottomRightMode, ChPos.X1, ChPos.Y1, ChPos.X2, ChPos.Y2);
            }
            
                unchecked
                {
                    Result.X = (Int32)GetCardinal(8); if (Result.X < 0) Result.X = 0; if (Result.X > 4000) Result.X = 4000;
                    Result.Y = (Int32)GetCardinal(12); if (Result.Y < 0) Result.Y = 0; if (Result.Y > 4000) Result.Y = 4000;
                    Result.Width = GetWord(16); if (Result.Width < 0) Result.Width = 0; if (Result.Width > 4000) Result.Width = 4000;
                    Result.Height = GetWord(20); if (Result.Height < 0) Result.Height = 0; if (Result.Height > 4000) Result.Height = 4000;
                }
            

			Result.Rotation = Data[30];

			return Result;
		}

		internal TChartTextOptions GetTextOptions(TWorkbookGlobals WorkbookGlobals, double FontScale)
		{
			TChartTextOptions Result = GetTextBaseOptions();
			if (Children == null) return Result;

			for (int i = 0; i < Children.Count; i++)
			{
				TxChartBaseRecord CR = (Children[i] as TxChartBaseRecord);
				if (CR == null) continue;
				switch (CR.Id)
				{
					case (int) xlr.ChartFontx:
						TChartFontXRecord FontX = CR as TChartFontXRecord;
						Result.Font = FontX.GetFont(WorkbookGlobals, FontScale);
						break;
				}
			}
			return Result;
		}

		private void AddLabelOptions(TDataLabel aLabel)
		{
			aLabel.LabelOptions = new TDataLabelOptions();
			TDataLabelOptions opt = aLabel.LabelOptions;

			int flags1 = GetWord(24);

			opt.AutoColor = (flags1 & 0x1) != 0;
			opt.ShowLegendKey = (flags1 & 0x2) != 0;
			if ((flags1 & 0x10) != 0)
			{
				opt.DataType = TLabelDataValue.SeriesInfo;
			}
			else
				opt.DataType = TLabelDataValue.Manual;

			opt.ShowValues = (flags1 & 0x4) != 0;

			opt.ShowCategories = (flags1 & 0x4000) != 0 || ((flags1 & 0x800) != 0);
			opt.ShowPercents = (flags1 & 0x1000) != 0 ;
			opt.ShowBubbles = (flags1 & 0x2000) != 0;

			opt.Deleted = ((flags1 & 0x40) != 0);

			int flags2 = GetWord(28);
			opt.Position = (TDataLabelPosition) (flags2 & 0x0F);

		}

		internal TDataLabel GetDataLabel(XlsFile xls, TCellList CellList, int SheetIndex, bool GetDefinitions, bool GetValues, double FontScale)
		{
			TDataLabel Result = new TDataLabel();
			Result.TextOptions = GetTextBaseOptions();
			AddLabelOptions(Result);

			if (Children == null) return Result;

			TRTFRun[] RTFRuns = null;
			for (int i = 0; i < Children.Count; i++)
			{
				TChartBaseRecord CR = (Children[i] as TChartBaseRecord);
				if (CR == null) continue;
				switch ((xlr)CR.Id)
				{
					case xlr.ChartFontx:
						TChartFontXRecord FontX = CR as TChartFontXRecord;
						Result.TextOptions.Font = FontX.GetFont(xls.InternalWorkbook.Globals, FontScale);
						break;

					case  xlr.ChartFrame:
						TChartFrameRecord FR = CR as TChartFrameRecord;
						Result.Frame = FR.GetFrameOptions();
						break;

					case xlr.ChartObjectLink:
						TChartObjectLinkRecord OL = CR as TChartObjectLinkRecord;
						Result.LinkedTo = (TLinkOption) OL.LinkType;
						Result.SeriesIndex = OL.SeriesIndex;
						Result.DataPointIndex = OL.DataPointIndex;
						break;

                    case xlr.ChartDataLabExtContent:
                        TChartDataLabExtContentsRecord DLext = CR as TChartDataLabExtContentsRecord;
                        Result.LabelOptions.ShowValues = DLext.ShowValues;
                        Result.LabelOptions.ShowSeriesName = DLext.ShowSeriesName;
                        Result.LabelOptions.ShowCategories = DLext.ShowCategories;
                        Result.LabelOptions.ShowPercents = DLext.ShowPercents;
                        Result.LabelOptions.ShowBubbles = DLext.ShowBubbles;
                        Result.LabelOptions.Separator = DLext.Separator;
                        break;

					case xlr.ChartAI:
						TChartAIRecord AI = CR as TChartAIRecord;
						if (AI != null && AI.LinkId == 0)
						{
							if (GetDefinitions)
							{
								Result.LabelDefinition = AI.GetDefinition(CellList);
							}

							if (GetValues && Result.LabelOptions.DataType == TLabelDataValue.Manual)
							{
								string[] TmpFmt = null;
								Result.LabelValues = AI.GetValues(AI.RefType, xls, SheetIndex, FChildren, false, out TmpFmt);
							}
						}

                        if ((AI.OptionFlags & 0x01) != 0)
                        {
                            Result.NumberFormat = xls.InternalWorkbook.Globals.Formats.Format(AI.FormatIndex);
                        }
						break;

					case xlr.ChartAlruns:
						TChartALRunsRecord ALRuns = (TChartALRunsRecord) CR;
						RTFRuns = ALRuns.Runs();
						break;
				}
			}

			if (RTFRuns != null && RTFRuns.Length > 0 && Result.LabelValues != null && Result.LabelOptions.DataType == TLabelDataValue.Manual)
			{
				for (int i = 0; i < Result.LabelValues.Length; i++)
				{
					string lbl = Result.LabelValues[i] as string;
					if (lbl != null)
					{
						Result.LabelValues[i] = new TRichString(lbl, RTFRuns, xls);
					}
				}

			}

			return Result;

		}


        internal int SeriesIndex() 
        {
            TChartObjectLinkRecord oLink = FindRec<TChartObjectLinkRecord>() as TChartObjectLinkRecord;
            if (oLink == null) return -1;
            return oLink.SeriesIndex;
        }
    }

    internal class TChartDataLabExtContentsRecord : TxChartBaseRecord
    {
		internal TChartDataLabExtContentsRecord(int aId, byte[] aData): base(aId, aData){}

        internal bool ShowSeriesName { get { return Data.Length > 12 && ((Data[12] & 0x01) != 0); } }
        internal bool ShowCategories { get { return Data.Length > 12 && ((Data[12] & 0x02) != 0); } }
        internal bool ShowValues { get { return Data.Length > 12 && ((Data[12] & 0x04) != 0); } }
        internal bool ShowPercents { get { return Data.Length > 12 && ((Data[12] & 0x08) != 0); } }
        internal bool ShowBubbles { get { return Data.Length > 12 && ((Data[12] & 0x10) != 0); } }

        internal string Separator
        {
            get
            {
                if (Data.Length <= 16) return null;
                string Result = null;
                long StSize = 0;
                StrOps.GetSimpleString(true, Data, 14, false, 0, ref Result, ref StSize);
                return Result;
            }
        }
    }

	internal class TChartFontXRecord: TxChartBaseRecord
	{
		internal TChartFontXRecord(int aId, byte[] aData): base(aId, aData){}
		internal int FontIndex {get{return GetWord(0);}}

		internal TFlxChartFont GetFont(TWorkbookGlobals WorkbookGlobals, double FontScaling)
		{
			int fi = FontIndex;
			if (fi < 0 || fi > WorkbookGlobals.Fonts.Count) fi = 0;

			TFlxChartFont Result =new TFlxChartFont();
			Result.Font = WorkbookGlobals.Fonts.GetFont(fi);
		

			//ChartFBI does not actually affect the size of the fonts, it is used only for resizing existing ones.
			/*if (Fbi == null || Fbi[fi] == null) Result.Scale = 1; else
			{
				float xScale = ((TChartFBIRecord)Fbi[fi]).Scale;
				if (xScale > 0) Result.Scale = xScale; else Result.Scale = FontScaling;
			}
            */
			Result.Scale = 1;
			return Result;
		}

		internal override TBaseRecord CopyTo(TChartCache Cache, TSheetInfo SheetInfo)
		{
			TChartFontXRecord Result = (TChartFontXRecord)base.CopyTo (Cache, SheetInfo);
			TWorkbookGlobals SourceGlobals = SheetInfo.SourceGlobals;
			TWorkbookGlobals DestGlobals = SheetInfo.DestGlobals;

			TCellList.FixCopyFont(0, SourceGlobals, DestGlobals, Data, Result.Data);
			return Result;
		}

	}

	internal class TChartIFmtRecord: TxChartBaseRecord
	{
		internal TChartIFmtRecord(int aId, byte[] aData): base(aId, aData){}
		internal int FormatIndex {get{return GetWord(0);}}

		internal override TBaseRecord CopyTo(TChartCache Cache, TSheetInfo SheetInfo)
		{
			TChartIFmtRecord Result = (TChartIFmtRecord)base.CopyTo (Cache, SheetInfo);
			TWorkbookGlobals SourceGlobals = SheetInfo.SourceGlobals;
			TWorkbookGlobals DestGlobals = SheetInfo.DestGlobals;
			if (SourceGlobals != DestGlobals && SourceGlobals != null && DestGlobals != null)
			{
				string fmt = SourceGlobals.Formats.Format(GetWord(0));
				Result.SetWord(0, DestGlobals.Formats.AddFormat(fmt));
			}
			return Result;
		}

	}

	internal class TChartSeriesTextRecord: TxChartBaseRecord
	{
		internal TChartSeriesTextRecord(int aId, byte[] aData): base(aId, aData){}
		
		internal static TChartSeriesTextRecord CreateFromData(string value)
		{
			TExcelString Xs= new TExcelString(TStrLenLength.is8bits, value, null, true);
			
			TChartSeriesTextRecord Result = new TChartSeriesTextRecord((int)xlr.ChartSeriestext, new byte[Xs.TotalSize() + 2]);
			Xs.CopyToPtr(Result.Data, 2);
			return Result;
		}

		internal string Text
		{
			get
			{
				TxBaseRecord MyRecord = this;
				int Ofs = 2;
				TExcelString Xs = new TExcelString(TStrLenLength.is8bits, ref MyRecord, ref Ofs);
				return Xs.Data;
			}
		}
	}

	internal class TChartDefaultTextRecord: TxChartBaseRecord
	{
		internal TChartDefaultTextRecord(int aId, byte[] aData): base(aId, aData){}

		internal int AppliesTo{get{return GetWord(0);}}
	}

	internal class TChartPosRecord: TxChartBaseRecord
	{
		internal TChartPosRecord(int aId, byte[] aData): base(aId, aData){}

        internal TChartLabelPositionMode TopLeftMode { get { return (TChartLabelPositionMode)GetWord(0); } set { SetWord(0, (int)value); } }
        internal TChartLabelPositionMode BottomRightMode { get { return (TChartLabelPositionMode)GetWord(2); } set { SetWord(2, (int)value); } }

        internal int X1 { get { return GetInt16(4); } }
        internal int Y1 { get { return GetInt16(8); } }
        internal int X2 { get { return GetInt16(12); } }
        internal int Y2 { get { return GetInt16(16); } }
	}

	internal class TChartAxisRecord: TxChartBaseRecord
	{
		internal TChartAxisRecord(int aId, byte[] aData): base(aId, aData){}

		internal int AxisType {get {return GetWord(0);} set{SetWord(0, value);}}

	}

	internal class TChartAxcExtRecord: TxChartBaseRecord
	{
		internal TChartAxcExtRecord(int aId, byte[] aData): base(aId, aData){}

		internal int Min {get {return GetWord(0);} set{SetWord(0, value);}}
		internal int Max {get {return GetWord(2);} set{SetWord(2, value);}}
		internal int MajorValue {get {return GetWord(4);} set{SetWord(4, value);}}
		internal int MajorUnits {get {return GetWord(6);} set{SetWord(6, value);}}
		internal int MinorValue {get {return GetWord(8);} set{SetWord(8, value);}}
		internal int MinorUnits {get {return GetWord(10);} set{SetWord(10, value);}}
		internal int BaseUnits {get {return GetWord(12);} set{SetWord(12, value);}}

		internal int CrossValueDate {get {return GetWord(14);} set{SetWord(14, value);}}

		internal int AxisOptions {get {return GetWord(16);} set{SetWord(16, value);}}

	}	
	
	internal class TChartValueRangeRecord: TxChartBaseRecord
	{
		internal TChartValueRangeRecord(int aId, byte[] aData): base(aId, aData){}

		internal double Min {get {return BitConverter.ToDouble(Data, 0);} }
		internal double Max {get {return BitConverter.ToDouble(Data, 8);} }
		internal double Major {get {return BitConverter.ToDouble(Data, 16);} }
		internal double Minor {get {return BitConverter.ToDouble(Data, 24);} }
		internal double CrossValue {get {return BitConverter.ToDouble(Data, 32);} }

		internal int AxisOptions {get {return GetWord(40);} set{SetWord(40, value);}}

	}

	internal class TChartAxisParentRecord: TxChartBaseRecord
	{
		internal TChartAxisParentRecord(int aId, byte[] aData): base(aId, aData){}

		internal int Index
		{
			get
			{
				return GetWord(0);
			}
		}
		internal Rectangle Rect
		{
			get
			{
				unchecked
				{
					return new Rectangle
						(
						(int)GetCardinal(2), (int)GetCardinal(6), (int)GetCardinal(10), (int)GetCardinal(14)
						);
				}
			}
		}
	}

	internal class TChartChartFormatRecord: TxChartBaseRecord
	{
		internal TChartChartFormatRecord(int aId, byte[] aData): base(aId, aData){}

		internal bool ChangeColorsOnEachSeries
		{
			get
			{
				return (GetWord(16) & 0x1) == 0x1;
			}
		}

		internal int ZOrder
		{
			get
			{
				return GetWord(18);
			}
		}
	}

	internal class TChartLegendRecord: TxChartBaseRecord
	{
		internal TChartLegendRecord(int aId, byte[] aData): base(aId, aData){}

		internal long xPos{get{return GetCardinal(0);}}
		internal long yPos{get{return GetCardinal(4);}}
		internal long xSize{get{return GetCardinal(8);}}
		internal long ySize{get{return GetCardinal(12);}}

		internal int LegendType{get{return GetWord(16);}}
		internal int Flags {get{return GetWord(18);}}
	}

	internal class TChartLegendXnRecord: TxChartBaseRecord
	{
		internal TChartLegendXnRecord(int aId, byte[] aData): base(aId, aData){}

		internal bool EntryDeleted{get{return (GetWord(2) & 0x1) != 0;}}
		internal bool EntryFormatted{get{return (GetWord(2) & 0x2) != 0;}}

		internal int SeriesId
		{
			get
			{
				unchecked
				{
					return (Int16) GetWord(0);
				}	
			}
		}

		internal TLegendEntryOptions GetLegendOptions(TWorkbookGlobals WorkbookGlobals, double FontScale)
		{
			TChartTextOptions TextFormat = null;
			if (EntryFormatted)
			{
				TChartTextRecord Tr;
				if (Children.FindRec(out Tr))
				{
					TextFormat = Tr.GetTextOptions(WorkbookGlobals, FontScale);
				}
			}

			return new TLegendEntryOptions(EntryDeleted, TextFormat);
		}
	}


	internal class TChartDataFormatRecord: TxChartBaseRecord
	{
		internal TChartDataFormatRecord(int aId, byte[] aData): base(aId, aData){}
	
		/// <summary>
		/// ffff means this applies to all the series
		/// </summary>
		internal int PointNumber
		{
			get
			{
				int Result = GetWord(0);
				if (Result == 0xFFFF) return -1;
				return Result;
			}
		}

		internal int SeriesIndex
		{
			get
			{
				return GetWord(2);
			}
		}

		internal int SeriesNumber
		{
			get
			{
				return GetWord(4);
			}
		}
	}

	internal class TChartAreaFormatRecord: TxChartBaseRecord
	{
		internal TChartAreaFormatRecord(int aId, byte[] aData): base(aId, aData){}

		internal int ForeColor
		{
			get
			{
				return (Int32)GetCardinal(0);
			}
		}

		internal int BackColor
		{
			get
			{
				return (Int32)GetCardinal(4);
			}
		}

		internal int Pattern
		{
			get
			{
				return GetWord(8);
			}
		}

		internal int Flags
		{
			get
			{
				return GetWord(10);
			}
		}

		internal int ForeColorIndex
		{
			get
			{
				return GetWord(12);
			}
		}

		internal int BackColorIndex
		{
			get
			{
				return GetWord(14);
			}
		}

	}

	internal class TChartLineFormatRecord: TxChartBaseRecord
	{
		internal TChartLineFormatRecord(int aId, byte[] aData): base(aId, aData){}

		internal int LineColor
		{
			get
			{
				return (Int32)GetCardinal(0);
			}
		}

		internal TChartLineStyle LineStyle
		{
			get
			{
				return (TChartLineStyle)GetWord(4);
			}
		}

		internal int Weight
		{
			get
			{
				return GetWord(6);
			}
		}

		internal int Flags
		{
			get
			{
				return GetWord(8);
			}
		}

		internal int LineColorIndex
		{
			get
			{
				return GetWord(10);
			}
		}

		internal ChartLineOptions GetLineFormat()
		{
			return new ChartLineOptions(ColorFromLong(LineColor), LineStyle, (TChartLineWeight)Weight);
		}
	}

	internal class TChartAxisLineFormatRecord: TxChartBaseRecord
	{
		internal TChartAxisLineFormatRecord(int aId, byte[] aData): base(aId, aData){}

		internal int AxisType
		{
			get
			{
				return GetWord(0);
			}
		}
	}

	internal class TChartTickRecord: TxChartBaseRecord
	{
		internal TChartTickRecord(int aId, byte[] aData): base(aId, aData){}

		internal byte MajorType{get{return Data[0];}}
		internal byte MinorType{get{return Data[1];}}
		internal byte LabelPosition{get{return Data[2];}}
		internal byte BackgroundMode{get{return Data[3];}}
		internal long LabelColor{get{return GetCardinal(4);}}

		internal int Rotation{get{return GetWord(28);}} //not documented.

	}

	internal class TChartAttachedLabelRecord: TxChartBaseRecord
	{
		internal TChartAttachedLabelRecord(int aId, byte[] aData): base(aId, aData){}

		internal bool ShowValue {get{return (GetWord(0) & 0x1) != 0; }}
		internal bool ShowPercent {get{return (GetWord(0) & 0x2) != 0; }}
		internal bool ShowCategory {get{return (GetWord(0) & 0x10) != 0; }}
		internal bool ShowBubbleSizes {get{return (GetWord(0) & 0x20) != 0; }}

	}

	internal class TChartCatSerRangeRecord: TxChartBaseRecord
	{
		internal TChartCatSerRangeRecord(int aId, byte[] aData): base(aId, aData){}

		internal int CatCross{get{return GetWord(0);}}
		internal int LabelFrequency{get{return GetWord(2);}}
		internal int TickFrequency{get{return GetWord(4);}}
		internal int Flags{get{return GetWord(6);}}

	}

	internal class TChartPieFormatRecord: TxChartBaseRecord
	{
		internal TChartPieFormatRecord(int aId, byte[] aData): base(aId, aData){}

		internal int PercentSliceDistance
		{
			get
			{
				return GetWord(0);
			}
		}
	}

	internal class TChartMarkerFormatRecord: TxChartBaseRecord
	{
		internal TChartMarkerFormatRecord(int aId, byte[] aData): base(aId, aData){}

		internal int ForeColor
		{
			get
			{
				return (Int32)GetCardinal(0);
			}
		}

		internal int BackColor
		{
			get
			{
				return (Int32)GetCardinal(4);
			}
		}

		internal int MarkerType
		{
			get
			{
				return GetWord(8);
			}
		}

		internal int Flags
		{
			get
			{
				return GetWord(10);
			}
		}

		internal int ForeColorIndex
		{
			get
			{
				return GetWord(12);
			}
		}

		internal int BackColorIndex
		{
			get
			{
				return GetWord(14);
			}
		}

		internal int MarkerSize
		{
			get
			{
				return GetWord(16);
			}
		}

	}


	internal class TChartSerFmtRecord: TxChartBaseRecord
	{
		internal TChartSerFmtRecord(int aId, byte[] aData): base(aId, aData){}

		internal int Flags
		{
			get
			{
				return GetWord(0);
			}
		}
	}


	internal class TChartGelFrameRecord: TxChartBaseRecord
	{
		internal TChartGelFrameRecord(int aId, byte[] aData): base(aId, aData){}

		internal TShapeOptionList ShapeProperties(TChartRecordList NextRecords, int Next)
		{
			byte[] RecordData = new byte[XlsEscherConsts.SizeOfTEscherRecordHeader];
			Array.Copy(Data, 0, RecordData, 0, 8);
			TEscherOPTRecord Opt = new TEscherOPTRecord(new TEscherRecordHeader(RecordData), null, null, null);

			TxBaseRecord MyRecord = this;
			int aPos = XlsEscherConsts.SizeOfTEscherRecordHeader;

			while (!Opt.Loaded())
			{
				Opt.Load(ref MyRecord, ref aPos, false, true);
				if (!Opt.Loaded())
				{
					if (Next >= NextRecords.Count) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
					MyRecord = NextRecords[Next] as TxBaseRecord;
					if (MyRecord == null)XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
					Next++;
					aPos = 0; //XlsEscherConsts.SizeOfTEscherRecordHeader;  After the first continue, there are no more Headers.
				}
			}
            Opt.AfterCreate();
			return Opt.ShapeOptions();
		}
	}


	internal class TChartPlotAreaRecord: TxChartBaseRecord
	{
		internal TChartPlotAreaRecord(int aId, byte[] aData): base(aId, aData){}
	}

	internal class TChartDropBarRecord: TxChartBaseRecord
	{
		internal TChartDropBarRecord(int aId, byte[] aData): base(aId, aData){}

		internal int GapWidth
		{
			get { return GetWord(0);}
		}

		internal TChartFrameOptions GetFrame()
		{
			return TChartFrameRecord.GetFrameOptions(Children);
		}
	}

	internal class TChartChartLineRecord: TxChartBaseRecord
	{
		internal TChartChartLineRecord(int aId, byte[] aData): base(aId, aData){}

		internal bool HasDropLines
		{
			get
			{
				return GetWord(0) == 0;
			}
		}
		internal bool HasHiLoLines
		{
			get
			{
				return GetWord(0) == 1;
			}
		}
		internal bool HasSeriesLines
		{
			get
			{
				return GetWord(0) == 2;
			}
		}

	}

	#endregion

	#region Format Chart Records
	internal class TChartFormatBaseRecord: TxChartBaseRecord
	{
		internal TChartFormatBaseRecord(int aId, byte[] aData): base(aId, aData){}
	}

	internal class TChartAreaRecord: TChartFormatBaseRecord
	{
		internal TChartAreaRecord(int aId, byte[] aData): base(aId, aData){}
		internal int Flags { get { return GetWord(0); } set { SetWord(0, value); } }
	}

	internal class TChartBarRecord: TChartFormatBaseRecord
	{
		internal TChartBarRecord(int aId, byte[] aData): base(aId, aData){}

		internal int BarOverlap	{get {return GetInt16(0);}	set	{SetWord(0, value);}}
		internal int CategoriesOverlap	{get {return GetInt16(2);}	set	{SetWord(2, value);}}
		internal int Flags	{get {return GetWord(4);}	set	{SetWord(4, value);}}
	}

	internal class TChartLineRecord: TChartFormatBaseRecord
	{
		internal TChartLineRecord(int aId, byte[] aData): base(aId, aData){}
        internal int Flags { get { return GetWord(0); } set { SetWord(0, value); } }
    }

	internal class TChartPieRecord: TChartFormatBaseRecord
	{
		internal TChartPieRecord(int aId, byte[] aData): base(aId, aData){}

		internal int FirstSliceAngle { get { return GetWord(0); } set { SetWord(0, value); } }
		internal int DonutRadius { get { return GetWord(2); } set { SetWord(2, value); } }
		internal int Flags { get { return GetWord(4); } set { SetWord(4, value); } }
	}

	internal class TChartRadarRecord: TChartFormatBaseRecord
	{
		internal TChartRadarRecord(int aId, byte[] aData): base(aId, aData){}
	}

	internal class TChartScatterRecord: TChartFormatBaseRecord
	{
		internal TChartScatterRecord(int aId, byte[] aData): base(aId, aData){}
		
		internal int PercentOfLargestBubble { get { return GetWord(0); } set { SetWord(0, value); } }
		internal int BubbleSize { get { return GetWord(2); } set { SetWord(2, value); } }
		internal int Flags { get { return GetWord(4); } set { SetWord(4, value); } }
	}

	internal class TChartSurfaceRecord: TChartFormatBaseRecord
	{
		internal TChartSurfaceRecord(int aId, byte[] aData): base(aId, aData){}
	}

	#endregion

	#region Series
	internal class TChartSeriesRecord: TxChartBaseRecord
	{
		internal TChartSeriesRecord(int aId, byte[] aData): base(aId, aData){}

		internal static TChartSeriesRecord CreateFromData(XlsFile xls, ChartSeries Value, int SheetIndex, TChartCache MasterCache)
		{
			TChartSeriesRecord Result = new TChartSeriesRecord((int) xlr.ChartSeries, new byte[12]);
			Result.CreateChildren(MasterCache);
			Result.FChildren.Add(new TBeginRecord((int)xlr.BEGIN, new byte[0]));
			Result.FChildren.Add(TChartAIRecord.CreateFromData(xls, 0, Value.TitleDefinition, MasterCache));
			
			TChartAIRecord YAi = TChartAIRecord.CreateFromData(xls, 1, Value.DataDefinition, MasterCache); Result.FChildren.Add(YAi);
			TChartAIRecord XAi = TChartAIRecord.CreateFromData(xls, 2, Value.CategoriesDefinition, MasterCache); Result.FChildren.Add(XAi);

            if (Value.ChartOptionsIndex > 0 && Value.ChartOptionsIndex < UInt16.MaxValue)
            {
                Result.FChildren.Add(new TxChartBaseRecord((int)xlr.ChartSertocrt, 
                    BitConverter.GetBytes((UInt16)Value.ChartOptionsIndex)));
            }
            Result.FChildren.Add(new TEndRecord((int)xlr.END, new byte[0]));

			string[] TmpFmt;
			object[] XValues = XAi.GetValues(XAi.RefType, xls, SheetIndex, Result.Children, false, out TmpFmt);
			object[] YValues = YAi.GetValues(YAi.RefType, xls, SheetIndex, Result.Children, false, out TmpFmt);
			
			bool XHasNumbers = XValues != null;
			if (XValues != null)
			{
				foreach (object o in XValues)
				{
					if (!(o is double))
					{
						XHasNumbers = false;
						break;
					}
				}
			}

			Result.XDataType = XHasNumbers? 1: 3;
			Result.YDataType = 1;
			Result.XDataCount = XValues == null? 0: XValues.Length;
			Result.YDataCount = YValues == null? 0: YValues.Length;
			Result.BubbleSize = 1;

			return Result;
		}

		internal int XDataType {get{return GetWord(0);} set {SetWord(0, value);}}
		internal int YDataType {get{return GetWord(2);} set {SetWord(2, value);}} //This one must always be 1.
		internal int XDataCount {get{return GetWord(4);} set {SetWord(4, value);}}
		internal int YDataCount {get{return GetWord(6);} set {SetWord(6, value);}} 
		internal int BubbleSize {get{return GetWord(8);} set {SetWord(8, value);}}//This one must always be 1. 

		internal static TChartPlotArea GetPlotArea(TChartAxisParentRecord AxisParent)
		{
			int PlotAreaPos = AxisParent.FindRecPos(typeof(TChartPlotAreaRecord));
			if (PlotAreaPos < 0 || PlotAreaPos >= AxisParent.Children.Count - 1) return null;

			TChartFrameRecord FR = AxisParent.Children[PlotAreaPos + 1] as TChartFrameRecord;
			if (FR == null) return null;

			return new TChartPlotArea(FR.GetFrameOptions());
		}

		internal static ChartSeriesOptions GetSeriesOptions(TChartDataFormatRecord DataFormat)
		{
			if (DataFormat == null) return null;
			ChartSeriesOptions Result = new ChartSeriesOptions(DataFormat.PointNumber, null, null, null, null, null);
			if (DataFormat.Children == null) return Result;

			for (int i = 0; i < DataFormat.Children.Count; i++)
			{
				TxChartBaseRecord CR = (DataFormat.Children[i] as TxChartBaseRecord);
				if (CR == null) continue;
				switch (CR.Id)
				{
					case (int) xlr.ChartLineformat:
						TChartLineFormatRecord Lf = CR as TChartLineFormatRecord;
						Result.LineOptions = new ChartSeriesLineOptions(ColorFromLong(Lf.LineColor), Lf.LineStyle, (TChartLineWeight)Lf.Weight, (Lf.Flags & 0x01) != 0);
						break;
					case (int) xlr.ChartAreaformat:
						TChartAreaFormatRecord Af = CR as TChartAreaFormatRecord;
						Result.FillOptions = new ChartSeriesFillOptions(ColorFromLong(Af.ForeColor), ColorFromLong(Af.BackColor), (TChartPatternStyle)Af.Pattern, (Af.Flags & 0x01) != 0, (Af.Flags & 0x02) != 0);
						break;
						
					case (int) xlr.ChartPieformat:
						TChartPieFormatRecord Pf = CR as TChartPieFormatRecord;
						Result.PieOptions = new ChartSeriesPieOptions(Pf.PercentSliceDistance);
						break;
					case (int) xlr.ChartMarkerformat:
						TChartMarkerFormatRecord Mf = CR as TChartMarkerFormatRecord;
						Result.MarkerOptions = new ChartSeriesMarkerOptions(ColorFromLong(Mf.ForeColor), ColorFromLong(Mf.BackColor), (TChartMarkerType)Mf.MarkerType, (Mf.Flags & 0x01) != 0, (Mf.Flags & 0x010) != 0, (Mf.Flags & 0x20) != 0, Mf.MarkerSize);
						break;
					case (int) xlr.ChartSerfmt:
						TChartSerFmtRecord Sf = CR as TChartSerFmtRecord;
						Result.MiscOptions = new ChartSeriesMiscOptions((Sf.Flags & 0x1) != 0, (Sf.Flags & 0x2) != 0, (Sf.Flags & 0x4) != 0);
						break;

					case (int) xlr.ChartGelframe:
						if (Result.ExtraOptions == null)
						{
							TChartGelFrameRecord Gf = CR as TChartGelFrameRecord;
							Result.ExtraOptions = Gf.ShapeProperties(DataFormat.Children, i + 1);
						}
						break;

				}
			}
			return Result;
		}

		internal ChartSeries GetValuesAndDefinition(XlsFile xls, TCellList CellList, int SheetIndex, 
			bool GetDefinitions, bool GetValues, bool GetOptions, double FontScale)
		{
			ChartSeries Result = new ChartSeries();
			if (FChildren == null) return Result;
            bool MultiCategoriesDown = true;
			for (int i = 0; i < FChildren.Count; i++) //must be done in ascending order, so values are read before categories, and we can know how to use multiple column categories.
			{
				TChartBaseRecord CR = (FChildren[i] as TChartBaseRecord);
				if (CR == null) continue;
				switch (CR.Id)
				{
					case (int)xlr.ChartAI:
						TChartAIRecord AI = CR as TChartAIRecord;
						if (AI != null)
						{
							switch (AI.LinkId)
							{
								case 0:
									if (GetDefinitions)
									{
										Result.TitleDefinition = AI.GetDefinition(CellList);
									}
									if (GetValues)
									{
										string[] TmpFmt;
										object[] R = AI.GetValues(AI.RefType, xls, SheetIndex, FChildren, true, out TmpFmt);
										if (R != null && R.Length>0) 
										{
											StringBuilder sb = new StringBuilder();
											for (int a = 0; a < R.Length; a++)
											{
												if (a > 0) sb.Append(" ");
												Color c = ColorUtil.Empty;
												string Fmt = TmpFmt == null || a >= TmpFmt.Length? String.Empty: TmpFmt[a];
												sb.Append(TFlxNumberFormat.FormatValue(R[a], Fmt, ref c, xls) );
											}
											Result.TitleValue = sb.ToString();
										}

										break;
									}
									break;
								case 1: 
									if (GetDefinitions)
									{
										Result.DataDefinition = AI.GetDefinition(CellList);
									}
									if (GetValues)
									{
										string[] TmpFmt;
										Result.DataValues = AI.GetValues(AI.RefType, xls, SheetIndex, FChildren, true, out TmpFmt, false, ref MultiCategoriesDown);
										Result.DataFormats = TmpFmt;
									}
									break;
								case 2: 
									if (GetDefinitions)
									{
										Result.CategoriesDefinition = AI.GetDefinition(CellList);
									}
									if (GetValues)
									{
										string[] TmpFmt;
										Result.CategoriesValues = AI.GetValues(AI.RefType, xls, SheetIndex, FChildren, true, out TmpFmt, true, ref MultiCategoriesDown);
										Result.CategoriesFormats = TmpFmt;
									}
									break;
							}
						}
						break;

					case (int)xlr.ChartSertocrt:
						Result.ChartOptionsIndex = BitOps.GetWord(CR.Data,0);
						break;

					case (int)xlr.ChartDataformat:
						TChartDataFormatRecord DF = CR as TChartDataFormatRecord;
						if (DF.PointNumber == -1)
						{
							Result.SeriesIndex = DF.SeriesIndex;
							Result.SeriesNumber = DF.SeriesNumber;
						}
						if (GetOptions) Result.Options.Add(GetSeriesOptions(DF));
						break;

					case (int)xlr.ChartLegendxn:
						TChartLegendXnRecord LegendXn = CR as TChartLegendXnRecord;
						if (GetOptions) 
						{
							Result.LegendOptions.Add(LegendXn.SeriesId, LegendXn.GetLegendOptions(xls.InternalWorkbook.Globals, FontScale));
						}

						break;
				}
			}
			return Result;
		}

		internal void SetDefinition(ExcelFile xls, ChartSeries Series, TChartCache MasterCache)
		{
			if (FChildren == null) CreateChildren(MasterCache);
			bool TitleSet = false;
			bool DataSet = false;
			bool CategoriesSet = false;
			for (int i = FChildren.Count - 1; i >=0 ; i--)
			{
				TChartBaseRecord CR = (FChildren[i] as TChartBaseRecord);
				if (CR == null) continue;
				switch (CR.Id)
				{
					case (int)xlr.ChartAI:
						TChartAIRecord AI = FChildren[i] as TChartAIRecord;
						if (AI != null)
						{
							switch (AI.LinkId)
							{
								case 0: 
									AI.SetDefinition(xls, Series.TitleDefinition, MasterCache);
									TitleSet =true;
									break;
								case 1: 
									AI.SetDefinition(xls, Series.DataDefinition, MasterCache);
									DataSet = true;
									break;
								case 2: 
									AI.SetDefinition(xls, Series.CategoriesDefinition, MasterCache);
									CategoriesSet = true;
									break;
							}
						}
						break;

					case (int)xlr.ChartSertocrt:
						BitOps.SetWord(CR.Data,0, Series.ChartOptionsIndex);
						break;
				}
			}

			if (!TitleSet)      FChildren.Add(TChartAIRecord.CreateFromData(xls, 0, Series.TitleDefinition, MasterCache));
			if (!DataSet)       FChildren.Add(TChartAIRecord.CreateFromData(xls, 1, Series.DataDefinition , MasterCache));
			if (!CategoriesSet) FChildren.Add(TChartAIRecord.CreateFromData(xls, 2, Series.CategoriesDefinition, MasterCache));
		}


	}
	#endregion
}
