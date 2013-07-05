using System;
using System.Diagnostics;
using System.Text;
using System.Globalization;

using FlexCel.Core;

namespace FlexCel.XlsAdapter
{
	internal class TWhatIfData
	{
		internal int WhatIfRow;
		internal int WhatIfCol;
		internal int WhatIfSheet;
		internal bool Recalculated;
		internal bool Recalculating;
		internal object FormulaValue;
	}

    /// <summary>
    /// Implements a formula for a cell.
    /// </summary>
    internal class TFormulaRecord : TCellRecord, IFormulaRecord
    {
        /// <summary>
        /// We need it in formulas, for example to calculate =Row()
        /// </summary>
        internal int FRow;

        private TParsedTokenList Data;

        internal object FormulaValue;
        byte[] Chn;
        internal UInt32 OptionFlags;
        internal bool bx;

        private bool Recalculating;
        private bool Recalculated;
        internal bool HasSubtotal;
        internal bool HasAggregate;

        internal TFormulaRecord Next;
        internal TFormulaRecord Prev;

		internal TWhatIfData WhatIf;

        private TTableRecord FTableRecord;
        private TArrayRecord FArrayRecord;

        internal TTableRecord TableRecord { get { return FTableRecord; } set { FTableRecord = value; } }
        internal TArrayRecord ArrayRecord { get { return FArrayRecord; } set { FArrayRecord = value; } }

        private TFormulaBounds Bounds;

        private TFormulaRecord(TNameRecordList Names, int aId, byte[] aData, TBiff8XFMap XFMap)
            : base(aId, aData, XFMap)
        {
            Data = TTokenManipulator.CreateFromBiff8(Names, FRow, Col, aData, 22, BitOps.GetWord(aData, 20), false, out HasSubtotal, out HasAggregate);

            FArrayRecord = null;
            Recalculated = false;
            Recalculating = false;
            Next = null;
            Prev = null;
            FRow = BitOps.GetWord(aData, 0);
            OptionFlags = (UInt32)BitOps.GetWord(aData, 14);
            Chn = new byte[4];
            Array.Copy(aData, 16, Chn, 0, Chn.Length);

            //Save the formula result
            FormulaValue = null;
            if (BitOps.GetWord(aData, 12) != 0xFFFF) //it's a number
            {
                FormulaValue = BitConverter.ToDouble(aData, 6);
            }
            else
            {
                switch (aData[6])
                {
                    case 0: FormulaValue = String.Empty; //It's a string. We will fill it later when we read the string record
                        break;
                    case 1: FormulaValue = (aData[8] == 1); //boolean
                        break;
                    //2 is error. we will return an error object.
                    case 2:
                        byte b = aData[8];
                        //Dont use here foreach ( TFlxFormulaErrorValue s in Enum.GetValues(typeof(TFlxFormulaErrorValue)) ) because it doesnt work on CF
                        if (b == (byte)TFlxFormulaErrorValue.ErrNull) FormulaValue = TFlxFormulaErrorValue.ErrNull;
                        else
                            if (b == (byte)TFlxFormulaErrorValue.ErrDiv0) FormulaValue = TFlxFormulaErrorValue.ErrDiv0;
                            else
                                if (b == (byte)TFlxFormulaErrorValue.ErrValue) FormulaValue = TFlxFormulaErrorValue.ErrValue;
                                else
                                    if (b == (byte)TFlxFormulaErrorValue.ErrRef) FormulaValue = TFlxFormulaErrorValue.ErrRef;
                                    else
                                        if (b == (byte)TFlxFormulaErrorValue.ErrName) FormulaValue = TFlxFormulaErrorValue.ErrName;
                                        else
                                            if (b == (byte)TFlxFormulaErrorValue.ErrNum) FormulaValue = TFlxFormulaErrorValue.ErrNum;
                                            else
                                                if (b == (byte)TFlxFormulaErrorValue.ErrNA) FormulaValue = TFlxFormulaErrorValue.ErrNA;
                        break;
                    case 3: FormulaValue = String.Empty; //Not documented but at least happens on pocket excel.  Update: now documented, it is an empty string.
                        break;
                }
                
            }
        }

       
		/// <summary>
        /// Create from data. Creates a formula record from stratch.
        /// This method wil *not* clone aData or aArrayData, so make sure the data can be used.
        /// </summary>
        internal TFormulaRecord(int aId, int aRow, int aCol, TXlsCellRange FmlaArrayRange, int aXF, TParsedTokenList aData, TParsedTokenList aArrayData, object FmlaResult, int aOptionFlags, bool Dates1904, int aArrayOptionFlags, bool abx)
            : base(aId, aCol, aXF)
        {
            Recalculating = false;
            Next = null;
            Prev = null;
            FRow = aRow;

            FormulaValue = null;
            bx = abx;

            SetFormulaResult(FmlaResult, Dates1904);
            Recalculated = true;
            OptionFlags = (UInt16)aOptionFlags;

            Data = aData; //no clone

            if (aArrayData != null)
            {
                FArrayRecord = new TArrayRecord((int)xlr.ARRAY, FmlaArrayRange, aArrayData, aArrayOptionFlags);
            }
            else
                FArrayRecord = null;

			if (Data!= null) FindSubtotal(out HasSubtotal, out HasAggregate);
        }

        internal static TFormulaRecord CreateFromBiff8(TNameRecordList Names, int aId, byte[] aData, TBiff8XFMap XFMap)
        {
            return new TFormulaRecord(Names, aId, aData, XFMap);
        }

        internal static TFormulaRecord CreateFromBiff4(TNameRecordList Names, byte[] aData, TBiff8XFMap XFMap)
        {
            //here is the real implementation of a biff4 formula.
            //But when reading from biff8, it looks like the formula format must be the same.
            /*byte[] NewData = new byte[aData.Length + 4];
            Array.Copy(aData, 0, NewData, 0, 16);
            Array.Copy(aData, 16, NewData, 20, aData.Length - 16);
            return new TFormulaRecord((int)xlr.FORMULA, NewData);*/

            return new TFormulaRecord(Names, (int)xlr.FORMULA, aData, XFMap);
        }

		private void FindSubtotal(out bool HasSubtotal, out bool HasAggregate)
		{
            HasSubtotal = false;
            HasAggregate = false;
			Data.ResetPositionToLast();
			while (!Data.Bof())
			{
				TBaseFunctionToken f = Data.LightPop() as TBaseFunctionToken;
                if (f != null)
                {
                    if (f.GetFunctionData().Index == 344) HasSubtotal = true;
                    if (f.GetFunctionData().Index == (int)TFutureFunctions.Aggregate) HasAggregate = true;
                }
			}
		}

        internal void ClearResult()
        {
            ClearChn();
			FormulaValue = null;
        }

        internal void ClearChn()
        {
            if (Chn != null) Array.Clear(Chn, 0, Chn.Length);
        }

        internal void ForceAutoRecalc()
        {
            // For automatic recalc on Excel97...
            OptionFlags = (byte)(OptionFlags | 2);
        }

        private bool HasStringRecord
        {
            get
            {
                string s = FormulaValue as string;
                return s != null && s.Length > 0;
            }
        }

        private TStringRecordData StringRecord
        {
            get
            {
                string s = FormulaValue as string;
                if (s != null && s.Length > 0) return new TStringRecordData(s);
                return null;
            }
        }

        internal bool HasExternRefs()
        {
            return TTokenManipulator.HasExternRefs(Data);
        }

        internal void NotRecalculated()
        {
            Recalculated = false;
            if (ArrayRecord != null) ArrayRecord.Recalculated = false;
        }

        internal void NotRecalculating()
        {
            Recalculating = false;
            if (ArrayRecord != null) ArrayRecord.Recalculating = false;
        }

        internal override bool AllowCopyOnOnlyFormula
        {
            get
            {
                return true;
            }
        }


        protected override TBaseRecord DoCopyTo(TSheetInfo SheetInfo)
        {
            TFormulaRecord Result = (TFormulaRecord)base.DoCopyTo(SheetInfo);
            Result.FRow = FRow;
            Result.Data = Data.Clone();
            //No need to clone FormulaValue, as it is immutable.
            Result.FormulaValue = null; //In fact, we shouldn't clone formula value at all.
            Result.FTableRecord = (TTableRecord)TTableRecord.Clone(FTableRecord, SheetInfo);
            Result.FArrayRecord = (TArrayRecord)TArrayRecord.Clone(FArrayRecord, SheetInfo);
            Result.NotRecalculated();
            Result.NotRecalculating();
            Result.Next = null;
            Result.Prev = null;
            Result.Bounds = null; //We will clear this cache.
            return Result;
        }

        internal TParsedTokenList CloneData()
        {
            return Data.Clone();
        }

        internal override void ArrangeInsertRange(int Row, TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
        {
            if (!((SheetInfo.InsSheet < 0) || (SheetInfo.SourceFormulaSheet != SheetInfo.InsSheet)))
            {
                if ((aRowCount != 0) && (Row >= CellRange.Top) && (Col <= CellRange.Right) && (Col >= CellRange.Left))
                    IncRef(ref FRow, aRowCount * CellRange.RowCount, FlxConsts.Max_Rows, XlsErr.ErrTooManyRows);  //row;
            }

            base.ArrangeInsertRange(Row, CellRange, aRowCount, aColCount, SheetInfo);
            ArrangeTokensInsertRange(Row, CellRange, aRowCount, aColCount, 0, 0, SheetInfo);
            if ((FTableRecord != null) && (SheetInfo.SourceFormulaSheet == SheetInfo.InsSheet)) FTableRecord.ArrangeInsertRange(CellRange, aRowCount, aColCount);
            if ((FArrayRecord != null) && (SheetInfo.SourceFormulaSheet == SheetInfo.InsSheet)) FArrayRecord.ArrangeInsertRange(Row, Col, CellRange, aRowCount, aColCount, SheetInfo);
        }

        internal void ArrangeMoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
        {
            if (SheetInfo.SourceFormulaSheet == SheetInfo.InsSheet)
            {
                if (CellRange.HasRow(FRow))
                    IncRef(ref FRow, NewRow - CellRange.Top, FlxConsts.Max_Rows, XlsErr.ErrTooManyRows);  //row;
            }

            ArrangeTokensMoveRange(CellRange, NewRow, NewCol, SheetInfo);
            if ((FTableRecord != null) && (SheetInfo.SourceFormulaSheet == SheetInfo.InsSheet)) FTableRecord.ArrangeMoveRange(CellRange, NewRow, NewCol);
            if ((FArrayRecord != null) && (SheetInfo.SourceFormulaSheet == SheetInfo.InsSheet)) FArrayRecord.ArrangeMoveRange(FRow, Col, CellRange, NewRow, NewCol, SheetInfo);
        }

        internal override void ArrangeCopyRange(TXlsCellRange SourceRange, int Row, int RowOffset, int ColOffset, TSheetInfo SheetInfo)
        {

            ArrangeTokensInsertRange(Row, SourceRange, 0, 0, RowOffset, ColOffset, SheetInfo); //Sheet info doesn't have meaning on copy
            if ((FTableRecord != null) && (SheetInfo.SourceFormulaSheet == SheetInfo.InsSheet)) FTableRecord.ArrangeCopyRange(RowOffset, ColOffset);
            if ((FArrayRecord != null) && (SheetInfo.SourceFormulaSheet == SheetInfo.InsSheet)) FArrayRecord.ArrangeCopyRange(Row, Col, RowOffset, ColOffset, SheetInfo);
            base.ArrangeCopyRange(SourceRange, Row, RowOffset, ColOffset, SheetInfo);   //should be last, so we don't modify Row or Col
            FRow += RowOffset;
        }

        private void ArrangeTokensInsertRange(int Row, TXlsCellRange CellRange, int aRowCount, int aColCount, int CopyRowOffset, int CopyColOffset,
            TSheetInfo SheetInfo)
        {
            if (Bounds != null && CopyRowOffset == 0 && CopyColOffset == 0   //When copying we will always fix the formula.
                && Bounds.OutBounds(CellRange, SheetInfo, aRowCount, aColCount)) return;

            if (Bounds == null) Bounds = new TFormulaBounds(); else Bounds.Clear();

            try
            {
                TTokenManipulator.ArrangeInsertAndCopyRange(Data, CellRange, Row, Col, aRowCount, aColCount, CopyRowOffset, CopyColOffset, SheetInfo, true, Bounds);
            }
            catch (ETokenException e)
            {
                XlsMessages.ThrowException(XlsErr.ErrBadFormula, Row + 1, Col + 1, e.Token);
            }
        }


        private void ArrangeTokensMoveRange(TXlsCellRange CellRange, int NewRow, int NewCol,
            TSheetInfo SheetInfo)
        {
            if (Bounds == null) Bounds = new TFormulaBounds(); else Bounds.Clear();

            try
            {
                TTokenManipulator.ArrangeMoveRange(Data, CellRange, FRow, Col, NewRow, NewCol, SheetInfo, Bounds);
            }
            catch (ETokenException e)
            {
                XlsMessages.ThrowException(XlsErr.ErrBadFormula, FRow + 1, Col + 1, e.Token);
            }
        }

        internal void ArrangeInsertSheet(TSheetInfo SheetInfo)
        {
            {
                TTokenManipulator.ArrangeInsertSheets(Data, SheetInfo); //ChangeGlobalToLocalNames
            }
        }

        private void ArrangeSharedTokens(int RowOffs, int ColOffs, bool FromBiff8)
        {
            try
			{
				TTokenManipulator.ArrangeSharedFormulas(Data, RowOffs, ColOffs, FromBiff8);
			}
			catch (ETokenException e)
			{
				XlsMessages.ThrowException(XlsErr.ErrBadFormula, RowOffs+1, ColOffs+1, e.Token);
			}
        }

        internal void MixShared(TParsedTokenList SharedTokens, int Row, bool Biff8)
        {
            MixShared(SharedTokens, Row, Col, Biff8);
        }

        internal void MixShared(TParsedTokenList SharedTokens, int RowOffs, int ColOffs, bool FromBiff8)
        {
            Data = SharedTokens.Clone();
            unchecked
            {
                OptionFlags &= (UInt16)~0x08; //Clear shared formula flag.
            }
            ArrangeSharedTokens(RowOffs, ColOffs, FromBiff8);
			FindSubtotal(out HasSubtotal, out HasAggregate);
        }

        internal bool IsExp(out UInt64 Key)
        {
            Key = 0;
            if (Data.Count != 1) return false; 
			Data.ResetPositionToLast();
            TBaseParsedToken Token = Data.LightPop();

            TExp_Token Exp = Token as TExp_Token;

			if (Exp != null) Key = (((UInt64)Exp.Row) << 32) + ((UInt32)Exp.Col);
			return Exp != null;
        }


        /// <summary>
        /// Used to delay-load string records.
        /// </summary>
        /// <param name="aValue"></param>
        internal void SetFormulaValue(TStringRecord aValue)
        {
            FormulaValue = aValue.Value();
        }

        private void SetFormulaResult(object aValue, bool Dates1904, ExcelFile xls)
        {
            object NewValue = aValue;
            if (xls.OptionsPrecisionAsDisplayed && TExcelTypes.ObjectToCellType(aValue) == TCellType.Number)
            {
                NewValue = TrimPrecisionAsDisplayed(aValue, xls, NewValue);
            }

            SetFormulaResult(NewValue, Dates1904);
        }

        private object TrimPrecisionAsDisplayed(object aValue, ExcelFile xls, object NewValue)
        {
            string Format = xls.GetCellVisibleFormatDef(FRow + 1, Col + 1).Format;
            if (Format != null && Format.Length > 0) //Not general
            {
                double V = Convert.ToDouble(aValue);

                bool SectionMatches; bool SuppressNegativeSign;
                string Fmt = TFlxNumberFormat.GetSection(Format, V, out SectionMatches, out SuppressNegativeSign);

                if (!SectionMatches)
                {
                    NewValue = TFlxNumberFormat.EmptySection;
                }
                else
                {
                    if (Fmt != null && Fmt.Length > 0)
                    {
                        string DecSep = ".";  //do not localize.
                        int comma = Fmt.LastIndexOf(DecSep);
                        if (comma >= 0)
                        {
                            int decimals = 0;
                            for (int i = comma + 1; i < Fmt.Length; i++)
                            {
                                if (Fmt[i] == '0' || Fmt[i] == '#') decimals++;
                            }

                            NewValue = TBaseParsedToken.UpRound(V, decimals);
                        }
                        else
                        {
                            if (Fmt.IndexOfAny(new char[] { '0', '#' }) >= 0) NewValue = TBaseParsedToken.UpRound(V, 0);
                        }
                    }
                }
            }
            return NewValue;
        }

        private void SetFormulaResult(object aValue, bool Dates1904)
        {
            FormulaValue = null; //clear result

            if (aValue is TMissingArg) aValue = null;
            switch (TExcelTypes.ObjectToCellType(aValue))
            {
                case TCellType.DateTime:
                    {
                        double d = FlxDateTime.ToOADate((DateTime)aValue, Dates1904);
                        FormulaValue = d;
                        break;
                    }
                case TCellType.Empty:
                    {
                        double d = 0;
                        FormulaValue = d;
                        break;
                    }
                case TCellType.Number:
                    {
                        double d = Convert.ToDouble(aValue);
                        FormulaValue = d;
                        break;
                    }
                case TCellType.String:
                    {
                        FormulaValue = aValue.ToString();
                        break;
                    }
                case TCellType.Bool:
                    {
                        FormulaValue = aValue;
                        break;
                    }
                case TCellType.Error:
                    {
                        FormulaValue = (TFlxFormulaErrorValue)aValue;
                        break;
                    }
                default:
                    {
                        //formulas should never return something else
                        //an exception thrown here would be catched, so we don't throw it.
                        FormulaValue = TFlxFormulaErrorValue.ErrNA;
                        break;
                    }
            }
        }

        private byte[] FormulaResultAsByteArray()
        {
            if (FormulaValue is double)
            {
                return BitConverter.GetBytes((double)FormulaValue);
            }

            byte[] b = new byte[8];

            if (FormulaValue is String)
            {
                BitOps.SetWord(b, 6, 0xFFFF);
				if (((string)FormulaValue).Length == 0) b[0] = 3;
            }
            else if (FormulaValue is bool)
            {
                b[0] = 1;
                if (Convert.ToBoolean(FormulaValue, CultureInfo.CurrentCulture)) b[2] = 1; else b[2] = 0;
                //no need to set 0s really. Formula result has been cleared.
                BitOps.SetWord(b, 6, 0xFFFF);
            }
            else
            {
                TFlxFormulaErrorValue ev = (FormulaValue is TFlxFormulaErrorValue) ? (TFlxFormulaErrorValue)FormulaValue : TFlxFormulaErrorValue.ErrNA;
                b[0] = 2;
                b[2] = (byte)(ev);
                //no need to set 0s really. Formula result has been cleared.
                BitOps.SetWord(b, 6, 0xFFFF);
            }

            return b;
        }

        internal static object EvaluateFormula(TParsedTokenList Fmla, TWorkbookInfo wi)
        {
            return EvaluateFormula(null, Fmla, wi, false);
        }

        internal static object EvaluateFormula(TBaseAggregate f, TParsedTokenList Tokens, TWorkbookInfo wi, bool EvalRef)
        {
            try
            {
                if (EvalRef)
                    return Tokens.EvaluateAllRef(wi, new TCalcState(), new TCalcStack());
                else
                    return Tokens.EvaluateAll(wi, f, new TCalcState(), new TCalcStack());
            }
            catch (FlexCelException)
            {
                return TFlxFormulaErrorValue.ErrNA;
            }
            catch (FormatException)
            {
                return TFlxFormulaErrorValue.ErrValue;
            }
            catch (ArithmeticException)
            {
                return TFlxFormulaErrorValue.ErrValue;
            }
            catch (ArgumentOutOfRangeException)
            {
                return TFlxFormulaErrorValue.ErrNum;
            }
        }

        internal override object GetValue(ICellList Cells)
        {
            TParsedTokenList ArrayData = ArrayRecord == null ? null : ArrayRecord.Data;
            string FmlaText = null;
            try
            {
                if (Cells.Workbook == null || !Cells.Workbook.IgnoreFormulaText)
                {
                    FmlaText = TFormulaConvertInternalToText.AsString(Data, FRow, Col, Cells);
                }
            }
            catch (FlexCelException)
            {
                FmlaText = TFormulaMessages.TokenString(TFormulaToken.fmStartFormula) + TFormulaMessages.ErrString(TFlxFormulaErrorValue.ErrNA);
            }

            TFormulaSpan Span = GetSpan(Cells);
            return new TFormula(FmlaText, FormulaValue, Data, ArrayData, true, Span);
        }

        private TFormulaSpan GetSpan(ICellList Cells)
        {
            TFormulaSpan Span = new TFormulaSpan();

            if (Cells == null) return Span;
            
            TArrayRecord ArrData; int RowArr; int ColArr; int RowCount; int ColCount;
            if (HasArrayFormula(Cells, out ArrData, out RowArr, out ColArr, out RowCount, out ColCount))
            {
                Span = new TFormulaSpan(RowCount, ColCount, ArrayRecord != null);
            }
            else
            {
                TTableRecord TableData;
                if (HasTableFormula(Cells, out TableData, out RowArr, out ColArr, out RowCount, out ColCount))
                {
                    Span = new TFormulaSpan(RowCount, ColCount, TableRecord != null);
                }
            }

            return Span;
        }

        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
        {
            base.SaveToStream(Workbook, SaveData, Row);

            SaveResult(Workbook);
            Workbook.Write16((UInt16)OptionFlags);
            if (Recalculated || Chn == null)  //If recalc=manual, chn has been cleared. If recalc=auto and !recalculated -> it is the original formula.
            {
                Workbook.Write(new byte[4], 4); //Cleared chn.
            }
            else
            {
                Workbook.Write(Chn, Chn.Length);
            }

            TTokenManipulator.SaveToStream(SaveData.Globals.Names, Workbook, TFormulaType.Normal, Data, true);

            if (FArrayRecord != null) FArrayRecord.SaveToStream(Workbook, Row, Col, SaveData);
            if (FTableRecord != null) FTableRecord.SaveToStream(Workbook, SaveData, Row);
            if (HasStringRecord)
            {
                StringRecord.SaveToStream(Workbook);
            }
        }

        private void SaveResult(IDataStream Workbook)
        {
            byte[] b = FormulaResultAsByteArray();
            Workbook.Write(b, b.Length);
        }

        internal override void SaveToPxl(TPxlStream PxlStream, int Row, TPxlSaveData SaveData)
        {
            base.SaveToPxl(PxlStream, Row, SaveData);
            if (!PxlRecordIsValid(Row)) return;

            TBiff7FormulaConverter ResultList = new TBiff7FormulaConverter();
            int FmlaLenWithoutArray;
            if (ResultList.LoadBiff8(TFormulaConvertInternalToBiff8.GetTokenData(SaveData.Globals.Names, Data, TFormulaType.Normal, out FmlaLenWithoutArray), 0, SaveData.Globals.References) != null)
            {
                object Result = FormulaValue;
                if (Result is Double)
                {
                    TNumberRecord Nr = new TNumberRecord(Col, XF, (Double)Result);
                    Nr.SaveToPxl(PxlStream, Row, SaveData);
                }
                else if (Result is bool)
                {
                    TBoolErrRecord Br = new TBoolErrRecord(Col, XF, (bool)Result);
                    Br.SaveToPxl(PxlStream, Row, SaveData);
                }
                else if (Result is TFlxFormulaErrorValue)
                {
                    TBoolErrRecord Br = new TBoolErrRecord(Col, XF, (TFlxFormulaErrorValue)Result);
                    Br.SaveToPxl(PxlStream, Row, SaveData);
                }
                else if (Result == null)
                {
                    //nothing here.
                }
                else
                {
                    TLabelSSTRecord.SaveToPxl(PxlStream, Row, Col, XF, Convert.ToString(Result), SaveData);
                }
                return;
            }

            PxlStream.WriteByte((byte)pxl.FORMULA);
            PxlStream.Write16((UInt16)Row);
            PxlStream.WriteByte((byte)Col);
            PxlStream.Write16(SaveData.GetBiff8FromCellXF(XF));

            PxlStream.Write(FormulaResultAsByteArray(), 0, 8); //Fmla result.
            PxlStream.WriteByte((byte)(OptionFlags & 0x1)); //OptionFlags.

            int ResultSize = ResultList.Size;
            PxlStream.Write16((UInt16)ResultSize);

            byte[] Biff7Data = new byte[ResultSize];
            ResultList.CopyToPtr(Biff7Data, 0);
            PxlStream.Write(Biff7Data, 0, ResultSize);

            if (HasStringRecord)
            {
                StringRecord.SaveToPxl(PxlStream);
            }
        }


        internal override int TotalSizeNoHeaders()
        {
            int Result = base.TotalSizeNoHeaders() + 16 + TTokenManipulator.TotalSizeWithArray(Data, TFormulaType.Normal);
            return Result;
        }

        internal override int TotalSize()
        {
            int Result = TotalSizeNoHeaders() + 4;
            if (FTableRecord != null) Result += FTableRecord.TotalSize();
            if (FArrayRecord != null) Result += FArrayRecord.TotalSize();
            if (HasStringRecord) Result += StringRecord.Length;
            return Result;

        }

        private static object ArrayToNum(object o)
        {
            object[,] arr = o as object[,];
            if (arr != null)
            {
                if (arr.GetLength(0) > 0 && arr.GetLength(1) > 0) return arr[0, 0];
                return TFlxFormulaErrorValue.ErrNA;
            }
            return o;
        }

        internal bool IsArrayFormula
        {
            get
            {
                if (Data.Count != 1) return false;
                Data.ResetPositionToLast();
                return Data.LightPop() is TExp_Token; 
            }
        }

        internal bool IsTableFormula
        {
            get
            {
                if (Data.Count != 1) return false;
                Data.ResetPositionToLast();
                return Data.LightPop() is TTableToken;
            }
        }

        private bool HasArrayFormula(ICellList CellList, out TArrayRecord ArrData, out int RowArr, out int ColArr, out int RowCount, out int ColCount)
        {
            RowArr = 0; ColArr = 0; ArrData = null; RowCount = 0; ColCount = 0;
            
            if (Data.Count != 1) return false;
            Data.ResetPositionToLast();
            TExp_Token range = Data.LightPop() as TExp_Token;
            if (range == null) return false;
            
            RowArr = range.Row;
            ColArr = range.Col;

            bool Found = CellList.FoundArrayFormula(RowArr, ColArr, out ArrData);
            if (!Found || ArrData == null) return false;
            RowCount = ArrData.RowCount;
            ColCount = ArrData.ColCount;
            return true;
        }

        private bool HasTableFormula(ICellList CellList, out TTableRecord TableData, out int TopRow, out int LeftCol, out int RowCount, out int ColCount)
        {
            TopRow = 0; LeftCol = 0; TableData = null; RowCount = 0; ColCount = 0;

            if (Data.Count != 1) return false;
            Data.ResetPositionToLast();
            TTableToken range = Data.LightPop() as TTableToken;
            if (range == null) return false;

            TopRow = range.Row;
            LeftCol = range.Col;

            bool Found = CellList.FoundTableFormula(TopRow, LeftCol, out TableData);
            if (!Found || TableData == null) return false;
            RowCount = TableData.LastRow - TableData.FirstRow + 1;
            ColCount = TableData.LastCol - TableData.FirstCol + 1;
            return true;
        }

        private static object GetItem(object[,] ArrResult, int Row, int Col)
        {
            if (ArrResult == null || Row < 0 || Col < 0) return TFlxFormulaErrorValue.ErrNA;
            if (ArrResult.GetLength(0) == 1) Row = 0;
            if (ArrResult.GetLength(1) == 1) Col = 0;
            if (Row >= ArrResult.GetLength(0) || Col >= ArrResult.GetLength(1)) return TFlxFormulaErrorValue.ErrNA;
            return ArrResult[Row, Col];
        }

		internal void SetWhatIf(int aRow, int aCol, int aSheet)
		{
			if (aRow <= 0 || aCol <= 0) 
			{
				WhatIf = null;
				return;
			}
			
			if (WhatIf == null || WhatIf.WhatIfRow != aRow || WhatIf.WhatIfCol != aCol || WhatIf.WhatIfSheet != aSheet)
			{
				WhatIf = new TWhatIfData();
			}
			WhatIf.WhatIfRow = aRow;
			WhatIf.WhatIfCol = aCol;
			WhatIf.WhatIfSheet = aSheet;
		}

		private bool IsRecalculating
		{
			get
			{
				if (WhatIf == null) return Recalculating;
				return WhatIf.Recalculating;
			}
			set
			{
				if (WhatIf == null) Recalculating = value;
				else WhatIf.Recalculating = value;
			}
		}

		private bool IsRecalculated
		{
			get
			{
				if (WhatIf == null) return Recalculated;
				return WhatIf.Recalculated;
			}
			set
			{
				if (WhatIf == null) Recalculated = value;
				else WhatIf.Recalculated = value;
			}
		}

        internal void SetWhatIfFormulaResult(object v, bool Dates1904, ExcelFile xls)
        {
            if (WhatIf != null) WhatIf.FormulaValue = TExcelTypes.ConvertToAllowedObject(v, Dates1904);
            else SetFormulaResult(v, Dates1904, xls);
        }
        
        internal void SetWhatIfFormulaResult(object v, bool Dates1904)
        {
            if (WhatIf != null) WhatIf.FormulaValue = TExcelTypes.ConvertToAllowedObject(v, Dates1904);
            else SetFormulaResult(v, Dates1904);
        }

		internal object WhatIfFormulaValue
		{
			get
			{
				if (WhatIf != null) return WhatIf.FormulaValue; else return FormulaValue;
			}
		}

        internal void Recalc(TCellList CellList, ExcelFile aXls, int SheetIndexBase1, TCalcState CalcState, TCalcStack CalcStack)
        {
            if (IsRecalculated) return;

            if (/*CalcStack.Level > 500 ||*/ CalcState.Aborted) //We should do some tests in CalcStack.Level to avoid overflows, but we can't until we have a way to do a re-recalc.
            {
                CalcState.Aborted = true;
                SetWhatIfFormulaResult(TFlxFormulaErrorValue.ErrNA, false);
                IsRecalculated = true;
                IsRecalculating = false;
                return;
            }
            
            CalcStack.Level++;
            if (IsRecalculating) //Circular reference... FlxMessages.ThrowException(FlxErr.ErrCircularReference, new TCellAddress(Row,Col).CellRef);
            {
                SetWhatIfFormulaResult(TFlxFormulaErrorValue.ErrRef, CellList.Dates1904);
                aXls.AddUnsupported(TUnsupportedFormulaErrorType.CircularReference, null);
                IsRecalculating = false;
                IsRecalculated = true;
                return;
            }

            if (CalcStack.ParentSheetBase1 == SheetIndexBase1 && CalcStack.ParentXls == aXls)
            {
                //This is more complex than what it looks. In order to do a correct reorder we need to be sure we are not moving ParentFmla Up.
                //bool Reordered = aXls.ReorderCalcChain(SheetIndexBase1, CalcStack.ParentFmla, this);
                //if (Reordered) CalcState.Reordered = true;
            }
            CalcStack.ParentFmla = this; //this should be assigned after ReorderCalcChain
            CalcStack.ParentSheetBase1 = SheetIndexBase1;
            CalcStack.ParentXls = aXls;

            IsRecalculating = true;

            try
            {
                TArrayRecord ArrData; int RowArr; int ColArr; int RowCount; int ColCount;
                if (HasArrayFormula(CellList, out ArrData, out RowArr, out ColArr, out RowCount, out ColCount))
                {
                    TWorkbookInfo wi = new TWorkbookInfo(aXls, SheetIndexBase1, FRow, Col, RowCount, ColCount, 0, 0, IsArrayFormula);
                    object[,] ArrResult = ArrData.GetValueAndRecalc(wi, CalcState, CalcStack);
                    if (!IsRecalculated) //When evaluating, the formula could be recalculated, and we should honor the first recalc.
                        SetWhatIfFormulaResult(GetItem(ArrResult, FRow - RowArr, Col - ColArr), CellList.Dates1904, aXls);
                }
                else
                {
                    TWorkbookInfo wi = new TWorkbookInfo(aXls, SheetIndexBase1, FRow, Col, 0, ColCount, 0, 0, IsArrayFormula);
                    object res = Data.EvaluateAll(wi, TErr2Aggregate.Instance, CalcState, CalcStack);
                    if (!IsRecalculated) //When evaluating, the formula could be recalculated, and we should honor the first recalc.
                        SetWhatIfFormulaResult(ArrayToNum(res), CellList.Dates1904, aXls);
                }
            }
            catch (FlexCelException)
            {
                SetWhatIfFormulaResult(TFlxFormulaErrorValue.ErrNA, false);
            }
            catch (FormatException)
            {
                SetWhatIfFormulaResult(TFlxFormulaErrorValue.ErrValue, false);
            }
            catch (ArithmeticException)
            {
                SetWhatIfFormulaResult(TFlxFormulaErrorValue.ErrValue, false);
            }
            catch (ArgumentOutOfRangeException)
            {
                SetWhatIfFormulaResult(TFlxFormulaErrorValue.ErrNum, false);
            }

            if (CalcState.Aborted) SetWhatIfFormulaResult(TFlxFormulaErrorValue.ErrNA, false);

            IsRecalculating = false;
            IsRecalculated = true;
        }


        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            Loader.LastFormula = this;
            Loader.LastFormulaRow = rRow;

            UInt64 Key;
            if (IsExp(out Key)) //Might be a shared formula
            {
                TSharedFormula ShraredFmla;
                //If not found, it might be an Array record or the first shared formula. The first one will be taken care of in the LoadIntoWorksheet method in the SharedFormula record.
                if (Loader.ShrFmlas.TryGetValue(Key, out ShraredFmla))
                {
                    MixShared(ShraredFmla.Data, FRow, true); //only biff8 has a TBaseRecordLoader
                }
            }


            base.LoadIntoSheet(ws, rRow, RecordLoader, ref Loader);
        }

        #region Named Ranges
        internal void UpdateDeletedRanges(TDeletedRanges DeletedRanges)
        {
            TTokenManipulator.UpdateDeletedRanges(Data, DeletedRanges);
            if (ArrayRecord != null) TTokenManipulator.UpdateDeletedRanges(ArrayRecord.Data, DeletedRanges);
        }
        #endregion

#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
        internal override void SaveToXlsx(TOpenXmlWriter DataStream, int Row, TCellList CellList, bool Dates1904)
        {
            string t = GetXlsxType();
            if (t != null) DataStream.WriteAtt("t", t);

            if ((IsArrayFormula && ArrayRecord == null) || (IsTableFormula && TableRecord == null))
            {
                //nothing, we don't save the formula here.
            }
            else
            {
                DataStream.WriteStartElement("f", false);

                string FormulaText;
                if (ArrayRecord != null)
                {
                    DataStream.WriteAtt("t", "array");
                    WriteArrayRef(DataStream, Row, Col, ArrayRecord.RowCount, ArrayRecord.ColCount);
                    DataStream.WriteAtt("aca", (ArrayRecord.OptionFlags & 0x01) != 0, false);

                    FormulaText = TFormulaConvertInternalToText.AsString(ArrayRecord.Data, Row, Col, CellList, CellList.Globals, FlxConsts.Max_FormulaStringConstant, true);

                }
                else
                {
                    if (TableRecord != null)
                    {
                        DataStream.WriteAtt("t", "dataTable");
                        WriteArrayRef(DataStream, TableRecord.FirstRow, TableRecord.FirstCol, TableRecord.LastRow - TableRecord.FirstRow + 1, TableRecord.LastCol - TableRecord.FirstCol + 1);
                        FormulaText = null;
                        WriteTableDef(DataStream);

                    }
                    else
                    {
                        FormulaText = TFormulaConvertInternalToText.AsString(Data, Row, Col, CellList, CellList.Globals, FlxConsts.Max_FormulaStringConstant, true);
                    }
                }


                DataStream.WriteAtt("ca", (OptionFlags & 0x01) != 0, false);
                DataStream.WriteAtt("bx", bx, false);

                if (FormulaText != null) DataStream.WriteString(FormulaText);
                DataStream.WriteEndElement();
            }
            WriteFormulaValue(DataStream, Dates1904);
        }

        private void WriteTableDef(TOpenXmlWriter DataStream)
        {
            DataStream.WriteAtt("dtr", TableRecord.CellInputIsRow, false);
            DataStream.WriteAtt("dt2D", TableRecord.Has2Entries ,false);
            DataStream.WriteAtt("del1", TableRecord.IsDeleted1, false);
            DataStream.WriteAtt("del2", TableRecord.IsDeleted2, false);
            TCellAddress FirstCell = new TCellAddress(TableRecord.RwInpRw + 1, TableRecord.ColInpRw + 1);
            DataStream.WriteAtt("r1", FirstCell.CellRef);

            if (TableRecord.Has2Entries)
            {
                TCellAddress SecondCell = new TCellAddress(TableRecord.RwInpCol + 1, TableRecord.ColInpCol + 1);
                DataStream.WriteAtt("r2", SecondCell.CellRef);
            }            
        }

        private void WriteFormulaValue(TOpenXmlWriter DataStream, bool Dates1904)
        {
            switch (TExcelTypes.ObjectToCellType(FormulaValue))
            {
                case TCellType.Number:
                    DataStream.WriteElement("v", Convert.ToDouble(FormulaValue));
                    break;

                case TCellType.DateTime:
                    DataStream.WriteElement("v", FlxDateTime.ToOADate(Convert.ToDateTime(FormulaValue), Dates1904));
                    break;

                case TCellType.String:
                    DataStream.WriteElement("v", Convert.ToString(FormulaValue, CultureInfo.InvariantCulture));
                    break;

                case TCellType.Bool:
                    DataStream.WriteElement("v", Convert.ToBoolean(FormulaValue, CultureInfo.InvariantCulture));
                    break;

                case TCellType.Error:
                    DataStream.WriteElement("v", TFormulaMessages.ErrString((TFlxFormulaErrorValue)FormulaValue));
                    break;

                case TCellType.Empty:
                    break;

                case TCellType.Formula:
                    break;

                case TCellType.Unknown:
                    break;

                default:
                    break;
            }
        }

        private static void WriteArrayRef(TOpenXmlWriter DataStream, int aRow, int aCol, int RowCount, int ColCount)
        {
            TCellAddress Addr1 = new TCellAddress(aRow + 1, aCol + 1);
            TCellAddress Addr2 = new TCellAddress(aRow + RowCount, aCol + ColCount);
            if (Addr1.CellRef == Addr2.CellRef)
            {
                DataStream.WriteAtt("ref", Addr1.CellRef);
            }
            else
            {
                DataStream.WriteAtt("ref", Addr1.CellRef + TFormulaMessages.TokenString(TFormulaToken.fmRangeSep) + Addr2.CellRef);
            }
        }

        private string GetXlsxType()
        {
            switch (TExcelTypes.ObjectToCellType(FormulaValue))
            {
                case TCellType.Number:
                    return null;

                case TCellType.DateTime:
                    return null;

                case TCellType.String:
                    return "str";

                case TCellType.Bool:
                    return "b";

                case TCellType.Error:
                    return "e";

                case TCellType.Empty:
                    return null;

                case TCellType.Formula:
                    FlxMessages.ThrowException(FlxErr.ErrInternal);
                    break;
            }
            return null;            
        }
#endif


    }

    /// <summary>
    /// a name comment.
    /// </summary>
    internal class TNameCmtRecord : TBaseRecord
    {
        public string Comment;
        public string Name;

        internal TNameCmtRecord(int aId, byte[] aData)
        {
            long StSize = 0;
            StrOps.GetSimpleString(true, aData, 16, true, BitOps.GetWord(aData, 12), ref Name, ref StSize);
            StrOps.GetSimpleString(true, aData, 16 + (int)StSize, true, BitOps.GetWord(aData, 14), ref Comment, ref StSize);
        }

        protected override TBaseRecord DoCopyTo(TSheetInfo SheetInfo)
        {
            return null;
        }

        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
        {
        }

        internal override int TotalSize()
        {
            return 0;
        }

        internal override int TotalSizeNoHeaders()
        {
            return 0;
        }

        internal override int GetId
        {
            get { return (int)xlr.NAMECMT; }
        }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.Names.AddComment(this);
        }
    }

    /// <summary>
    /// Implements a Named Range.
    /// </summary>
    internal class TNameRecord : TBaseRecord
    {
        internal int Id;
        internal TParsedTokenList Data;
        private string FName;

        int FOptionFlags;
        internal string KeyboardShortcut;
        internal int RangeSheet; //0 based, negative means global
        internal string FMenu;
        internal string FDescription;
        internal string FHelp;
        internal string FStatusBar;
        internal string FComment;

        private TNameRecord(TNameRecordList Names, int aId, byte[] aData)
        {
            Id = aId;

            FOptionFlags = BitOps.GetWord(aData, 0);
            if (aData[2] != 0) KeyboardShortcut = Convert.ToString((char)aData[2]);

            long StSize = 0;
            int PosOfs = 0;
             if (aData[3] > 0) StrOps.GetSimpleString(false, aData, 14, true, aData[3], ref FName, ref StSize);
            PosOfs += (int)StSize;

            bool HasSubtotal; bool HasAggregate;
            if (14 + PosOfs >= aData.Length) Data = new TParsedTokenList(null);
            else
                Data = TTokenManipulator.CreateFromBiff8(Names, -1, -1, aData, 14 + PosOfs, BitOps.GetWord(aData, 4), true, out HasSubtotal, out HasAggregate);


            RangeSheet = BitOps.GetWord(aData, 8) - 1;  //data 6 might not have the right value.

            if (14 + PosOfs < aData.Length && aData[10] > 0)
            {
                StrOps.GetSimpleString(false, aData, 14 + PosOfs, true, aData[10], ref FMenu, ref StSize);
                PosOfs += (int)StSize;
            }
            if (14 + PosOfs < aData.Length && aData[11] > 0)
            {
                StrOps.GetSimpleString(false, aData, 14 + PosOfs, true, aData[11], ref FDescription, ref StSize);
                PosOfs += (int)StSize;
            }
            if (14 + PosOfs < aData.Length && aData[12] > 0)
            {
                StrOps.GetSimpleString(false, aData, 14 + PosOfs, true, aData[12], ref FHelp, ref StSize);
                PosOfs += (int)StSize;
            }

            if (14 + PosOfs < aData.Length && aData[13] > 0)
            {
                StrOps.GetSimpleString(false, aData, 14 + PosOfs, true, aData[13], ref FStatusBar, ref StSize);
                PosOfs += (int)StSize;
            }

        }

        internal static TNameRecord CreateFromBiff8(TNameRecordList Names, int aId, byte[] aData)
        {
            return new TNameRecord(Names, aId, aData);
        }

        internal TNameRecord(TXlsNamedRange Range, TWorkbookGlobals Globals, TCellList CellList)
        {
            Id = (int)xlr.NAME;

            if (Range.RangeFormula != null)
            {
                int DefaultSheet = 0;
                if (Range.NameSheetIndex >= 0)
                    DefaultSheet = Range.NameSheetIndex;
                string DefaultSheetName = Globals.GetSheetName(DefaultSheet);

                if (Range.RangeFormula.Trim().Length == 0)
                {
                    Data = new TParsedTokenList(new TBaseParsedToken[0]);
                }
                else
                {
                    ExcelFile Workbook = CellList == null ? null : CellList.Workbook;
                    TFormulaConvertTextToInternal Ps = new TFormulaConvertTextToInternal(Workbook, Workbook.ActiveSheet, true, Range.RangeFormula, true, true, true, DefaultSheetName, TFmReturnType.Ref, false);
                    Ps.Parse();
                    Data = Ps.GetTokens();
                }
            }
            else
                if (Range.Left == Range.Right && Range.Top == Range.Bottom)
                {
                    TBaseParsedToken[] Tokens = new TBaseParsedToken[1];
                    Tokens[0] = new TRef3dToken(ptg.Ref3d, Globals.References.AddSheet(Globals.SheetCount, Range.SheetIndex), Range.Top, Range.Left, true, true);
                    Data = new TParsedTokenList(Tokens);
                }
                else
                {
                    TBaseParsedToken[] Tokens = new TBaseParsedToken[1];
                    Tokens[0] = new TArea3dToken(ptg.Area3d, Globals.References.AddSheet(Globals.SheetCount, Range.SheetIndex), Range.Top, Range.Left, true, true, Range.Bottom, Range.Right, true, true);
                    Data = new TParsedTokenList(Tokens);
                }

            FOptionFlags = Range.OptionFlags;
            RangeSheet = Range.NameSheetIndex;
            FName = Range.Name;
            FComment = Range.Comment;
        }

        internal TNameRecord(int aId, TParsedTokenList aData, int aOptionFlags, string functionName, 
            string aKeyboardShortcut, int aRangeSheet0Based, string aMenu, string aDescription, string aHelp, string aStatusBar, string aComment)
        {
            Id = aId;
            Data = aData;
            FOptionFlags = aOptionFlags;
            FName = functionName;
            Data = aData;
            KeyboardShortcut = aKeyboardShortcut;
            RangeSheet = aRangeSheet0Based;
            FMenu = aMenu;
            FDescription = aDescription;
            FHelp = aHelp;
            FStatusBar = aStatusBar;
            FComment = aComment;
        }

        internal static TNameRecord CreateTempName(string NameStr, int Sheet)
        {
            TParsedTokenList Tokens = new TParsedTokenList(null);
            return new TNameRecord((int)xlr.NAME, Tokens, 0, NameStr, null, Sheet, null, null, null, null, null);
        }

        internal static TNameRecord CreateAddin(string functionName, bool AddErrorDataToFormula)
        {
            TParsedTokenList Tokens;
            int Opts = 0x0E; //macro
            if (AddErrorDataToFormula)
            {
                Tokens = new TParsedTokenList(new TBaseParsedToken[]{new TErrDataToken(TFlxFormulaErrorValue.ErrName)});
                Opts = 0x0B; //no macro, this is an internal func, and if left at 0x0e, file will open fine, but fail to load after saevd by Excel.
            }
            else
            {
                Tokens = new TParsedTokenList(null);
            }
            return new TNameRecord((int)xlr.NAME, Tokens, Opts, functionName, null, -1, null, null, null, null, null);
        }

		internal override int GetId	{ get {	return Id; }}

        internal void ArrangeInsertSheets(int FirstSheet, int SheetCount)
        {
            if ((RangeSheet != 0xFFFF) && (RangeSheet >= FirstSheet))
            {
                RangeSheet += SheetCount; //NewSheet is 0 based, Data[8] is one-based;
                if (RangeSheet > FlxConsts.Max_Sheets + 1) XlsMessages.ThrowException(XlsErr.ErrTooManySheets, RangeSheet, FlxConsts.Max_Sheets + 1);
            }
        }

        void ArrangeTokensInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount, int CopyRowOffset, int CopyColOffset, TSheetInfo SheetInfo)
        {
            try
            {
                TTokenManipulator.ArrangeInsertAndCopyRange(Data, CellRange, -1, -1, aRowCount, aColCount, CopyRowOffset, CopyColOffset, SheetInfo, true, null);
            }
            catch (ETokenException e)
            {
                XlsMessages.ThrowException(XlsErr.ErrBadName, Name, e.Token);
            }
        }


        internal void ArrangeInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
        {
            ArrangeTokensInsertRange(CellRange, aRowCount, aColCount, 0, 0, SheetInfo);
        }

        internal void ArrangeCopyRange(int RowOffset, int ColOffset, TSheetInfo SheetInfo)
        {
            ArrangeTokensInsertRange(new TXlsCellRange(0, 0, -1, -1), 0, 0, RowOffset, ColOffset, SheetInfo);
        }

        private void ArrangeTokensMoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
        {
            try
            {
                TTokenManipulator.ArrangeMoveRange(Data, CellRange, -1, -1, NewRow, NewCol, SheetInfo, null);
            }
            catch (ETokenException e)
            {
                XlsMessages.ThrowException(XlsErr.ErrBadName, Name, e.Token);
            }
        }

        internal void ArrangeMoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
        {
            ArrangeTokensMoveRange(CellRange, NewRow, NewCol, SheetInfo);
        }

        internal void UpdateDeletedRanges(TDeletedRanges DeletedRanges)
        {
            try
            {
                TTokenManipulator.UpdateDeletedRanges(Data, DeletedRanges);
            }
            catch (ETokenException e)
            {
                XlsMessages.ThrowException(XlsErr.ErrBadName, Name, e.Token);
            }
        }

        internal TNameRecord CopyTo(int DestSheet, int CopyRowOffset, int CopyColOffset, TSheetInfo SheetInfo)
        {
            TNameRecord Result = (TNameRecord)TNameRecord.Clone(this, SheetInfo);

            Result.RangeSheet = DestSheet;
            Result.ArrangeCopyRange(CopyRowOffset, CopyColOffset, SheetInfo);

            return Result;
        }


        internal override void SaveToPxl(TPxlStream PxlStream, int Row, TPxlSaveData SaveData)
        {
            base.SaveToPxl(PxlStream, Row, SaveData);
            TBiff7FormulaConverter ResultList = new TBiff7FormulaConverter();
            int FmlaLenWithoutArray;
            if (ResultList.LoadBiff8(TFormulaConvertInternalToBiff8.GetTokenData(SaveData.Globals.Names, Data, TFormulaType.Name, out FmlaLenWithoutArray), 0, SaveData.Globals.References) != null)
            {
                return;
            }

            PxlStream.WriteByte((byte)pxl.NAME);

            PxlStream.Write16((UInt16)(OptionFlags & 0x1)); //OptionFlags.
            byte[] NameText = Encoding.Unicode.GetBytes(Name);
            PxlStream.WriteByte((byte)(NameText.Length / 2));


            int ResultSize = ResultList.Size;
            PxlStream.Write16((UInt16)ResultSize);

            UInt16 ixals = (UInt16)(RangeSheet + 1);
            if (ixals == 0) ixals = 0xFFFF;  //references to local book.
            PxlStream.Write16(ixals);

            PxlStream.Write(NameText, 0, NameText.Length);

            byte[] Biff7Data = new byte[ResultSize];
            ResultList.CopyToPtr(Biff7Data, 0);
            PxlStream.Write(Biff7Data, 0, ResultSize);
        }

        internal string Name
        {
            get
            {
                if (FName == null) return String.Empty;
                return FName;
            }
            set
            {
                if (String.Equals(Name, value, StringComparison.InvariantCultureIgnoreCase)) return;
                FName = value;
            }
        }

        internal string Comment
        {
            get
            {
                return FComment;
            }
            set
            {
                FComment = value;
            }
        }

        internal int NameLength
        {
            get
            {
                return Name.Length;
            }
        }

        internal int OptionFlags
        {
            get
            {
                return FOptionFlags;
            }
        }

        internal bool IsAddin
        {
            get
            {
                return (OptionFlags & 0x08) != 0;
            }
        }

        internal TParsedTokenList FormulaData
        {
            get
            {
                return Data;
            }
            set
            {
                Data = value;
            }
        }

        internal bool HasFormulaData
        {
            get
            {
                return Data.Count > 0;
            }
        }

        internal TNameRecord CopyWithoutFormulaData()
        {
            return new TNameRecord(Id, new TParsedTokenList(null), OptionFlags, Name, KeyboardShortcut, RangeSheet, FMenu, FDescription, FHelp, FStatusBar, FComment);
        }

        internal TNameRecord ArrangeCopySheet(TSheetInfo SheetInfo)
        {
            try
            {
                TTokenManipulator.ArrangeInsertSheets(Data, SheetInfo);
            }
            catch (ETokenException e)
            {
                XlsMessages.ThrowException(XlsErr.ErrBadName, Name, e.Token);
            }

            RangeSheet = SheetInfo.InsSheet; //InsSheet is 0 based, Data[8] is one-based;
            return this;
        }

        internal int R1
        {
            get
            {
                if (Data.Count != 1) return -1;
                Data.ResetPositionToLast();
                TBaseParsedToken basetoken = Data.LightPop();
                TArea3dToken at = basetoken as TArea3dToken;
                if (at != null && !at.IsErr()) return at.GetRow1(0);

                TRef3dToken rf = basetoken as TRef3dToken;
                if (rf != null && !rf.IsErr()) return rf.GetRow1(0);
                return -1;
            }
        }

        internal int R2
        {
            get
            {
                if (Data.Count != 1) return -1;
                Data.ResetPositionToLast();
                TBaseParsedToken basetoken = Data.LightPop();
                TArea3dToken at = basetoken as TArea3dToken;
                if (at != null && !at.IsErr()) return at.GetRow2(0);

                TRef3dToken rf = basetoken as TRef3dToken;
                if (rf != null && !rf.IsErr()) return rf.GetRow1(0);
                return -1;
            }
        }

        internal int C1
        {
            get
            {
                if (Data.Count != 1) return -1;
                Data.ResetPositionToLast();
                TBaseParsedToken basetoken = Data.LightPop();
                TArea3dToken at = basetoken as TArea3dToken;
                if (at != null && !at.IsErr()) return at.GetCol1(0);

                TRef3dToken rf = basetoken as TRef3dToken;
                if (rf != null && !rf.IsErr()) return rf.GetCol1(0);
                return -1;
            }
        }

        internal int C2
        {
            get
            {
                if (Data.Count != 1) return -1;
                Data.ResetPositionToLast();
                TBaseParsedToken basetoken = Data.LightPop();
                TArea3dToken at = basetoken as TArea3dToken;
                if (at != null && !at.IsErr()) return at.GetCol2(0);

                TRef3dToken rf = basetoken as TRef3dToken;
                if (rf != null && !rf.IsErr()) return rf.GetCol1(0);
                return -1;
            }
        }

        internal int RefersToSheet(TReferences References)
        {
            if (Data.Count != 1) return -1;
            Data.ResetPositionToLast();
            TBaseParsedToken basetoken = Data.LightPop();
            TArea3dToken at = basetoken as TArea3dToken;
            if (at != null && ! at.IsErr())
            {
                return References.GetJustOneSheet(at.FExternSheet);
            }

            TRef3dToken rf = basetoken as TRef3dToken;
            if (rf != null && !rf.IsErr())
            {
                return References.GetJustOneSheet(rf.FExternSheet);
            }
            return -1;
        }

        internal bool HasExternRefs(TReferences References)
        {
            return TTokenManipulator.HasExternLinks(Data, References);
        }

        internal override int TotalSize()
        {
            return TotalSizeNoHeaders() + XlsConsts.SizeOfTRecordHeader + CommentTotalSize();
        }

        internal override int TotalSizeNoHeaders()
        {
            int Result = 14 + TTokenManipulator.TotalSizeWithArray(Data, TFormulaType.Name) +
                GetByteLen(FName) + GetByteLen(FMenu) + GetByteLen(FDescription) + GetByteLen(FHelp) + GetByteLen(FStatusBar);

			return Result;
        }

        private int CommentTotalSize() //there will never be a continue record here.
        {
            if (FComment == null || FComment.Length == 0) return 0;
            return XlsConsts.SizeOfTRecordHeader + CommentTotalSizeNoHeaders();
        }
        private int CommentTotalSizeNoHeaders()
        {
            if (FComment == null || FComment.Length == 0) return 0;
            return 16 + GetWordLen(Trim255(FComment)) + GetWordLen(FName);
        }

        private static string Trim255(string a)
        {
            if (a == null || a.Length <= 255) return a;
            return a.Substring(0, 255);
        }

        private int GetByteLen(string variable)
        {
            if (variable == null || variable.Length == 0) return 0;
            TExcelString Xs = new TExcelString(TStrLenLength.is8bits, variable, null, false);
            return Xs.TotalSize() - 1;
        }

        private int GetWordLen(string variable)
        {
            if (variable == null || variable.Length == 0) return 0;
            TExcelString Xs = new TExcelString(TStrLenLength.is16bits, variable, null, false);
            return Xs.TotalSize() - 2;
        }

        private void AddLen(ref UInt32 Lens, string variable, int offs)
        {
            if (variable == null || variable.Length == 0) return;

            if (variable.Length > 255) FlxMessages.ThrowException(FlxErr.ErrNameTooLong, 255);
            Lens += ((UInt32)variable.Length) << offs; //here we write the char count, not the byte len
        }

        private static void WriteString(IDataStream Workbook, string variable)
        {
            if (variable == null || variable.Length == 0) return;
            
            TExcelString Xs = new TExcelString(TStrLenLength.is8bits, variable, null, false);

            byte[] bData = new byte[Xs.TotalSize() - 1];
            Xs.CopyToPtr(bData, 0, false);

            Workbook.Write(bData, bData.Length);
        }


        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
        {
            Workbook.WriteHeader((UInt16)Id, (UInt16)TotalSizeNoHeaders());

            unchecked
            {
                Workbook.Write16((UInt16)FOptionFlags);
            }
            if (NameLength > 255) FlxMessages.ThrowException(FlxErr.ErrNameTooLong, 255);
            byte kbs = 0;
            if (KeyboardShortcut != null && KeyboardShortcut.Length == 1)
            {
                char kbc = KeyboardShortcut[0];
                if ((kbc >= 0x41 && kbc <= 0x5A) || (kbc >= 0x61 && kbc <= 0x7A))
                {
                    kbs = (byte)kbc;
                }
            }
            Workbook.Write16((UInt16)(kbs + NameLength << 8));

            int FmlaNoArrayLen;
            byte[] bData = TFormulaConvertInternalToBiff8.GetTokenData(SaveData.Globals.Names, Data, TFormulaType.Name, out FmlaNoArrayLen);
            Workbook.Write16((UInt16)FmlaNoArrayLen);

            Workbook.Write16(0);  //Workbook.Write16((UInt16)(RangeSheet + 1));  It is not what Excel Saves.
            Workbook.Write16((UInt16)(RangeSheet + 1));

            UInt32 Lens = 0;
            AddLen(ref Lens, FMenu, 0);
            AddLen(ref Lens, FDescription, 8);
            AddLen(ref Lens, FHelp, 16);
            AddLen(ref Lens, FStatusBar, 24);

            Workbook.Write32(Lens);

            WriteString(Workbook, Name);
            Workbook.Write(bData, bData.Length);

            WriteString(Workbook, FMenu);
            WriteString(Workbook, FDescription);
            WriteString(Workbook, FHelp);
            WriteString(Workbook, FStatusBar);

            if (FComment != null && FComment.Length > 0)
            {
                Workbook.WriteHeader((UInt16)xlr.NAMECMT, (UInt16)CommentTotalSizeNoHeaders());
                Workbook.Write16((UInt16)xlr.NAMECMT);
                Workbook.Write(new byte[10], 10);

                UInt32 Lens2 = 0;
                AddLen(ref Lens2, FName, 0);
                AddLen(ref Lens2, FComment, 16);
                Workbook.Write32(Lens2);
                WriteString(Workbook, Name);
                WriteString(Workbook, FComment);
            }
        }

        protected override TBaseRecord DoCopyTo(TSheetInfo SheetInfo)
        {
            return new TNameRecord(Id, Data.Clone(), OptionFlags, Name, KeyboardShortcut, RangeSheet, FMenu, FDescription, FHelp, FStatusBar, FComment);           
        }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.Names.Add(this);
        }

    }

    internal class TTableRecord : TBaseRecord
    {
        int Id;
        internal int OptionFlags;
        internal int FirstRow;
        internal int LastRow;
        internal int FirstCol;
        internal int LastCol;
        internal int RwInpRw;
        internal int ColInpRw;
        internal int RwInpCol;
        internal int ColInpCol;

        internal const int FlagDeleted1 = 0x10;
        internal const int FlagDeleted2 = 0x20;



        private TTableRecord(int aId, byte[] aData)
            : base()
        {
            Id = aId;
            FirstRow = BitOps.GetWord(aData, 0);  //Don't use Biff8Word here since we don't want to change the table even if it is in the last row.
            LastRow = BitOps.GetWord(aData, 2);
            FirstCol = aData[4];
            LastCol = aData[5];

            OptionFlags = BitOps.GetWord(aData, 6) & ~2; // Calc on load will crash Excel 2010

            RwInpRw = BitOps.GetWord(aData, 8);
            ColInpRw = BitOps.GetWord(aData, 10);
            RwInpCol = BitOps.GetWord(aData, 12);
            ColInpCol = BitOps.GetWord(aData, 14);

            CheckCol(ref RwInpRw, ref ColInpRw, FlxConsts.Max_Columns97_2003, FlagDeleted1);
            CheckCol(ref RwInpCol, ref ColInpCol, FlxConsts.Max_Columns97_2003, FlagDeleted2);
        }

        internal TTableRecord(int aId, int aOptionFlags, int aFirstRow, int aFirstCol, int aLastRow, int aLastCol, int aRwInpRw, int aColInpRw, int aRwInpCol, int aColInpCol)
        {
            Id = aId;
            OptionFlags = aOptionFlags;
            FirstRow = aFirstRow;
            LastRow = aLastRow;
            FirstCol = aFirstCol;
            LastCol = aLastCol;
            RwInpRw = aRwInpRw;
            ColInpRw = aColInpRw;
            RwInpCol = aRwInpCol;
            ColInpCol = aColInpCol;
        }

        internal static TTableRecord CreateFromBiff8(int aId, byte[] aData)
        {
            return new TTableRecord(aId, aData);
        }

		internal override int GetId	{ get {	return Id; }}

        internal void ArrangeCopyRange(int DeltaRow, int DeltaCol)
        {
            IncRef(ref FirstRow, DeltaRow, FlxConsts.Max_Rows, XlsErr.ErrTooManyRows); //Here we raise an error, can't insert past the bound of a sheet.
            IncRef(ref LastRow, DeltaRow, FlxConsts.Max_Rows, XlsErr.ErrTooManyRows);
            IncRef(ref FirstCol, DeltaCol, FlxConsts.Max_Columns, XlsErr.ErrTooManyColumns);
            IncRef(ref LastCol, DeltaCol, FlxConsts.Max_Columns, XlsErr.ErrTooManyColumns);

            IncRowToMax(ref RwInpRw, ref ColInpRw, DeltaRow, FlagDeleted1);  //here, we create an invalid ref
            IncRowToMax(ref RwInpCol, ref ColInpCol, DeltaRow, FlagDeleted2);
            IncColToMax(ref RwInpRw, ref ColInpRw, DeltaCol, FlagDeleted1);
            IncColToMax(ref RwInpCol, ref ColInpCol, DeltaCol, FlagDeleted2);
        }

        internal void ArrangeInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount)
        {
            //Increment the position of the table. Here we give an error if we pass the maximum value, or we would be losing data
            if (CellRange.HasCol(FirstCol) && CellRange.HasCol(LastCol))
            {
                if (FirstRow >= CellRange.Top) IncRef(ref FirstRow, aRowCount, FlxConsts.Max_Rows, XlsErr.ErrTooManyRows);
                if (LastRow >= CellRange.Top) IncRef(ref LastRow, aRowCount, FlxConsts.Max_Rows, XlsErr.ErrTooManyRows);
            }
            if (CellRange.HasRow(FirstRow) && CellRange.HasRow(LastRow))
            {
                if (FirstCol >= CellRange.Left) IncRef(ref FirstCol, aColCount, FlxConsts.Max_Columns, XlsErr.ErrTooManyColumns);
                if (LastCol >= CellRange.Left) IncRef(ref LastCol, aColCount, FlxConsts.Max_Columns, XlsErr.ErrTooManyColumns);
            }

            //Increment the entry cells. If they go out of limits, we should replace them with #ref
            if (CellRange.HasCol(ColInpRw))
            {
                if (RwInpRw >= CellRange.Top) IncRowToMax(ref RwInpRw, ref ColInpRw, aRowCount, FlagDeleted1);
            }
            if (CellRange.HasCol(ColInpCol))
            {
                if (RwInpCol >= CellRange.Top) IncRowToMax(ref RwInpCol, ref ColInpCol, aRowCount, FlagDeleted2);
            }

            if (CellRange.HasRow(RwInpRw))
            {
                if (ColInpRw >= CellRange.Left) IncColToMax(ref RwInpRw, ref ColInpRw, aRowCount, FlagDeleted1);
            }
            if (CellRange.HasRow(RwInpCol))
            {
                if (ColInpCol >= CellRange.Left) IncColToMax(ref RwInpCol, ref ColInpCol, aRowCount, FlagDeleted2);
            }
        }

        internal void ArrangeMoveRange(TXlsCellRange CellRange, int NewRow, int NewCol)
        {
            //Increment the position of the table. Here we give an error if we pass the maximum value, or we would be losing data
            if (CellRange.HasCol(FirstCol) && CellRange.HasCol(LastCol) && CellRange.HasRow(FirstRow) && CellRange.HasRow(LastRow))
            {
                IncRef(ref FirstRow, NewRow - CellRange.Top, FlxConsts.Max_Rows, XlsErr.ErrTooManyRows);
                IncRef(ref LastRow, NewRow - CellRange.Top, FlxConsts.Max_Rows, XlsErr.ErrTooManyRows);
                IncRef(ref FirstCol, NewCol - CellRange.Left, FlxConsts.Max_Columns, XlsErr.ErrTooManyColumns);
                IncRef(ref LastCol, NewCol - CellRange.Left, FlxConsts.Max_Columns, XlsErr.ErrTooManyColumns);
            }
            else
            {
                if (CellRange.HasCol(FirstCol) || CellRange.HasCol(LastCol) || CellRange.HasRow(FirstRow) || CellRange.HasRow(LastRow))
                {
                    XlsMessages.ThrowException(XlsErr.ErrCantMovePartOfTable);
                }
            }

            //Increment the entry cells. If they go out of limits, we should replace them with #ref
            if (CellRange.HasRow(RwInpRw) && CellRange.HasCol(ColInpRw))
            {
                IncRowToMax(ref RwInpRw, ref ColInpRw, NewRow - CellRange.Top, FlagDeleted1);
                IncColToMax(ref RwInpRw, ref ColInpRw, NewCol - CellRange.Left, FlagDeleted1);
            }
            if (CellRange.HasRow(RwInpCol) && CellRange.HasCol(ColInpCol))
            {
                IncRowToMax(ref RwInpCol, ref ColInpCol, NewRow - CellRange.Top, FlagDeleted2);
                IncColToMax(ref RwInpCol, ref ColInpCol, NewCol - CellRange.Left, FlagDeleted2);
            }
        }

        private void IncRef(ref int w, int Delta, int Max, XlsErr ErrWhenTooMany)
        {
            w += Delta;
            if ((w < 0) || (w > Max)) XlsMessages.ThrowException(ErrWhenTooMany, w + 1, Max + 1);
        }

        private void CheckCol(ref int Rw, ref int Col, int Max, int Flag)
        {
            if ((Col > Max) || (Col < 0) || ((OptionFlags & Flag) != 0)) { Rw = -1; Col = -1; OptionFlags |= Flag; }  //Invalid ref
        }
        private void IncColToMax(ref int Rw, ref int Col, int Offset, int Flag)
        {
            if (Col < 0) return;
            Col += Offset;
            CheckCol(ref Rw, ref Col, FlxConsts.Max_Columns, Flag);
        }

        private void CheckRow(ref int Rw, ref int Col, int Max, int Flag)
        {
            if ((Rw > Max) || (Rw < 0) || ((OptionFlags & Flag) != 0)) { Rw = -1; Col = -1; OptionFlags |= Flag; }  //Invalid ref
        }

        private void IncRowToMax(ref int Rw, ref int Col, int Offset, int Flag)
        {
            if (Col < 0) return; //we test always the col
            Rw += Offset;
            CheckRow(ref Rw, ref Col, FlxConsts.Max_Rows, Flag);
        }

        internal bool IsDeleted1 { get { return ColInpRw < 0; } }
        internal bool IsDeleted2 { get { return ColInpCol < 0; } }

		internal bool Has2Entries { get { return (OptionFlags & 0x08) != 0; } }
		internal bool CellInputIsRow { get { return (OptionFlags & 0x04) != 0; } }

        protected override TBaseRecord DoCopyTo(TSheetInfo SheetInfo)
        {
            return (TBaseRecord)MemberwiseClone();
        }

        internal override int TotalSizeNoHeaders()
        {
            return 16;
        }
        internal override int TotalSize()
        {
            return TotalSizeNoHeaders() + XlsConsts.SizeOfTRecordHeader;
        }
        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
        {
            Workbook.WriteHeader((UInt16)Id, (UInt16)TotalSizeNoHeaders());
            Workbook.WriteRow(FirstRow);
            Workbook.WriteRow(LastRow);
            Workbook.WriteColByte(FirstCol);
            Workbook.WriteColByte(LastCol);

            int Flags = OptionFlags;
            try
            {
                int rw1 = RwInpRw;
                int cl1 = ColInpRw;
                CheckRow(ref rw1, ref cl1, FlxConsts.Max_Rows97_2003, FlagDeleted1);
                CheckCol(ref rw1, ref cl1, FlxConsts.Max_Columns97_2003, FlagDeleted1);

                int rw2 = RwInpCol;
                int cl2 = ColInpCol;
                CheckRow(ref rw2, ref cl2, FlxConsts.Max_Rows97_2003, FlagDeleted2);
                CheckCol(ref rw2, ref cl2, FlxConsts.Max_Columns97_2003, FlagDeleted2);

                unchecked
                {
                    Workbook.Write16((UInt16)OptionFlags);
                    Workbook.Write16((UInt16)rw1);
                    Workbook.Write16((UInt16)cl1);
                    Workbook.Write16((UInt16)rw2);
                    Workbook.Write16((UInt16)cl2);
                }
            }
            finally
            {
                OptionFlags = Flags;
            }
        }
    }

    internal class TArrayRecord : TBaseRecord
    {
        internal TParsedTokenList Data;
        internal UInt16 Id;
        internal object[,] ArrResult;
        internal bool Recalculated;
        internal bool Recalculating;

        private int FirstRow, LastRow, FirstColumn, LastColumn;

        byte[] Chn;
        internal UInt16 OptionFlags;

        private TArrayRecord(TNameRecordList Names, int aId, byte[] aData)
            : base()
        {
            Id = (UInt16)aId;

            FirstRow = BitOps.GetWord(aData, 0);
            LastRow = BitOps.GetWord(aData, 2);
            FirstColumn = aData[4];
            LastColumn = aData[5];

            OptionFlags = (UInt16)BitOps.GetWord(aData, 6);
            Chn = new byte[4];
            Array.Copy(aData, 8, Chn, 0, Chn.Length);

            bool HasSubtotal; bool HasAggregate;
            Data = TTokenManipulator.CreateFromBiff8(Names, -1, -1, aData, 14, BitOps.GetWord(aData, 12), false, out HasSubtotal, out HasAggregate);
        }

        internal static TArrayRecord CreateFromBiff8(TNameRecordList Names, int aId, byte[] aData)
        {
            return new TArrayRecord(Names, aId, aData);
        }

        /// <summary>
        /// This method will *not* clone data.
        /// </summary>
        internal TArrayRecord(int aId, TXlsCellRange FmlaArrayRange, TParsedTokenList aData, int aOptionFlags)
            : base()
        {

            Data = aData;
            Id = (UInt16)aId;

            FirstRow = FmlaArrayRange.Top;
            LastRow = FmlaArrayRange.Bottom;
            FirstColumn = FmlaArrayRange.Left;
            LastColumn = FmlaArrayRange.Right;

            OptionFlags = (UInt16)aOptionFlags;
        }

        internal override int GetId
        {
            get { return Id; }
        }

        internal void ArrangeCopyRange(int Row, int Col, int RowOffset, int ColOffset, TSheetInfo SheetInfo)
        {
            ArrangeTokensInsertRange(Row, Col, new TXlsCellRange(0, 0, -1, -1), 0, 0, RowOffset, ColOffset, SheetInfo); //Sheet info doesn't have meaning on copy
        }

        internal void ArrangeInsertRange(int Row, int Col, TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
        {
            ArrangeTokensInsertRange(Row, Col, CellRange, aRowCount, aColCount, 0, 0, SheetInfo);
        }

        private void ArrangeTokensInsertRange(int Row, int Col, TXlsCellRange CellRange, int aRowCount, int aColCount, int CopyRowOffset, int CopyColOffset,
            TSheetInfo SheetInfo)
        {
            try
            {
                TTokenManipulator.ArrangeInsertAndCopyRange(Data, CellRange, Row, Col, aRowCount, aColCount, CopyRowOffset, CopyColOffset, SheetInfo, true, null);
            }
            catch (ETokenException e)
            {
                XlsMessages.ThrowException(XlsErr.ErrBadFormula, Row + 1, Col + 1, e.Token);
            }
        }

        internal void ArrangeMoveRange(int Row, int Col, TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
        {
            ArrangeTokensMoveRange(Row, Col, CellRange, NewRow, NewCol, SheetInfo);
        }

        private void ArrangeTokensMoveRange(int Row, int Col, TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
        {
            try
            {
                TTokenManipulator.ArrangeMoveRange(Data, CellRange, Row, Col, NewRow, NewCol, SheetInfo, null);
            }
            catch (ETokenException e)
            {
                XlsMessages.ThrowException(XlsErr.ErrBadFormula, Row + 1, Col + 1, e.Token);
            }
        }

        internal int RowCount
        {
            get
            {
                return LastRow - FirstRow + 1;
            }
        }

        internal int ColCount
        {
            get
            {
                return LastColumn - FirstColumn + 1;
            }
        }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            if (Loader.LastFormula == null) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
            Loader.LastFormula.ArrayRecord = this;
        }


        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
        {
            //we need the column.
            FlxMessages.ThrowException(FlxErr.ErrInternal);
        }

        internal void SaveToStream(IDataStream Workbook, int Row, int Col, TSaveData SaveData)
        {
            Workbook.WriteHeader((UInt16)Id, (UInt16)TotalSizeNoHeaders());
            Workbook.Write16((UInt16)Row);
            Workbook.Write16((UInt16)(Row + RowCount - 1));
            Workbook.Write16((UInt16)(Col + ((Col + ColCount - 1) << 8)));

            Workbook.Write16(OptionFlags);
            if (Recalculated || Chn == null)  //If recalc=manual, chn has been cleared. If recalc=auto and !recalculated -> it is the original formula.
            {
                Workbook.Write(new byte[4], 4); //Cleared chn.
            }
            else
            {
                Workbook.Write(Chn, Chn.Length);
            }
            TTokenManipulator.SaveToStream(SaveData.Globals.Names, Workbook, TFormulaType.Normal, Data, true);
        }

        protected override TBaseRecord DoCopyTo(TSheetInfo SheetInfo)
        {
            TArrayRecord Result = (TArrayRecord)MemberwiseClone();
            Result.Data = Data.Clone();
            Result.Recalculated = false;
            Result.Recalculating = false;
            Result.ArrResult = null;
            return Result;
        }


        private void SetArrResult(TFlxFormulaErrorValue Err)
        {
            ArrResult = new object[1, 1];
            ArrResult[0, 0] = Err;
        }

        internal object[,] GetValueAndRecalc(TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack)
        {
            if (Recalculated) return ArrResult;
            if (Recalculating) SetArrResult(TFlxFormulaErrorValue.ErrRef);

            Recalculating = true;
            try
            {
                object Result = Data.EvaluateAll(wi, CalcState, CalcStack);
                ArrResult = Result as object[,];
                if (ArrResult == null)
                {
                    ArrResult = new object[1, 1];
                    ArrResult[0, 0] = Result;
                }
            }
            catch (FlexCelException)
            {
                SetArrResult(TFlxFormulaErrorValue.ErrNA);
            }
            catch (FormatException)
            {
                SetArrResult(TFlxFormulaErrorValue.ErrValue);
            }
            catch (ArithmeticException)
            {
                SetArrResult(TFlxFormulaErrorValue.ErrValue);
            }
            catch (ArgumentOutOfRangeException)
            {
                SetArrResult(TFlxFormulaErrorValue.ErrNum);
            }

            Recalculating = false;
            Recalculated = true;
            return ArrResult;
        }

        internal override int TotalSizeNoHeaders()
        {
            int Result = 14 + TTokenManipulator.TotalSizeWithArray(Data, TFormulaType.Normal);
            return Result;
        }

        internal override int TotalSize()
        {
            int Result = TotalSizeNoHeaders() + 4;
            return Result;
        }

    }

	internal class TBiff8ShrFmlaRecord : TxBaseRecord
	{
        internal TBiff8ShrFmlaRecord(int aId, byte[] aData) : base(aId, aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            if (Loader.LastFormula == null) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
            UInt64 K;
            if (!Loader.LastFormula.IsExp(out K)) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
            
            TSharedFormula SharedFmla = new TSharedFormula(RecordLoader.Names, K, Data, 10, GetWord(8));  //Note that we use K from last record here, not Key from this shared formula record. Key might be wrong.
			Loader.ShrFmlas.Add(SharedFmla);

			Loader.LastFormula.MixShared(SharedFmla.Data, rRow, true);
        }

    }

    /// <summary>
    /// To be able to easily know if a formula applies to an insertandcopy.
    /// </summary>
    internal class TFormulaBounds
    {
        internal int Sheet1;
        internal int Sheet2;
        internal int Row1;
        internal int Row2;
        internal int Col1;
        internal int Col2;

        internal TFormulaBounds()
        {
            Clear();
        }

        internal void Clear()
        {
            Sheet1 = -1;
            Sheet2 = -1;
            Row1 = -1;
            Row2 = -1;
            Col1 = -1;
            Col2 = -1;
        }

        internal void AddSheet(int Sheet)
        {
            if (Sheet1 == -1)
            {
                Sheet1 = Sheet;
                Sheet2 = Sheet;
            }
            else
            {
                if (Sheet < Sheet1) Sheet1 = Sheet;
                if (Sheet > Sheet2) Sheet2 = Sheet;
            }
        }

        internal void AddSheets(TSheetRange Sheets)
        {
            for (int i = Sheets.FirstSheet; i <= Sheets.LastSheet; i++)
            {
                AddSheet(i);
            }
        }

        internal void AddRow(int Row)
        {
            if (Row1 == -1)
            {
                Row1 = Row;
                Row2 = Row;
            }
            else
            {
                if (Row < Row1) Row1 = Row;
                if (Row > Row2) Row2 = Row;
            }
        }

        internal void AddCol(int Col)
        {
            if (Col1 == -1)
            {
                Col1 = Col;
                Col2 = Col;
            }
            else
            {
                if (Col < Col1) Col1 = Col;
                if (Col > Col2) Col2 = Col;
            }
        }


        internal bool OutBounds(TXlsCellRange CellRange, TSheetInfo SheetInfo, int aRowCount, int aColCount)
        {
            if (Sheet1 == -2) return false; //Should not be ignored. It contains a formula array.
            if (SheetInfo.InsSheet != SheetInfo.SourceFormulaSheet)
            {
                if (Sheet1 == -1)
                {
                    return true; // this formula is local to the sheet, and we are inserting in other sheet.
                }
                else
                {
                    if (Sheet1 > SheetInfo.InsSheet || Sheet2 < SheetInfo.InsSheet) return true; //This formula does not reference the inserted sheet.
                }
            }

            if (aColCount == 0) //Inserting rows.
            {
                if (CellRange.Top > Row2) return true;
                if (CellRange.Left > Col2) return true;
                if (CellRange.Right < Col1) return true;
            }
            if (aRowCount == 0) //Inserting cols.
            {
                if (CellRange.Left > Col2) return true;
                if (CellRange.Top > Row2) return true;
                if (CellRange.Bottom < Row1) return true;
            }

            return false;
        }
    }

}
