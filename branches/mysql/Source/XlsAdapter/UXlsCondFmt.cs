using System;
using FlexCel.Core;
using System.Diagnostics;

namespace FlexCel.XlsAdapter
{
    /// <summary>
    /// Conditional format. This records holds a lot of CFs together.
    /// </summary>
    internal class TCondFmtRecord: TRangeRecord
    {
        internal TCondFmtRecord(int aId, byte[] aData): base(aId, aData){}

		internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
		{
            TCondFmt cf = new TCondFmt();
            ws.ConditionalFormats.Add(cf);
            cf.LoadFromStream(RecordLoader, this);
		}

    }

	internal struct TCellCondFmt: IComparable
	{
		internal int C1;
		internal int C2;
		internal TCondFmt Fmt;

		internal TCellCondFmt(int aC1, int aC2, TCondFmt aFmt)
		{
			C1 = aC1;
			C2 = aC2;
			Fmt = aFmt;
		}

		public int CompareTo(object o)
		{
			if (!(o is TCellCondFmt)) return 1;

			return C1.CompareTo(((TCellCondFmt)o).C1);
		}
	}

	internal class TCFRecord: TBaseRecord
    {
        #region Privates
        private int Id;
        private bool UsesFormula;
        private TConditionType ConditionType;
        //private bool StopIfTrue;
		private TParsedTokenList Fmla1 = null;
		private TParsedTokenList Fmla2 = null;

		private TConditionalFormatDef Fmt;
		private byte[] Biff8Format; //Used to keep exactly the same settings... there are too many undocumented things here and we don't want to lose them.

        #endregion

        #region Constructors
        private TCFRecord(int aId)
            : base()
        {
            Id = aId;
        }

        private TCFRecord(TNameRecordList Names, int aId, byte[] aData)
            : this(aId) 
        { 
            LoadBiff8(Names, aData);
        }

        internal static TCFRecord LoadFromBiff8(TNameRecordList Names, int aId, byte[] aData)
        {
            return new TCFRecord(Names, aId, aData);
        }

		internal override int GetId	{ get {	return Id; }}

        #endregion

        #region ToBiff8
        private static long GetOptionFlags(TConditionalFormatDefStandard Fmt)
        {
            if (!Fmt.HasFormat) return 0;
            int Result = 0x3FF;
            if (!Fmt.ApplyBorderLeft) Result |= 0x400;
            if (!Fmt.ApplyBorderRight) Result |= 0x800;
            if (!Fmt.ApplyBorderTop) Result |= 0x1000;
            if (!Fmt.ApplyBorderBottom) Result |= 0x2000;

            Result |= 0xC000;

            if (!Fmt.ApplyPatternStyle) Result |= 0x10000;
            if (!Fmt.ApplyPatternFg) Result |= 0x20000;
            if (!Fmt.ApplyPatternBg) Result |= 0x40000;

            Result |= 0x380000;

            if (Fmt.HasFontBlock) Result |= 0x4000000;
            if (Fmt.HasBorderBlock) Result |= 0x10000000;
            if (Fmt.HasPatternBlock) Result |= 0x20000000;

            return Result;
        }

        private static int SaveBiff8Format(IFlexCelPalette xls, TConditionalFormatDefStandard StdFmt, bool IncludeFontBlock, bool IncludeBorderBlock, bool IncludePatternBlock, Byte[] NewData)
        {
            BitOps.SetCardinal(NewData, 6, GetOptionFlags(StdFmt));

            int Pos = 12;
            if (IncludeFontBlock)
            {
                if (StdFmt.ApplyFontSize20)
                {
                    BitOps.SetCardinal(NewData, Pos + 64, StdFmt.FontSize20);
                }
                else
                {
                    BitOps.SetCardinal(NewData, Pos + 64, 0xFFFFFFFF);
                }

                if (StdFmt.ApplyFontStyleBoldAndItalic)
                {
                    if ((StdFmt.FontStyle & TFlxFontStyles.Italic) != 0) NewData[Pos + 68] = 0x2;
                    if ((StdFmt.FontStyle & TFlxFontStyles.Bold) != 0) BitOps.SetWord(NewData, Pos + 72, 0x2BC); else BitOps.SetWord(NewData, Pos + 72, 0x190);
                }

                if (StdFmt.ApplyFontStyleStrikeout)
                {
                    if ((StdFmt.FontStyle & TFlxFontStyles.StrikeOut) != 0) NewData[Pos + 68] |= 0x80;
                }

                if (StdFmt.ApplyFontStyleSubSuperscript)
                {
                    if ((StdFmt.FontStyle & TFlxFontStyles.Superscript) != 0) NewData[Pos + 74] = 0x1;
                    if ((StdFmt.FontStyle & TFlxFontStyles.Subscript) != 0) NewData[Pos + 74] = 0x2;
                }

                if (StdFmt.ApplyFontUnderline)
                {
                    switch (StdFmt.FontUnderline)
                    {
                        case TFlxUnderline.Single: NewData[Pos + 76] = 0x1; break;
                        case TFlxUnderline.Double: NewData[Pos + 76] = 0x2; break;
                        case TFlxUnderline.SingleAccounting: NewData[Pos + 76] = 0x21; break;
                        case TFlxUnderline.DoubleAccounting: NewData[Pos + 76] = 0x22; break;
                    }
                }

                if (StdFmt.ApplyFontColor)
                {
                    BitOps.SetCardinal(NewData, Pos + 80, StdFmt.FontColor.GetBiff8ColorIndex(xls, TAutomaticColor.Font));
                }
                else
                    BitOps.SetCardinal(NewData, Pos + 80, 0xFFFFFFFF);

                NewData[Pos + 88] = 0x18;
                if (!StdFmt.ApplyFontStyleBoldAndItalic) NewData[Pos + 88] |= 0x2;
                if (!StdFmt.ApplyFontStyleStrikeout) NewData[Pos + 88] |= 0x80;

                if (!StdFmt.ApplyFontStyleSubSuperscript) NewData[Pos + 92] = 0x1;
                if (!StdFmt.ApplyFontUnderline) NewData[Pos + 96] = 0x1;
                NewData[Pos + 116] = 0x1;


                Pos += 118;
            }

            if (IncludeBorderBlock)
            {
                int b1 = 0;
                if (StdFmt.ApplyBorderLeft) b1 |= (int)StdFmt.BorderLeft.Style & 0xF;
                if (StdFmt.ApplyBorderRight) b1 |= ((int)StdFmt.BorderRight.Style & 0xF) << 4;
                NewData[Pos] = (byte)b1;

                int b2 = 0;
                if (StdFmt.ApplyBorderTop) b2 |= ((int)StdFmt.BorderTop.Style & 0xF) << 0;
                if (StdFmt.ApplyBorderBottom) b2 |= ((int)StdFmt.BorderBottom.Style & 0xF) << 4;
                NewData[Pos + 1] = (Byte)b2;

                int LRColors = 0;
                if (StdFmt.ApplyBorderLeft) LRColors = StdFmt.BorderLeft.Color.GetBiff8ColorIndex(xls, TAutomaticColor.DefaultForeground);
                if (StdFmt.ApplyBorderRight) LRColors |= StdFmt.BorderRight.Color.GetBiff8ColorIndex(xls, TAutomaticColor.DefaultForeground) << 7;

                BitOps.SetWord(NewData, Pos + 2, LRColors);

                int TBColors = 0;
                if (StdFmt.ApplyBorderTop) TBColors = StdFmt.BorderTop.Color.GetBiff8ColorIndex(xls, TAutomaticColor.DefaultForeground);
                if (StdFmt.ApplyBorderBottom) TBColors |= StdFmt.BorderBottom.Color.GetBiff8ColorIndex(xls, TAutomaticColor.DefaultForeground) << 7;

                BitOps.SetWord(NewData, Pos + 4, TBColors);

                Pos += 8;
            }

            if (IncludePatternBlock)
            {
                if (StdFmt.ApplyPatternStyle) NewData[Pos + 1] = (byte)((((uint)StdFmt.PatternStyle - 1) << 2) & 0xFC);
                int PatColor = 0;
                if (StdFmt.ApplyPatternFg) PatColor = (StdFmt.PatternFgColor.GetBiff8ColorIndex(xls, TAutomaticColor.DefaultForeground) & 0x7F);
                if (StdFmt.ApplyPatternBg) PatColor += (((StdFmt.PatternBgColor.GetBiff8ColorIndex(xls, TAutomaticColor.DefaultBackground) & 0x7F)) << 7);
                BitOps.SetWord(NewData, Pos + 2, PatColor);
                Pos += 4;
            }
            return Pos;
        }

        private byte[] ToBiff8(TNameRecordList Names, IFlexCelPalette xls, bool OnlyCalcLen, out int Len)
        {
            Len = 0;
            byte[] F1 = null;
            byte[] F2 = null;
            TConditionalFormatDefStandard StdFmt = null;
            byte ConditionTypeB = 0;
            byte CFType = 0;

            int Fmla1LenNoArray = 0;
            int Fmla2LenNoArray = 0;

            if (!UsesFormula)
            {
                if (Fmla1 != null)
                {
                    F1 = TFormulaConvertInternalToBiff8.GetTokenData(Names, Fmla1, TFormulaType.CondFmt, out Fmla1LenNoArray);
                }
                if (Fmla2 != null)
                {
                    F2 = TFormulaConvertInternalToBiff8.GetTokenData(Names, Fmla2, TFormulaType.CondFmt, out Fmla2LenNoArray);
                }

                CFType = 1;
                unchecked
                {
                    ConditionTypeB = (byte)ConditionType;
                }
            }
            else
            {
                if (Fmla1 != null)
                {
                    F1 = TFormulaConvertInternalToBiff8.GetTokenData(Names, Fmla1, TFormulaType.CondFmt, out Fmla1LenNoArray);
                }
                CFType = 2;
            }

			StdFmt = Fmt as TConditionalFormatDefStandard;
			if (CFType == 0) return null; //no valid cf.
			bool IncludeFontBlock = StdFmt.HasFontBlock;
            bool IncludeBorderBlock = StdFmt.HasBorderBlock;
            bool IncludePatternBlock = StdFmt.HasPatternBlock;

			if (Biff8Format == null)
			{
				int VariableSize = 0;
				if (F1 != null) VariableSize += F1.Length;
				if (F2 != null) VariableSize += F2.Length;
				if (IncludeFontBlock) VariableSize += 118;
				if (IncludeBorderBlock) VariableSize += 8;
				if (IncludePatternBlock) VariableSize += 4;

				Len = 6 + 6 + VariableSize;
			}
			else
			{
				Len = 6 + Biff8Format.Length;
				if (F1 != null) Len += F1.Length;
				if (F2 != null) Len += F2.Length;
			}

            if (OnlyCalcLen) return null;

            Byte[] NewData = new byte[Len];
            NewData[0] = CFType;
            NewData[1] = ConditionTypeB;
            if (F1 != null) BitOps.SetWord(NewData, 2, Fmla1LenNoArray);  
            if (F2 != null) BitOps.SetWord(NewData, 4, Fmla2LenNoArray);

            int Pos;
            if (Biff8Format != null)
            {
                Pos = 6 + Biff8Format.Length;
                Array.Copy(Biff8Format, 0, NewData, 6, Biff8Format.Length);
            }
            else
            {
                Pos = SaveBiff8Format(xls, StdFmt, IncludeFontBlock, IncludeBorderBlock, IncludePatternBlock, NewData);
            }

            if (F1 != null)
            {
                Array.Copy(F1, 0, NewData, Pos, F1.Length);
                Pos += F1.Length;
            }

            if (F2 != null)
            {
                Array.Copy(F2, 0, NewData, Pos, F2.Length);
                Pos += F2.Length;
            }
            return NewData;
        }

        #endregion

        #region From Biff8
        private static void LoadFontBlock(byte[] Data, TConditionalFormatDefStandard Fmt)
        {
            int Pos = FontBlockPos;
            long FontSize = BitOps.GetCardinal(Data, Pos + 64);
            if (FontSize != 0xFFFFFFFF)
            {
                Fmt.ApplyFontSize20 = true;
                Fmt.FontSize20 = (int)FontSize;
            }

            long FontOptions = BitOps.GetCardinal(Data, Pos + 68);
            long FontOptionFlags = BitOps.GetCardinal(Data, Pos + 88);

            if ((FontOptionFlags & 0x2) == 0)  //Font style  modified
            {
                Fmt.ApplyFontStyleBoldAndItalic = true;

                if ((FontOptions & 0x2) != 0)
                    Fmt.FontStyle |= TFlxFontStyles.Italic;

                int FontWeight = BitOps.GetWord(Data, Pos + 72);
                if (FontWeight > 600)
                    Fmt.FontStyle |= TFlxFontStyles.Bold;

            }

            if ((FontOptionFlags & 0x80) == 0)  //Font strikeout modified
            {
                Fmt.ApplyFontStyleStrikeout = true;
                if ((FontOptions & 0x80) != 0)
                    Fmt.FontStyle |= TFlxFontStyles.StrikeOut;
            }

            long SubSuperScriptModified = BitOps.GetCardinal(Data, Pos + 92);
            if (SubSuperScriptModified == 0)  //Font sub/superscript modified
            {
                Fmt.ApplyFontStyleSubSuperscript = true;
                int FontSubSuper = BitOps.GetWord(Data, Pos + 74);
                if (FontSubSuper == 0x01)
                    Fmt.FontStyle |= TFlxFontStyles.Superscript;

                if (FontSubSuper == 0x02)
                    Fmt.FontStyle |= TFlxFontStyles.Subscript;
            }

            Fmt.ApplyFontUnderline = BitOps.GetCardinal(Data, Pos + 96) == 0;
            if (Fmt.ApplyFontUnderline)
            {
                byte FontUnderline = Data[Pos + 76];
                switch (FontUnderline)
                {
                    case 0x01: Fmt.FontUnderline = TFlxUnderline.Single; break;
                    case 0x02: Fmt.FontUnderline = TFlxUnderline.Double; break;
                    case 0x21: Fmt.FontUnderline = TFlxUnderline.SingleAccounting; break;
                    case 0x22: Fmt.FontUnderline = TFlxUnderline.DoubleAccounting; break;
                    default: Fmt.FontUnderline = TFlxUnderline.None; break;
                }//case
            }

            long FontColorIndex = BitOps.GetCardinal(Data, Pos + 80);
            Fmt.ApplyFontColor = (FontColorIndex != 0xFFFFFFFF);
            if (Fmt.ApplyFontColor) Fmt.FontColor = TExcelColor.FromBiff8ColorIndex((int)FontColorIndex);

        }

        private void LoadBorderBlock(byte[] Data, TConditionalFormatDefStandard Fmt)
        {
            int Pos = BorderBlockPos(Data);
            long Flags = OptionFlags(Data);

            byte LeftRight = Data[Pos + 0];
            byte TopBottom = Data[Pos + 1];
            long ColorIndexes = BitOps.GetCardinal(Data, Pos + 2);

            if ((Flags & 0x400) == 0)  //left border
            {
                Fmt.ApplyBorderLeft = true;
                Fmt.BorderLeft = new TFlxOneBorder((TFlxBorderStyle)(LeftRight & 0xF), TExcelColor.FromBiff8ColorIndex(((ColorIndexes >> 0) & 0x7F)));
            }

            if ((Flags & 0x800) == 0)  //right border
            {
                Fmt.ApplyBorderRight = true;
                Fmt.BorderRight = new TFlxOneBorder((TFlxBorderStyle)((LeftRight >> 4) & 0xF), TExcelColor.FromBiff8ColorIndex(((ColorIndexes >> 7) & 0x7F)));
            }

            if ((Flags & 0x1000) == 0) //top border
            {
                Fmt.ApplyBorderTop = true;
                Fmt.BorderTop = new TFlxOneBorder((TFlxBorderStyle)(TopBottom & 0xF), TExcelColor.FromBiff8ColorIndex(((ColorIndexes >> 16) & 0x7F)));
            }

            if ((Flags & 0x2000) == 0)  //bottom border
            {
                Fmt.ApplyBorderBottom = true;
                Fmt.BorderBottom = new TFlxOneBorder((TFlxBorderStyle)((TopBottom >> 4) & 0xF), TExcelColor.FromBiff8ColorIndex(((ColorIndexes >> 23) & 0x7F)));
            }
        }

        private void LoadPatternBlock(byte[] Data, TConditionalFormatDefStandard Fmt)
        {
            int Pos = PatternBlockPos(Data);
            long Flags = OptionFlags(Data);
            int ColorIndexes = BitOps.GetWord(Data, Pos + 2);

            if ((Flags & 0x10000) == 0)  //pattern style
            {
                Fmt.ApplyPatternStyle = true;
                int CellPattern = Data[Pos + 1] >> 2;
                Fmt.PatternStyle = (TFlxPatternStyle)(CellPattern + 1);
            }

            if ((Flags & 0x20000) == 0)  //pattern fg color
            {
                Fmt.ApplyPatternFg = true;
                Fmt.PatternFgColor = TExcelColor.FromBiff8ColorIndex((ColorIndexes >> 0) & 0x7F);
            }

            if ((Flags & 0x40000) == 0)  //pattern bg color
            {
                Fmt.ApplyPatternBg = true;
                Fmt.PatternBgColor = TExcelColor.FromBiff8ColorIndex((ColorIndexes >> 7) & 0x7F);
            }
        }

        private TConditionalFormatDefStandard CreateConditionalFormatDef(byte[] Data)
        {
            TConditionalFormatDefStandard Result = new TConditionalFormatDefStandard();

            if (HasFontBlock(Data)) LoadFontBlock(Data, Result);
            if (HasBorderBlock(Data)) LoadBorderBlock(Data, Result);
            if (HasPatternBlock(Data)) LoadPatternBlock(Data, Result);

            return Result;
        }

        internal void LoadBiff8(TNameRecordList Names, byte[] Data)
        {
            TFormulaConvertBiff8ToInternal f1 = new TFormulaConvertBiff8ToInternal();
            Fmla1 = Cce1(Data) == 0 ? null : f1.ParseRPN(Names, -1, -1, Data, Fmla1Start(Data), Cce1(Data), true); //no real need for relative since shared formulas can't be 3d, and we only need relative for the non-existing tokens ptgarea3dn and ptgref3dn.

            //this.StopIfTrue = false;
            Fmla2 = null;
            Fmt = CreateConditionalFormatDef(Data);
            Biff8Format = new byte[Data.Length - 6 - Cce1(Data) - Cce2(Data)];
            Array.Copy(Data, 6, Biff8Format, 0, Biff8Format.Length);

            UsesFormula = false;

            switch (CfType(Data))
            {
                case 1:
                    TFormulaConvertBiff8ToInternal f2 = new TFormulaConvertBiff8ToInternal();
                    Fmla2 = Cce2(Data) == 0 ? null : f2.ParseRPN(Names, -1, -1, Data, Fmla2Start(Data), Cce2(Data), true); //no real need for relative since shared formulas can't be 3d, and we only need relative for the non-existing tokens ptgarea3dn and ptgref3dn.
                    ConditionType = (TConditionType)Op(Data);
                    UsesFormula = false;
                    break;
                case 2:
                    UsesFormula = true;
                    break;
            }

        }

        private static byte CfType(byte[] Data) { return Data[0]; }
        private static byte Op(byte[] Data) { return Data[1]; }
        private static int Cce1(byte[] Data) { return BitOps.GetWord(Data, 2); }
        private static int Cce2(byte[] Data) { return BitOps.GetWord(Data, 4); }
        private static long OptionFlags(byte[] Data) { return BitOps.GetCardinal(Data, 6); }

        private static bool HasFontBlock(byte[] Data) { return (OptionFlags(Data) & 0x04000000) != 0; }
        private static bool HasBorderBlock(byte[] Data) { return (OptionFlags(Data) & 0x10000000) != 0; }
        private static bool HasPatternBlock(byte[] Data) { return (OptionFlags(Data) & 0x20000000) != 0; }

        private static int FontBlockPos
        {
            get
            {
                return 12;
            }
        }

        private static int BorderBlockPos(byte[] Data)
        {
            if (HasFontBlock(Data)) return FontBlockPos + 118;
            return FontBlockPos;
        }

        private int PatternBlockPos(byte[] Data)
        {
            if (HasBorderBlock(Data)) return BorderBlockPos(Data) + 8;
            return BorderBlockPos(Data);
        }

        private int Fmla1Start(byte[] Data)
        {
            if (HasPatternBlock(Data)) return PatternBlockPos(Data) + 4;
            return PatternBlockPos(Data);
        }

        private int Fmla2Start(byte[] Data) { return Fmla1Start(Data) + Cce1(Data); }
        #endregion

        #region From API Structure
        internal static TCFRecord LoadFrom(TConditionalFormatRule FmtRule, ExcelFile Xls)
		{
            TCFRecord Result = new TCFRecord((int)xlr.CF);
            Result.Fmt = (TConditionalFormatDef)FmtRule.FormatDef.Clone();
            //Result.StopIfTrue = FmtRule.StopIfTrue;

            bool IsValidRule = false;
			TConditionalCellValueRule CVRule = FmtRule as TConditionalCellValueRule;
			if (CVRule != null)
			{
				if (CVRule.Formula1 != null)
				{
                    TFormulaConvertTextToInternal f = new TFormulaConvertTextToInternal(Xls, Xls.ActiveSheet, true, CVRule.Formula1, true, true);
					f.Parse();
                    Result.Fmla1 = f.GetTokens();
                }
				if (CVRule.Formula2 != null)
				{
                    TFormulaConvertTextToInternal f = new TFormulaConvertTextToInternal(Xls, Xls.ActiveSheet, true, CVRule.Formula2, true, true);
                    f.Parse();
                    Result.Fmla2 = f.GetTokens();
                }

				Result.UsesFormula = false;
                IsValidRule = true;
                Result.ConditionType = CVRule.ConditionType;
			}
			else
			{
				TConditionalFormulaRule FmRule = FmtRule as TConditionalFormulaRule;
				if (FmRule != null)
				{
					if (FmRule.Formula != null)
					{
                        TFormulaConvertTextToInternal f = new TFormulaConvertTextToInternal(Xls, Xls.ActiveSheet, true, FmRule.Formula, true, true);
                        f.Parse();
                        Result.Fmla1 = f.GetTokens();
                    }
					Result.UsesFormula = true;
                    IsValidRule = true;
				}
			}
			if (!IsValidRule) return null; //no valid cf.
            return Result;
        }
        #endregion

        #region To API Structure
        internal TConditionalFormatRule ToCFRule(TCellList CellList)
        {
            string F1 = Fmla1 == null || CellList.Workbook.IgnoreFormulaText ? null : TFormulaConvertInternalToText.AsString(Fmla1, 0, 0, CellList); //We will use always A1 as base for the formulas.
            TConditionalFormatDefStandard FormatDef = (TConditionalFormatDefStandard)Fmt.Clone();

            if (!UsesFormula)
            {
                string F2 = Fmla2 == null || CellList.Workbook.IgnoreFormulaText ? null : TFormulaConvertInternalToText.AsString(Fmla2, 0, 0, CellList);
                return new TConditionalCellValueRule(FormatDef, true, ConditionType, F1, F2);
            }

            return new TConditionalFormulaRule(FormatDef, true, F1);

        }
        #endregion

        #region SaveToStream
        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
        {
            int len;
            byte[] b = ToBiff8(SaveData.Globals.Names, SaveData.Palette, false, out len);	
			if (b == null) return;
			Workbook.WriteHeader((UInt16)Id, (UInt16)b.Length);
			Workbook.Write(b, b.Length);
        }

        internal override int TotalSizeNoHeaders()
        {
            int len;
            ToBiff8(null, null, true, out len);
            return len;
        }

        internal override int TotalSize()
        {
            return TotalSizeNoHeaders() + XlsConsts.SizeOfTRecordHeader;
        }

        protected override TBaseRecord DoCopyTo(TSheetInfo SheetInfo)
        {
            TCFRecord Result = (TCFRecord)MemberwiseClone();
            if (Fmla1 != null) Result.Fmla1 = Fmla1.Clone();
            if (Fmla2 != null) Result.Fmla2 = Fmla2.Clone();
            //There is no need to clone Biff8Format because it is invariant.
            return Result;
        }
        #endregion

        #region Compare
        internal bool EqualsDef(TConditionalFormatRule Fmt, TCellList CellList)
		{
			TConditionalFormatRule Myself = ToCFRule(CellList);
			return Myself.Equals(Fmt);
        }
        #endregion

        #region InsertAndCopy
        private static void ArrangeTokensInsertRange(TParsedTokenList Fmla, TXlsCellRange CellRange, int FmlaRow, int FmlaCol, int aRowCount, int aColCount, int CopyRowOffset, int CopyColOffset, TSheetInfo SheetInfo)
		{
			try
			{
				TTokenManipulator.ArrangeInsertAndCopyRange(Fmla, CellRange, FmlaRow, FmlaCol, aRowCount, aColCount, CopyRowOffset, CopyColOffset, SheetInfo, true, null);
			}			
			catch (ETokenException e)
			{
				XlsMessages.ThrowException(XlsErr.ErrBadCF, e.Token);
			}
		}

		internal void ArrangeInsertRange(TXlsCellRange CellRange, int FmlaRow, int FmlaCol, int aRowCount, int aColCount, TSheetInfo SheetInfo)
		{
			if (Fmla1 != null) ArrangeTokensInsertRange(Fmla1, CellRange, FmlaRow, FmlaCol, aRowCount, aColCount, 0, 0, SheetInfo);
			if (Fmla2 != null) ArrangeTokensInsertRange(Fmla2,  CellRange, FmlaRow, FmlaCol, aRowCount, aColCount, 0, 0, SheetInfo);
		}
        
		internal static void ArrangeCopyRange(int RowOffset, int ColOffset)
		{
			//  No need to arrange anything... ranges are relative to the cells
		}

		private static void ArrangeTokensMoveRange(TParsedTokenList Fmla, TXlsCellRange CellRange, int FmlaRow, int FmlaCol, int NewRow, int NewCol, TSheetInfo SheetInfo)
		{
			try
			{
				TTokenManipulator.ArrangeMoveRange(Fmla, CellRange, FmlaRow, FmlaCol, NewRow, NewCol, SheetInfo, null);
			}			
			catch (ETokenException e)
			{
				XlsMessages.ThrowException(XlsErr.ErrBadCF, e.Token);
			}
		}

		internal void ArrangeMoveRange(TXlsCellRange CellRange, int FmlaRow, int FmlaCol, int NewRow, int NewCol, TSheetInfo SheetInfo)
		{
			if (Fmla1 != null) ArrangeTokensMoveRange(Fmla1, CellRange, FmlaRow, FmlaCol, NewRow, NewCol, SheetInfo);
			if (Fmla2 != null) ArrangeTokensMoveRange(Fmla2,  CellRange, FmlaRow, FmlaCol, NewRow, NewCol, SheetInfo);
        }
        #endregion

        #region Evaluate condition
		private object Fmla1Value(ExcelFile aXls, int aSheetIndex, int RowOffset, int ColOffset)
		{
            if (Fmla1 == null) return TFlxFormulaErrorValue.ErrNA;
            TWorkbookInfo wi = new TWorkbookInfo(aXls, aSheetIndex, 0, 0, 0, 0, RowOffset, ColOffset, true);
			return TFormulaRecord.EvaluateFormula(Fmla1, wi); 
		}

		private object Fmla2Value(ExcelFile aXls, int aSheetIndex, int RowOffset, int ColOffset)
		{
            if (Fmla2 == null) return TFlxFormulaErrorValue.ErrNA;
            TWorkbookInfo wi = new TWorkbookInfo(aXls, aSheetIndex, 0, 0, 0, 0, RowOffset, ColOffset, true);
            return TFormulaRecord.EvaluateFormula(Fmla2, wi);
		}

		private static bool ToBool(object res)
		{
			if (res == null) return false;
			if (res is TFlxFormulaErrorValue) return false;
			bool Result = false;
			try
			{
				if (!TBaseParsedToken.ExtToBool(res, out Result)) return false;
			}
			catch(FormatException)
			{
				return false;
			}
			return Result;
		}

		private static bool Compare(object res1, object CellValue, out int ResultValue)
		{
			ResultValue = 0;
			if (res1 is TFlxFormulaErrorValue) return false;

			object cmp1 = TBaseParsedToken.CompareValues(res1, CellValue);
			if (cmp1 is TFlxFormulaErrorValue) return false;

			ResultValue = Convert.ToInt32(cmp1);
			return true;
		}

		internal bool Evaluate(ExcelFile aXls, int aSheetIndex, int RowOffset, int ColOffset)
		{
			if (UsesFormula)
			{
				object res = Fmla1Value(aXls, aSheetIndex, RowOffset, ColOffset);
				return ToBool(res);
			}

            object CellValue = aXls.GetCellValueAndRecalc(aSheetIndex, RowOffset + 1, ColOffset + 1, new TCalcState(), new TCalcStack());
			if (CellValue is TFlxFormulaErrorValue) return false;

			int Cmp1, Cmp2; 
			switch (ConditionType)
			{
				case TConditionType.Between:
					if (!Compare(Fmla1Value(aXls, aSheetIndex, RowOffset, ColOffset), CellValue, out Cmp1)) return false;
					if (!Compare(Fmla2Value(aXls, aSheetIndex, RowOffset, ColOffset), CellValue, out Cmp2)) return false;

					return Cmp1 * Cmp2 <= 0;

                case TConditionType.NotBetween:
					if (!Compare(Fmla1Value(aXls, aSheetIndex, RowOffset, ColOffset), CellValue, out Cmp1)) return false;
					if (!Compare(Fmla2Value(aXls, aSheetIndex, RowOffset, ColOffset), CellValue, out Cmp2)) return false;

					return Cmp1 * Cmp2 > 0;

                case TConditionType.Equal:
					if (!Compare(Fmla1Value(aXls, aSheetIndex, RowOffset, ColOffset), CellValue, out Cmp1)) return false;
					return Cmp1 == 0;

                case TConditionType.NotEqual:
					if (!Compare(Fmla1Value(aXls, aSheetIndex, RowOffset, ColOffset), CellValue, out Cmp1)) return false;
					return Cmp1 != 0;

                case TConditionType.GreaterThan:
					if (!Compare(Fmla1Value(aXls, aSheetIndex, RowOffset, ColOffset), CellValue, out Cmp1)) return false;
					return Cmp1 < 0;

                case TConditionType.LessThan:
					if (!Compare(Fmla1Value(aXls, aSheetIndex, RowOffset, ColOffset), CellValue, out Cmp1)) return false;
					return Cmp1 > 0;

                case TConditionType.GreaterOrEqual:
					if (!Compare(Fmla1Value(aXls, aSheetIndex, RowOffset, ColOffset), CellValue, out Cmp1)) return false;
					return Cmp1 <= 0;

                case TConditionType.LessOrEqual:
					if (!Compare(Fmla1Value(aXls, aSheetIndex, RowOffset, ColOffset), CellValue, out Cmp1)) return false;
					return Cmp1 >= 0;

			}
			return false;
		}
		#endregion

		#region Modify XF
		internal void ModifyFormat(TFlxFormat Format)
		{
			TConditionalFormatDefStandard StdFmt = Fmt as TConditionalFormatDefStandard;

			if (StdFmt.HasFontBlock)
				ModifyFont(StdFmt, Format.Font);
			if (StdFmt.HasBorderBlock)
				ModifyBorders(StdFmt, Format.Borders);
			if (StdFmt.HasPatternBlock)
				ModifyPattern(StdFmt, ref Format.FillPattern);
		}

		private static void ModifyFont(TConditionalFormatDefStandard Fmt, TFlxFont Font)
		{
			if (Fmt.ApplyFontSize20) Font.Size20 = Fmt.FontSize20;
			if (Fmt.ApplyFontStyleBoldAndItalic)
			{
				if ((Fmt.FontStyle & TFlxFontStyles.Italic) != 0)
					Font.Style |= TFlxFontStyles.Italic;
				else
					Font.Style &= ~TFlxFontStyles.Italic;

				if ((Fmt.FontStyle & TFlxFontStyles.Bold) != 0)
					Font.Style |= TFlxFontStyles.Bold;
				else
					Font.Style &= ~TFlxFontStyles.Bold;
			}

			if (Fmt.ApplyFontStyleStrikeout)  //Font strikeout modified
			{
				if ((Fmt.FontStyle & TFlxFontStyles.StrikeOut) != 0)
					Font.Style |= TFlxFontStyles.StrikeOut;
				else
					Font.Style &= ~TFlxFontStyles.StrikeOut;
			}

			if (Fmt.ApplyFontStyleSubSuperscript)  //Font sub/superscript modified
			{
				if ((Fmt.FontStyle & TFlxFontStyles.Superscript) != 0)
					Font.Style |= TFlxFontStyles.Superscript;
				else
					Font.Style &= ~TFlxFontStyles.Superscript;

				if ((Fmt.FontStyle & TFlxFontStyles.Subscript) != 0)
					Font.Style |= TFlxFontStyles.Subscript;
				else
					Font.Style &= ~TFlxFontStyles.Subscript;
			}

			if (Fmt.ApplyFontUnderline)  //Font underline modified
			{
				Font.Underline = Fmt.FontUnderline;
			}
            
			if (Fmt.ApplyFontColor)
			{
				Font.Color = Fmt.FontColor;
			}

		}

		private static void ModifyBorders(TConditionalFormatDefStandard Fmt, TFlxBorders Borders)
		{
			if (Fmt.ApplyBorderLeft)  //left border
			{
                Borders.Left = Fmt.BorderLeft;
			}

			if (Fmt.ApplyBorderTop)  //top border
			{
				Borders.Top = Fmt.BorderTop;
			}

			if (Fmt.ApplyBorderRight)  //right border
			{
				Borders.Right = Fmt.BorderRight;
			}
			if (Fmt.ApplyBorderBottom)  //bottom border
			{
				Borders.Bottom = Fmt.BorderBottom;
			}
		}

		private static void ModifyPattern (TConditionalFormatDefStandard Fmt, ref TFlxFillPattern Pattern)
		{
			if (Fmt.ApplyPatternStyle) Pattern.Pattern = Fmt.PatternStyle;
			if (Fmt.ApplyPatternFg) Pattern.FgColor = Fmt.PatternFgColor;
			if (Fmt.ApplyPatternBg) Pattern.BgColor = Fmt.PatternBgColor;

			if (!Fmt.ApplyPatternStyle && (Fmt.ApplyPatternFg || Fmt.ApplyPatternBg)) 
			{
				Pattern.Pattern = TFlxPatternStyle.Automatic;
			}

		}
		#endregion

        #region Deleted Ranges
        internal void UpdateDeletedRanges(TDeletedRanges DeletedRanges)
        {
            if (Fmla1 != null) TTokenManipulator.UpdateDeletedRanges(Fmla1, DeletedRanges);
            if (Fmla2 != null) TTokenManipulator.UpdateDeletedRanges(Fmla2, DeletedRanges);
        }
        #endregion
    }

    /// <summary>
    /// A list with CF records.
    /// </summary>
    internal class TCFRecordList : TBaseRecordList<TCFRecord>
    {
		internal void ArrangeInsertRange(TXlsCellRange CellRange, int FmlaRow, int FmlaCol, int aRowCount, int aColCount, TSheetInfo SheetInfo)
		{
			for (int i=0; i< Count;i++)
				this[i].ArrangeInsertRange(CellRange, FmlaRow, FmlaCol, aRowCount, aColCount, SheetInfo);
		}

		internal void ArrangeMoveRange(TXlsCellRange CellRange, int FmlaRow, int FmlaCol, int NewRow, int NewCol, TSheetInfo SheetInfo)
		{
			for (int i=0; i< Count;i++)
				this[i].ArrangeMoveRange(CellRange, FmlaRow, FmlaCol, NewRow, NewCol, SheetInfo);
		}

        internal void UpdateDeletedRanges(TDeletedRanges DeletedRanges)
        {
            for (int i = 0; i < Count; i++)
                this[i].UpdateDeletedRanges(DeletedRanges);
        }


		internal void ConditionallyModifyFormat(TFlxFormat Format, ExcelFile aXls, int aSheetIndex, int RowOffset, int ColOffset)
		{
			for (int i=0; i < Count; i++)
			{
				TCFRecord it = this[i];
				if (it.Evaluate(aXls, aSheetIndex, RowOffset, ColOffset))
				{
					it.ModifyFormat(Format);
					return; // CFs do not add.
				}
			}
		}

		internal void LoadFrom(TConditionalFormatRule[] Fmt, TCellList CellList)
		{
			if (Fmt.Length > XlsConsts.MaxCFRules) XlsMessages.ThrowException(XlsErr.ErrTooManyCFRules);			
			for (int i=0; i < Fmt.Length; i++)
			{
				Add (TCFRecord.LoadFrom(Fmt[i], CellList.Workbook));
			}
		}

    }

    /// <summary>
    /// An internal representation of a Conditional format. 
    /// </summary>
    internal class TCondFmt: TRangeEntry
    {
        private int Flag;
        private TExcelRange AllRange;
        private TCFRecordList CFs;
        
        internal TCondFmt()
        {
            Init();
        }

		internal TCondFmt(int Row1, int Col1, int Row2, int Col2, TConditionalFormatRule[] Fmt, TCellList CellList): this()
		{
			RangeValuesList.AddForced(new TExcelRange(Row1, Col1, Row2, Col2));
			for (int i = 0; i < Fmt.Length; i++)
			{
				if (Fmt[i] is TConditionalFormulaRule) 
				{
					Flag = 1; //needs recalc
					break;
				}
			}
			CFs.LoadFrom(Fmt, CellList);
			FixAllRange();
		}

        private void Init()
        {
            RangeValuesList = new TRangeValuesList(513, 4 + TExcelRange.Length, true, true);
            AllRange=new TExcelRange();
            CFs= new TCFRecordList();
        }

        protected override TRangeEntry DoCopyTo(TSheetInfo SheetInfo)
        {
            TCondFmt Result=(TCondFmt)base.DoCopyTo(SheetInfo);
            Result.Flag=Flag;
            Result.AllRange=(TExcelRange)AllRange.Clone();
            Result.CFs= new TCFRecordList();
            Result.CFs.CopyFrom(CFs, SheetInfo);
            return Result;
        }

		/// <summary>
		/// Debug use only
		/// </summary>
		private void TestAllRange()
		{
		}

		internal void FixAllRange()
		{
			if (RangeValuesList.Count<=0) return;
			AllRange = (TExcelRange)RangeValuesList[0].Clone();


			for (int i = RangeValuesList.Count - 1; i>=0; i--)
			{
				TExcelRange ex = RangeValuesList[i];
				if (ex.C1<AllRange.C1) AllRange.C1 = ex.C1;
				if (ex.C2>AllRange.C2) AllRange.C2 = ex.C2;
				if (ex.R1<AllRange.R1) AllRange.R1 = ex.R1;
				if (ex.R2>AllRange.R2) AllRange.R2 = ex.R2;
			}
		}

        internal void Clear()
        {
            if (CFs!=null) CFs.Clear();
            if (RangeValuesList!=null) RangeValuesList.Clear();
        }

        internal override void LoadFromStream(TBaseRecordLoader RecordLoader, TRangeRecord First)
        {
            Clear();
            TxBaseRecord MyRecord= First;

            int CFCount=BitOps.GetWord(First.Data, 0);
            Flag=BitOps.GetWord(First.Data,2);
            int aPos=4;
            byte[] TempData= new byte[TExcelRange.Length];
            BitOps.ReadMem(ref MyRecord, ref aPos, TempData);
            AllRange.LoadFromBiff8(TempData, true);
            RangeValuesList.LoadFromBiff8(First, aPos);
			FixAllRange(); //The spreadsheet might have not been saved by Excel and have an incorrect AllRange.

            //Load corresponding CFs
            for (int i=0; i< CFCount;i++)
            {
                TBaseRecord R=RecordLoader.LoadRecord(false);
                if (!(R is TCFRecord)) XlsMessages.ThrowException(XlsErr.ErrInvalidCF);
                CFs.Add((TCFRecord)R);
            }

        }

        internal override void SaveToStream(IDataStream DataStream, TSaveData SaveData)
        {
            if (RangeValuesList.Count==0) return; //Don't save empty CF's
            int aCount = RangeValuesList.RepeatCountR(RangeValuesList.Count);
            for (int i=0;i< aCount;i++)
            {
                DataStream.WriteHeader((UInt16)xlr.CONDFMT, (UInt16)RangeValuesList.RecordSizeR(i, RangeValuesList.Count));

                int CFCount= CFs.Count;
                DataStream.Write16((UInt16)CFCount);
                DataStream.Write16((UInt16)Flag);
                DataStream.Write(AllRange.Data(true), TExcelRange.Length);

                RangeValuesList.SaveToStreamR(DataStream, SaveData, i);
                CFs.SaveToStream(DataStream, SaveData,  0);
            }
        }
        
        internal override void SaveRangeToStream(IDataStream DataStream, TSaveData SaveData, TXlsCellRange CellRange)
        {
            int Rc=RangeValuesList.CountRangeRecords(CellRange);
            if (Rc==0) return; //Don't save empty CF's
            int aCount = RangeValuesList.RepeatCountR(Rc);
            for (int i = 0; i < aCount; i++)
            {
                DataStream.WriteHeader((UInt16)xlr.CONDFMT, (UInt16)RangeValuesList.RecordSizeR(i, Rc));

                int CFCount = CFs.Count;
                DataStream.Write16((UInt16)CFCount);
                DataStream.Write16((UInt16)Flag);
                DataStream.Write(AllRange.Data(true), TExcelRange.Length);

                RangeValuesList.SaveRangeToStreamR(DataStream, SaveData, i, Rc, CellRange);
                CFs.SaveToStream(DataStream, SaveData, 0);
            }
        }

        internal override long TotalSize()
        {
            if (RangeValuesList.Count==0) return 0;
            return RangeValuesList.TotalSizeR(RangeValuesList.Count) + CFs.TotalSize*RangeValuesList.RepeatCountR(RangeValuesList.Count);
        }

        internal override long TotalRangeSize(TXlsCellRange CellRange)
        {
            int i= RangeValuesList.CountRangeRecords(CellRange);
            if (RangeValuesList.Count==0) return 0; 
            else
                return RangeValuesList.TotalSizeR(i)
                    + CFs.TotalSize*RangeValuesList.RepeatCountR(i);
        }

        internal override void ArrangeInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
        {
            base.ArrangeInsertRange(CellRange, aRowCount, aColCount, SheetInfo);

            CFs.ArrangeInsertRange(CellRange, AllRange.R1, AllRange.C1, aRowCount, aColCount, SheetInfo );
			FixAllRange();  //As ranges might be inserted only by half a row, it is too complex to calculate what should be included. So we will just adapt this here.
			TestAllRange();
        }

        internal override void InsertAndCopyRange(TXlsCellRange SourceRange, int DestRow, int DestCol, int aRowCount, int aColCount, TRangeCopyMode CopyMode, TFlxInsertMode InsertMode, TSheetInfo SheetInfo)
        {
            bool RangeIntersects = (AllRange.R1 <= SourceRange.Bottom) && (AllRange.R2 >= SourceRange.Top)
                && (AllRange.C1 <= SourceRange.Right) && (AllRange.C2 >= SourceRange.Left);
            base.InsertAndCopyRange(SourceRange, DestRow, DestCol, aRowCount, aColCount, CopyMode, InsertMode, SheetInfo);
            if (RangeIntersects && CopyMode != TRangeCopyMode.None)
                RangeValuesList.CopyRangeInclusive(SourceRange, DestRow, DestCol, aRowCount, aColCount, ref AllRange.R1, ref AllRange.R2, ref AllRange.C1, ref AllRange.C2);
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

			CFs.ArrangeMoveRange(CellRange, AllRange.R1, AllRange.C1, NewRow, NewCol, SheetInfo );
			FixAllRange();  //As ranges might be inserted only by half a row, it is too complex to calculate what should be included. So we will just adapt this here.
			TestAllRange();
		}

        internal override void UpdateDeletedRanges(TDeletedRanges DeletedRanges)
        {
            CFs.UpdateDeletedRanges(DeletedRanges);
        }

		internal TXlsCellRange CellRange1Based()
		{
			return new TXlsCellRange(AllRange.R1+1, AllRange.C1+1, AllRange.R2+1, AllRange.C2+1);
		}

		internal TConditionalFormatRule[] GetCFDef(TCellList CellList)
		{
			TConditionalFormatRule[] Result = new TConditionalFormatRule[CFs.Count];
			
			for(int i = 0; i < CFs.Count; i++)
			{
				Result[i] = CFs[i].ToCFRule(CellList);
			}
			return Result;
		}

        internal void UpdateCFRows(TCellList Cells)
        {
            for (int i = RangeValuesList.Count - 1; i >= 0; i--)
            {
                TExcelRange xr = RangeValuesList[i];
                int MaxR2 = Math.Min(xr.R2, Cells.Count);
                for (int aRow = xr.R1; aRow <= MaxR2; aRow++)
                {
                    TCellCondFmt[] RowFmt = Cells.GetRowCondFmt(aRow);

                    if (RowFmt == null)
                    {
                        Cells.SetRowCondFmt(aRow, new TCellCondFmt[] { new TCellCondFmt(xr.C1, xr.C2, this) });
                    }
                    else
                    {
                        int OldLen = RowFmt.Length;
                        TCellCondFmt[] NewRow = new TCellCondFmt[OldLen + 1];
                        int k = 0;
                        for (int z = 0; z < OldLen; z++)
                        {
                            TCellCondFmt OldCf = RowFmt[z];

                            if (k == 0 && xr.C1 < OldCf.C1)
                            {
                                NewRow[z] = new TCellCondFmt(xr.C1, xr.C2, this);
                                k++;
                            }
                            NewRow[z + k] = OldCf;
                        }
                        if (k == 0) NewRow[OldLen] = new TCellCondFmt(xr.C1, xr.C2, this);
                        Cells.SetRowCondFmt(aRow, NewRow);
                    }
                }
            }
        }

		internal void ConditionallyModifyFormat(TFlxFormat Format, ExcelFile aXls, int aSheetIndex, int RowOffset, int ColOffset)
		{
			CFs.ConditionallyModifyFormat(Format, aXls, aSheetIndex, RowOffset, ColOffset);
		}

		internal bool EqualsDef(TConditionalFormatRule[] Fmt, TCellList CellList)
		{
			if (Fmt == null || Fmt.Length == 0) return false; //it is not equal to anything.
			if (CFs.Count != Fmt.Length) return false;
			for (int i = Fmt.Length - 1; i >= 0; i--)
			{
				if (!CFs[i].EqualsDef(Fmt[i], CellList)) return false;
			}
			return true;
		}

		/// <summary>
		/// Removes all references to this range on this cf.
		/// </summary>
		/// <param name="aRow1"></param>
		/// <param name="aCol1"></param>
		/// <param name="aRow2"></param>
		/// <param name="aCol2"></param>
		internal void ClearRange(int aRow1, int aCol1, int aRow2, int aCol2)
		{
			bool Modified = false;
			if (Modified) FixAllRange();
		}
	
		internal static void AddRange(int aRow1, int aCol1, int aRow2, int aCol2)
		{
			//if (Modified) FixAllRange();
		}
	}
}
