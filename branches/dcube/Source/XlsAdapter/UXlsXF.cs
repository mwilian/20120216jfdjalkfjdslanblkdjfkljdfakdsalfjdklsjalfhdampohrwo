using System;
using System.Globalization;
using FlexCel.Core;
using System.Text;
using System.Diagnostics;
using System.Collections.Generic;

#if (MONOTOUCH)
using Color = MonoTouch.UIKit.UIColor;
#endif

#if (WPF)
using real = System.Double;
using System.Windows.Media;
#else
using System.Drawing;
using real = System.Single;
#endif

using System.IO;


namespace FlexCel.XlsAdapter
{

    #region XF
    /// <summary>
    /// A map from a byte array to a struct.
    /// </summary>
    internal class TXFDat
    {
        byte[] Data = null;
        internal TXFDat(byte[] aData)
        {
            Data = aData;
        }

        internal int Font { get { return BitOps.GetWord(Data, 0); } set { BitOps.SetWord(Data, 0, value); } }
        internal int Format { get { return BitOps.GetWord(Data, 2); } set { BitOps.SetWord(Data, 2, value); } }
        internal int Options4 { get { return BitOps.GetWord(Data, 4); } set { BitOps.SetWord(Data, 4, value); } }
        internal int Options6 { get { return BitOps.GetWord(Data, 6); } set { BitOps.SetWord(Data, 6, value); } }
        internal int Options8 { get { return BitOps.GetWord(Data, 8); } set { BitOps.SetWord(Data, 8, value); } }
        internal int Options10 { get { return BitOps.GetWord(Data, 10); } set { BitOps.SetWord(Data, 10, value); } }
        internal int Options12 { get { return BitOps.GetWord(Data, 12); } set { BitOps.SetWord(Data, 12, value); } }
        internal long Options14 { get { return BitOps.GetCardinal(Data, 14); } set { BitOps.SetCardinal(Data, 14, value); } }
        internal int Options18 { get { return BitOps.GetWord(Data, 18); } set { BitOps.SetWord(Data, 18, value); } }

        internal static int Length { get { return 20; } }

        internal bool IsStyle
        {
            get
            {
                return (Options4 & 0x0004) != 0;
            }
            set
            {
                if (value)
                    Options4 |= 0x0004;
                else
                    Options4 &= ~0x0004;
            }
        }

        internal int Parent
        {
            get
            {
                return (Options4 >> 4) & 0xFFF;
            }
            set
            {
                Options4 = (Options4 & 0xF) | ((value << 4) & 0xFFF0);
            }
        }

        internal int CellPattern { get { return (int)((Options14 & 0xFC000000) >> 26); } }
        internal int CellFgColorIndex { get { return (Options18 & 0x7F); } }
        internal int CellBgColorIndex { get { return ((Options18 & 0x3F80) >> 7); } }

        internal TFlxBorderStyle GetBorderStyle(int aPos, byte FirstBit)
        {
            return (TFlxBorderStyle)((Data[aPos] >> FirstBit) & 0xF);
        }
        internal int GetBorderColorIndex(int aPos, byte FirstBit)
        {
            int Result = ((BitOps.GetWord(Data, aPos) >> FirstBit) & 0x7F);
            //if (Result<1) Result=1;
            return Result;
        }
        internal TFlxBorderStyle GetBorderStyleExt(int aPos, byte FirstBit)
        {
            return (TFlxBorderStyle)((BitOps.GetCardinal(Data, aPos) >> FirstBit) & 0xF);
        }
        internal int GetBorderColorIndexExt(int aPos, byte FirstBit)
        {
            int Result = (int)(((BitOps.GetCardinal(Data, aPos) >> FirstBit) & 0x7F));
            //if (Result<1) Result=1;
            return Result;
        }

        internal TFlxDiagonalBorder DiagonalStyle
        {
            get
            {
                return (TFlxDiagonalBorder)((BitOps.GetCardinal(Data, 12) >> 14) & 3);
            }
        }

        internal TVFlxAlignment VAlign
        {
            get
            {
                return (TVFlxAlignment)((Options6 & 0x70) >> 4);
            }
        }
        internal THFlxAlignment HAlign
        {
            get
            {
                return (THFlxAlignment)((Options6 & 0x7));
            }
        }

        internal bool WrapText
        {
            get
            {
                return (Options6 & 0x8) == 0x8;
            }
        }

        internal int Rotation
        {
            get
            {
                return Data[7];
            }
        }

    }

    #region ColorDictionary
    internal class TUsedColorDictionary
    {
        Dictionary<int, bool> Dict;

        public TUsedColorDictionary()
        {
            Dict = new Dictionary<int, bool>();
        }

        public bool this[int aColor]
        {
            get
            {
                return Dict[aColor];
            }
            set
            {
                Dict[aColor] = value;
            }
        }

        /// <summary>
        /// Will not overwrite values where IsIndexed is true.
        /// </summary>
        /// <param name="aColor"></param>
        /// <param name="IsIndexed"></param>
        internal void AddColor(int aColor, bool IsIndexed)
        {
            if (IsIndexed || !Dict.ContainsKey(aColor))
            {
                Dict[aColor] = IsIndexed;
            }
        }

        #region IEnumerable<KeyValuePair<int,bool>> Members

        public IEnumerator<int> GetEnumerator()
        {
            return Dict.Keys.GetEnumerator();
        }

        #endregion
    }
    #endregion


    /// <summary>
    /// XF format record. This class should be inmutable once it has been fully loaded.
    /// It is used inside a dictionary (and if we update the values, keys in the dictionary will become invalid).
    /// </summary>
    internal class TXFRecord : TBaseRecord
    {
        #region Variables
        private int Id;
        private int FFontIndex;
        private int FBorders;
        private int FFillPattern;
        private int NumberFormat;
        private THFlxAlignment FHAlignment;
        private TVFlxAlignment FVAlignment;
        private bool FLocked;
        private bool FHidden;
        private bool FWrapText;
        private bool FShrinkToFit;
        private bool F123Prefix;
        private byte FRotation;
        private bool FJustLast;
        private byte FIReadOrder;
        private bool FMergeCell;
        private byte FIndent;
        private bool FSxButton;

        private TLinkedStyle FLinkedStyle;
        private bool FIsStyle;
        private int FParent;

        private int CachedHashCode;
        private TColorIndexCache FgColor;
        private TColorIndexCache BgColor;
        private TColorIndexCache BLeft; private TColorIndexCache BTop;
        private TColorIndexCache BRight; private TColorIndexCache BBottom; private TColorIndexCache BDiag;

        internal TFutureStorage FutureStorage;

        #endregion

        #region Constructors
        internal TXFRecord(int aId, byte[] aData, TBorderList BorderList, TPatternList PatternList, TBiff8XFMap XFMap)
        {
            Init(aId);
            LoadFromBiff8(aData, BorderList, PatternList, XFMap);
        }

        public void Init(int aId)
        {
            Id = aId;
            FLinkedStyle = new TLinkedStyle();
            FLinkedStyle.AutomaticChoose = false; //it should always be false, so equals is consistent.
        }

        /// <summary>
        /// CreateFromFormat.
        /// </summary>
        internal TXFRecord(TFlxFormat Fmt, bool IsZero, TWorkbookGlobals Globals, bool AddParentStyleIfNeeded)
        {
            Init((int)xlr.XF);

            if (IsZero) //Format 0 must be in font 0.
            {
                FFontIndex = 0;
                if (Globals.Fonts.Count == 0) Globals.Fonts.Add(new TFontRecord(Fmt.Font)); 
                else Globals.Fonts[0] = new TFontRecord(Fmt.Font);
            }
            else
            {
                FFontIndex = Globals.Fonts.AddFont(Fmt.Font);
            }

            NumberFormat = Globals.Formats.AddFormat(Fmt.Format);

            if (Fmt.IsStyle)
            {
                FParent = 0xFFF;
            }
            else
            {
                FParent = Globals.Styles.GetStyle(Fmt.NotNullParentStyle);
                if (FParent < 0 || FParent >= Globals.StyleXF.Count)
                {
                    FParent = 0;
                    if (AddParentStyleIfNeeded) FParent = AddBuiltIntStyle(Fmt.NotNullParentStyle, Globals);
                }
            }

            FBorders = Globals.Borders.Add(Fmt.Borders);
            FFillPattern = Globals.Patterns.Add(Fmt.FillPattern);
            FIsStyle = Fmt.IsStyle;

            FHAlignment = Fmt.HAlignment;
            FVAlignment = Fmt.VAlignment;
            FLocked = Fmt.Locked;
            FHidden = Fmt.Hidden;
            FWrapText = Fmt.WrapText;
            FShrinkToFit = Fmt.ShrinkToFit;
            F123Prefix = Fmt.Lotus123Prefix;
            FRotation = Fmt.Rotation;
            FIndent = Fmt.Indent;

            FLinkedStyle.Assign(Fmt.LinkedStyle);
            //docs say this doesn't matter for style xfs, but is used by Excel to know what to apply next time you select the style.

            //Here Options8 is not ready yet, so we can't use it.
            if (!Fmt.IsStyle && Fmt.LinkedStyle.AutomaticChoose) //In this case, we need to select and set to 0 all properties that are the same than the parent style. Properties that are different should be 1.
            {
                TXFRecord ParentXF = Globals.StyleXF[FParent];
                FLinkedStyle.LinkedAlignment =
                    FHAlignment == ParentXF.FHAlignment
                    && FWrapText == ParentXF.WrapText
                    && FVAlignment == ParentXF.FVAlignment
                    && FJustLast == ParentXF.FJustLast
                    && FRotation == ParentXF.FRotation
                    && FIndent == ParentXF.FIndent
                    && FShrinkToFit == ParentXF.FShrinkToFit
                    && FIReadOrder == ParentXF.FIReadOrder;

                FLinkedStyle.LinkedBorder =
                    Globals.Borders[FBorders].Equals(Globals.Borders[ParentXF.FBorders]);

                FLinkedStyle.LinkedFill =
                    Globals.Patterns[FFillPattern].Equals(Globals.Patterns[ParentXF.FFillPattern]);

                FLinkedStyle.LinkedFont = FFontIndex == ParentXF.FFontIndex; //Font not equal to parent
                FLinkedStyle.LinkedNumericFormat = NumberFormat == ParentXF.NumberFormat; //fmt not equal to parent
                FLinkedStyle.LinkedProtection = FLocked == ParentXF.FLocked && FHidden == ParentXF.FHidden;

            }

            FLinkedStyle.AutomaticChoose = false; //it wasn't copied in the assign, but anyway, just to make sure.

        }

        private int AddBuiltIntStyle(string name, TWorkbookGlobals Globals)
        {
            int Level;
            int BuiltinId = TBuiltInStyles.GetIdAndLevel(name, out Level);
            if (BuiltinId < 0) return 0;

            int Result = Globals.AddStyleFormat(TBuiltInStyles.GetDefaultStyle(BuiltinId, Level), name);
            Globals.Styles.SetStyle(name, Result);
            return Result;
        }

        /// <summary>
        /// CreateFromFormat.
        /// </summary>
        internal TXFRecord(int aFontIndex, int aNumberFormat, int aFillPattern, int aBorders, bool aIsStyle, int aParent,
            THFlxAlignment aHAlignment, TVFlxAlignment aVAlignment, bool aLocked, bool aHidden, bool aWrapText, bool aShrinkToFit,
            bool aLotus123Prefix, int aRotation, int aIndent, bool aFsxButton, TLinkedStyle aLinkedStyle)
        {
            Init((int)xlr.XF);
            FFontIndex = aFontIndex;
            NumberFormat = aNumberFormat;
            FFillPattern = aFillPattern;
            FBorders = aBorders;

            FIsStyle = aIsStyle;
            if (aIsStyle)
            {
                FParent = 0xFFF;
            }
            else
                FParent = aParent;

            FHAlignment = aHAlignment;
            FVAlignment = aVAlignment;
            FLocked = aLocked;
            FHidden = aHidden;
            FWrapText = aWrapText;
            FShrinkToFit = aShrinkToFit;
            if (aIsStyle) F123Prefix = false; else F123Prefix = aLotus123Prefix;
            FRotation = (byte)aRotation;
            FIndent = (byte)aIndent;
            FSxButton = aFsxButton;

            FLinkedStyle = aLinkedStyle;
            //docs say this doesn't matter for style xfs, but is used by Excel to know what to apply next time you select the style.

            FLinkedStyle.AutomaticChoose = false;
        }


        internal void MoveBordersAndPatternsToOtherFile(TBorderList SourceBorders, TBorderList DestBorders, TPatternList SourcePatterns, TPatternList DestPatterns)
        {
            FBorders = DestBorders.Add(SourceBorders[FBorders]);
            FFillPattern = DestPatterns.Add(SourcePatterns[FFillPattern]);
        }

        internal void MergeWithParentStyle(TXFRecord ParentRecord)
        {
            if (FLinkedStyle.LinkedNumericFormat)
            {
                NumberFormat = ParentRecord.NumberFormat;
            }

            if (FLinkedStyle.LinkedFont)
            {
                FFontIndex = ParentRecord.FFontIndex;
            }

            if (FLinkedStyle.LinkedAlignment)
            {
                FHAlignment = ParentRecord.FHAlignment;
                FWrapText = ParentRecord.WrapText;
                FVAlignment = ParentRecord.FVAlignment;
                FJustLast = ParentRecord.FJustLast;
                FRotation = ParentRecord.FRotation;
                FIndent = ParentRecord.FIndent;
                FShrinkToFit = ParentRecord.FShrinkToFit;
                FIReadOrder = ParentRecord.FIReadOrder;
            }

            if (FLinkedStyle.LinkedBorder)
            {
                FBorders = ParentRecord.FBorders;
            }

            if (FLinkedStyle.LinkedFill)
            {
                FFillPattern = ParentRecord.FFillPattern;
            }

            if (FLinkedStyle.LinkedProtection)
            {
                FLocked = ParentRecord.FLocked;
                FHidden = ParentRecord.FHidden;
            }
        }

        protected override TBaseRecord DoCopyTo(TSheetInfo SheetInfo)
        {
            TXFRecord Result = (TXFRecord)MemberwiseClone();
            Result.Init(Id);
            if (FLinkedStyle != null) Result.FLinkedStyle = (TLinkedStyle)FLinkedStyle.Clone();
            return Result;
        }

        #endregion

        #region public props
        public int FontIndex { get { return FFontIndex; } set { FFontIndex = value; } }
        public int FormatIndex { get { return NumberFormat; } set { NumberFormat = value; } }
        public int Borders { get { return FBorders; } set { FBorders = value; } }
        public int FillPattern { get { return FFillPattern; } set { FFillPattern = value; } }
        public bool WrapText { get { return FWrapText; } set { FWrapText = value; } }
        public byte Rotation { get { return FRotation; } set { FRotation = value; } }
        public bool IsStyle { get { return FIsStyle; } }
        public int Parent { get { return FParent; } set { FParent = value; } }
        public THFlxAlignment HAlignment { get { return FHAlignment; } set { FHAlignment = value; } }
        public TVFlxAlignment VAlignment { get { return FVAlignment; } set { FVAlignment = value; } }
        public bool ShrinkToFit { get { return FShrinkToFit; } set { FShrinkToFit = value; } }
        public byte Indent { get { return FIndent; } set { FIndent = value; } }
        public bool JustLast { get { return FJustLast; } set { FJustLast = value; } }
        public byte IReadOrder { get { return FIReadOrder; } set { FIReadOrder = value; } }
        public bool Locked { get { return FLocked; } set { FLocked = value; } }
        public bool Hidden { get { return FHidden; } set { FHidden = value; } }
        public bool Lotus123Prefix { get { return F123Prefix; } set { F123Prefix = value; } }
        public bool SxButton { get { return FSxButton; } set { FSxButton = value; } }
        public TLinkedStyle LinkedStyle { get { return FLinkedStyle; } }

        #endregion

        #region Utility
        internal int GetActualFontIndex(TFontRecordList FontList)
        {
            int FontIdx = FFontIndex;
            if (FontIdx >= 4) FontIdx--; //Font number 4 does not exist
            if ((FontIdx < 0) || (FontIdx >= FontList.Count)) FontIdx = 0;
            return FontIdx;
        }

        private static void DoOneColor(bool[] UsedColors, TExcelColor aColor, TUsedColorDictionary ColorDictionary, IFlexCelPalette xls)
        {
            if (ColorDictionary != null) ColorDictionary.AddColor(aColor.ToColor(xls).ToArgb(), aColor.ColorType == TColorType.Indexed);

            if (aColor.ColorType != TColorType.Indexed) return;
            int i = aColor.Index;
            if (i >= 0 && i < UsedColors.Length) UsedColors[i] = true;
        }

        private static void DoOneBorder(bool[] UsedColors, TFlxOneBorder bs, TUsedColorDictionary ColorDictionary, IFlexCelPalette xls)
        {
            if (bs.Style != TFlxBorderStyle.None)
            {
                DoOneColor(UsedColors, bs.Color, ColorDictionary, xls);
            }
        }

        internal void FillUsedColors(bool[] UsedColors, TFontRecordList FontList, TBorderList BorderList, TPatternList PatternList, TUsedColorDictionary ColorDictionary, IFlexCelPalette xls)
        {
            TFlxBorders Borders = BorderList[FBorders];
            DoOneBorder(UsedColors, Borders.Left, ColorDictionary, xls);
            DoOneBorder(UsedColors, Borders.Right, ColorDictionary, xls);
            DoOneBorder(UsedColors, Borders.Top, ColorDictionary, xls);
            DoOneBorder(UsedColors, Borders.Bottom, ColorDictionary, xls);
            DoOneBorder(UsedColors, Borders.Diagonal, ColorDictionary, xls);

            TFlxFillPattern Pattern = PatternList[FFillPattern];

            if (Pattern.Pattern == TFlxPatternStyle.Gradient)
            {
                if (Pattern.Gradient.Stops.Length > 0) DoOneColor(UsedColors, Pattern.Gradient.Stops[0].Color, ColorDictionary, xls);
            }
            else
            {
                if (Pattern.Pattern != TFlxPatternStyle.None)
                {
                    DoOneColor(UsedColors, Pattern.FgColor, ColorDictionary, xls);

                    if (Pattern.Pattern != TFlxPatternStyle.Solid)  //bgColor does not matter here.
                    {
                        DoOneColor(UsedColors, Pattern.BgColor, ColorDictionary, xls);
                    }
                }
            }

            DoOneColor(UsedColors, FontList[GetActualFontIndex(FontList)].Color, ColorDictionary, xls);
        }
        #endregion

        #region Biff8
        public byte[] SaveToBiff8(TSaveData SaveData, int[] StylesSavedAtPos)
        {
            byte[] Data = new byte[20];
            TXFDat XFDat = new TXFDat(Data);
            XFDat.Font = FFontIndex;
            XFDat.Format = NumberFormat;

            int Biff8Parent = 0xFFF;
            if (!IsStyle && StylesSavedAtPos != null)
            {
                Biff8Parent = StylesSavedAtPos[FParent];
            }

            XFDat.Options4 = BitOps.BoolToBit(FLocked, 0) +
                BitOps.BoolToBit(FHidden, 1) +
                (0 << 2) + //Cell style
                BitOps.BoolToBit(F123Prefix, 3) + //123 lotus
                ((Biff8Parent << 4) & 0xFFF0);

            if (FIsStyle)
            {
                XFDat.Options4 = XFDat.Options4 | (0x0004); // style
            }

            XFDat.Options6 = (int)(FHAlignment) +
                BitOps.BoolToBit(FWrapText, 3) +
                (((int)FVAlignment) << 4) +
                BitOps.BoolToBit(FJustLast, 7) +
                ((FRotation << 8) & 0xFF00);


            TFlxBorders Borders = SaveData.BorderList[FBorders];

            XFDat.Options10 = (int)Borders.Left.Style +
                ((int)(Borders.Right.Style) << 4) +
                ((int)(Borders.Top.Style) << 8) +
                ((int)(Borders.Bottom.Style) << 12);

            int bl = (Borders.Left.Color.GetBiff8ColorIndex(SaveData.Palette, TAutomaticColor.DefaultForeground, ref BLeft)) & 0x7F;
            int br = (Borders.Right.Color.GetBiff8ColorIndex(SaveData.Palette, TAutomaticColor.DefaultForeground, ref BRight)) & 0x7F;

            int bt = (Borders.Top.Color.GetBiff8ColorIndex(SaveData.Palette, TAutomaticColor.DefaultForeground, ref BTop)) & 0x7F;
            int bb = (Borders.Bottom.Color.GetBiff8ColorIndex(SaveData.Palette, TAutomaticColor.DefaultForeground, ref BBottom)) & 0x7F;

            int bd = (Borders.Diagonal.Color.GetBiff8ColorIndex(SaveData.Palette, TAutomaticColor.DefaultForeground, ref BDiag)) & 0x7F;
            //check if a border is set then the color too. If not Excel disables the "format cell" dialog.
            
            if (Borders.Left.Style != TFlxBorderStyle.None && bl == 0) bl = 0x40;
            if (Borders.Top.Style != TFlxBorderStyle.None && bt == 0) bt = 0x40;
            if (Borders.Bottom.Style != TFlxBorderStyle.None && bb == 0) bb = 0x40;
            if (Borders.Right.Style != TFlxBorderStyle.None && br == 0) br = 0x40;
            if (Borders.Diagonal.Style != TFlxBorderStyle.None && bd == 0) bd = 0x40;

            XFDat.Options12 = bl +
                ((br) << 7) +
                ((int)(Borders.DiagonalStyle) << 14);

            TFlxFillPattern Pattern = SaveData.Globals.Patterns[FFillPattern];
            TFlxPatternStyle fp = Pattern.Pattern == TFlxPatternStyle.Gradient ? TFlxPatternStyle.Solid : Pattern.Pattern;
            XFDat.Options14 = bt +
                (bb << 7) +
                (bd << 14) +
                ((int)(Borders.Diagonal.Style) << 21) +
                BitOps.BoolToBit(NeedsXFExt(SaveData.Globals), 25) +
                (((int)fp - 1) << 26);

            TExcelColor FinalFgColor = Pattern.FgColor;
            TExcelColor FinalBgColor = Pattern.BgColor;
            if (Pattern.Pattern == TFlxPatternStyle.Gradient)
            {
                FinalBgColor = TExcelColor.Automatic;
                if (Pattern.Gradient.Stops.Length > 0) FinalFgColor = Pattern.Gradient.Stops[0].Color;
            }

            TAutomaticColor BgAutColor = Pattern.Pattern == TFlxPatternStyle.Solid || Pattern.Pattern == TFlxPatternStyle.Gradient ?
                TAutomaticColor.DefaultForeground : TAutomaticColor.DefaultBackground;


            XFDat.Options18 = ((FinalFgColor.GetBiff8ColorIndex(SaveData.Palette, TAutomaticColor.DefaultForeground, ref FgColor)) & 0x7F) +
                (((FinalBgColor.GetBiff8ColorIndex(SaveData.Palette, BgAutColor, ref BgColor)) & 0x7F) << 7) +
                BitOps.BoolToBit(FSxButton, 14); //Attached to pivot table


            TLinkedStyle cx = FLinkedStyle;
            //docs say this doesn't matter for style xfs, but is used by Excel to know what to apply next time you select the style.

            //Here Options8 is not ready yet, so we can't use it.
            int f1, f2, f3, f4, f5, f6;
            f1 = cx.LinkedNumericFormat ? 0 : 0x0400; //fmt not equal to parent
            f2 = cx.LinkedFont ? 0 : 0x0800; //Font not equal to parent
            f3 = cx.LinkedAlignment ? 0 : 0x1000;
            f4 = cx.LinkedBorder ? 0 : 0x2000;
            f5 = cx.LinkedFill ? 0 : 0x4000;
            f6 = cx.LinkedProtection ? 0 : 0x8000;

            int mIndent = FIndent; if (mIndent > 15) mIndent = 15; //biff8 can only save 15 levels in the main xf record.
            XFDat.Options8 = (mIndent & 0xF) +
                BitOps.BoolToBit(FShrinkToFit, 4) +
                BitOps.BoolToBit(FMergeCell, 5) + //mergecell
                ((FIReadOrder << 6) & 0xC0) + //readingOrder
                f1 + f2 + f3 + f4 + f5 + f6;

            return Data;
        }

        private void LoadFromBiff8(byte[] aData, TBorderList BorderList, TPatternList PatternList, TBiff8XFMap XFMap)
        {
            TXFDat XFDat = new TXFDat(aData);
            FFontIndex = XFDat.Font;

            TFlxBorders Borders = new TFlxBorders();
            Borders.Left.Style = XFDat.GetBorderStyle(10, 0);
            Borders.Right.Style = XFDat.GetBorderStyle(10, 4);
            Borders.Top.Style = XFDat.GetBorderStyle(11, 0);
            Borders.Bottom.Style = XFDat.GetBorderStyle(11, 4);

            Borders.Left.Color = TExcelColor.FromBiff8ColorIndex(XFDat.GetBorderColorIndex(12, 0));
            Borders.Right.Color = TExcelColor.FromBiff8ColorIndex(XFDat.GetBorderColorIndex(12, 7));
            Borders.Top.Color = TExcelColor.FromBiff8ColorIndex(XFDat.GetBorderColorIndex(14, 0));
            Borders.Bottom.Color = TExcelColor.FromBiff8ColorIndex(XFDat.GetBorderColorIndex(14, 7));

            Borders.Diagonal.Style = XFDat.GetBorderStyleExt(14, 21);
            Borders.Diagonal.Color = TExcelColor.FromBiff8ColorIndex(XFDat.GetBorderColorIndexExt(14, 14));

            Borders.DiagonalStyle = XFDat.DiagonalStyle;

            FBorders = BorderList.Add(Borders);

            NumberFormat = XFDat.Format;

            TFlxFillPattern Pattern = new TFlxFillPattern();
            Pattern.Pattern = (TFlxPatternStyle)(XFDat.CellPattern + 1);
            Pattern.FgColor = TExcelColor.FromBiff8ColorIndex(XFDat.CellFgColorIndex);
            Pattern.BgColor = TExcelColor.FromBiff8ColorIndex(XFDat.CellBgColorIndex);
            FFillPattern = PatternList.Add(Pattern);

            FSxButton = (XFDat.Options18 & 0x4000) != 0;

            FHAlignment = XFDat.HAlign;
            FVAlignment = XFDat.VAlign;

            FLocked = (XFDat.Options4 & 0x1) == 0x1;
            FHidden = (XFDat.Options4 & 0x2) == 0x2;
            F123Prefix = (XFDat.Options4 & 0x8) == 0x8;


            FWrapText = (XFDat.Options6 & 0x08) == 0x08;
            FJustLast = (XFDat.Options6 & 0x80) == 0x80;

            FShrinkToFit = (XFDat.Options8 & 0x10) == 0x10;
            FMergeCell = (XFDat.Options8 & 0x20) == 0x20;
            FIReadOrder = (byte)((XFDat.Options8 & 0xC0) >> 6);

            FRotation = aData[7];
            FIndent = (byte)(XFDat.Options8 & 0xF);

            FIsStyle = XFDat.IsStyle;
            FParent = XFMap == null || FIsStyle ? XFDat.Parent : XFMap.GetStyleXF2007(XFDat.Parent);

            int Link = XFDat.Options8;
            FLinkedStyle = new TLinkedStyle();
            FLinkedStyle.AutomaticChoose = false;
            FLinkedStyle.LinkedNumericFormat = (Link & 0x0400) == 0;
            FLinkedStyle.LinkedFont = (Link & 0x0800) == 0;
            FLinkedStyle.LinkedAlignment = (Link & 0x1000) == 0;
            FLinkedStyle.LinkedBorder = (Link & 0x2000) == 0;
            FLinkedStyle.LinkedFill = (Link & 0x4000) == 0;
            FLinkedStyle.LinkedProtection = (Link & 0x8000) == 0;
        }

        internal static TExcelColor ReplaceIndexedByRGB(IFlexCelPalette xls, TExcelColor aColor)
        {
            if (aColor.ColorType != TColorType.Indexed) return aColor;
            return aColor.ToColor(xls);
        }

        #endregion

        #region Save and Convert
        internal TFlxFormat FlxFormat(TStyleRecordList StyleList, TFontRecordList FontList, TFormatRecordList FormatList, TBorderList BorderList, TPatternList PatternList)
        {
            TFlxFormat Result = new TFlxFormat();
            Result.Borders = (TFlxBorders)BorderList[FBorders].Clone();
            Result.FillPattern = (TFlxFillPattern)PatternList[FFillPattern].Clone();
            Result.Font = FontList[GetActualFontIndex(FontList)].FlxFont();
            Result.Format = FormatList.Format(NumberFormat);

            Result.HAlignment = FHAlignment;
            Result.Hidden = FHidden;
            Result.Indent = FIndent;
            Result.IsStyle = FIsStyle;
            Result.LinkedStyle.Assign(FLinkedStyle);
            Result.Lotus123Prefix = F123Prefix;

            Result.Locked = FLocked;

            Result.ParentStyle = null;
            if (!FIsStyle)
            {
                Result.ParentStyle = StyleList.GetStyleNameFromXF(FParent);
            }

            Result.Rotation = FRotation;
            Result.ShrinkToFit = FShrinkToFit;
            Result.VAlignment = FVAlignment;
            Result.WrapText = FWrapText;

            return Result;
        }

        internal override void SaveToPxl(TPxlStream PxlStream, int Row, TPxlSaveData SaveData)
        {
            base.SaveToPxl(PxlStream, Row, SaveData);
            PxlStream.WriteByte((byte)pxl.XF);
            int TmpFontIndex = FFontIndex;
            if (TmpFontIndex > 4) TmpFontIndex--; //Font 4 does not exist on Biff8.
            PxlStream.Write16((UInt16)TmpFontIndex); //Font Index

            //Formats on BIFF8 have a format id, and you search for that id on the format list.
            //Formats on PXL are sequential, < 233 means internal and >= 233 means the position on the list.
            int FormatIndex = NumberFormat;
            string FmtStr = SaveData.Globals.Formats.Format(FormatIndex);  //gets the string format including internals.
            PxlStream.Write16((UInt16)SaveData.Globals.Formats.GetPxlIndex(FmtStr));//Format Index

            byte[] ff = { 0xFF, 0xFF, 0xFF, 0xFF };
            PxlStream.Write(ff, 0, ff.Length);  //reserved

            UInt16 BaseAttrs = 0;
            if (FLocked) BaseAttrs |= 0x2; //locked
            if (FHidden) BaseAttrs |= 0x4; //hidden
            if (F123Prefix) BaseAttrs |= 0x8; // 1-2-3 prefix
            PxlStream.Write16(BaseAttrs);

            byte HAlignAndWordWrap = (byte)((int)FHAlignment + BitOps.BoolToBit(FWrapText, 3));
            byte VAlign = (byte)(((int)FVAlignment & 0x03) + 1);  //VAlign is +1

            TFlxBorders Borders = SaveData.Globals.Borders[FBorders];
            byte bBorders = 0;
            if (Borders.Left.Style != TFlxBorderStyle.None) bBorders |= 2; //left border
            if (Borders.Right.Style != TFlxBorderStyle.None) bBorders |= 8; //right border
            if (Borders.Top.Style != TFlxBorderStyle.None) bBorders |= 1; //top border
            if (Borders.Bottom.Style != TFlxBorderStyle.None) bBorders |= 4; //bottom border

            PxlStream.Write16((UInt16)(HAlignAndWordWrap | (VAlign << 4) | (bBorders << 8))); //Text Attributes.

            PxlStream.Write16(0xFF);  //reserved. doc is wrong here, this goes first.

            TFlxFillPattern Pattern = SaveData.Globals.Patterns[FFillPattern];
            if (Pattern.Pattern != TFlxPatternStyle.None && Pattern.Pattern != TFlxPatternStyle.Automatic)
            {
                PxlStream.Write16(Pattern.FgColor.GetPxlColorIndex(SaveData.Palette, TAutomaticColor.DefaultForeground));  //bg color
            }
            else
            {
                PxlStream.Write16(0xFF); //no fill
            }

            PxlStream.WriteByte(Borders.Top.Color.GetPxlColorIndex(SaveData.Palette, TAutomaticColor.DefaultForeground));
            PxlStream.WriteByte(Borders.Left.Color.GetPxlColorIndex(SaveData.Palette, TAutomaticColor.DefaultForeground));
            PxlStream.WriteByte(Borders.Bottom.Color.GetPxlColorIndex(SaveData.Palette, TAutomaticColor.DefaultForeground));
            PxlStream.WriteByte(Borders.Right.Color.GetPxlColorIndex(SaveData.Palette, TAutomaticColor.DefaultForeground));


            PxlStream.Write16(0);  //reserved
        }

        internal override int GetId
        {
            get { return Id; }
        }

        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
        {
            FlxMessages.ThrowException(FlxErr.ErrInternal);
        }

        internal void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row, int[] StylesSavedAtPos, ref uint XFCRC)
        {
            Workbook.WriteHeader((UInt16)Id, (UInt16)TotalSizeNoHeaders());
            byte[] Data = SaveToBiff8(SaveData, StylesSavedAtPos);
            Workbook.Write(Data, Data.Length);
            TXFCRCRecord.UpdateCRC(Data, ref XFCRC);
        }

        internal static void SaveDefaultOutline(IDataStream Workbook, ref uint XFCRC)
        {
            byte[] Data = {0x00, 0x00, 0x00, 0x00, 0xF5, 0xFF, 0x20, 0x00, 0x00, 0xF4, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x20} ;
            Workbook.WriteHeader((UInt16)xlr.XF, (UInt16)Data.Length);
            Workbook.Write(Data, Data.Length);
            TXFCRCRecord.UpdateCRC(Data, ref XFCRC);
        }

        internal override int TotalSizeNoHeaders()
        {
            return 20;
        }

        internal override int TotalSize()
        {
            return TotalSizeNoHeaders() + XlsConsts.SizeOfTRecordHeader;
        }

        #endregion

        #region Compare and Hash

        public override bool Equals(object obj)
        {
            TXFRecord o = obj as TXFRecord;
            if (o == null) return false;

            if (FFontIndex != o.FFontIndex) return false;
            if (NumberFormat != o.NumberFormat) return false;
            if (FHAlignment != o.FHAlignment) return false;
            if (FVAlignment != o.FVAlignment) return false;
            if (FLocked != o.FLocked) return false;
            if (FHidden != o.FHidden) return false;
            if (FWrapText != o.FWrapText) return false;
            if (FShrinkToFit != o.FShrinkToFit) return false;
            if (F123Prefix != o.F123Prefix) return false;
            if (FRotation != o.FRotation) return false;
            if (FJustLast != o.FJustLast) return false;
            if (FIReadOrder != o.FIReadOrder) return false;
            if (FMergeCell != o.FMergeCell) return false;
            if (FIndent != o.FIndent) return false;
            if (FSxButton != o.FSxButton) return false;

            if (FIsStyle != o.FIsStyle) return false;
            if (FParent != o.FParent) return false;

            if (FFillPattern != o.FFillPattern) return false;
            if (!FLinkedStyle.Equals(o.FLinkedStyle)) return false;

            if (FBorders != o.FBorders) return false;
            if (!TFutureStorage.Equals(FutureStorage, o.FutureStorage)) return false;
            return true;
        }

        public override int GetHashCode()
        {
            if (CachedHashCode != 0) return CachedHashCode; //It could be that this is 0 even if calcluated, but even in that case, it will just slow down a little, the thing will still work. So we don't need to add a guard variable here.
            //Also note that this class is immutable once the hashcode has been generated, so we can cache it.
            CachedHashCode = HashCoder.GetHash(
            FFontIndex.GetHashCode(),
            FBorders.GetHashCode(),
            NumberFormat.GetHashCode(),
            FHAlignment.GetHashCode(),
            FVAlignment.GetHashCode(),
            FLocked.GetHashCode(),
            FHidden.GetHashCode(),
            FWrapText.GetHashCode(),
            FShrinkToFit.GetHashCode(),
            F123Prefix.GetHashCode(),
            FRotation.GetHashCode(),
            FJustLast.GetHashCode(),
            FIReadOrder.GetHashCode(),
            FMergeCell.GetHashCode(),
            FIndent.GetHashCode(),
            FSxButton.GetHashCode(),

            FFillPattern.GetHashCode(),
            FLinkedStyle.GetHashCode(),
            FIsStyle.GetHashCode(),
            FParent.GetHashCode());

            return CachedHashCode;
        }
        #endregion

        #region XFExt

        private static TExcelColor AddTrueColor(TExcelColor OldColor, TXFExtProp p, TWorkbookGlobals Globals, ref TColorIndexCache IndexCache)
        {
            TExcelColor Result = TExcelColor.FromBiff8(p.Data);
            CacheTrueColor(OldColor, Globals, ref IndexCache, Result);
            return Result;
        }

        private static void CacheTrueColor(TExcelColor OldColor, TWorkbookGlobals Globals, ref TColorIndexCache IndexCache, TExcelColor NewColor)
        {
            if (OldColor.ColorType == TColorType.Indexed)
            {
                IndexCache.LastColorStored = NewColor;
                IndexCache.LastColorInPalette = OldColor.ToColor(Globals.Workbook).ToArgb();
                IndexCache.Index = OldColor.InternalIndex;
                if (NewColor.ColorType == TColorType.Theme)
                {
                    IndexCache.LastColorInTheme = Globals.Workbook.GetColorTheme(NewColor.Theme);
                }
            }
        }

        internal void AddExt(TXFExtRecord xfe, TWorkbookGlobals Globals)
        {
            int ofs = 20;
            TFlxBorders NewBorder = (TFlxBorders)Globals.Borders[FBorders].Clone();
            TFlxFillPattern NewPattern = Globals.Patterns[FFillPattern];

            for (int i = 0; i < xfe.Count; i++)
            {
                TXFExtProp p = xfe.GetProp(ref ofs);

                switch (p.PropType)
                {
                    case TXFExtType.CellFgColor:
                        NewPattern.FgColor = AddTrueColor(NewPattern.FgColor, p, Globals, ref FgColor);
                        break;

                    case TXFExtType.CellBgColor:
                        NewPattern.BgColor = AddTrueColor(NewPattern.BgColor, p, Globals, ref BgColor);
                        break;

                    case TXFExtType.CellGradient:
                        NewPattern.Gradient = TExcelGradientFromBiff8(p.Data);
                        if (NewPattern.Gradient.Stops.Length > 0)
                        {
                            CacheTrueColor(NewPattern.FgColor, Globals, ref FgColor, NewPattern.Gradient.Stops[0].Color); //cache the indexed colors so they aren't recalculated
                        }
                        break;

                    case TXFExtType.TopBorderColor:
                        NewBorder.Top.Color = AddTrueColor(NewBorder.Top.Color, p, Globals, ref BTop);
                        break;

                    case TXFExtType.BottomBorderColor:
                        NewBorder.Bottom.Color = AddTrueColor(NewBorder.Bottom.Color, p, Globals, ref BBottom);
                        break;

                    case TXFExtType.LeftBorderColor:
                        NewBorder.Left.Color = AddTrueColor(NewBorder.Left.Color, p, Globals, ref BLeft);
                        break;

                    case TXFExtType.RightBorderColor:
                        NewBorder.Right.Color = AddTrueColor(NewBorder.Right.Color, p, Globals, ref BRight);
                        break;

                    case TXFExtType.DiagBorderColor:
                        NewBorder.Diagonal.Color = AddTrueColor(NewBorder.Diagonal.Color, p, Globals, ref BDiag);
                        break;

                    case TXFExtType.TextColor:
                        TFontRecord XFFont = Globals.Fonts[GetActualFontIndex(Globals.Fonts)];
                        XFFont.Color = AddTrueColor(XFFont.Color, p, Globals, ref XFFont.FontColor); //Different fonts are saved differently in biff8. So we don't need to look if 2 diff XF used the same font.
                        break;

                    case TXFExtType.FontScheme:
                        TFontScheme FontScheme = TFontScheme.None;
                        switch (p.Data[0])
                        {
                            case 1: FontScheme = TFontScheme.Major; break;
                            case 2: FontScheme = TFontScheme.Minor; break;
                        }

                        Globals.Fonts[GetActualFontIndex(Globals.Fonts)].Scheme = FontScheme;
                        break;

                    case TXFExtType.Indent:
                        unchecked
                        {
                            FIndent = (byte)BitOps.GetWord(p.Data, 0);
                        }
                        break;
                }
            }

            FBorders = Globals.Borders.AddOrReplaceBorder(FBorders, NewBorder);
            FFillPattern = Globals.Patterns.AddOrReplacePattern(FFillPattern, NewPattern);
        }

        private TExcelGradient TExcelGradientFromBiff8(byte[] p)
        {
            TExcelGradient Result = null;
            int GrType = BitOps.GetWord(p, 0);

            if (GrType == 0)
            {
                Result = new TExcelLinearGradient();
                ((TExcelLinearGradient)Result).RotationAngle = BitConverter.ToDouble(p, 4);
            }
            else if (GrType == 1)
            {
                TExcelRectangularGradient rg = new TExcelRectangularGradient();
                rg.Left = BitConverter.ToDouble(p, 12);
                rg.Right = BitConverter.ToDouble(p, 20);
                rg.Top = BitConverter.ToDouble(p, 28);
                rg.Bottom = BitConverter.ToDouble(p, 36);
                Result = rg;
            }
            else return null; //invalid.

            Result.Stops = new TGradientStop[BitOps.GetWord(p, 44)];
            int ofs = 48;
            for (int i = 0; i < Result.Stops.Length; i++)
            {
                Result.Stops[i].Position = BitConverter.ToDouble(p, ofs + 6);
                Result.Stops[i].Color = TExcelColor.FromBiff8(p, ofs, ofs + 14, ofs + 2, false);
                ofs += 22;
            }


            return Result;
        }

        private bool SpecialColor(TExcelColor aColor, IFlexCelPalette palette)
        {
            if (aColor.ColorType == TColorType.Automatic || aColor.ColorType == TColorType.Indexed) return false;
            
            /*No, RGB color should be saved as such, not as indexed
            if (aColor.ColorType == TColorType.RGB)
            {
                if (palette.PaletteContainsColor(aColor)) return false;
            }*/

            return true;
        }


        internal bool NeedsXFExt(TWorkbookGlobals Globals)
        {
            if (Globals.Workbook.XlsBiffVersion == TXlsBiffVersion.Excel2003) return false;

            if (SpecialColor(Globals.Patterns[FFillPattern].FgColor, Globals.Workbook)) return true;
            if (SpecialColor(Globals.Patterns[FFillPattern].BgColor, Globals.Workbook)) return true;
            if (Globals.Patterns[FFillPattern].Pattern == TFlxPatternStyle.Gradient) return true;

            if (SpecialColor(Globals.Borders[FBorders].Top.Color, Globals.Workbook)) return true;
            if (SpecialColor(Globals.Borders[FBorders].Bottom.Color, Globals.Workbook)) return true;
            if (SpecialColor(Globals.Borders[FBorders].Left.Color, Globals.Workbook)) return true;
            if (SpecialColor(Globals.Borders[FBorders].Right.Color, Globals.Workbook)) return true;
            if (SpecialColor(Globals.Borders[FBorders].Diagonal.Color, Globals.Workbook)) return true;

            if (SpecialColor(Globals.Fonts[GetActualFontIndex(Globals.Fonts)].Color, Globals.Workbook)) return true;

            if (Globals.Fonts[GetActualFontIndex(Globals.Fonts)].Scheme != TFontScheme.None) return true;
            if (FIndent > 15) return true;

            return false;
        }

        private void CheckSpecialColor(TExcelColor c, IFlexCelPalette Palette, ref int size, ref int PropCount)
        {
            if (SpecialColor(c, Palette)) { PropCount++; size += 20; };

        }

        internal int XFExtLen(TWorkbookGlobals Globals)
        {
            int PropCount;
            return XFExtLen(Globals, out PropCount);
        }

        private int XFExtLen(TWorkbookGlobals Globals, out int PropCount)
        {
            int Result = 0;
            PropCount = 0;
            CheckSpecialColor(Globals.Patterns[FFillPattern].FgColor, Globals.Workbook, ref Result, ref PropCount);
            CheckSpecialColor(Globals.Patterns[FFillPattern].BgColor, Globals.Workbook, ref Result, ref PropCount);
            if (Globals.Patterns[FFillPattern].Gradient != null) { PropCount++; Result += CalcGradientLength(Globals.Patterns[FFillPattern].Gradient); }

            CheckSpecialColor(Globals.Borders[FBorders].Top.Color, Globals.Workbook, ref Result, ref PropCount);
            CheckSpecialColor(Globals.Borders[FBorders].Bottom.Color, Globals.Workbook, ref Result, ref PropCount);
            CheckSpecialColor(Globals.Borders[FBorders].Left.Color, Globals.Workbook, ref Result, ref PropCount);
            CheckSpecialColor(Globals.Borders[FBorders].Right.Color, Globals.Workbook, ref Result, ref PropCount);
            CheckSpecialColor(Globals.Borders[FBorders].Diagonal.Color, Globals.Workbook, ref Result, ref PropCount);

            CheckSpecialColor(Globals.Fonts[GetActualFontIndex(Globals.Fonts)].Color, Globals.Workbook, ref Result, ref PropCount);

            if (Globals.Fonts[GetActualFontIndex(Globals.Fonts)].Scheme != TFontScheme.None) { PropCount++; Result += 5; }
            if (FIndent > 15) { PropCount++; Result += 6; };

            if (PropCount == 0) return 0;
            return Result + 20;
        }


        private int CalcGradientLength(TExcelGradient aGradient)
        {
            int StopCount = aGradient.Stops == null ? 0 : aGradient.Stops.Length;
            return 4  //length of the prop
                 + 44  //gradient props
                 + 4  //Stop count
                 + 22 * StopCount; //stop size
        }

        internal void SaveXFExt(int Position, IDataStream DataStream, TSaveData SaveData)
        {
            int PropCount;
            int len = XFExtLen(SaveData.Globals, out PropCount);
            if (PropCount <= 0) return;
            DataStream.WriteHeader((UInt16)xlr.XFEXT, (UInt16)len);
            DataStream.Write16((UInt16)xlr.XFEXT);
            DataStream.Write(new byte[12], 12);

            DataStream.Write16((UInt16)Position);
            DataStream.Write16(0);
            DataStream.Write16((UInt16)PropCount);

            TWorkbookGlobals Globals = SaveData.Globals;

            WriteSpecialColor(DataStream, TXFExtType.CellFgColor, Globals.Patterns[FFillPattern].FgColor, Globals.Workbook);
            WriteSpecialColor(DataStream, TXFExtType.CellBgColor, Globals.Patterns[FFillPattern].BgColor, Globals.Workbook);
            if (Globals.Patterns[FFillPattern].Gradient != null) WriteGradient(DataStream, Globals.Patterns);

            WriteSpecialColor(DataStream, TXFExtType.TopBorderColor, Globals.Borders[FBorders].Top.Color, Globals.Workbook);
            WriteSpecialColor(DataStream, TXFExtType.BottomBorderColor, Globals.Borders[FBorders].Bottom.Color, Globals.Workbook);
            WriteSpecialColor(DataStream, TXFExtType.LeftBorderColor, Globals.Borders[FBorders].Left.Color, Globals.Workbook);
            WriteSpecialColor(DataStream, TXFExtType.RightBorderColor, Globals.Borders[FBorders].Right.Color, Globals.Workbook);
            WriteSpecialColor(DataStream, TXFExtType.DiagBorderColor, Globals.Borders[FBorders].Diagonal.Color, Globals.Workbook);

            WriteSpecialColor(DataStream, TXFExtType.TextColor, Globals.Fonts[GetActualFontIndex(Globals.Fonts)].Color, Globals.Workbook);

            WriteSpecialFont(DataStream, Globals.Fonts[GetActualFontIndex(Globals.Fonts)].Scheme);
            WriteSpecialIdent(DataStream);
        }

        private void WriteGradient(IDataStream DataStream, TPatternList Patterns)
        {
            TExcelGradient gr = Patterns[FFillPattern].Gradient;

            DataStream.Write16((UInt16)TXFExtType.CellGradient);
            DataStream.Write16((UInt16)CalcGradientLength(gr));

            TExcelLinearGradient lgr = gr as TExcelLinearGradient;

            if (lgr != null)
            {
                DataStream.Write32(0); //Gradient type
                DataStream.Write(BitConverter.GetBytes(lgr.RotationAngle), 8);
                DataStream.Write(new byte[32], 32);
            }
            else
            {
                TExcelRectangularGradient rgr = gr as TExcelRectangularGradient;
                if (rgr == null) XlsMessages.ThrowException(XlsErr.ErrInternal);

                DataStream.Write32(1);
                DataStream.Write(new byte[8], 8); //rotation angle                
                DataStream.Write(BitConverter.GetBytes(rgr.Left), 8);
                DataStream.Write(BitConverter.GetBytes(rgr.Right), 8);
                DataStream.Write(BitConverter.GetBytes(rgr.Top), 8);
                DataStream.Write(BitConverter.GetBytes(rgr.Bottom), 8);
            }

            if (gr.Stops == null)
            {
                DataStream.Write32(0);
                return;
            }

            DataStream.Write32((UInt32)gr.Stops.Length);

            for (int i = 0; i < gr.Stops.Length; i++)
            {
                TExcelColor aColor = gr.Stops[i].Color;
                switch (aColor.ColorType)
                {
                    case TColorType.RGB:
                        DataStream.Write16(0x02);
                        UInt32 RGB = (UInt32)(aColor.RGB);
                        UInt32 BGR = 0xFF000000 | (RGB & 0x00FF00) | ((RGB & 0xFF0000) >> 16) | ((RGB & 0x0000FF) << 16);
                        DataStream.Write32((UInt32)(BGR));
                        break;

                    case TColorType.Automatic:
                        DataStream.Write16(0x00);
                        DataStream.Write32(0);
                        break;

                    case TColorType.Theme:
                        DataStream.Write16(0x03);
                        unchecked
                        {
                            DataStream.Write32((UInt32)aColor.Theme);
                        }
                        break;

                    case TColorType.Indexed:
                        DataStream.Write16(0x01);
                        unchecked
                        {
                            DataStream.Write32((UInt32)aColor.Index - 1 + 8);  //findex is 1-based too.
                        }
                        break;
                }
                DataStream.Write(BitConverter.GetBytes(gr.Stops[i].Position), 8);
                DataStream.Write(BitConverter.GetBytes(aColor.Tint), 8);

            }

        }

        private void WriteSpecialIdent(IDataStream DataStream)
        {
            if (FIndent < 16) return;
            DataStream.Write16((UInt16)TXFExtType.Indent);
            DataStream.Write16(6);

            UInt16 i = FIndent;
            if (i > 250) i = 250;
            DataStream.Write16(i);
        }

        private void WriteSpecialFont(IDataStream DataStream, TFontScheme aFontScheme)
        {
            if (aFontScheme != TFontScheme.None)
            {
                DataStream.Write16((UInt16)TXFExtType.FontScheme);
                DataStream.Write16(5);
                switch (aFontScheme)
                {
                    case TFontScheme.Minor:
                        DataStream.Write(new byte[] { 2 }, 1);
                        break;

                    case TFontScheme.Major:
                        DataStream.Write(new byte[] { 1 }, 1);
                        break;

                    default:
                        DataStream.Write(new byte[] { 0 }, 1);
                        break;
                }
            }
        }

        private void WriteSpecialColor(IDataStream DataStream, TXFExtType ExtType, TExcelColor aColor, IFlexCelPalette Palette)
        {
            if (!SpecialColor(aColor, Palette)) return;
            DataStream.Write16((UInt16)ExtType);
            DataStream.Write16(20);
            switch (aColor.ColorType)
            {
                case TColorType.RGB:
                    DataStream.Write16(0x02);
                    WriteNTint(DataStream, aColor);
                    UInt32 RGB = (UInt32)(aColor.RGB);
                    UInt32 BGR = 0xFF000000 | (RGB & 0x00FF00) | ((RGB & 0xFF0000) >> 16) | ((RGB & 0x0000FF) << 16);
                    DataStream.Write32((UInt32)(BGR));
                    break;

                case TColorType.Automatic:
                    DataStream.Write16(0x00);
                    WriteNTint(DataStream, aColor);
                    DataStream.Write32(0);
                    break;

                case TColorType.Theme:
                    DataStream.Write16(0x03);
                    WriteNTint(DataStream, aColor);
                    unchecked
                    {
                        DataStream.Write32((UInt32)aColor.Theme);
                    }
                    break;

                case TColorType.Indexed:
                    DataStream.Write16(0x01);
                    WriteNTint(DataStream, aColor);
                    unchecked
                    {
                        DataStream.Write32((UInt32)aColor.Index - 1 + 8);  //findex is 1-based too.
                    }
                    break;
            }
            DataStream.Write32(0);
            DataStream.Write32(0);

        }

        private static void WriteNTint(IDataStream DataStream, TExcelColor aColor)
        {
            Int16 nTint;
            if (aColor.Tint > 1) nTint = Int16.MaxValue;
            else if (aColor.Tint < -1) nTint = -Int16.MaxValue;
            else nTint = (Int16)Math.Round(aColor.Tint * Int16.MaxValue);
            unchecked
            {
                DataStream.Write16((UInt16)nTint);
            }
        }
        #endregion

        #region Load
        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.Biff8XF.Add(this);
        }
        #endregion

        internal void AddFutureStorage(TFutureStorageRecord R)
        {
            TFutureStorage.Add(ref FutureStorage, R);
        }
    }



#if FRAMEWORK20
    internal class TSortedXFCache : Dictionary<TXFRecord, int>
    {
#else
    internal class TSortedXFCache: Hashtable
    {
        internal bool TryGetValue(TXFRecord key, out int Result)
        {
            object oResult = this[key];
            if (oResult == null) 
            {
                Result = -1;
                return false;
            }
            Result = (int) oResult;
            return true;
        }

#endif

        internal void Remove(TXFRecord item, TXFRecordList RecordList)
        {
            //An XF record might be in more than one place, and have a single entry in the cache.
            //We should remove this only if there is no other place where this cache points to.

            int pos = (int)this[item];
            for (int i = 0; i < RecordList.Count; i++)
            {
                if (pos != i && RecordList[i].Equals(item))
                {
                    this[item] = i;
                    return;
                }

                Remove(item);
            }
        }
    }

    /// <summary>
    /// A list with XF Records.
    /// </summary>
    internal class TXFRecordList : TBaseRecordList<TXFRecord>          
    {
        private TSortedXFCache FSortedXfCache;

        internal TXFRecordList()
        {
            FSortedXfCache = new TSortedXFCache();
        }

        internal override void OnAdd(TXFRecord r, int index)
        {
            base.OnAdd(r, index);
            if (index < FList.Count - 1) UpgradeCache(index, 1);
            FSortedXfCache[r] = index;
        }

        internal override void OnDelete(TXFRecord r, int index)
        {
            base.OnDelete(r, index);
            TXFRecord RemoveKey = r;
            if (index <= FList.Count) UpgradeCache(index + 1, -1);

            int Pos;
            if (FSortedXfCache.TryGetValue(RemoveKey, out Pos))
            {
                if (Pos == index) FSortedXfCache.Remove(RemoveKey, this);
            }
        }

        private void UpgradeCache(int Pos, int ofs)
        {
#if FRAMEWORK20
            foreach (TXFRecord key in new List<TXFRecord>(FSortedXfCache.Keys))
#else
                foreach (TXFRecord key in new ArrayList(FSortedXfCache.Keys))
#endif
            {
                if (((int)FSortedXfCache[key]) >= Pos) FSortedXfCache[key] = ((int)FSortedXfCache[key]) + ofs;
            }
        }

        public override void Clear()
        {
            FSortedXfCache.Clear(); //Not really needed, only for performance.
            base.Clear();
        }

        internal bool FindFormat(TXFRecord XF, ref int Index)
        {
            if (!FSortedXfCache.TryGetValue(XF, out Index))
            {
                Index = -1;
            }

            return Index >= 0;
        }

        internal void MergeFromPxlXF(TXFRecordList SourceXFs, int BaseFont, TWorkbookGlobals Globals, TWorkbookGlobals SourceGlobals)
        {
            for (int i = 0; i < SourceXFs.Count; i++)
            {
                TXFRecord XF = SourceXFs[i];
                if (XF.FontIndex > 0)
                {
                    XF.FontIndex += BaseFont;
                    if (XF.FontIndex >= 4) XF.FontIndex++;
                }

                string Format = String.Empty;
                int FormatIndex = XF.FormatIndex;
                Format = SourceGlobals.Formats.FormatFromPxl(FormatIndex);

                XF.FormatIndex = Globals.Formats.AddFormat(Format);

                XF.MoveBordersAndPatternsToOtherFile(SourceGlobals.Borders, Globals.Borders, SourceGlobals.Patterns, Globals.Patterns);
                Add(XF);
            }
        }

        internal bool[] GetUsedColors(int ColorCount, TFontRecordList FontList, TBorderList BorderList, TPatternList PatternList)
        {
            bool[] Result = new bool[ColorCount];
            int aCount = Count;
            for (int i = 0; i < aCount; i++)
            {
                this[i].FillUsedColors(Result, FontList, BorderList, PatternList, null, null);
            }

            return Result;
        }

        internal void OptimizeColorPalette(int ColorCount,TWorkbookGlobals Globals, IFlexCelPalette xls)
        {
            bool[] UsedColors = new bool[ColorCount];
            TUsedColorDictionary ColorsInFile = new TUsedColorDictionary();
            int aCount = Count;
            for (int i = 0; i < aCount; i++)
            {
                this[i].FillUsedColors(UsedColors, Globals.Fonts, Globals.Borders, Globals.Patterns, ColorsInFile, xls);
            }

            IEnumerator<int> CurrentColor = ColorsInFile.GetEnumerator();
            

            for (int i = 3; i < ColorCount; i++) //first colors will be kept as black and white
            {
                if (UsedColors[i]) continue;

                if (!CurrentColor.MoveNext()) break;
                Color PaletteColor =  ColorUtil.FromArgb(CurrentColor.Current);
                Globals.SetColorPalette(i, PaletteColor);                   
            }

        }

        /// <summary>
        /// Note that this method will change the XF structure, so we need to remove it from the cache and add it again.
        /// </summary>
        internal void UpdateChangedStyleInCellXF(int StyleXF, TXFRecord StyleRecord, bool RemoveParent)
        {
            for (int i = Count - 1; i >= 0; i--)
            {
                if (this[i].Parent == StyleXF)
                {
                    int SearchPos;
                    if (FSortedXfCache.TryGetValue(this[i], out SearchPos) && SearchPos == i) FSortedXfCache.Remove(this[i], this); //only remove it if it was actually pointing there.
                    if (RemoveParent)
                    {
                        this[i].Parent = 0; //no need to merge here, since styles are kept merged when we modify the styles.
                    }
                    else
                    {
                        this[i].MergeWithParentStyle(StyleRecord);
                    }
                    FSortedXfCache[this[i]] = i;
                }
            }
        }

        public override void SaveToStream(IDataStream DataStream, TSaveData SaveData, int Row)
        {
            XlsMessages.ThrowException(XlsErr.ErrInternal); //we should call SaveAllToStream
        }

        public void SaveAllToStream(IDataStream DataStream, ref TSaveData SaveData, TXFRecordList CellXFs)
        {
            Biff8Utils.CheckXF(Count + CellXFs.Count);
            uint XFCRC = 0;

            ((TXFRecord)FList[0]).SaveToStream(DataStream, SaveData, 0, null, ref XFCRC); //Normal style. Always exists.

            int[] StylesSavedAtPos = new int[Count];
            SaveData.StylesSavedAtPos = StylesSavedAtPos;

            SaveData.AddedRecords[0] = 0;
            //Order is: 14 style records, normal xf record, rest of style records, rest of xf records.
            for (int i = 1; i < FlxConsts.DefaultFormatIdBiff8; i++)
            {
                string StyleName = TBuiltInStyles.GetName((byte)(1 + (i - 1) % 2), (i - 1) / 2);
                int StyleIndex = SaveData.Globals.Styles.GetStyle(StyleName);
                if (StyleIndex < 0)
                {
                    SaveData.AddedRecords[0]++;
                    TXFRecord.SaveDefaultOutline(DataStream, ref XFCRC); //outline styles are all the same and just apply the font.
                }
                else
                {
                    this[StyleIndex].SaveToStream(DataStream, SaveData, 0, null, ref XFCRC);
                    if (StylesSavedAtPos[StyleIndex] != 0) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid); //2 style records can't point to the same xf. This algorithm would leak in that case too.
                    StylesSavedAtPos[StyleIndex] = i;
                }
            }

            ((TXFRecord)CellXFs.FList[0]).SaveToStream(DataStream, SaveData, 0, StylesSavedAtPos, ref XFCRC);

            int CurrentPos = 16;
            int aStyleCount = FList.Count;
            for (int i = 1; i < aStyleCount; i++)
            {
                if (StylesSavedAtPos[i] == 0) //if it was saved then it is at least 1.
                {
                    StylesSavedAtPos[i] = CurrentPos;
                    ((TXFRecord)FList[i]).SaveToStream(DataStream, SaveData, 0, null, ref XFCRC);
                    CurrentPos++;
                }
            }

            int aCellCount = CellXFs.FList.Count;
            for (int i = 1; i < aCellCount; i++)
            {
                ((TXFRecord)CellXFs.FList[i]).SaveToStream(DataStream, SaveData, 0, StylesSavedAtPos, ref XFCRC);
            }

            if (NeedsXFExt(SaveData.Globals) || CellXFs.NeedsXFExt(SaveData.Globals))
            {
                SaveXFExt(DataStream, SaveData, CellXFs, XFCRC, StylesSavedAtPos);
            }

        }


        internal override long TotalSize
        {
            get
            {
                FlxMessages.ThrowException(FlxErr.ErrInternal); //this should not be called, sizeWithXFExt should be instead.
                return 0; //to compile
            }
        }

        internal long SizeWithoutExt
        {
            get
            {
                return base.TotalSize;
            }
        }

        internal long SizeWithXFExt(TWorkbookGlobals Globals, TXFRecordList CellXFs)
        {
            long Result = base.TotalSize + CellXFs.SizeWithoutExt;

            for (int i = 1; i < FlxConsts.DefaultFormatIdBiff8; i++)
            {
                string StyleName = TBuiltInStyles.GetName((byte)(1 + (i - 1) % 2), (i - 1) / 2);
                int StyleIndex = Globals.Styles.GetStyle(StyleName);
                if (StyleIndex < 0)
                {
                    Result += 20 + XlsConsts.SizeOfTRecordHeader; //We will add a fake style to complete the required 15 style records.
                }
            }


            if (NeedsXFExt(Globals) || CellXFs.NeedsXFExt(Globals))
            {
                Result += 20 + XlsConsts.SizeOfTRecordHeader; //XFCRC
                for (int i = 0; i < Count; i++)
                {
                    int len = this[i].XFExtLen(Globals);
                    if (len > 0)
                    {
                        Result += len + XlsConsts.SizeOfTRecordHeader;
                    }
                }

                for (int i = 0; i < CellXFs.Count; i++)
                {
                    int len = CellXFs[i].XFExtLen(Globals);
                    if (len > 0)
                    {
                        Result += len + XlsConsts.SizeOfTRecordHeader;
                    }
                }

            }
            return Result;
        }

        private bool NeedsXFExt(TWorkbookGlobals Globals)
        {
            for (int i = 0; i < Count; i++)
            {
                if (this[i].NeedsXFExt(Globals)) return true;
            }
            return false;
        }

        private void SaveXFExt(IDataStream DataStream, TSaveData SaveData, TXFRecordList CellXFs, uint XFCRC,int[] StylesSavedAtPos)
        {
            TXFCRCRecord.Save(DataStream, Count + CellXFs.Count + SaveData.AddedRecords[0], XFCRC);

            int aStyleCount = FList.Count;
            int MaxSavedPos = 0x0F;
            for (int i = 0; i < aStyleCount; i++)
            {
                ((TXFRecord)FList[i]).SaveXFExt(StylesSavedAtPos[i], DataStream, SaveData);
                if (StylesSavedAtPos[i] > MaxSavedPos) MaxSavedPos = StylesSavedAtPos[i];
            }

            ((TXFRecord)CellXFs.FList[0]).SaveXFExt(0x0F, DataStream, SaveData);

            int aCellCount = CellXFs.FList.Count;
            for (int i = 1; i < aCellCount; i++)
            {
                ((TXFRecord)CellXFs.FList[i]).SaveXFExt(i + MaxSavedPos, DataStream, SaveData);
            }
        }

        internal void AddExt(TXFExtRecord xfe, int xf, TWorkbookGlobals Globals)
        {
            //Need to invalidate the cache
            int SearchPos;
            if (FSortedXfCache.TryGetValue(this[xf], out SearchPos) && SearchPos == xf) FSortedXfCache.Remove(this[xf], this); //only remove it if it was actually pointing there.
            this[xf].AddExt(xfe, Globals);
            FSortedXfCache[this[xf]] = xf;
        }

        internal void EnsureMinimumCellXFs(TWorkbookGlobals Globals)
        {
            if (Count <= 0) Add(new TXFRecord(TFlxFormat.CreateStandard2007(), true, Globals, true));  //this will add style 0 if needed
        }

        internal void EnsureMinimumStyles(TStyleRecordList Styles, TWorkbookGlobals Globals)
        {
            //EnsureMinimumCellXFs will create normal style XF. Style the "STYLE" record might not have been created.
            string NormalStr = TBuiltInStyles.GetName((byte)TBuiltInStyle.Normal, 0);
            if (!Globals.Styles.HasStyle(NormalStr)) Globals.Styles.SetStyle(NormalStr, 0);
        }
    }

    #endregion

    #region XFExt
    internal enum TXFExtType
    {
        CellFgColor = 0x0004,
        CellBgColor = 0x0005,
        CellGradient = 0x0006,
        TopBorderColor = 0x0007,
        BottomBorderColor = 0x0008,
        LeftBorderColor = 0x0009,
        RightBorderColor = 0x000A,
        DiagBorderColor = 0x000B,
        TextColor = 0x000D,
        FontScheme = 0x000E,
        Indent = 0x000F
    }

    internal struct TXFExtProp
    {
        internal TXFExtType PropType;
        internal byte[] Data;
    }

    /// <summary>
    /// Extended properties in 2007.
    /// </summary>
    internal class TXFExtRecord : TxBaseRecord
    {
        public TXFExtRecord(byte[] aData)
            : base((int)xlr.XFEXT, aData)
        {
        }

        public int XF { get { return GetWord(14); } }
        public int Count { get { return GetWord(18); } }

        public TXFExtProp GetProp(ref int index)
        {
            TXFExtProp Result = new TXFExtProp();
            Result.PropType = (TXFExtType)GetWord(index);
            index += 2;
            Result.Data = new byte[GetWord(index) - 4];
            index += 2;
            Array.Copy(Data, index, Result.Data, 0, Result.Data.Length);
            index += Result.Data.Length;

            return Result;
        }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            WorkbookLoader.XFExtList.Add(this);
        }
    }

    internal class TXFExtRecordList : List<TXFExtRecord>
    {
    }


    /// <summary>
    /// CRC to know if XFExt records are valid.
    /// </summary>
    internal class TXFCRCRecord : TBaseRecord
    {
        private static UInt32[] CRCCache = InitCrcCache();
        internal UInt32 CRC;
        internal int XFCount;

        public TXFCRCRecord(byte[] Data)
        {
            XFCount = BitOps.GetWord(Data, 14);
            CRC = (UInt32)BitOps.GetCardinal(Data, 16);
        }

        internal static void UpdateCRC(byte[] Data, ref uint XFCRC)
        {
            foreach (byte b in Data)
            {
                unchecked
                {
                    UInt32 Index = XFCRC;
                    Index >>= 24;
                    Index ^= b;
                    XFCRC <<= 8;
                    XFCRC ^= CRCCache[Index];
                }
            }
        }

        internal static void Save(IDataStream DataStream, int Count, uint XFCRC)
        {
            DataStream.WriteHeader((UInt16)xlr.XFCRC, 20);
            DataStream.Write16((UInt16)xlr.XFCRC);
            DataStream.Write(new byte[12], 12);
            DataStream.Write16((UInt16)Count);
            DataStream.Write32(XFCRC);
        }

        private static UInt32[] InitCrcCache()
        {
            UInt32[] Result = new UInt32[256];
            for (UInt32 i = 0; i < Result.Length; i++)
            {
                unchecked
                {
                    UInt32 Value = i;
                    Value <<= 24;
                    for (int b = 0; b < 8; b++)
                    {
                        if ((Value & 0x80000000) != 0)
                        {
                            Value <<= 1;
                            Value ^= 0xAF;
                        }
                        else
                        {
                            Value <<= 1;
                        }

                    }
                    Result[i] = Value & 0xFFFF;
                }
            }

            return Result;
        }


        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            WorkbookLoader.XFCRC = CRC;
            WorkbookLoader.XFCount = XFCount;
        }

        #region Not implemented
        protected override TBaseRecord DoCopyTo(TSheetInfo SheetInfo)
        {
            FlxMessages.ThrowException(FlxErr.ErrInternal);
            return null;
        }


        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
        {
            FlxMessages.ThrowException(FlxErr.ErrInternal);
        }

        internal override int TotalSize()
        {
            FlxMessages.ThrowException(FlxErr.ErrInternal);
            return -1;
        }

        internal override int TotalSizeNoHeaders()
        {
            FlxMessages.ThrowException(FlxErr.ErrInternal);
            return -1;
        }

        internal override int GetId
        {

            get { FlxMessages.ThrowException(FlxErr.ErrInternal); return -1; }
        }
        #endregion

    }

    #endregion

    #region Font


    internal class TFontDat
    {
        byte[] Data = null;
        internal TFontDat(byte[] aData)
        {
            Data = aData;
        }

        internal int Height { get { return BitOps.GetWord(Data, 0); } set { BitOps.SetWord(Data, 0, value); } }
        internal int GrBit { get { return BitOps.GetWord(Data, 2); } set { BitOps.SetWord(Data, 2, value); } }
        internal int ColorIndex { get { return BitOps.GetWord(Data, 4); } set { BitOps.SetWord(Data, 4, value); } }
        internal int BoldStyle { get { return BitOps.GetWord(Data, 6); } set { BitOps.SetWord(Data, 6, value); } }
        internal int SuperSub { get { return BitOps.GetWord(Data, 8); } set { BitOps.SetWord(Data, 8, value); } }
        internal byte Underline { get { return Data[10]; } set { Data[10] = value; } }
        internal byte Family { get { return Data[11]; } set { Data[11] = value; } }
        internal byte CharSet { get { return Data[12]; } set { Data[12] = value; } }
        internal byte Reserved { get { return Data[13]; } set { Data[13] = value; } }

        //Font name is not included
        internal static int StaticLength { get { return 14; } }

        internal bool Reserved1 { get { return (BitOps.GetWord(Data, 2) & 0x1) != 0; } }
        internal bool Reserved2 { get { return (BitOps.GetWord(Data, 2) & 0x4) != 0; } }
        internal TFlxFontStyles GetStyle()
        {
            TFlxFontStyles Result = TFlxFontStyles.None;
            if (BitOps.GetWord(Data, 6) == 0x2BC) Result |= TFlxFontStyles.Bold;
            int Flags = BitOps.GetWord(Data, 2);
            if ((Flags & 0x02) != 0) Result |= TFlxFontStyles.Italic;
            if ((Flags & 0x08) != 0) Result |= TFlxFontStyles.StrikeOut;
            if ((Flags & 0x10) != 0) Result |= TFlxFontStyles.Outline;
            if ((Flags & 0x20) != 0) Result |= TFlxFontStyles.Shadow;
            if ((Flags & 0x40) != 0) Result |= TFlxFontStyles.Condense;
            if ((Flags & 0x80) != 0) Result |= TFlxFontStyles.Extend;
            switch (BitOps.GetWord(Data, 8))
            {
                case 1: Result |= TFlxFontStyles.Superscript; break;
                case 2: Result |= TFlxFontStyles.Subscript; break;
            } //case

            return Result;
        }

        internal TFlxUnderline GetUnderline()
        {
            switch (Data[10])
            {
                case 0x01: return TFlxUnderline.Single;
                case 0x02: return TFlxUnderline.Double;
                case 0x21: return TFlxUnderline.SingleAccounting;
                case 0x22: return TFlxUnderline.DoubleAccounting;
                default: return TFlxUnderline.None;
            }//case
        }

        internal string GetName()
        {
            string s = null;
            long ssize = 0;
            StrOps.GetSimpleString(false, Data, 14, false, 0, ref s, ref ssize);
            return s;
        }

    }


    internal class TFontRecord : TBaseRecord
    {
        private int Id;

        internal bool Reuse; // Font records used in chartfbi should not be used for anything else.
        internal int CopiedTo; //to fix FBIs

        private string FName;
        private int FSize20;
        internal TExcelColor Color;
        private TFlxFontStyles FStyle;
        private TFlxUnderline FUnderline;
        private byte FFamily;
        private byte FCharSet;
        private TFontScheme FScheme;
        private bool Reserved1, Reserved2;
        private byte Reserved;
        internal TColorIndexCache FontColor;

        internal TFontRecord(int aId, byte[] aData)
        {
            Init(aId);
            LoadFromBiff8(aData);
        }

        private void Init(int aId)
        {
            Id = aId;
            Reuse = true;
            CopiedTo = -1;
        }

        internal TFontRecord(TFlxFont aFont)
            : this(aFont, true)
        {
        }

        /// <summary>
        /// CreateFromFlxFont
        /// </summary>
        internal TFontRecord(TFlxFont aFont, bool aReuse)
        {
            Init((int)xlr.FONT);
            Reuse = aReuse;

            FName = aFont.Name;
            FSize20 = aFont.Size20;
            Color = aFont.Color;
            FStyle = aFont.Style;
            FUnderline = aFont.Underline;
            FFamily = aFont.Family;
            FCharSet = aFont.CharSet;
            FScheme = aFont.Scheme;
        }


        private byte[] SaveToBiff8(IFlexCelPalette xls)
        {
            TExcelString Xs = new TExcelString(TStrLenLength.is8bits, FName, null, true);
            Byte[] Data = new byte[TFontDat.StaticLength + Xs.TotalSize()];
            TFontDat FontDat = new TFontDat(Data);

            FontDat.Height = FSize20;
            int GrBit = 0;
            if ((FStyle & TFlxFontStyles.Italic) != 0) GrBit |= 0x02;
            if ((FStyle & TFlxFontStyles.StrikeOut) != 0) GrBit |= 0x08;

            if ((FStyle & TFlxFontStyles.Outline) != 0) GrBit |= 0x10;
            if ((FStyle & TFlxFontStyles.Shadow) != 0) GrBit |= 0x20;
            if ((FStyle & TFlxFontStyles.Condense) != 0) GrBit |= 0x40;
            if ((FStyle & TFlxFontStyles.Extend) != 0) GrBit |= 0x80;

            if (Reserved1) GrBit |= 0x01;
            if (Reserved2) GrBit |= 0x04;

            FontDat.GrBit = GrBit;

            FontDat.ColorIndex = Color.GetBiff8ColorIndex(xls, TAutomaticColor.Font, ref FontColor);
            if ((FStyle & TFlxFontStyles.Bold) != 0) FontDat.BoldStyle = 0x2BC; else FontDat.BoldStyle = 0x190;
            if ((FStyle & TFlxFontStyles.Subscript) != 0) FontDat.SuperSub = 2;
            else if ((FStyle & TFlxFontStyles.Superscript) != 0) FontDat.SuperSub = 1;
            else FontDat.SuperSub = 0;
            switch (FUnderline)
            {
                case TFlxUnderline.Single: FontDat.Underline = 0x01; break;
                case TFlxUnderline.Double: FontDat.Underline = 0x02; break;
                case TFlxUnderline.SingleAccounting: FontDat.Underline = 0x21; break;
                case TFlxUnderline.DoubleAccounting: FontDat.Underline = 0x22; break;
                default: FontDat.Underline = 0; break;
            } //case

            FontDat.Family = FFamily;
            FontDat.CharSet = FCharSet;
            FontDat.Reserved = Reserved;

            Xs.CopyToPtr(Data, TFontDat.StaticLength);

            return Data;
        }

        private void LoadFromBiff8(byte[] Data)
        {
            TFontDat FontDat = new TFontDat(Data);
            FName = FontDat.GetName();
            FSize20 = FontDat.Height;
            Color = TExcelColor.FromBiff8ColorIndex(FontDat.ColorIndex);
            FStyle = FontDat.GetStyle();
            FUnderline = FontDat.GetUnderline();
            FFamily = FontDat.Family;
            FCharSet = FontDat.CharSet;
            Reserved1 = FontDat.Reserved1;
            Reserved2 = FontDat.Reserved2;
            Reserved = FontDat.Reserved;

        }

        internal bool SameFont(TFontRecord aFont)
        {
            if (FName != aFont.FName) return false;
            if (FSize20 != aFont.FSize20) return false;
            if (Color != aFont.Color) return false;
            if (FStyle != aFont.FStyle) return false;
            if (FUnderline != aFont.FUnderline) return false;
            if (FFamily != aFont.FFamily) return false;
            if (FCharSet != aFont.FCharSet) return false;
            if (FScheme != aFont.FScheme) return false;
            return true;
        }

        protected override TBaseRecord DoCopyTo(TSheetInfo SheetInfo)
        {
            TFontRecord Result = (TFontRecord)MemberwiseClone();
            Result.Init(Id);
            return Result;
        }

        internal string Name
        {
            get
            {
                return FName;
            }
        }

        internal TFontScheme Scheme { get { return FScheme; } set { FScheme = value; } }

        internal TFlxFont FlxFont()
        {
            TFlxFont Result = new TFlxFont();
            Result.Name = Name;
            Result.Size20 = FSize20;
            Result.Color = Color;
            Result.Style = FStyle;
            Result.Underline = FUnderline;
            Result.Family = FFamily;
            Result.CharSet = FCharSet;
            Result.Scheme = FScheme;

            return Result;
        }

        internal override int GetId
        {
            get { return Id; }
        }

        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
        {
            byte[] Data = SaveToBiff8(SaveData.Palette);
            Workbook.WriteHeader((UInt16)Id, (UInt16)Data.Length);
            Workbook.Write(Data, Data.Length);
        }

        internal override int TotalSize()
        {
            return TotalSizeNoHeaders() + XlsConsts.SizeOfTRecordHeader;
        }

        internal override int TotalSizeNoHeaders()
        {
            TExcelString Xs = new TExcelString(TStrLenLength.is8bits, FName, null, true);
            return TFontDat.StaticLength + Xs.TotalSize();
        }

        internal override void SaveToPxl(TPxlStream PxlStream, int Row, TPxlSaveData SaveData)
        {
            base.SaveToPxl(PxlStream, Row, SaveData);
            PxlStream.WriteByte((byte)pxl.FONT);
            byte[] NewData = SaveToBiff8(SaveData.Palette);
            BitOps.SetWord(NewData, 4, Color.GetPxlColorIndex(SaveData.Palette, TAutomaticColor.Font));
            PxlStream.Write(NewData, 0, 14);

            PxlStream.WriteString8(Name);
        }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.Fonts.Add(this);
        }

    }

    internal class TFontRecordList : TBaseRecordList<TFontRecord>
    {
        internal int AddFont(TFlxFont aFont)
        {
            int Result = -1;
            TFontRecord TempFont = new TFontRecord(aFont);  //.CreateFromFlxFont
            for (int i = 0; i < Count; i++)
                if (this[i].Reuse && TempFont.SameFont(this[i]))
                {
                    Result = i; if (Result >= 4) Result++; //Font number 4 does not exist
                    return Result;
                }

            Result = Count;
            Add(TempFont);
            if (Result >= 4) Result++; //Font number 4 does not exist
            return Result;
        }

        internal int AddNotReusableFont(TFontRecord aFont)
        {
            Add(aFont);
            int Result = Count - 1;
            if (Result >= 4) Result++; //Font number 4 does not exist
            return Result;
        }

        public TFlxFont GetFont(int fontIndex)
        {
            if (fontIndex == 4) fontIndex = 0;  //font 4 does not exist
            if (fontIndex > 4) fontIndex--;
            if ((fontIndex < 0) || (fontIndex >= Count)) fontIndex = 0;
            return this[fontIndex].FlxFont();
        }

        public TFontRecord GetFontRecord(int fontIndex)
        {
            if (fontIndex == 4) fontIndex = 0;  //font 4 does not exist
            if (fontIndex > 4) fontIndex--;
            if ((fontIndex < 0) || (fontIndex >= Count)) fontIndex = 0;
            return this[fontIndex];
        }

        public void SetFont(int fontIndex, TFlxFont aFont)
        {
            if (fontIndex == 4) return;  //font 4 does not exist
            if (fontIndex > 4) fontIndex--;
            if ((fontIndex < 0) || (fontIndex >= Count)) return;

            this[fontIndex] = new TFontRecord(aFont, this[fontIndex].Reuse);
        }

        internal void MergeFromPxlFont(TFontRecordList SourceFonts)
        {
            if (SourceFonts.Count > 0)  //copy format 0.
            {
                this[0] = SourceFonts[0];
            }
            for (int i = 1; i < SourceFonts.Count; i++)
            {
                TFontRecord Font = SourceFonts[i];
                Add(Font);
            }
        }
    }

    #endregion

    #region Style

    internal sealed class TBuiltInStyles
    {
        #region Internal Names
        static string[] InternalNames =
			{
				FlxConsts.NormalStyleName,
				"RowLevel_",
				"ColLevel_",
				"Comma",
				"Currency",
				"Percent",
				"Comma [0]",
				"Currency [0]",
				"Hyperlink",
				"Followed Hyperlink",
				"Note",
				"Warning Text",
				"Emphasis 1 (obsolete)",
				"Emphasis 2 (obsolete)",
				"Emphasis 3 (obsolete)",
				"Title",
				"Heading 1",
				"Heading 2",
				"Heading 3",
				"Heading 4",
				"Input",
				"Output",
				"Calculation",
				"Check Cell",
				"Linked Cell",
				"Total",
				"Good",
				"Bad",
				"Neutral",
				"Accent1",
				"20% - Accent1",
				"40% - Accent1",
				"60% - Accent1",
				"Accent2",
				"20% - Accent2",
				"40% - Accent2",
				"60% - Accent2",
				"Accent3",
				"20% - Accent3",
				"40% - Accent3",
				"60% - Accent3",
				"Accent4",
				"20% - Accent4",
				"40% - Accent4",
				"60% - Accent4",
				"Accent5",
				"20% - Accent5",
				"40% - Accent5",
				"60% - Accent5",
				"Accent6",
				"20% - Accent6",
				"40% - Accent6",
				"60% - Accent6",
				"Explanatory Text"
			};
        #endregion

        #region Internal Categories
        static TStyleCategory[] Categories =
        {
            TStyleCategory.GoodBadNeutral,
            TStyleCategory.DataModel,
            TStyleCategory.DataModel,
            TStyleCategory.NumberFormat,
            TStyleCategory.NumberFormat,
            TStyleCategory.NumberFormat,
            TStyleCategory.NumberFormat,
            TStyleCategory.NumberFormat,
            TStyleCategory.DataModel,
            TStyleCategory.DataModel,
            TStyleCategory.DataModel,
            TStyleCategory.DataModel,
            TStyleCategory.ThemedCell,
            TStyleCategory.ThemedCell,
            TStyleCategory.ThemedCell,
            TStyleCategory.TitleHeading,
            TStyleCategory.TitleHeading,
            TStyleCategory.TitleHeading,
            TStyleCategory.TitleHeading,
            TStyleCategory.TitleHeading,
            TStyleCategory.DataModel,
            TStyleCategory.DataModel,
            TStyleCategory.DataModel,
            TStyleCategory.DataModel,
            TStyleCategory.DataModel,
            TStyleCategory.TitleHeading,
            TStyleCategory.GoodBadNeutral,
            TStyleCategory.GoodBadNeutral,
            TStyleCategory.GoodBadNeutral,
            TStyleCategory.ThemedCell,
            TStyleCategory.ThemedCell,
            TStyleCategory.ThemedCell,
            TStyleCategory.ThemedCell,
            TStyleCategory.ThemedCell,
            TStyleCategory.ThemedCell,
            TStyleCategory.ThemedCell,
            TStyleCategory.ThemedCell,
            TStyleCategory.ThemedCell,
            TStyleCategory.ThemedCell,
            TStyleCategory.ThemedCell,
            TStyleCategory.ThemedCell,
            TStyleCategory.ThemedCell,
            TStyleCategory.ThemedCell,
            TStyleCategory.ThemedCell,
            TStyleCategory.ThemedCell,
            TStyleCategory.ThemedCell,
            TStyleCategory.ThemedCell,
            TStyleCategory.ThemedCell,
            TStyleCategory.ThemedCell,
            TStyleCategory.ThemedCell,
            TStyleCategory.ThemedCell,
            TStyleCategory.ThemedCell,
            TStyleCategory.ThemedCell,
            TStyleCategory.DataModel
        };
        #endregion

        #region Internal definitions

        private static TFlxFormat GetInternalNameDef(int BuiltInId)
        {
            if (BuiltInId < 0) return null;
            TFlxFormat StyleFmt = TFlxFormat.CreateStandard2007();
            StyleFmt.IsStyle = true;
            switch ((TBuiltInStyle)BuiltInId)
            {
                case TBuiltInStyle.Good:
                    StyleFmt.Font.Color = ColorUtil.FromArgb(0x00, 0x61, 0x00);
                    StyleFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
                    StyleFmt.FillPattern.FgColor = ColorUtil.FromArgb(0xC6, 0xEF, 0xCE);
                    break;

                case TBuiltInStyle.Comma:
                    StyleFmt.Format = "_(* #,##0.00_);_(* \\(#,##0.00\\);_(* \"-\"??_);_(@_)";
                    break;

                case TBuiltInStyle.Accent1_40_percent:
                    StyleFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
                    StyleFmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent1, 0.599993896298105);
                    break;

                case TBuiltInStyle.Accent2_40_percent:
                    StyleFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
                    StyleFmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent2, 0.599993896298105);
                    break;

                case TBuiltInStyle.Accent3_40_percent:
                    StyleFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
                    StyleFmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
                    break;

                case TBuiltInStyle.Accent4_40_percent:
                    StyleFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
                    StyleFmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
                    break;

                case TBuiltInStyle.Accent5_40_percent:
                    StyleFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
                    StyleFmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.599993896298105);
                    break;

                case TBuiltInStyle.Accent6_40_percent:
                    StyleFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
                    StyleFmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.599993896298105);
                    break;

                case TBuiltInStyle.Accent2_60_percent:
                    StyleFmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
                    StyleFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
                    StyleFmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent2, 0.399975585192419);
                    break;

                case TBuiltInStyle.Accent3_60_percent:
                    StyleFmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
                    StyleFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
                    StyleFmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
                    break;

                case TBuiltInStyle.Accent1_60_percent:
                    StyleFmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
                    StyleFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
                    StyleFmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent1, 0.399975585192419);
                    break;

                case TBuiltInStyle.Accent6_60_percent:
                    StyleFmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
                    StyleFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
                    StyleFmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
                    break;

                case TBuiltInStyle.Accent4_60_percent:
                    StyleFmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
                    StyleFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
                    StyleFmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
                    break;

                case TBuiltInStyle.Accent5_60_percent:
                    StyleFmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
                    StyleFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
                    StyleFmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.399975585192419);
                    break;

                case TBuiltInStyle.Normal:
                    break;

                case TBuiltInStyle.Neutral:
                    StyleFmt.Font.Color = ColorUtil.FromArgb(0x9C, 0x65, 0x00);
                    StyleFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
                    StyleFmt.FillPattern.FgColor = ColorUtil.FromArgb(0xFF, 0xEB, 0x9C);
                    break;

                case TBuiltInStyle.Heading_1:
                    StyleFmt.Font.Size20 = 300;
                    StyleFmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2);
                    StyleFmt.Font.Style = TFlxFontStyles.Bold;
                    StyleFmt.Borders.Bottom.Style = TFlxBorderStyle.Thick;
                    StyleFmt.Borders.Bottom.Color = TExcelColor.FromTheme(TThemeColor.Accent1);
                    break;

                case TBuiltInStyle.Bad:
                    StyleFmt.Font.Color = ColorUtil.FromArgb(0x9C, 0x00, 0x06);
                    StyleFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
                    StyleFmt.FillPattern.FgColor = ColorUtil.FromArgb(0xFF, 0xC7, 0xCE);
                    break;

                case TBuiltInStyle.Comma0:
                    StyleFmt.Format = "_(* #,##0_);_(* \\(#,##0\\);_(* \"-\"_);_(@_)";
                    break;

                case TBuiltInStyle.Calculation:
                    StyleFmt.Font.Color = ColorUtil.FromArgb(0xFA, 0x7D, 0x00);
                    StyleFmt.Font.Style = TFlxFontStyles.Bold;
                    StyleFmt.Borders.Left.Style = TFlxBorderStyle.Thin;
                    StyleFmt.Borders.Left.Color = ColorUtil.FromArgb(0x7F, 0x7F, 0x7F);
                    StyleFmt.Borders.Right.Style = TFlxBorderStyle.Thin;
                    StyleFmt.Borders.Right.Color = ColorUtil.FromArgb(0x7F, 0x7F, 0x7F);
                    StyleFmt.Borders.Top.Style = TFlxBorderStyle.Thin;
                    StyleFmt.Borders.Top.Color = ColorUtil.FromArgb(0x7F, 0x7F, 0x7F);
                    StyleFmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
                    StyleFmt.Borders.Bottom.Color = ColorUtil.FromArgb(0x7F, 0x7F, 0x7F);
                    StyleFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
                    StyleFmt.FillPattern.FgColor = ColorUtil.FromArgb(0xF2, 0xF2, 0xF2);
                    break;

                case TBuiltInStyle.Title:
                    StyleFmt.Font.Name = "Cambria";
                    StyleFmt.Font.Size20 = 360;
                    StyleFmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2);
                    StyleFmt.Font.Style = TFlxFontStyles.Bold;
                    StyleFmt.Font.Scheme = TFontScheme.Major;
                    break;

                case TBuiltInStyle.Hyperlink:
                    StyleFmt.Font.Color = TExcelColor.FromTheme(TThemeColor.HyperLink);
                    StyleFmt.Font.Underline = TFlxUnderline.Single;
                    StyleFmt.VAlignment = TVFlxAlignment.top;
                    StyleFmt.Locked = false;
                    break;

                case TBuiltInStyle.Followed_Hyperlink:
                    StyleFmt.Font.Color = TExcelColor.FromTheme(TThemeColor.FollowedHyperLink);
                    StyleFmt.Font.Underline = TFlxUnderline.Single;
                    StyleFmt.VAlignment = TVFlxAlignment.top;
                    StyleFmt.Locked = false;
                    break;

                case TBuiltInStyle.Linked_Cell:
                    StyleFmt.Font.Color = ColorUtil.FromArgb(0xFA, 0x7D, 0x00);
                    StyleFmt.Borders.Bottom.Style = TFlxBorderStyle.Double;
                    StyleFmt.Borders.Bottom.Color = ColorUtil.FromArgb(0xFF, 0x80, 0x01);
                    break;

                case TBuiltInStyle.Heading_2:
                    StyleFmt.Font.Size20 = 260;
                    StyleFmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2);
                    StyleFmt.Font.Style = TFlxFontStyles.Bold;
                    StyleFmt.Borders.Bottom.Style = TFlxBorderStyle.Thick;
                    StyleFmt.Borders.Bottom.Color = TExcelColor.FromTheme(TThemeColor.Accent1, 0.499984740745262);
                    break;

                case TBuiltInStyle.Heading_3:
                    StyleFmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2);
                    StyleFmt.Font.Style = TFlxFontStyles.Bold;
                    StyleFmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
                    StyleFmt.Borders.Bottom.Color = TExcelColor.FromTheme(TThemeColor.Accent1, 0.399975585192419);
                    break;

                case TBuiltInStyle.Warning_Text:
                    StyleFmt.Font.Color = ColorUtil.FromArgb(0xFF, 0x00, 0x00);
                    break;

                case TBuiltInStyle.Currency0:
                    StyleFmt.Format = "_(\"$\"\\ * #,##0_);_(\"$\"\\ * \\(#,##0\\);_(\"$\"\\ * \"-\"_);_(@_)";
                    break;

                case TBuiltInStyle.Check_Cell:
                    StyleFmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
                    StyleFmt.Font.Style = TFlxFontStyles.Bold;
                    StyleFmt.Borders.Left.Style = TFlxBorderStyle.Double;
                    StyleFmt.Borders.Left.Color = ColorUtil.FromArgb(0x3F, 0x3F, 0x3F);
                    StyleFmt.Borders.Right.Style = TFlxBorderStyle.Double;
                    StyleFmt.Borders.Right.Color = ColorUtil.FromArgb(0x3F, 0x3F, 0x3F);
                    StyleFmt.Borders.Top.Style = TFlxBorderStyle.Double;
                    StyleFmt.Borders.Top.Color = ColorUtil.FromArgb(0x3F, 0x3F, 0x3F);
                    StyleFmt.Borders.Bottom.Style = TFlxBorderStyle.Double;
                    StyleFmt.Borders.Bottom.Color = ColorUtil.FromArgb(0x3F, 0x3F, 0x3F);
                    StyleFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
                    StyleFmt.FillPattern.FgColor = ColorUtil.FromArgb(0xA5, 0xA5, 0xA5);
                    break;

                case TBuiltInStyle.Heading_4:
                    StyleFmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2);
                    StyleFmt.Font.Style = TFlxFontStyles.Bold;
                    break;

                case TBuiltInStyle.Output:
                    StyleFmt.Font.Color = ColorUtil.FromArgb(0x3F, 0x3F, 0x3F);
                    StyleFmt.Font.Style = TFlxFontStyles.Bold;
                    StyleFmt.Borders.Left.Style = TFlxBorderStyle.Thin;
                    StyleFmt.Borders.Left.Color = ColorUtil.FromArgb(0x3F, 0x3F, 0x3F);
                    StyleFmt.Borders.Right.Style = TFlxBorderStyle.Thin;
                    StyleFmt.Borders.Right.Color = ColorUtil.FromArgb(0x3F, 0x3F, 0x3F);
                    StyleFmt.Borders.Top.Style = TFlxBorderStyle.Thin;
                    StyleFmt.Borders.Top.Color = ColorUtil.FromArgb(0x3F, 0x3F, 0x3F);
                    StyleFmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
                    StyleFmt.Borders.Bottom.Color = ColorUtil.FromArgb(0x3F, 0x3F, 0x3F);
                    StyleFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
                    StyleFmt.FillPattern.FgColor = ColorUtil.FromArgb(0xF2, 0xF2, 0xF2);
                    break;

                case TBuiltInStyle.Note:
                    StyleFmt.Borders.Left.Style = TFlxBorderStyle.Thin;
                    StyleFmt.Borders.Left.Color = ColorUtil.FromArgb(0xB2, 0xB2, 0xB2);
                    StyleFmt.Borders.Right.Style = TFlxBorderStyle.Thin;
                    StyleFmt.Borders.Right.Color = ColorUtil.FromArgb(0xB2, 0xB2, 0xB2);
                    StyleFmt.Borders.Top.Style = TFlxBorderStyle.Thin;
                    StyleFmt.Borders.Top.Color = ColorUtil.FromArgb(0xB2, 0xB2, 0xB2);
                    StyleFmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
                    StyleFmt.Borders.Bottom.Color = ColorUtil.FromArgb(0xB2, 0xB2, 0xB2);
                    StyleFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
                    StyleFmt.FillPattern.FgColor = ColorUtil.FromArgb(0xFF, 0xFF, 0xCC);
                    break;

                case TBuiltInStyle.Currency:
                    StyleFmt.Format = "_(\"$\"\\ * #,##0.00_);_(\"$\"\\ * \\(#,##0.00\\);_(\"$\"\\ * \"-\"??_);_(@_)";
                    break;

                case TBuiltInStyle.Accent4:
                    StyleFmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
                    StyleFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
                    StyleFmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4);
                    break;

                case TBuiltInStyle.Accent5:
                    StyleFmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
                    StyleFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
                    StyleFmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5);
                    break;

                case TBuiltInStyle.Accent6:
                    StyleFmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
                    StyleFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
                    StyleFmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6);
                    break;

                case TBuiltInStyle.Accent1:
                    StyleFmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
                    StyleFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
                    StyleFmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent1);
                    break;

                case TBuiltInStyle.Accent2:
                    StyleFmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
                    StyleFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
                    StyleFmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent2);
                    break;

                case TBuiltInStyle.Accent3:
                    StyleFmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
                    StyleFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
                    StyleFmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3);
                    break;

                case TBuiltInStyle.Total:
                    StyleFmt.Font.Style = TFlxFontStyles.Bold;
                    StyleFmt.Borders.Top.Style = TFlxBorderStyle.Thin;
                    StyleFmt.Borders.Top.Color = TExcelColor.FromTheme(TThemeColor.Accent1);
                    StyleFmt.Borders.Bottom.Style = TFlxBorderStyle.Double;
                    StyleFmt.Borders.Bottom.Color = TExcelColor.FromTheme(TThemeColor.Accent1);
                    break;

                case TBuiltInStyle.Input:
                    StyleFmt.Font.Color = ColorUtil.FromArgb(0x3F, 0x3F, 0x76);
                    StyleFmt.Borders.Left.Style = TFlxBorderStyle.Thin;
                    StyleFmt.Borders.Left.Color = ColorUtil.FromArgb(0x7F, 0x7F, 0x7F);
                    StyleFmt.Borders.Right.Style = TFlxBorderStyle.Thin;
                    StyleFmt.Borders.Right.Color = ColorUtil.FromArgb(0x7F, 0x7F, 0x7F);
                    StyleFmt.Borders.Top.Style = TFlxBorderStyle.Thin;
                    StyleFmt.Borders.Top.Color = ColorUtil.FromArgb(0x7F, 0x7F, 0x7F);
                    StyleFmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
                    StyleFmt.Borders.Bottom.Color = ColorUtil.FromArgb(0x7F, 0x7F, 0x7F);
                    StyleFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
                    StyleFmt.FillPattern.FgColor = ColorUtil.FromArgb(0xFF, 0xCC, 0x99);
                    break;

                case TBuiltInStyle.Percent:
                    StyleFmt.Font.Family = 2;
                    break;

                case TBuiltInStyle.Accent6_20_percent:
                    StyleFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
                    StyleFmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
                    break;

                case TBuiltInStyle.Accent5_20_percent:
                    StyleFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
                    StyleFmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
                    break;

                case TBuiltInStyle.Accent4_20_percent:
                    StyleFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
                    StyleFmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
                    break;

                case TBuiltInStyle.Explanatory_Text:
                    StyleFmt.Font.Color = ColorUtil.FromArgb(0x7F, 0x7F, 0x7F);
                    StyleFmt.Font.Style = TFlxFontStyles.Italic;
                    break;

                case TBuiltInStyle.Accent2_20_percent:
                    StyleFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
                    StyleFmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent2, 0.799981688894314);
                    break;

                case TBuiltInStyle.Accent3_20_percent:
                    StyleFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
                    StyleFmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
                    break;

                case TBuiltInStyle.Accent1_20_percent:
                    StyleFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
                    StyleFmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent1, 0.799981688894314);
                    break;
            }
            return StyleFmt;
        }
        #endregion

        private TBuiltInStyles() { }

        internal static string GetName(byte Id, int Level)
        {
            if (Id >= InternalNames.Length) return "INTERNAL__" + Id.ToString(CultureInfo.InvariantCulture);
            if (Id == 1 || Id == 2) return InternalNames[Id] + (Level + 1).ToString(CultureInfo.InvariantCulture);
            return InternalNames[Id];
        }

        internal static bool IsBuiltInGlobal(int Id)
        {
            return Id >= 0;
        }

        internal static int GetIdAndLevel(string Name, out int Level)
        {
            Level = 0;
            if (Name == null) return 0;
            Name = Name.Trim();
            if (string.Equals(InternalNames[0], Name, StringComparison.InvariantCultureIgnoreCase)) return 0;

            for (int i = 3; i < InternalNames.Length; i++)
            {
                if (string.Equals(InternalNames[i], Name, StringComparison.InvariantCultureIgnoreCase)) return i;
            }

            for (int i = 1; i < 3; i++)
            {
                if (string.Compare(InternalNames[i], 0, Name, 0, InternalNames[i].Length, StringComparison.InvariantCultureIgnoreCase) == 0)
                {
                    if (Name.Length > InternalNames[i].Length)
                    {
                        string rest = Name.Substring(InternalNames[i].Length);
                        if (rest.Length == 1 && rest[0] > '0' && rest[0] <= '9')
                        {
                            Level = (int)(rest[0] - '1');
                            return i;
                        }
                    }
                }
            }

            return -1;
        }

        internal static TStyleCategory GetCategory(int StyleId)
        {
            if (StyleId < 0 || StyleId >= Categories.Length) return TStyleCategory.Custom;
            return Categories[StyleId];
        }

        internal static TFlxFormat GetDefaultStyle(int BuiltinId, int Level)
        {
            return GetInternalNameDef(BuiltinId);
        }
    }

    internal class TStyleRecord : TBaseRecord, IComparable
    {
        private string FName;
        internal int XF;
        internal int BuiltinId;
        internal int iLevel;
        internal bool Hidden;
        internal bool CustomBuiltin;
        private TStyleCategory Category;
        internal byte[] XFProps; //We really should apply this to the corresponding XF record at loading and save it depending on the XF record too. In any case, a XFExt takes priority over what's here, so if the file wasn't invalidated by Excel 2003, this is useless info. For the future...

        internal bool IgnoreIt;
        internal TFutureStorage FutureStorage;

        internal TStyleRecord(string aName, int aXF, int aBuiltInId, int aLevel, bool aHidden, bool aCustombuiltIn, TStyleCategory aCategory)
        {
            FName = ValidateName(aName, true);
            XF = aXF;
            BuiltinId = aBuiltInId;
            iLevel = aLevel;
            Hidden = aHidden;
            CustomBuiltin = aCustombuiltIn;
            Category = aCategory;
        }

        internal TStyleRecord(int aId, byte[] aData, TBiff8XFMap XFMap)
        {
            XF = BitOps.GetWord(aData, 0) & 0xFFF;
            if (XFMap != null) XF = XFMap.GetStyleXF2007(XF);

            bool IsBuiltIn = (aData[1] & 0x80) != 0;

            if (IsBuiltIn)
            {
                BuiltinId = aData[2];
                iLevel = aData[3];
                FName = TBuiltInStyles.GetName(aData[2], aData[3]);
            }
            else
            {
                BuiltinId = -1;
                iLevel = -1;

                string StrValue = String.Empty;
                long StrSize = 0;
                StrOps.GetSimpleString(true, aData, 2, false, 0, ref StrValue, ref StrSize);
                FName = StrValue;
            }
        }

        internal void AddStyleEx(TStyleExRecord aStyleExt)
        {
            if (FName != aStyleExt.Name) return; //invalid
            Hidden = aStyleExt.Hidden;
            CustomBuiltin = aStyleExt.CustomBuitIn;

            Category = (TStyleCategory)aStyleExt.Category;
            if (aStyleExt.IsBuiltIn) BuiltinId = aStyleExt.BuiltInId; else BuiltinId = -1;
            XFProps = aStyleExt.XFProps;
        }

        internal string Name
        {
            get { return FName; }
            set { ValidateName(value, true); FName = value; }
        }

        internal void AddFutureStorage(TFutureStorageRecord R)
        {
            TFutureStorage.Add(ref FutureStorage, R);
        }


        internal static TStyleRecord CreateBuiltIn(int StyleId, int XF, int LevelId)
        {
            return new TStyleRecord(TBuiltInStyles.GetName((byte)StyleId, LevelId), XF, StyleId, LevelId, false, false, TBuiltInStyles.GetCategory(StyleId));
        }

        internal static TStyleRecord CreateUserDefined(string Name, int XF, bool Hidden)
        {
            return new TStyleRecord(Name, XF, -1, -1, Hidden, false, TStyleCategory.Custom);
        }

        internal static TStyleRecord CreateFromXlsx(string aName, int aXF, int aBuiltInId, int aLevel, bool aHidden, bool aCustombuiltIn)
        {
            //We will convert internal names to english. So "Titulo" will become "Title". This is what Excel does too.
            //If we stored the real name, we should be cautious when comparing, to use the id in those cases.
            TStyleCategory Category = TBuiltInStyles.GetCategory(aBuiltInId);//note that alevel is 0-based.
            return new TStyleRecord(NormalizedName(aBuiltInId, aName, aLevel), aXF, aBuiltInId, aLevel, aHidden, aCustombuiltIn, Category);
        }


        internal static string ValidateName(string Name, bool ThrowEx)
        {
            if (Name == null)
            {
                if (!ThrowEx) return null;
                FlxMessages.ThrowException(FlxErr.ErrInvalidEmptyName);
            }

            string tName = Name.Trim(); //a style like "  test" can be defined... even when the UI won't show it.
            if (tName.Length == 0 && !IsInternal(Name)) //if (Name = filterdatabase, Trim will delete it).
            {
                if (!ThrowEx) return null;
                FlxMessages.ThrowException(FlxErr.ErrInvalidEmptyName);
            }
            if (Name.Length > 255)
            {
                if (!ThrowEx) return null;
                FlxMessages.ThrowException(FlxErr.ErrNameTooLong, 255);
            }
            return Name;
        }

        private static bool IsInternal(string Name)
        {
            return Name.Length == 1 && Name[0] < 32; 
        }

        protected override TBaseRecord DoCopyTo(TSheetInfo SheetInfo)
        {
            return (TStyleRecord)MemberwiseClone();
        }

        internal override int GetId
        {
            get { return (int)xlr.STYLE; }
        }

        internal bool IsBiff8BuiltIn
        {
            get
            {
                return BuiltinId >= 0 && BuiltinId <= 0x09;
            }
        }

        //For comparing, so 2 internal names are the same, even if localized or whatever.
        internal static string NormalizedName(int aBuiltinId, string aName, int aLevel)
        {
            return aBuiltinId >= 0 ? TBuiltInStyles.GetName((byte)aBuiltinId, aLevel) : aName;
        }

        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
        {
            if (IgnoreIt) return;
            Workbook.WriteHeader((UInt16)xlr.STYLE, (UInt16)StyleSize());

            int Biff8XF = SaveData.StylesSavedAtPos[XF];
            Biff8Utils.CheckXF(Biff8XF);
            int BuiltInBit = IsBiff8BuiltIn ? 0x8000 : 0;
            Workbook.Write16((UInt16)(Biff8XF | BuiltInBit));

            if (IsBiff8BuiltIn)
            {
                unchecked
                {
                    UInt16 BuiltinData = IsBiff8BuiltIn ? (UInt16)BuiltinId : (UInt16)0xFFFF;

                    if (BuiltinId == 1 || BuiltinId == 2) BuiltinData |= (UInt16)(iLevel << 8); else BuiltinData |= 0xFF00;
                    Workbook.Write16(BuiltinData);
                }
            }
            else
            {
                TExcelString es = new TExcelString(TStrLenLength.is16bits, FName, null, false);
                int NameSize = es.TotalSize();

                byte[] NameData = new byte[NameSize];
                es.CopyToPtr(NameData, 0, true);

                Workbook.Write(NameData, NameData.Length);
            }

            if (NeedsStyleEx) WriteStyleEx(Workbook);
        }

        internal bool NeedsStyleEx
        {
            get
            {
                return BuiltinId > 0x09 || XFProps != null;
            }
        }

        private void WriteStyleEx(IDataStream Workbook)
        {
            Workbook.WriteHeader((UInt16)xlr.STYLEEX, (UInt16)SizeExNoHeaders);
            Workbook.Write16(0x0892);
            Workbook.Write(new byte[10], 10);

            Workbook.Write16((UInt16)(BitOps.GetBool(IsBuiltInStyle(), Hidden, CustomBuiltin) | ((int)Category << 8)));
            unchecked
            {
                Workbook.Write16((UInt16)(BuiltinId | (iLevel << 8)));
            }

            byte[] NameEx = Encoding.Unicode.GetBytes(FName);
            Workbook.Write16((UInt16)(NameEx.Length / 2));
            Workbook.Write(NameEx, NameEx.Length);

            if (XFProps != null) Workbook.Write(XFProps, XFProps.Length);
            else Workbook.Write(new byte[4], 0, 4); //no props.
        }


        internal int SizeEx
        {
            get
            {
                return SizeExNoHeaders + XlsConsts.SizeOfTRecordHeader;
            }
        }

        internal int SizeExNoHeaders
        {
            get
            {
                int XFPropsSize = XFProps == null ? 4 : XFProps.Length;
                int NameExSize = Encoding.Unicode.GetByteCount(FName);

                return 18 + NameExSize + XFPropsSize;
            }
        }

        private int StyleSize()
        {
            if (IsBiff8BuiltIn) return 4;
            TExcelString es = new TExcelString(TStrLenLength.is16bits, FName, null, false);
            return 2 + es.TotalSize();
        }

        internal override int TotalSize()
        {
            if (IgnoreIt) return 0;
            return StyleSize() + XlsConsts.SizeOfTRecordHeader + (NeedsStyleEx ? SizeEx : 0);
        }

        internal override int TotalSizeNoHeaders()
        {
            if (IgnoreIt) return 0;
            return StyleSize() + (NeedsStyleEx ? SizeExNoHeaders : 0);
        }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.Styles.Add(this);
        }

        #region IComparable Members

        public int CompareTo(object obj)
        {
            TStyleRecord o2 = obj as TStyleRecord;
            if (o2 == null) return 1;
            return String.Compare(Name, o2.Name, StringComparison.InvariantCultureIgnoreCase); //built in and not built in styles are similar here.
        }

        #endregion

        internal bool IsBuiltInStyle()
        {
            return BuiltinId >= 0;
        }
    }

    internal class TStyleRecordList : TBaseRecordList<TStyleRecord>
    {
        TStyleRecord Normal;

        internal override void OnAdd(TStyleRecord r, int index)
        {
            if (r.XF == 0) Normal = r;
        }

        internal override void OnDelete(TStyleRecord r, int index)
        {
            if (r.XF == 0) Normal = null;
        }

        internal void AddStyleEx(TStyleExRecord StyleExRecord)
        {
            int Index = -1;
            if (Find(TStyleRecord.CreateUserDefined(StyleExRecord.Name, 0, false), ref Index))
            {
                this[Index].AddStyleEx(StyleExRecord);
            }
        }

        internal string GetStyleName(int Index)
        {
            if (Index < 0 || Index >= Count) FlxMessages.ThrowException(FlxErr.ErrStyleDoesntExists, Index);
            return this[Index].Name;
        }

        internal string GetStyleNameFromXF(int XF)
        {
            if (XF == 0) //Having a full cache for all possible parents is overkill, and will in most cases slow down, not increase the speed. But a cache for normal is justified.
            {
                if (Normal != null) return Normal.Name;
                return null;
            }
            for (int i = Count - 1; i >= 0; i--)
            {
                TStyleRecord st = this[i];
                if (st.XF == XF) return st.Name;
            }

            return null;
        }


        internal int GetStyle(int Index)
        {
            if (Index < 0 || Index >= Count) FlxMessages.ThrowException(FlxErr.ErrStyleDoesntExists, Index);
            return this[Index].XF;
        }

        internal int GetStyle(string Name)
        {
            Name = TStyleRecord.ValidateName(Name, false);
            if (Name == null) return -1;

            int Index = -1;
            if (Find(TStyleRecord.CreateUserDefined(Name, 0, false), ref Index))
            {
                return this[Index].XF;
            }

            return -1;
        }

        internal void RenameStyle(string OldName, string NewName)
        {
            OldName = TStyleRecord.ValidateName(OldName, true);
            NewName = TStyleRecord.ValidateName(NewName, true);

            int index = -1;
            if (Find(TStyleRecord.CreateUserDefined(NewName, 0, false), ref index)) FlxMessages.ThrowException(FlxErr.ErrStyleAlreadyExists, NewName);
            if (!Find(TStyleRecord.CreateUserDefined(OldName, 0, false), ref index)) FlxMessages.ThrowException(FlxErr.ErrStyleDoesntExists, OldName);

            int level;
            if (TBuiltInStyles.GetIdAndLevel(NewName, out level) >= 0) FlxMessages.ThrowException(FlxErr.ErrCantRenameBuiltInStyle, NewName);

            TStyleRecord R = this[index];
            if (R.IsBuiltInStyle()) FlxMessages.ThrowException(FlxErr.ErrCantRenameBuiltInStyle, OldName);

            R.Name = NewName;
        }

        internal bool HasStyle(string Name)
        {
            int index = -1;
            return Find(TStyleRecord.CreateUserDefined(Name, 0, false), ref index);
        }

        internal void SetStyle(string Name, int XF)
        {
            Name = TStyleRecord.ValidateName(Name, true);

            int index = -1;
            if (Find(TStyleRecord.CreateUserDefined(Name, 0, false), ref index))
            {
                TStyleRecord R = (TStyleRecord)this[index];
                R.XF = XF;
                R.XFProps = null;
                R.IgnoreIt = false;
            }
            else
            {
                int Level;
                int BuiltInId = TBuiltInStyles.GetIdAndLevel(Name, out Level);
                if (TBuiltInStyles.IsBuiltInGlobal(BuiltInId))
                {
                    Insert(index, TStyleRecord.CreateBuiltIn(BuiltInId, XF, Level));
                }
                else
                {
                    Insert(index, TStyleRecord.CreateUserDefined(Name, XF, false));
                }
            }
        }

        internal void DeleteStyle(string Name, TXFRecordList CellXFList)
        {
            Name = TStyleRecord.ValidateName(Name, true);

            int index = -1;
            if (Find(TStyleRecord.CreateUserDefined(Name, 0, false), ref index))
            {
                TStyleRecord R = (TStyleRecord)this[index];
                if (R.IsBuiltInStyle()) FlxMessages.ThrowException(FlxErr.ErrCantDeleteBuiltInStyle, Name);
                Delete(index);

                CellXFList.UpdateChangedStyleInCellXF(R.XF, null, true);
            }
            else
            {
                FlxMessages.ThrowException(FlxErr.ErrStyleDoesntExists, Name);
            }

        }

        internal void AddBiff8Outlines()
        {
            //order is row1 / col1 / row2 / col2...
            for (int i = 1; i < FlxConsts.DefaultFormatIdBiff8; i++)
            {
                string StyleName = TBuiltInStyles.GetName((byte)(1 + (i - 1) % 2), (i - 1)/ 2);
                int StyleIndex = GetStyle(StyleName);
                if (StyleIndex < 0)
                {
                    TStyleRecord st = TStyleRecord.CreateBuiltIn(1 + (i - 1) % 2, i, (i - 1)/ 2);
                    st.IgnoreIt = true;
                    Add(st);
                }
            }
        }

    }

    internal class TStyleExRecord : TxBaseRecord
    {
        internal TStyleExRecord(int aId, byte[] aData) : base(aId, aData) { }

        internal bool IsBuiltIn
        {
            get
            {
                return (Data[12] & 0x1) != 0;
            }
        }

        internal bool Hidden
        {
            get
            {
                return (Data[12] & 0x2) != 0;
            }
        }

        internal bool CustomBuitIn
        {
            get
            {
                return (Data[12] & 0x4) != 0;
            }
        }

        internal byte Category
        {
            get
            {
                return (Data[13]);
            }
        }

        internal byte BuiltInId
        {
            get
            {
                return (Data[14]);
            }
        }

        internal byte[] XFProps
        {
            get
            {
                int Len = GetWord(16) * 2;
                int p = 18 + Len;
                if (p < Data.Length && Data.Length - p < 7000) //if there are continue records, discard everything.
                {
                    byte[] Result = new byte[Data.Length - p];
                    Array.Copy(Data, p, Result, 0, Result.Length);
                    return Result;
                }

                return null;
            }
        }

        internal string Name
        {
            get
            {
                if (IsBuiltIn && TBuiltInStyles.IsBuiltInGlobal(Data[14])) //even when the name is stored, we will match by our "real" name, to make sure they are the same. 
                {
                    return TBuiltInStyles.GetName(Data[14], Data[15]);
                }
                int Len = GetWord(16);
                return Encoding.Unicode.GetString(Data, 18, Len * 2);
            }
            set
            {
                int Len = GetWord(16) * 2;
                byte[] NewName = Encoding.Unicode.GetBytes(value);

                byte[] NewData = new byte[Data.Length - Len + NewName.Length];
                Array.Copy(Data, 0, NewData, 0, 14);
                BitOps.SetWord(NewData, 16, NewName.Length / 2);
                Array.Copy(NewName, 0, NewData, 18, NewName.Length);

                Debug.Assert(Data.Length - 18 - Len == NewData.Length - 18 - NewName.Length);
                Array.Copy(Data, 18 + Len, NewData, 18 + NewName.Length, Data.Length - 18 - Len);

                Data = NewData;

            }
        }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.Styles.AddStyleEx(this);
        }
    }
    #endregion

    #region Format

    internal class TFormatRecord : TBaseRecord
    {
        private int FFormatId;
        private string FFormatDef;

        internal TFormatRecord(int aId, byte[] aData)
        {
            FFormatId = BitOps.GetWord(aData, 0);
            long SSize = 0;
            StrOps.GetSimpleString(true, aData, 2, false, 0, ref FFormatDef, ref SSize);
        }

        /// <summary>
        /// CreateFromData
        /// </summary>
        internal TFormatRecord(string Fmt, int NewID)
        {
            if (Fmt == null || Fmt.Length > 255) XlsMessages.ThrowException(XlsErr.ErrInvalidFormatStringLength, Fmt);
            if (NewID < 0) XlsMessages.ThrowException(XlsErr.ErrInvalidFormatId);
            FFormatId = NewID;
            FFormatDef = Fmt;
        }

        internal int FormatId
        {
            get { return FFormatId; }
        }

        internal string FormatDef
        {
            get { return FFormatDef; }
        }

        #region Base Record
        internal override void SaveToPxl(TPxlStream PxlStream, int Row, TPxlSaveData SaveData)
        {
            if (FFormatId < 0 || FFormatId > XlsConsts.MaxNumFormatId) XlsMessages.ThrowException(XlsErr.ErrInvalidFormatId);
            base.SaveToPxl(PxlStream, Row, SaveData);
            PxlStream.WriteByte((byte)pxl.xFORMAT);
            PxlStream.WriteString8(FormatDef);
        }

        protected override TBaseRecord DoCopyTo(TSheetInfo SheetInfo)
        {
            return new TFormatRecord(FFormatDef, FFormatId);
        }

        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
        {
            if (FFormatId < 0 || FFormatId > XlsConsts.MaxNumFormatId) XlsMessages.ThrowException(XlsErr.ErrInvalidFormatId);
            TExcelString Xs = new TExcelString(TStrLenLength.is16bits, FormatDef, null, false);
            byte[] Str = new byte[Xs.TotalSize()];
            Xs.CopyToPtr(Str, 0);

            Workbook.WriteHeader((UInt16)xlr.xFORMAT, (UInt16)(Str.Length + 2));
            Workbook.Write16((UInt16)FFormatId);
            Workbook.Write(Str, 0, Str.Length);
        }

        internal override int TotalSizeNoHeaders()
        {
            TExcelString Xs = new TExcelString(TStrLenLength.is16bits, FormatDef, null, false);
            return 2 + Xs.TotalSize();
        }

        internal override int TotalSize()
        {
            return XlsConsts.SizeOfTRecordHeader + TotalSizeNoHeaders();
        }

        internal override int GetId
        {
            get { return (int)xlr.xFORMAT; }
        }
        #endregion

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.Formats.SetBucket(this);
        }

        #region Compare
        public override bool Equals(object obj)
        {
            TFormatRecord fmt = obj as TFormatRecord;
            if (fmt == null) return false;

            return fmt.FFormatDef == FFormatDef && fmt.FFormatId == FFormatId;
        }

        public override int GetHashCode()
        {
            return HashCoder.GetHash(FFormatId, FFormatDef.GetHashCode());
        }
        #endregion
    }

    internal class TFormatRecordList : ISaveBiff8 //Items are TFormatRecord
    {
        private int StartFmt;
        private StringIntHashtable RecordDictionary;
        private StringIntHashtable FBuiltInDictionary;

#if (FRAMEWORK20)
        private List<TFormatRecord> Records;
#else
		private ArrayList Records;
#endif

        /// <summary>
        /// Use Create() to create an instance. This way we avoid calling virtual methods on a constructor.
        /// </summary>
        private TFormatRecordList()
        {
#if (FRAMEWORK20)
            Records = new List<TFormatRecord>(300);
#else
            Records = new ArrayList(300);
#endif
            RecordDictionary = new StringIntHashtable();
        }

        internal static TFormatRecordList Create()
        {
            TFormatRecordList Result = new TFormatRecordList();
            return Result;
        }

        public void Clear()
        {
            Records.Clear();
            RecordDictionary.Clear();
            StartFmt = 0;
        }


        #region BuiltIn

        private static readonly string[] XlsBuiltInFormatsUs =  //STATIC*
            {
                "", "0", "0.00","#,##0","#,##0.00",                                 //0..4
                "", "", "", "",                                                     //5..8  Contained in file
                "0%","0.00%","0.00E+00","# ?/?","# ??/??",                            //9..13
                "mm/dd/YYYY","DD-MMM-YY","DD-MMM","MMM-YY",                         //14..17
                "h:mm AM/PM","h:mm:ss AM/PM","hh:mm","hh:mm:ss",                    //18..21
                "mm/dd/YYYY hh:mm",                                                 //22
                "","","","","","","","","","","","","","",                          //23..36 Reserved
                "#,##0 _$;-#,##0 _$","#,##0 _$;[Red]-#,##0 _$",             //37..38
                "#,##0.00 _$;-#,##0.00 _$","#,##0.00 _$;[Red]-#,##0.00 _$", //39..40
                "","","","",                                                        //41..44 contained in file
                "mm:ss","[h]:mm:ss","mm:ss,0","##0.0E+0","@"                //45..49
            };

        private static string XlsBuiltInFormats(int i)
        {
            if (i == 14) //Regional date.
            {
                return TFlxNumberFormat.RegionalDateString;
            }
            if (i == 22) //Regional date time
            {
                return TFlxNumberFormat.RegionalDateTimeString;
            }

            return XlsBuiltInFormatsUs[i];
        }


        private static readonly string[] PxlBuiltInFormats =  //STATIC*  (Why oh why does it have to be different from XlsBuiltInFormats??)
            {
                "", "0.00", "0.00_);(0.00)", "#,##0.00","#,##0.00_);(#,##0.00)", "#,##0", "#,##0_);(#,##0)",                                 
                "0", "0);(0)", "$#,##0.00","$#,##0.00_);($#,##0.00)",
                "$#,##0", "$#,##0_);($#,##0)",
                "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)",
                "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)",
                "_($* #,##0_);_($* (#,##0);_($* \"-\"_);_(@_)",
                "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)",
                
                "m/d", "m/d/yy", "mm/dd/yy", "d-mmm", "d-mmm-yy", "dd-mmm-yy", 
                "mmm-yy", "mmmm-yy", @"mmmm d\, yyyy", 
                "m/d/yy h:mm AM/PM", "m/d/yy h:mm", "h:mm AM/PM", "h:mm",
                "h:mm:ss AM/PM", "h:mm:ss",

                "0.00%","0%","?/?","# ??/??","# ???/???","0.00E+00", 
                "0.E+00","@","mm:ss.0","[hh]:mm:ss",   

                "","","","","","","","",

                "yyyy/m/d", "yyyy/m/d h:mm AM/PM", "yy-mm-dd", @"d\.m", 
                @"d\.m\.yy" , 
                "","","",

                @"d\. mmm", @"d\. mmm yy",
                "mmmm yy", @"d\. mmmm yyyy", 
                @"d\.m\.yy h:mm AM/PM" ,
                @"d\.m\.yy h:mm" , "dd mm yy", "d/m", "d/m/yy", "dd/mm/yy",
                "d mmmm yyyy", "d/m/yy h:mm AM/PM", "d/m/yy h:mm",
                @"d \d\e mmmm \d\e yyyy", 
                "","","","","","",
                "","","","","","","","","","",
                "","","","","","","","","","",
                "","","","","","","","","","",
                "","","","","","","","","","",
                "","","","","","","","","","",
                "","","","","","","","","","",
                "","","","","","","","","","",
                "","","","","","","",
                "mmmm", "m/d/yyyy", "m/d/yyyy h:mm AM/PM"
            };

        #endregion

        private StringIntHashtable BuiltInDictionary
        {
            get
            {
                if (FBuiltInDictionary == null)
                {
                    FBuiltInDictionary = new StringIntHashtable();

                    for (int i = XlsBuiltInFormatsUs.Length - 1; i >= 0; i--) //keep it reversed so "" goes to format 0.
                    {
                        FBuiltInDictionary[XlsBuiltInFormats(i)] = i;
                    }
                }

                return FBuiltInDictionary;
            }
        }

        internal static string GetInternalFormat(int index)
        {
            return XlsBuiltInFormats(index);
        }

        internal string Format(int FormatId)
        {
            if (FormatId < Records.Count && FormatId >= 0 && Records[FormatId] != null) return ((TFormatRecord)Records[FormatId]).FormatDef;
            if ((FormatId >= 0) && (FormatId < XlsBuiltInFormatsUs.Length)) return XlsBuiltInFormats(FormatId);
            return String.Empty;
        }

        internal string FormatFromPxl(int FormatId)
        {
            if (FormatId < Records.Count && FormatId >= 0 && Records[FormatId] != null) return ((TFormatRecord)Records[FormatId]).FormatDef;
            if ((FormatId >= 0) && (FormatId < PxlBuiltInFormats.Length)) return PxlBuiltInFormats[FormatId];
            return String.Empty;
        }

        internal int GetPxlIndex(string Format)
        {
            //See if an internal format is ok first.
            for (int i = 0; i < PxlBuiltInFormats.Length; i++)
            {
                if (String.Equals(PxlBuiltInFormats[i], Format, StringComparison.InvariantCulture)) return i;
            }

            //search on the list.
            int k;
            if (RecordDictionary.TryGetValue(Format, out k))
            {
                int z = 0;
                for (int i = StartFmt; i < k; i++)
                {
                    if (Records[i] != null) z++;
                }
                return z + 233; //custom formats are + 233
            }

            return 0; //should not come here, but if it does, just return general format.
        }

        internal void SetBucket(TFormatRecord Fmt)
        {
            if (Records.Count == 0) StartFmt = Fmt.FormatId;
            else
            {
                if (Fmt.FormatId < StartFmt) StartFmt = Fmt.FormatId;
            }

            for (int i = Records.Count; i <= Fmt.FormatId; i++)
            {
                Records.Add(null);
            }

            TFormatRecord OldRecord = (TFormatRecord)Records[Fmt.FormatId];
            if (OldRecord != null)
            {
                int Id = (int)RecordDictionary[OldRecord.FormatDef];
                if (Id == Fmt.FormatId) RecordDictionary.Remove(OldRecord.FormatDef);
            }
            Records[Fmt.FormatId] = Fmt;
            RecordDictionary[Fmt.FormatDef] = Fmt.FormatId;
        }

        internal int AddFormat(string Fmt)
        {
            int k;
            if (RecordDictionary.TryGetValue(Fmt, out k)) return k;
            if (BuiltInDictionary.TryGetValue(Fmt, out k)) return k;

            int Result = Math.Max(XlsConsts.MinNumFormatId, Records.Count); //0xA4 is the first used defined format.
            SetBucket(new TFormatRecord(Fmt, Result));
            return Result;
        }

        internal bool IsEmpty
        {
            get
            {
                return Records.Count == 0;
            }
        }

        internal int Count
        {
            get
            {
                return Records.Count;
            }
        }

        internal TFormatRecord this[int index]
        {
            get
            {
                return (TFormatRecord)Records[index];
            }
        }

        internal int TotalSize
        {
            get
            {
                int Result = 0;
                int aCount = Records.Count;
                for (int i = StartFmt; i < aCount; i++) //RecordDictionary could be better here because it has only non null entries, but then, there could be 2 entries with the same format string, and RecordDictionary would point only to one of them.
                {
                    if (Records[i] != null) Result += ((TFormatRecord)Records[i]).TotalSize();
                }

                return Result;
            }
        }

        #region ISaveBiff8 Members

        public void SaveToStream(IDataStream DataStream, TSaveData SaveData, int Row)
        {
            int aCount = Records.Count;
            for (int i = StartFmt; i < aCount; i++)
            {
                if (Records[i] != null) ((TFormatRecord)Records[i]).SaveToStream(DataStream, SaveData, Row);
            }
        }

        public void SaveToPxl(TPxlStream PxlStream, int Row, TPxlSaveData SaveData)
        {
            int aCount = Records.Count;
            for (int i = StartFmt; i < aCount; i++)
            {
                if (Records[i] != null) ((TFormatRecord)Records[i]).SaveToPxl(PxlStream, Row, SaveData);
            }
        }

        #endregion

    }
    #endregion

    #region Borders

    internal class TBorderList
    {
        private UInt32List RefCount;
#if (FRAMEWORK20)
        private List<TFlxBorders> FList;
        private Dictionary<TFlxBorders, int> FSearchList;


        internal TBorderList()
        {
            FList = new List<TFlxBorders>();
            FSearchList = new Dictionary<TFlxBorders, int>();
            RefCount = new UInt32List();
        }
#else
		private ArrayList FList;
		private Hashtable FSearchList;


		internal TBorderList() 
		{
			FList = new ArrayList();
			FSearchList = new Hashtable();
            RefCount = new UInt32List();
		}
#endif


        internal int AddOrReplaceBorder(int index, TFlxBorders Borders)
        {
            if (RefCount[index] <= 1)
            {
                FList[index] = Borders;
                return index;
            }

            int idx = Add(Borders);
            if (idx != index) RefCount[index]--;
            return idx;
        }


        #region Generics
        internal TFlxBorders this[int index]
        {
            get { return (TFlxBorders)FList[index]; }
        }

        public int Count
        {
            get { return FList.Count; }
        }

        public void Clear()
        {
            FList.Clear();
            FSearchList.Clear();
            RefCount.Clear();
        }
        #endregion

        /// <summary>
        /// It won't check for duplicates or clone the borders.
        /// </summary>
        /// <param name="aBorder"></param>
        public void AddForced(TFlxBorders aBorder)
        {
            FList.Add(aBorder);
            FSearchList[aBorder] = FList.Count - 1;
            RefCount.Add(1);
        }

        public int Add(TFlxBorders aBorder)
        {
#if (FRAMEWORK20)
            int index;
            if (FSearchList.TryGetValue(aBorder, out index))
            {
                RefCount[index]++;
                return index;
            }
#else
			object o = FSearchList[aBorder];
			if (o != null) 
			{
				RefCount[(int)o]++;
				return (int)o;
			}
#endif
            FList.Add((TFlxBorders)aBorder.Clone());
            FSearchList[aBorder] = FList.Count - 1;
            RefCount.Add(1);
            return FList.Count - 1;
        }

    }
    #endregion

    #region Patterns

    internal class TPatternList
    {
        private UInt32List RefCount;
#if (FRAMEWORK20)
        private List<TFlxFillPattern> FList;
        private Dictionary<TFlxFillPattern, int> FSearchList;


        internal TPatternList()
        {
            FList = new List<TFlxFillPattern>();
            FSearchList = new Dictionary<TFlxFillPattern, int>();
            RefCount = new UInt32List();
        }

#else
		private ArrayList FList;
		private Hashtable FSearchList;


		internal TPatternList() 
		{
			FList = new ArrayList();
			FSearchList = new Hashtable();
            RefCount = new UInt32List();         
		}
#endif

		internal void EnsureRequiredFills()
		{
			if (Count <= 0)
			{
				TFlxFillPattern pat0 = new TFlxFillPattern();
				pat0.Pattern = TFlxPatternStyle.None;
				AddForced(pat0);
			}

			if (Count <= 1)
			{
				TFlxFillPattern pat1 = new TFlxFillPattern();
				pat1.Pattern = TFlxPatternStyle.Gray16;
				AddForced(pat1);
			}
		}

        internal int AddOrReplacePattern(int FillPattern, TFlxFillPattern Pattern)
        {
            if (RefCount[FillPattern] <= 1 && FillPattern > 1) //patterns 0 and 1 are reserved.
            {
                if ((int)FSearchList[FList[FillPattern]] == FillPattern) FSearchList.Remove(Pattern);
                FList[FillPattern] = Pattern;
                FSearchList[Pattern] = FillPattern;
                RefCount[FillPattern] = 1;
                return FillPattern;
            }

            int idx = Add(Pattern); //this will fix the cache.
            if (idx != FillPattern) RefCount[FillPattern]--;

            return idx;
        }

        #region Generics
        internal TFlxFillPattern this[int index]
        {
            get { return ((TFlxFillPattern)FList[index]); }
        }

        public void Clear()
        {
            FList.Clear();
            FSearchList.Clear();
            RefCount.Clear();
        }

        public int Count
        {
            get
            {
                return FList.Count;
            }
        }
        #endregion

        /// <summary>
        /// It won't check for duplicates or clone the pattern.
        /// </summary>
        /// <param name="aPattern"></param>
        public void AddForced(TFlxFillPattern aPattern)
        {
            FList.Add(aPattern);
            FSearchList[aPattern] = FList.Count - 1;
            RefCount.Add(1);
        }

        public int Add(TFlxFillPattern aPattern)
        {
            if (aPattern.Pattern == TFlxPatternStyle.Gradient && aPattern.Gradient == null) XlsMessages.ThrowException(XlsErr.ErrNullGradient);

            int index;
            if (FSearchList.TryGetValue(aPattern, out index))
            {
                RefCount[index]++;
                return index;
            }

            FList.Add((TFlxFillPattern)aPattern.Clone());
            FSearchList[aPattern] = FList.Count - 1;
            RefCount.Add(1);
            return FList.Count - 1;
        }

    }
    #endregion

    #region Theme

    internal class TThemeRecord : TBaseRecord
    {
        const int DataOfs = 16;
        const int ContinueOfs = 12;
        const int MaxRecordDataSize = XlsConsts.MaxRecordDataSize;

        internal TFutureStorage FutureStorage;
        internal TFutureStorage ElementsFutureStorage;
        internal TFutureStorage ColorFutureStorage;
        internal TFutureStorage FontFutureStorage;
        internal TFutureStorage MajorFontFutureStorage;
        internal TFutureStorage MinorFontFutureStorage;

        internal byte[] Data;
        internal TTheme Theme;
        private TContinueRecord Continue; //used to temporary load only.

        internal TThemeRecord()
        {
            Theme = new TTheme();
        }

        internal TThemeRecord(byte[] aData)
        {
            Theme = null;
            Data = aData;
        }

        private TThemeRecord(TTheme aTheme, byte[] aData, 
            TFutureStorage aFutureStorage,
            TFutureStorage aElementsFutureStorage,
            TFutureStorage aColorFutureStorage,
            TFutureStorage aFontFutureStorage,
            TFutureStorage aMajorFontFutureStorage,
            TFutureStorage aMinorFontFutureStorage
            )
        {
            if (aTheme == null) Theme = null; else Theme = aTheme.Clone();
            if (aData == null) Data = null; else Data = (byte[])aData.Clone();
            FutureStorage = TFutureStorage.Clone(aFutureStorage);
            ElementsFutureStorage = TFutureStorage.Clone(aElementsFutureStorage);
            ColorFutureStorage = TFutureStorage.Clone(aColorFutureStorage);
            FontFutureStorage = TFutureStorage.Clone(aFontFutureStorage);
            MajorFontFutureStorage = TFutureStorage.Clone(aMajorFontFutureStorage);
            MinorFontFutureStorage = TFutureStorage.Clone(aMinorFontFutureStorage);
        }

        internal override void AddContinue(TContinueRecord aContinue)
        {
            Continue = aContinue;
        }

        internal void LoadFromBiff8()
        {
            Theme = new TTheme();
            unchecked
            {
                if (Data != null && Data.Length >= 16) Theme.ThemeVersion = (int)BitOps.GetCardinal(Data, 12);
            }
            JoinContinues();
            Continue = null;

#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
            if (Data != null && Data.Length > DataOfs)
            {
                using (MemoryStream DataStream = new MemoryStream(Data, DataOfs, Data.Length - DataOfs))
                {
                    using (TOpenXmlReader xml = new TOpenXmlReader(DataStream, false, null, String.Empty, TExcelFileErrorActions.All))
                    {
                        TXlsxRecordLoader Loader = new TXlsxRecordLoader(xml, null, null, null);
                        Loader.ReadTheme(this);
                    }
                }
            }
#endif
        }

        private void JoinContinues()
        {
            TContinueRecord aContinue = Continue;

            int ContinueLen = 0;
            while (aContinue != null)
            {
                if (aContinue.Data.Length < ContinueOfs) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
                ContinueLen += aContinue.Data.Length - ContinueOfs;
                aContinue = aContinue.Continue;
            }

            if (ContinueLen == 0) return;

            byte[] NewData = new byte[Data.Length + ContinueLen]; 
            Array.Copy(Data, 0, NewData, 0, Data.Length);

            aContinue = Continue;
            int NewDataPos = Data.Length;
            while (aContinue != null)
            {
                Array.Copy(aContinue.Data, ContinueOfs, NewData, NewDataPos, aContinue.Data.Length - ContinueOfs);
                NewDataPos += aContinue.Data.Length - ContinueOfs;
                aContinue = aContinue.Continue;
            }

            Data = NewData;
        }

        protected override TBaseRecord DoCopyTo(TSheetInfo SheetInfo)
        {
            TThemeRecord Result = new TThemeRecord(Theme, Data, FutureStorage, ElementsFutureStorage, ColorFutureStorage, FontFutureStorage, MajorFontFutureStorage, MinorFontFutureStorage);
            return Result;
        }

        internal int CalcContinues()
        {
            int Len = Data.Length;
            if (Len <= MaxRecordDataSize) return 0;
            Len -= MaxRecordDataSize;

            return (Len - 1) / (MaxRecordDataSize - ContinueOfs) + 1;
        }

        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
        {
            if (Data == null) return;

            UInt16 rs = (UInt16)Math.Min(Data.Length, MaxRecordDataSize);
            Workbook.WriteHeader((UInt16)xlr.THEME, rs);
            Workbook.Write(Data, rs);

            int Len = Data.Length - MaxRecordDataSize;
            int DataPos = rs;
            while (Len > 0)
            {
                rs = (UInt16)Math.Min(Len, MaxRecordDataSize - ContinueOfs);
                Workbook.WriteHeader((UInt16)xlr.CONTINUEFRT12, (UInt16)(rs + ContinueOfs));
                Workbook.Write16((UInt16)xlr.CONTINUEFRT12);
                Workbook.Write(new byte[10], 10);
                Workbook.Write(Data, DataPos, rs);

                Len -= rs;
                DataPos += rs;

            }
        }

        internal override int TotalSize()
        {
            if (Data != null) return TotalSizeNoHeaders() + XlsConsts.SizeOfTRecordHeader +
                (XlsConsts.SizeOfTRecordHeader + ContinueOfs) * CalcContinues();
            return 0;
        }

        internal override int TotalSizeNoHeaders()
        {
            if (Data != null) return Data.Length;
            return 0;
        }

        internal void CalcData(TWorkbook Workbook)
        {
            if (Data != null) return;
            if (Theme.IsStandard) return; //if theme is 124226 we won't save it either.

#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
            using (MemoryStream ms = new MemoryStream())
            {
                using (TOpenXmlWriter xlsx = new TOpenXmlWriter(ms, true))
                {
                    TXlsxDrawingWriter rw = new TXlsxDrawingWriter(xlsx, null, Workbook);
                    rw.WriteThemeManager();
                    rw.WriteTheme(true);
                }

                if (FlxUtils.IsMonoRunning())
                {
                    //Workarounds mono bug: https://bugzilla.novell.com/show_bug.cgi?id=591866
                    byte[] aData = ms.ToArray();
                    
                    Data = new byte[aData.Length + DataOfs];
                    BitOps.SetWord(Data, 0, (UInt16)xlr.THEME);
                    //Custom theme so "customtheme" is 0
                    Array.Copy(aData, 0, Data, DataOfs, aData.Length);
                    ms.Read(Data, DataOfs, (int)ms.Length);
                    return;
                }

                Data = new byte[ms.Length + DataOfs];
                BitOps.SetWord(Data, 0, (UInt16)xlr.THEME);
                //Custom theme so "customtheme" is 0

                ms.Position = 0;
                ms.Read(Data, DataOfs, (int)ms.Length);
            }
#endif
        }

        internal override int GetId
        {
            get { return (int)xlr.THEME; }
        }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.ThemeRecord = this;
        }

#if (FRAMEWORK30)
        internal void AddFutureStorage(TFutureStorageRecord R)
        {
            TFutureStorage.Add(ref FutureStorage, R);
        }

        internal void AddElementsFutureStorage(TFutureStorageRecord R)
        {
            TFutureStorage.Add(ref ElementsFutureStorage, R);
        }

        internal void AddColorFutureStorage(TFutureStorageRecord R)
        {
            TFutureStorage.Add(ref ColorFutureStorage, R);
        }

        internal void AddFontFutureStorage(TFutureStorageRecord R)
        {
            TFutureStorage.Add(ref FontFutureStorage, R);
        }

        internal void AddMajorMinorFontFutureStorage(TThemeFont ThemeFont, TFutureStorageRecord R)
        {
            if (ThemeFont == Theme.Elements.FontScheme.MajorFont)
            {
                TFutureStorage.Add(ref MajorFontFutureStorage, R);
            }
            else
                if (ThemeFont == Theme.Elements.FontScheme.MinorFont)
                {
                    TFutureStorage.Add(ref MinorFontFutureStorage, R);
                }
                else
                    XlsMessages.ThrowException(XlsErr.ErrInternal);
        }
#endif

        internal void InvalidateData()
        {
            //in frameworks that don't support it, we won't remove data.
#if (FRAMEWORK30)
            Theme.ThemeVersion = 0;
            Data = null;
#endif
        }
    }

    #endregion

    #region DXF
    internal class TDXFRecordList : TMiscRecordList
    {
        internal TFutureStorage Xlsx; //Should dissapear after DXF are supported.

        public override void Clear()
        {
            base.Clear();
            Xlsx = null;
        }
    }
    /// <summary>
    /// This should be updated once we support dxf.
    /// </summary>
    internal class TDXFRecord : TxBaseRecord
    {
        internal TDXFRecord(int aId, byte[] aData) : base(aId, aData) { }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.DXF.Add(this);
        }
    }
    #endregion

    #region Table Styles
    internal class TTableStyleRecordList : TMiscRecordList
    {
        internal TFutureStorage Xlsx;

        public override void Clear()
        {
            base.Clear();
            Xlsx = null;
        }

    }

    /// <summary>
    /// a Table Style
    /// </summary>
    internal class TTableStyleRecord : TxBaseRecord
    {
        internal TTableStyleRecord(int aId, byte[] aData) : base(aId, aData) { }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.TableStyles.Add(this);
        }
    }

    internal class TTableStylesRecord : TxBaseRecord
    {
        internal TTableStylesRecord(int aId, byte[] aData) : base(aId, aData) { }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.TableStyles.Add(this);
        }
    }

    internal class TTableStyleElementRecord : TxBaseRecord
    {
        internal TTableStyleElementRecord(int aId, byte[] aData) : base(aId, aData) { }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.TableStyles.Add(this);
        }
    }

    #endregion

    #region Biff8XFList

    struct TBiff8XFPos
    {
        internal int XF2007;
        internal TXFRecordList XFList;
        internal bool IsStyle;

        internal TBiff8XFPos(int aBiff8XF, TXFRecordList aXFList, bool aIsStyle)
        {
            XF2007 = aBiff8XF;
            XFList = aXFList;
            IsStyle = aIsStyle;
        }
    }

#if (FRAMEWORK20)
    internal class TBiff8XFPosList : List<TBiff8XFPos>
    {
    }
#else
	internal class TBiff8XFPosList : ArrayList
	{
		public new TBiff8XFPos this[int index]
		{
			get
			{
				return (TBiff8XFPos)base[index];
			}
		}
	}
#endif

    /// <summary>
    /// A class to hold the XF values as we load them from an xls file, to convert them to 2007 format.
    /// </summary>
    internal class TBiff8XFMap
    {
        private TXFRecordList StyleXFList;
        private TXFRecordList CellXFList;
        private TBiff8XFPosList Biff8XF22007XF;

        internal TBiff8XFMap(TXFRecordList aStyleXFList, TXFRecordList aCellXFList)
        {
            StyleXFList = aStyleXFList;
            CellXFList = aCellXFList;
            Biff8XF22007XF = new TBiff8XFPosList();
        }

        internal int Count
        {
            get
            {
                return StyleXFList.Count + CellXFList.Count;
            }
        }


        internal void AddExt(TXFExtRecordList XFExtList, TWorkbookGlobals Globals)
        {
            foreach (TXFExtRecord xfe in XFExtList)
            {
                int xf = xfe.XF;
                if (xf < 0 || xf >= Biff8XF22007XF.Count) continue; //invalid

                Biff8XF22007XF[xf].XFList.AddExt(xfe, Biff8XF22007XF[xf].XF2007, Globals); //biff8xf 0 means xf must be ignored.
            }
        }

        internal void Add(TXFRecord XFRecord)
        {
            if (XFRecord.IsStyle)
            {
                Biff8XF22007XF.Add(new TBiff8XFPos(StyleXFList.Count, StyleXFList, true));
                StyleXFList.Add(XFRecord);
            }
            else
            {
                Biff8XF22007XF.Add(new TBiff8XFPos(CellXFList.Count, CellXFList, false));
                CellXFList.Add(XFRecord);
            }
        }

        internal int GetCellXF2007(int biff8xf)
        {
            if (biff8xf == 0) return -1; //biff8 0 means xf doesn't matter.
            if (biff8xf < 0 || biff8xf >= Biff8XF22007XF.Count) return 0;
            if (Biff8XF22007XF[biff8xf].IsStyle) return 0;  // we won't throw an exception here since some third party tools can create invalid files.
            return Biff8XF22007XF[biff8xf].XF2007;
        }

        internal int GetStyleXF2007(int biff8xf)
        {
            if (biff8xf < 0 || biff8xf >= Biff8XF22007XF.Count) return 0;
            if (!Biff8XF22007XF[biff8xf].IsStyle) return 0;
            return Biff8XF22007XF[biff8xf].XF2007;
        }

        internal static int GetPxlCellXF2007(int pxlxf, int MainXFCount)
        {
            return pxlxf + MainXFCount;
        }


    }
    #endregion
}
