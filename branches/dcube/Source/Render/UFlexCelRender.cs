using System;
using System.Text;
using System.IO;
using FlexCel.Core;
using System.Globalization;
using System.Collections.Generic;

#if (MONOTOUCH)
using System.Drawing;
using real = System.Single;
using Color = MonoTouch.UIKit.UIColor;
using Font = MonoTouch.UIKit.UIFont;
#else
	#if (WPF)
	using RectangleF = System.Windows.Rect;
	using SizeF = System.Windows.Size;
	using real = System.Double;
	using System.Windows.Media;
	#else
	using real = System.Single;
	using Colors = System.Drawing.Color;
	using DashStyles = System.Drawing.Drawing2D.DashStyle;
	
	using System.Drawing;
	using System.Drawing.Imaging;
	using System.Drawing.Drawing2D;
	using System.Drawing.Text;
	using System.Diagnostics;
	#endif
#endif

namespace FlexCel.Render
{
    /// <summary>
    /// A class for viewing/printing Excel Files.
    /// </summary>
    internal class FlexCelRender: IDrawObjectMethods
    {
        #region Privates
        protected ExcelFile FWorkbook;

        protected bool FPrintFormulas;

        protected real ColMultDisplay;
        protected real RowMultDisplay;
        protected real FmlaMult;

        protected THidePrintObjects FHidePrintObjects;

        protected IFlxGraphics Canvas = null;
        protected real Zoom100;

        protected RectangleF MarginBounds;

        protected bool FDrawGridLines;

        protected TFontCache FontCache;
		
        private bool[] VisibleFormatsCache;
		protected TXFFormatCache PageFormatCache;

        protected Color GridLinesColor;

        private int CurrentPrintAreaRegion;
        private int ColumnStartCache;
        private int[] ColumnWidthsCache;  //To speed up measuring
        private TCellMergedCache[] CellMergedCache;
        private TShapesCache[] ShapesCache;

        protected TXlsMargins FMargins;

        internal const real DispMul = FlxConsts.DispMul;
		internal const real DoubleLineSep = 1f*DispMul/100f;
		internal const real DoubleLineSepDiag = 2f*DoubleLineSep;

        private bool FReverseRightToLeftStrings = false;
        const real LineJoinAdj = 0.5f * DispMul / 200f; //In an ideal world, this should be 0. But if we set it to 0, acrobat and gdi+ will sometimes "cross" the borders when drawing an exact box. (at low resoutions)

        #endregion

        #region Constructors
        public FlexCelRender()
        {
            Zoom100 = 1;
            PageFormatCache = new TXFFormatCache();
        }
        #endregion

        #region Properties
        /// <summary>
        /// The ExcelFile that we want to display.
        /// </summary>
        public ExcelFile Workbook { get { return FWorkbook; } set { FWorkbook = value; } }

        /// <summary>
        /// This is the maximum row that should be displayed. If there are images on a sheet, it can be bigger than XlsFile.RowCount.
        /// If a row or cell has format but the format is not visible (for example a different height)it will not show.
        /// </summary>
        public int MaxVisibleRow { get { return GetMaxVisibleRow(); } }

        /// <summary>
        /// This is the maximum column that should be displayed. If there are images on a sheet or a cell spans to the right, it can be bigger than XlsFile.ColCount.
        /// If a column has format but the format is not visible (for example a different width)it will not show.
        /// </summary>
        public int MaxVisibleCol { get { return GetMaxVisibleCol(); } }

        /// <summary>
        /// Sets the canvas used by the renderer object. You always have to set it before using this class.
        /// </summary>
        /// <param name="aCanvas"></param>
        public void SetCanvas(IFlxGraphics aCanvas) { Canvas = aCanvas; }

        /// <summary>
        /// Select which kind of objects should not be printed or exported to pdf.
        /// </summary>
        public THidePrintObjects HidePrintObjects {get {return FHidePrintObjects;} set { FHidePrintObjects = value;}}

        public bool ReverseRightToLeftStrings  {get{return FReverseRightToLeftStrings ;} set{FReverseRightToLeftStrings =value;}}

        #endregion

        #region Utilities
        protected static void FillRepeatingRange(TParsedTokenList data, TRepeatingRange Result)
        {
            data.ResetPositionToLast();
            while (!data.Bof())
            {
                TBaseParsedToken Token = data.LightPop();
                if (TBaseParsedToken.CalcBaseToken(Token.GetId) == ptg.MemFunc || Token is TUnionToken) continue;

                TArea3dToken ar = Token as TArea3dToken;
				if (ar != null && !ar.IsErr())
				{
					if (ar.GetRow1(0) != 0 || ar.GetRow2(0) != FlxConsts.Max_Rows)
					{
						Result.FirstRow = ar.GetRow1(0) + 1;
						Result.LastRow = ar.GetRow2(0) + 1;
					}
					else
						if (ar.GetCol1(0) != 0 || ar.GetCol2(0) != FlxConsts.Max_Columns)
					{
						Result.FirstCol = ar.GetCol1(0) + 1;
						Result.LastCol = ar.GetCol2(0) + 1;
					}
				}
				else
				{
					return; // there should be no other token here.
				}
            }
        }

        private TRepeatingRange GetRepeatingRange(TXlsCellRange PrintArea)
        {
            if (FWorkbook == null) return new TRepeatingRange(2, 1, 2, 1);
            bool IgnoreFormulaText = FWorkbook.IgnoreFormulaText;
            try
            {
                Workbook.IgnoreFormulaText = true; //even if printing formula text...

                TXlsNamedRange nr = FWorkbook.GetNamedRange(TXlsNamedRange.GetInternalName(InternalNameRange.Print_Titles), -1, FWorkbook.ActiveSheet);
                if (nr == null) return new TRepeatingRange(2, 1, 2, 1);

                TParsedTokenList data = nr.FormulaData;
                TRepeatingRange Result = new TRepeatingRange(2, 1, 2, 1);
                FillRepeatingRange(data, Result);

                if (Result.FirstRow < PrintArea.Top) Result.FirstRow = PrintArea.Top;
                if (Result.LastRow > PrintArea.Bottom) Result.LastRow = PrintArea.Bottom;
                if (Result.FirstCol < PrintArea.Left) Result.FirstCol = PrintArea.Left;
                if (Result.LastCol > PrintArea.Right) Result.LastCol = PrintArea.Right;

                return Result;
            }
            finally
            {
                FWorkbook.IgnoreFormulaText = IgnoreFormulaText;
            }
        }

        private Color GetFgColor(TExcelColor aColor)
        {
            return aColor.ToColor(Workbook, Color.Black);
        }

		private TFlxFormat GetCellVisibleFormatDef(int row, int col, bool Merged)
		{
			return PageFormatCache.GetCellVisibleFormatDef(FWorkbook, row, col, Merged);
		}

        /// <summary>
        /// A cache to speed up formats. As they are read-only, we can reutilize older format definitions and not create new ones.
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        private TFlxFormat GetCellVisibleFormatDef(int row, int col)
        {
            return GetCellVisibleFormatDef(row, col, false);
        }

        private bool HasString(int aRow, int aCol)
        {
            object o = GetCellValue(aRow, aCol);
            switch(TExcelTypes.ObjectToCellType(o))
            {
                case TCellType.String: 
		        case TCellType.Unknown: //Unknown cells will be converted to strings.
					return true;
                case TCellType.Formula:
                    TFormula Fmla = o as TFormula;
                    if (Fmla != null)
                        if (FPrintFormulas)
                            return false;
                        else
                            return TExcelTypes.ObjectToCellType(Fmla.Result) == TCellType.String;
                    break;
            }
            return false;
        }

        private bool CellCanSpawnRight(int aRow, int aCol)
        {
            if (FmlaMult < 0.9) return false; //When in formula mode cells do not span.
            TFlxFormat Fm = GetCellVisibleFormatDef(aRow, aCol);
            if (Fm.HAlignment == THFlxAlignment.right || Fm.WrapText || Fm.Rotation == 255) return false;
            if (!HasString(aRow, aCol))return false;

            bool Merged;
            XCellMergedBounds(aRow, aCol, out Merged);
            if (Merged) return false;
            XCellMergedBounds(aRow, aCol +1, out Merged);
            if (Merged) return false;

            return true;
        }

        private bool CellCanSpawnLeft(int aRow, int aCol)
        {
            if (FmlaMult < 0.9) return false; //When in formula mode cells do not span.
            TFlxFormat Fm = GetCellVisibleFormatDef(aRow, aCol);
            if (Fm.HAlignment == THFlxAlignment.left || Fm.WrapText || Fm.Rotation == 255) return false;  //align general can be right if it is right to left.            
            if (!HasString(aRow, aCol))return false;
            //if (Fm.HAlignment == THFlxAlignment.general && TextPainter.IsLeftToRight()) return false;

            bool Merged;
            XCellMergedBounds(aRow, aCol, out Merged);
            if (Merged) return false;
            XCellMergedBounds(aRow, aCol - 1, out Merged);
            if (Merged) return false;
            
            return true;
        }

        private void CacheVisibleFormats()
        {
            VisibleFormatsCache = new bool[FWorkbook.FormatCount];

            for (int i= VisibleFormatsCache.Length-1; i>=0; i--)
            {
                TFlxFormat fmt = FWorkbook.GetFormat(i);
                VisibleFormatsCache[i] = 
                    (fmt.FillPattern.Pattern != TFlxPatternStyle.None
                    || fmt.Borders.Top.Style != TFlxBorderStyle.None
                    || fmt.Borders.Left.Style != TFlxBorderStyle.None
                    || fmt.Borders.Right.Style != TFlxBorderStyle.None
                    || fmt.Borders.Bottom.Style != TFlxBorderStyle.None
                    || fmt.Borders.DiagonalStyle != TFlxDiagonalBorder.None);
            }
        }

        private bool IsVisibleFormat(int XF)
        {
            if (XF<0 || XF>= VisibleFormatsCache.Length) return false;
            return VisibleFormatsCache[XF];
        }

        private void LastMerged(ref int LastRow, ref int LastCol)
        {
            int aCount = FWorkbook.CellMergedListCount;
            for (int i = 1; i <= aCount; i++)
            {
                TXlsCellRange cr = FWorkbook.CellMergedList(i);
                if (LastCol < cr.Right) LastCol = cr.Right;
                if (LastRow < cr.Bottom) LastRow = cr.Bottom;
            }
        }

        private int PrintableColCount()
        {
            int Result=0;

			//Once you save, Excel does not care for column formats.
			/*
            for (int c= FlxConsts.Max_Columns+1; c>0; c--)
                if (!FWorkbook.GetColHidden(c) && IsVisibleFormat(FWorkbook.GetColFormat(c))) 
                {
                    Result=c;
                    break;
                }
		    */		

            for (int r= FWorkbook.RowCount; r>0; r--)
            {
                if (FWorkbook.GetRowHidden(r)) continue;

                int ci=FWorkbook.ColCountInRow(r);
                while (ci>0)
                {
                    int XF=-1;
                    object o = FWorkbook.GetCellValueIndexed(r, ci, ref XF);
                    if (o != null || IsVisibleFormat(XF))
                    {
                        int col = FWorkbook.ColFromIndex(r, ci);
                        if (!FWorkbook.GetColHidden(col) && col>Result) 
                        {
                            Result= col;
                            break;
                        }
                    }
                    ci--;
                }
            }

            return Result;
        }

        private static void GetAnchors(List<TClientAnchor> Result, TShapeProperties sp)
        {
            if (sp.Anchor != null) 
            {
                Result.Add(sp.Anchor);
                return;
            }

            //some shapes do not have anchors. We need to look inside them.
            for (int i = 1; i <= sp.ChildrenCount; i++)
            {
                GetAnchors(Result, sp.Children(i));
            }
        }

        private int GetMaxVisibleCol()
        {
            if (FWorkbook.SheetType == TSheetType.Chart) return 1;

            Canvas.CreateSFormat();
            try
            {
                InitRender(new TXlsCellRange(1, 1, 1, FlxConsts.Max_Columns + 1), 0, 1);

                RectangleF PaintClipRect = Canvas.ClipBounds;

                CacheVisibleFormats();
			
                int cmax = PrintableColCount();
                int rdummy=FlxConsts.Max_Rows;
                LastMerged(ref rdummy, ref cmax);           
                
				CacheMergedCells(new TXlsCellRange(1, 1, FWorkbook.RowCount, cmax + 1), 0);
				if (cmax > 0)
                {
                    real ALeft = 0;
                    for (int i = 1; i < cmax; i++) ALeft += RealColWidth(i);
                    real ARight = ALeft + RealColWidth(cmax);
                    real ATop = 0;
                    real ABottom = 0;
                    int rmax = FWorkbook.RowCount;
                    if (rmax>0)
                    {
                        for (int i = 1; i < rmax; i++) ATop += FWorkbook.GetRowHeight(i, true) / RowMultDisplay;
                        ABottom = ATop + FWorkbook.GetRowHeight(rmax, true) / RowMultDisplay;

                        TSpawnedCellList SpawnedCells = new TSpawnedCellList();

                        bool GetOut = false;
                        for (int r = 1; r <= rmax; r++)
                        {
                            if (FWorkbook.IsEmptyRow(r)) continue;
                            do
                            {
                                RectangleF ARect = RectangleXY(ALeft, ATop, ARight, ABottom);

                                real OldZoom100 = Zoom100; //Always use a 100% zoom to avoid rounding errors
                                try
                                {
                                    Zoom100 = 1;
                                    DrawCell(cmax, r, ARect, PaintClipRect, SpawnedCells, false, false, TSpanDirection.Left);
                                }
                                finally
                                {
                                    Zoom100 = OldZoom100;
                                }
                                GetOut = true;
                                if (cmax <= FlxConsts.Max_Columns && SpawnedCells.ContainsKey(FlxHash.MakeHash(r, cmax)))
                                {
                                    cmax++;
                                    ALeft = ARight;
                                    ARight += RealColWidth(cmax);
                                    GetOut = false;
                                }
                            }
                            while (!GetOut);
						
                            ATop = ABottom;
                            if (r < rmax)
                                ABottom += FWorkbook.GetRowHeight(r + 1, true) / RowMultDisplay;
                        }
                    }
                }

                int aCount = FWorkbook.ObjectCount;
                for (int i = 1; i <= aCount; i++) //No need to search inside grouped images
                {
                    TShapeProperties sp = FWorkbook.GetObjectProperties(i, false);  
                    if (sp.Print && sp.Visible && sp.ObjectType != TObjectType.Comment)
                    {
                        List<TClientAnchor> Anchors = new List<TClientAnchor>();
                        GetAnchors(Anchors, sp);
                        foreach (TClientAnchor Anchor in Anchors)
                        {
                            if (Anchor != null && Anchor.Col2 > cmax) cmax = Anchor.Col2;
                        }
                    }
                }

			
                return cmax;
            }
            finally
            {
                Canvas.DestroySFormat();
            }
        }

        private bool IsPrintingRow(int r)
        {
            if (FWorkbook.GetRowHidden(r)) return false;
            if (IsVisibleFormat(FWorkbook.GetRowFormat(r))) return true;

            for (int i= FWorkbook.ColCountInRow(r); i>0; i--)
            {
                int XF=-1;
                object o = FWorkbook.GetCellValueIndexed(r, i, ref XF);
                if (o != null || IsVisibleFormat(XF)) 
                {
                    int col = FWorkbook.ColFromIndex(r, i);
                    if (!FWorkbook.GetColHidden(col)) return true;
                }
            }

            return false;
        }

        private int GetMaxVisibleRow()
        {
            if (FWorkbook.SheetType == TSheetType.Chart) return 1;
            int Result = FWorkbook.RowCount;
            CacheVisibleFormats();

            while (Result > 0 && !IsPrintingRow(Result)) Result--;

            int cdummy = FlxConsts.Max_Columns;
            LastMerged(ref Result, ref cdummy);

            int aCount = FWorkbook.ObjectCount;
            for (int i = 1; i <= aCount; i++) //No need to search inside grouped images
            {
                TShapeProperties sp = FWorkbook.GetObjectProperties(i, false);
                if (sp.Print && sp.Visible && sp.ObjectType != TObjectType.Comment)
                {
                    List<TClientAnchor> Anchors = new List<TClientAnchor>();
                    GetAnchors(Anchors, sp);
                    foreach (TClientAnchor Anchor in Anchors)
                    {
                        if (Anchor != null && Anchor.Row2 > Result) Result = Anchor.Row2;
                    }
                }
            }
            return Result;
        }


		private object GetCellValue(int aRow, int aCol)
		{
			return GetCellValue(aRow, aCol, true);
		}

		private object GetCellValue(int aRow, int aCol, bool CheckCenterAcrossSelection)
		{
			bool Merged;
			TXlsCellRange Mb = XCellMergedBounds(aRow, aCol, CheckCenterAcrossSelection, out Merged);
			if (aRow != Mb.Top || aCol != Mb.Left) return null; //some merged cells might have data, but if it is not the top-left cell, it should be displayed empty.

			if (FWorkbook.GetColHidden(aCol)) 
			{
				if (!Merged)
					return null;
			}
			return FWorkbook.GetCellValue(aRow, aCol);
		}

		private bool IsEmptyCell(int aRow, int aCol)
		{
			return IsEmptyCell(aRow, aCol, true);
		}

        private bool IsEmptyCell(int aRow, int aCol, bool CheckCenterAcrossSelection)
        {
            if (aRow <= 0 || aCol <= 0 || aRow > FlxConsts.Max_Rows || aCol > FlxConsts.Max_Columns) return false;
            object v = GetCellValue(aRow, aCol, CheckCenterAcrossSelection);

            return v == null || ((v is string || v is TRichString) && v.ToString().Length == 0);
        }

        private bool IsCenterAcrossSelection(int Row, int Col)
        {
            if (FWorkbook.GetRowHeight(Row, true)<=0 ||
                FWorkbook.GetColWidth(Col, true)<=0 ) return true;

            return GetCellVisibleFormatDef(Row, Col).HAlignment == THFlxAlignment.center_across_selection;
        }

		private TXlsCellRange XCellMergedBounds(int aRow, int aCol, out bool Merged)
		{
            return XCellMergedBounds(aRow, aCol, true, out Merged);
		}

		public TXlsCellRange XCellMergedBounds(int aRow, int aCol, bool CheckCenterAcrossSelection, out bool Merged)
		{
			TXlsCellRange Result = null;
			Merged = false;

			if (aRow==0 || aCol==0) return new TXlsCellRange(aRow, aCol, aRow, aCol);
			Merged = (CellMergedCache[CurrentPrintAreaRegion].TryGetValue(FlxHash.MakeHash(aRow, aCol), out Result));
			if (!Merged) Result = new TXlsCellRange(aRow, aCol, aRow, aCol);
            
			if (CheckCenterAcrossSelection && GetCellVisibleFormatDef(aRow, aCol).HAlignment == THFlxAlignment.center_across_selection)
			{
				//Center Across selection only grows From the current cell.
				//while (Result.Left>1 && IsEmptyCell(aRow,Result.Left-1) && GetCellVisibleFormatDef(aRow, Result.Left-1).HAlignment==THFlxAlignment.center_across_selection) 
				//    Result.Left--;
				while (Result.Right <= FlxConsts.Max_Columns && IsEmptyCell(aRow, Result.Right + 1, false) && IsCenterAcrossSelection(aRow, Result.Right + 1))
					Result.Right++;
			}
			return Result;
		}

        protected real HeaderRowHeight()
        {
            return 250/ RowMultDisplay * Zoom100;
        }

        protected real RealRowHeight(int r)
        {
            return FWorkbook.GetRowHeight(r, true) / RowMultDisplay * Zoom100;
        }

        protected real CalcAcumRowHeight(int R1, int R2)
        {
            real Result = 0;
            for (int i = R1; i < R2; i++) Result += FWorkbook.GetRowHeight(i, true) / RowMultDisplay * Zoom100;
            for (int i = R1 - 1; i >= R2; i--) Result -= FWorkbook.GetRowHeight(i, true) / RowMultDisplay * Zoom100;

            return Result;
        }


        protected void CacheMergedCells(TXlsCellRange CellRange, int aPrintAreaRegion)
        {
			CellMergedCache[aPrintAreaRegion] = new TCellMergedCache();
            int aCount = FWorkbook.CellMergedListCount;
            for (int i = 1; i <= aCount; i++)
            {
                TXlsCellRange cr = FWorkbook.CellMergedList(i);
                int crLeft = Math.Max(cr.Left, CellRange.Left);
                int crRight = Math.Min(cr.Right, CellRange.Right);
                int crTop = Math.Max(cr.Top, CellRange.Top);
                int crBottom = Math.Min(cr.Bottom, CellRange.Bottom);
                for (int c = crLeft; c <= crRight; c++)
                    for (int r = crTop; r <= crBottom; r++)
                    {
                        CellMergedCache[aPrintAreaRegion][FlxHash.MakeHash(r, c)] = cr;
                    }
            }
        }

        protected void CacheColumns(TXlsCellRange CellRange)
        {
            int[] OldColumnCache = ColumnWidthsCache;
            int OldStart = ColumnStartCache;

            if (CellRange.Left < ColumnStartCache || CellRange.Right > ColumnStartCache + ColumnWidthsCache.Length)
            {
                ColumnStartCache = Math.Min(OldStart, CellRange.Left);
                int NewMax = Math.Max(CellRange.Right, OldStart + OldColumnCache.Length);
                ColumnWidthsCache = new int[NewMax - ColumnStartCache + 1];

                for (int i = ColumnStartCache; i < ColumnStartCache + ColumnWidthsCache.Length; i++)
                {
                    ColumnWidthsCache[i - ColumnStartCache] = FWorkbook.GetColWidth(i, true);
                }
            }
        }

        private void CacheShapes(int PrintAreaRegion, TXlsCellRange CellRange, int RowsInPage, int ColsInPage)
        {
            ShapesCache[PrintAreaRegion] = new TShapesCache(RowsInPage, ColsInPage, CellRange);
            ShapesCache[PrintAreaRegion].Fill(Workbook);

        }

        protected real HeaderColWidth()
        {
            return 920/ ColMultDisplay * Zoom100 * FmlaMult;
        }

        protected real RealColWidth(int c)
        {
            if (c< ColumnStartCache || c - ColumnStartCache >= ColumnWidthsCache.Length)
                return FWorkbook.GetColWidth(c, true)/ ColMultDisplay * Zoom100;

            return ColumnWidthsCache[c - ColumnStartCache] / ColMultDisplay * Zoom100;
        }

        protected real CalcAcumColWidth(int C1, int C2)
        {
            //We can't just have the width pre added because of rounding errors.
            real Result = 0;
            for (int i = C1; i < C2; i++) Result += RealColWidth(i);
            for (int i = C1 - 1; i >= C2; i--) Result -= RealColWidth(i);
            return Result;
        }

        internal static RectangleF RectangleXY(real l, real t, real r, real b)
        {
            return new RectangleF(l, t, r - l, b - t);
        }

        private RectangleF RectangleCell(int Row, int Col, ref RectangleF PaintClipRect, ref TXlsCellRange PagePrintRange)
        {
            return new  RectangleF(PaintClipRect.Left + CalcAcumColWidth(PagePrintRange.Left, Col),
                PaintClipRect.Top + CalcAcumRowHeight(PagePrintRange.Top, Row),
                RealColWidth(Col),
                RealRowHeight(Row));
        }


		/// <summary>
		/// Interface for sheetmethods
		/// </summary>
		/// <param name="A"></param>
        /// <param name="PagePrintRange"></param>
        /// <param name="PaintClipRect"></param>
		/// <returns></returns>
		public RectangleF GetImageRectangle(TXlsCellRange PagePrintRange, RectangleF PaintClipRect, TClientAnchor A)
		{
			return RectangleF.FromLTRB(
				PaintClipRect.Left + CalcAcumColWidth(PagePrintRange.Left, A.Col1) + A.Dx1 * RealColWidth(A.Col1) / 1024F,
				PaintClipRect.Top + CalcAcumRowHeight(PagePrintRange.Top, A.Row1) + A.Dy1  * RealRowHeight(A.Row1) / 255F,
				PaintClipRect.Left + CalcAcumColWidth(PagePrintRange.Left, A.Col2) + A.Dx2 * RealColWidth(A.Col2) / 1024F,           
				PaintClipRect.Top + CalcAcumRowHeight(PagePrintRange.Top, A.Row2) + A.Dy2  * RealRowHeight(A.Row2) / 255F);
		}

        #endregion

        #region Text drawing
        internal static real CalcAngle(int ExcelRotation, out bool Vertical)
        {
            Vertical = ExcelRotation == 255;
            if (ExcelRotation < 0) return 0;
            if (ExcelRotation <= 90) return ExcelRotation; //*2*Math.PI/360;
            if (ExcelRotation <= 180) return (90 - ExcelRotation); //*2*Math.PI/360;
            return 0;
        }

        private TRichString ColTitle(int ColNum)
        {
            string ColName = Workbook.OptionsR1C1 ? ColNum.ToString(CultureInfo.CurrentCulture) : TCellAddress.EncodeColumn(ColNum);
            return new TRichString(ColName, new TRTFRun[0], FWorkbook);
        }

        private TRichString RowTitle(int RowNum)
        {
            return new TRichString(RowNum.ToString(), new TRTFRun[0], FWorkbook);
        }

		private SizeF CalcTextExtent(Font AFont, TRichString Text, out real MaxDescent)
		{
			return RenderMetrics.CalcTextExtent(Canvas, FontCache, Zoom100, AFont, Text, out MaxDescent);
		}

        #endregion

        #region DrawCell
        private void DrawCell(int aCol, int aRow,
            RectangleF CellRect,
            RectangleF PaintClipRect,
            TSpawnedCellList SpawnedCells,
            bool ReallyDraw,
            bool OnlySpawned,
            TSpanDirection SpanDirection)
        {
            real Clp = 0;
            if (Zoom100 <= 0.50) Clp = 0; else if (Zoom100 <= 0.90) Clp = 1 * DispMul / 100f; else Clp = 1 * DispMul / 100f; //Do not use margins around the cell if zoom is small
            bool MultiLine = false;
            bool IsText = false;
            bool HAlignGeneral = false;
            THFlxAlignment HJustify = THFlxAlignment.general;
            TVFlxAlignment VJustify = TVFlxAlignment.bottom;
            //Reset Style

            real Alpha = 0;
            bool Vertical = false;
            double ActualValue = 0;

            Color DrawFontColor = Colors.Black;
            TFlxFont DrawFont = new TFlxFont();
            DrawFont.Name = "Arial"; DrawFont.Size20 = 200;
            TSubscriptData SubData = new TSubscriptData(TFlxFontStyles.None);

            THAlign HAlign = THAlign.Left;
            TVAlign VAlign = TVAlign.Bottom;
            real Indent = 0;
            TRichString OutText = new TRichString();

            if (FWorkbook == null) return;

            TAdaptativeFormats AdaptativeFormats = null;
            bool Merged = false;
            if (aRow > 0 && aCol > 0)
            {
                //Merged CELLS
                //We see this before anything else, because if it's Merged, we have to exit
                TXlsCellRange MergedBounds = XCellMergedBounds(aRow, aCol, out Merged);
                if (Merged && OnlySpawned) return;

                if (!Merged && MergedBounds.Right > aCol && SpawnedCells != null) //The cell is centered across selection.
                {
                    for (int c = aCol; c < MergedBounds.Right; c++)
                    {
                        SpawnedCells[FlxHash.MakeHash(aRow, c)] = null;
                    }
                }

                if (aCol > MergedBounds.Left) CellRect.X -= CalcAcumColWidth(MergedBounds.Left, aCol);
                if (MergedBounds.Left != MergedBounds.Right) CellRect.Width = CalcAcumColWidth(MergedBounds.Left, MergedBounds.Right + 1);
                if (aRow > MergedBounds.Top) CellRect.Y -= CalcAcumRowHeight(MergedBounds.Top, aRow);
                if (MergedBounds.Top != MergedBounds.Bottom) CellRect.Height = CalcAcumRowHeight(MergedBounds.Top, MergedBounds.Bottom + 1);

                if (aCol > MergedBounds.Left || aRow > MergedBounds.Top)
                {
                    DrawCell(MergedBounds.Left, MergedBounds.Top,
                        CellRect,
                        PaintClipRect, SpawnedCells, ReallyDraw, OnlySpawned, SpanDirection);
                    return;
                }

                //Value
                object OutValue = GetCellValue(aRow, aCol);
                TFormula Fmla = OutValue as TFormula;
                if (Fmla != null)//This should be before CellType=TExcelType.Object... so formula is converted to display.
                    if (FPrintFormulas)
                        OutValue = Fmla.Text;
                    else
                        OutValue = Fmla.Result;

                TFlxFormat Fm = GetCellVisibleFormatDef(aRow, aCol, Merged);

                GetDataAlign(OutValue, Fm, ref IsText, ref ActualValue, ref HAlign);

                //MULTILINE
                MultiLine = Fm.WrapText || Fm.HAlignment == THFlxAlignment.justify || Fm.VAlignment == TVFlxAlignment.justify;
                if (!IsText) MultiLine = false;

                //FONT
                DrawFontColor = Fm.Font.Color.ToColor(Workbook, Colors.Black);


                DrawFont = (TFlxFont)Fm.Font.Clone();

                SubData = new TSubscriptData(Fm.Font.Style);

                Alpha = CalcAngle(Fm.Rotation, out Vertical);

                //BORDERS
                /* Now handled on DrawLines()
                    */

                //ALIGN
                HJustify = Fm.HAlignment;
                GetHJustify(HJustify, ref HAlign, out HAlignGeneral);

                VJustify = Fm.VAlignment;
                GetVJustify(VJustify, Alpha, ref VAlign);

                //NUMERIC FORMAT
                bool HasDate, HasTime;
                OutText = TFlxNumberFormat.FormatValue(OutValue, Fm.Format, ref DrawFontColor, FWorkbook, out HasDate, out HasTime, out AdaptativeFormats);
                if (AdaptativeFormats != null && !AdaptativeFormats.IsEmpty) OutText = new TRichString(AdaptativeFormats.ApplySeparators(OutText.ToString()));
                OutText = TextPainter.ArabicShape(OutText, ReverseRightToLeftStrings);

                if (Fm.Indent != 0) Indent = Fm.Indent * 256f / ColMultDisplay * 1.74f;

            }

            //Header Row and col.
            if (aCol == 0 && aRow > 0 & !Workbook.GetAutoRowHeight(aRow))
            {
                DrawFont.Style |= TFlxFontStyles.Bold;
                DrawFontColor = Colors.Navy;
            }

            if (aCol == 0 || aRow == 0)
            {
                HAlign = THAlign.Center;
                //DrawBrush = new SolidBrush(Colors.Gray);
                //BottomColor=Colors.Gray;
                //RightColor=BottomColor;
                /*if (!HideCursor) 
                        if (aRow==Row || aCol=Col) DrawBrush= new SolidBrush(0x00F2BEAA);
                    //else if PointInGridRange(aCol, aRow, Selection) then ABrush:= $00F2BEAA;
                    */
                if (aRow == 0 && aCol != 0) OutText = ColTitle(aCol);
                else if (aRow != 0) OutText = RowTitle(aRow);
            }
            /*create a shadow for multiselect
                 * else
                    if (!HideCursor)
                {
                    if (AState==TGridDrawState.Selected!=0 && First ) 
                    {
                        ABrush=ColorToRGB(ABrush);
                        if ((ABrush & 0xFFFFFF)==0) ABrush=0x5A4942; else
                        {
                            R= ABrush & $ff;
                            G= ABrush & $ff00 shr 8;
                            B= ABrush & $ff0000 shr 16;
                            ABrush=Round(33+B*0.68)shl 16+Round(12+G*0.65) shl 8+ Round(R*0.68) ;
                        }
                    }
                }*/

            //Support for drawing a continued cell on an empty one
            if (IsEmptyCell(aRow, aCol))
            {
                if (SpanDirection != TSpanDirection.Right)
                {
                    //Search for the previous non empty cell
                    int i = FWorkbook.ColToIndex(aRow, aCol) - 1;
                    while (i > 0 && IsEmptyCell(aRow, FWorkbook.ColFromIndex(aRow, i))) i--;
                    if (i > 0)
                    {
                        int k = FWorkbook.ColFromIndex(aRow, i);
                        if (CellCanSpawnRight(aRow, k))
                            DrawCell(k, aRow, RectangleXY(CellRect.Left + CalcAcumColWidth(aCol, k), CellRect.Top, CellRect.Left + CalcAcumColWidth(aCol, k + 1), CellRect.Bottom), PaintClipRect, SpawnedCells, ReallyDraw, OnlySpawned, SpanDirection);
                    }
                }

                if (SpanDirection != TSpanDirection.Left)
                {
                    //Search for next non empty cell
                    int i = FWorkbook.ColToIndex(aRow, aCol);
                    while (i > 0 && i <= FWorkbook.ColCountInRow(aRow) && IsEmptyCell(aRow, FWorkbook.ColFromIndex(aRow, i))) i++;
                    if (i > 0 && i <= FWorkbook.ColCountInRow(aRow))
                    {
                        int k = FWorkbook.ColFromIndex(aRow, i);
                        if (CellCanSpawnLeft(aRow, k))
                            DrawCell(k, aRow, RectangleXY(CellRect.Left + CalcAcumColWidth(aCol, k), CellRect.Top, CellRect.Left + CalcAcumColWidth(aCol, k + 1), CellRect.Bottom), PaintClipRect, SpawnedCells, ReallyDraw, OnlySpawned, SpanDirection);
                    }
                }

                return; //nothing to draw
            }

            DrawText(FWorkbook, Canvas, FontCache, Zoom100, ReverseRightToLeftStrings, this,
                aCol, aRow, ref CellRect, ref PaintClipRect, SpawnedCells, ReallyDraw, OnlySpawned, Clp,
                MultiLine, HAlignGeneral, HJustify, VJustify, Alpha, Vertical, DrawFont, ref DrawFontColor,
                ref SubData, ref HAlign, VAlign, Indent, OutText, Merged, IsText, ActualValue, AdaptativeFormats);
        }

       
		public static void GetDataAlign(object Value, TFlxFormat Fm, ref bool IsText, ref double ActualValue, ref THAlign HAlign)
        {
            TCellType CellType = TExcelTypes.ObjectToCellType(Value);
            switch (CellType)
            {
                case TCellType.Bool:
                case TCellType.Error:
                    HAlign = THAlign.Center;
                    break;

                case TCellType.String:
				case TCellType.Unknown:
                    HAlign = THAlign.Left;
                    IsText = true;
                    break;

                case TCellType.Number:
                    if (Fm.Format == null || Fm.Format.Length == 0 || Fm.Format == "@"
						|| String.Equals(Fm.Format, "general", StringComparison.InvariantCultureIgnoreCase)) ActualValue = Convert.ToDouble(Value, CultureInfo.CurrentCulture);  //General format. We will need to try to fit this number on the column
                    HAlign = THAlign.Right;
                    break;

                default:
                    HAlign = THAlign.Right;
                    break;
            }
        }

        internal static void GetVJustify(TVFlxAlignment VJustify, real Alpha, ref TVAlign VAlign)
        {
            switch (VJustify)
            {
                case TVFlxAlignment.justify:
                case TVFlxAlignment.distributed:
                    if (Alpha == 0)
                        VAlign = TVAlign.Top;
                    break;
                case TVFlxAlignment.top:
                    VAlign = TVAlign.Top; break;
                case TVFlxAlignment.center: VAlign = TVAlign.Center; break;
                case TVFlxAlignment.bottom: VAlign = TVAlign.Bottom; break;
            } //case

        }

        internal static void GetHJustify(THFlxAlignment HAlignment, ref THAlign HAlign, out bool HAlignGeneral)
        {
            HAlignGeneral = false;
            switch (HAlignment)
            {
                case THFlxAlignment.left: HAlign = THAlign.Left; break;
                case THFlxAlignment.center:
                case THFlxAlignment.center_across_selection:
                    HAlign = THAlign.Center;
                    break;
                case THFlxAlignment.right: HAlign = THAlign.Right; break;
                case THFlxAlignment.general: HAlignGeneral = true; break;
            }//case
        }

        private static TXRichString SharpLine(IFlxGraphics Canvas, TFontCache FontCache, real Zoom100, Font AFont, real CellWidth)
        {
            TRichString s = new TRichString("#");
            real md;
            SizeF tm = RenderMetrics.CalcTextExtent(Canvas, FontCache, Zoom100, AFont, s, out md);

            int SharpCount = tm.Width == 0? 15: (int)(CellWidth / tm.Width) + 1;
            if (SharpCount < 1) SharpCount = 1;
            return new TXRichString(new TRichString(new string('#', SharpCount)), false, tm.Width * SharpCount, tm.Height, null);
        }

        public static TXRichString TryToFit(IFlxGraphics Canvas, TFontCache FontCache, real Zoom100, Font AFont, double ActualValue, real CellWidth, real Clp)
        {
            if (ActualValue == 0) return SharpLine(Canvas, FontCache, Zoom100, AFont, CellWidth);   // 0 means no fit. Anyway, it is impossible to fit a 0 into something smaller

            //Color does not matter, it has already been found out. We can use standard .NET formats here
            
            SizeF SaveTm = new SizeF(0,0);
            TRichString SaveS = null;
            int Max = 15; int Min = 0;
            while (Min <= Max)
            {
                int i = (Max + Min) / 2;
                string sf = ActualValue.ToString("G" + (i+1).ToString(CultureInfo.CurrentCulture), CultureInfo.CurrentCulture);

                TRichString s = new TRichString(sf);
                real md;
                SizeF tm = RenderMetrics.CalcTextExtent(Canvas, FontCache, Zoom100, AFont, s, out md);
                if (tm.Width > CellWidth - 2* Clp)
                {
                    Max = i - 1; 
                }
                else
                {
                    SaveS = s;
                    SaveTm = tm;
                    Min = i + 1;
                }
            }

            if (SaveS != null) return new TXRichString(SaveS, false, SaveTm.Width, SaveTm.Height, null); //TryToFit is only called with general, no need for adaptative formats.
            return SharpLine(Canvas, FontCache, Zoom100, AFont, CellWidth);
        }

		internal static void DrawText(ExcelFile Workbook, IFlxGraphics Canvas, TFontCache FontCache, real Zoom100, bool ReverseRightToLeftStrings, IDrawObjectMethods SheetMethods,
			int aCol, int aRow, ref RectangleF CellRect, ref RectangleF PaintClipRect, TSpawnedCellList SpawnedCells, bool ReallyDraw, bool OnlySpawned, real Clp, bool MultiLine, bool HAlignGeneral, THFlxAlignment HJustify, TVFlxAlignment VJustify, 
			real Alpha, bool Vertical, TFlxFont DrawFont, ref Color DrawFontColor,  
			ref TSubscriptData SubData, ref THAlign HAlign, TVAlign VAlign, 
			real Indent, TRichString OutText, bool Merged, bool IsText, double ActualValue, TAdaptativeFormats AdaptativeFormats)
		{
			Font AFont = FontCache.GetFont(DrawFont, SubData.Factor * Zoom100);
		{
			real SinAlpha = (real)Math.Sin(Alpha * Math.PI / 180); real CosAlpha = (real)Math.Cos(Alpha * Math.PI / 180);
			
            SizeF TextExtent;
            TXRichStringList TextLines;
            TFloatList MaxDescent;
            TAdaptativeFormats AF = ReallyDraw ? AdaptativeFormats : null;

            RectangleF CellRect1 = new RectangleF(CellRect.Left + Indent, CellRect.Top, CellRect.Width - Indent, CellRect.Height);
            TextPainter.CalcTextBox(Canvas, FontCache, Zoom100, CellRect1, Clp, MultiLine, Alpha, Vertical, OutText, AFont, AF, out TextExtent, out TextLines, out MaxDescent);
			if (TextLines.Count <= 0) return;

			if (!IsText && TextLines.Count == 1)
			{
				real CellWidth = CellRect.Width;
				if (Alpha != 0)
				{
					real d = CellRect.Height / Math.Abs(SinAlpha);

					CellWidth = d - TextLines[0].YExtent * CosAlpha / Math.Abs(SinAlpha);
				}
				if (TextLines[0].XExtent > CellWidth)
				{
					TextLines[0] = TryToFit(Canvas, FontCache, Zoom100, AFont, ActualValue, CellWidth, Clp);
				}
			}

			real[] X;
			real[] Y;

			RectangleF ContainingRect = TextPainter.CalcTextCoords(out X, out Y, OutText, VAlign, ref HAlign, Indent, Alpha, CellRect, Clp, 
                TextExtent, HAlignGeneral, Vertical, SinAlpha, CosAlpha, TextLines, Workbook.Linespacing, VJustify);


			//Uncomment following to test boxes.
			//Canvas.ResetClip();
			//Canvas.FillRectangle(Brushes.LightBlue, ContainingRect);
			
			//Now we have the coords for the cell and for the text. Check if we need to span the cell.
			//Cell can span to both sides at the same time, when centered.
			if ((MultiLine || Merged) && OnlySpawned) return;

			//RectangleF TextRect = new RectangleF(CellRect.Left + Clp, CellRect.Top + Clp, CellRect.Width - 2 * Clp, CellRect.Height - 2 * Clp);
			RectangleF TextRect = new RectangleF(CellRect.Left + Clp, CellRect.Top, CellRect.Width - 2 * Clp, CellRect.Height);

			if (!MultiLine && !Merged && aCol >= 0)
				SheetMethods.SpawnCell(aCol, aRow, SpawnedCells, HAlign, ref TextRect, ref ContainingRect);

			if (ReallyDraw && OutText.Length > 0)
				TextPainter.DrawRichText(Workbook, Canvas, FontCache, Zoom100, ReverseRightToLeftStrings, ref CellRect, ref PaintClipRect, ref TextRect, ref ContainingRect, 
					Clp, HJustify, VJustify, Alpha, DrawFontColor, SubData, OutText, TextExtent, TextLines, AFont, MaxDescent, X, Y);
		}
		}

 
        public void SpawnCell(int aCol, int aRow, TSpawnedCellList SpawnedCells, THAlign HAlign, ref RectangleF TextRect, ref RectangleF ContainingRect)
        {
            int LastCol = aCol + 1;
            if (CellCanSpawnRight(aRow, aCol) && HAlign != THAlign.Right)
            {
                while (ContainingRect.Right > TextRect.Right
                    && LastCol < ColumnWidthsCache.Length + ColumnStartCache && LastCol >= ColumnStartCache
                    && LastCol <= FlxConsts.Max_Columns + 1
                    && IsEmptyCell(aRow, LastCol))
                {
                    bool Merged;
                    XCellMergedBounds(aRow, LastCol, out Merged);
                    if (Merged) break;
                    TextRect.Width += RealColWidth(LastCol);
                    if (SpawnedCells != null) SpawnedCells[FlxHash.MakeHash(aRow, LastCol - 1)] = null;
                    LastCol++;
                }
            }

            LastCol = aCol - 1;
            if (CellCanSpawnLeft(aRow, aCol) && HAlign != THAlign.Left)
            {
                while (ContainingRect.Left < TextRect.Left
                    && LastCol < ColumnWidthsCache.Length + ColumnStartCache && LastCol >= ColumnStartCache
                    && LastCol <= FlxConsts.Max_Columns + 1
                    && IsEmptyCell(aRow, LastCol))
                {
                    bool Merged;
                    XCellMergedBounds(aRow, LastCol, out Merged);
                    if (Merged) break;

                    real cw = RealColWidth(LastCol);
                    TextRect.Width += cw;
                    TextRect.X -= cw;
                    if (SpawnedCells != null) SpawnedCells[FlxHash.MakeHash(aRow, LastCol)] = null;
                    LastCol--;
                }
            }
        }


        #endregion

        #region Calculate pages

		protected void CalcPrintedPage(int StartRow, int StartCol, out int EndRow, out int EndCol, TXlsCellRange PrintRange, RectangleF PaintClipRect, TRepeatingRange RepeatingRange)
		{
			int dummy;
			CalcRowsInPage(StartRow, out EndRow, PrintRange, PaintClipRect, RepeatingRange, 100, out dummy);
			CalcColsInPage(StartCol, out EndCol, PrintRange, PaintClipRect, RepeatingRange, 100, out dummy);
		}


		private void CalcRowsInPage(int StartRow, out int EndRow, TXlsCellRange PrintRange, RectangleF PaintClipRect, TRepeatingRange RepeatingRange)
		{
			int dummy;
			CalcRowsInPage(StartRow, out EndRow, PrintRange, PaintClipRect, RepeatingRange, 100, out dummy);
		}

		private void CalcColsInPage(int StartCol, out int EndCol, TXlsCellRange PrintRange, RectangleF PaintClipRect, TRepeatingRange RepeatingRange)
		{
			int dummy;
			CalcColsInPage(StartCol, out EndCol, PrintRange, PaintClipRect, RepeatingRange, 100, out dummy);
		}

		private void CalcRowsInPage(int StartRow, out int EndRow, TXlsCellRange PrintRange, RectangleF PaintClipRect, TRepeatingRange RepeatingRange, int PagePercent, out int LastMinLevel)
		{
			CalcPrintedPageAndPageBreaks(StartRow, out EndRow, PrintRange.Top, PrintRange.Bottom, PaintClipRect.Top, PaintClipRect.Height, PaintClipRect.Bottom,
				RepeatingRange.FirstRow, RepeatingRange.MaxRow(StartRow), PagePercent, out LastMinLevel, true);
			
		}

		private void CalcColsInPage(int StartCol, out int EndCol, TXlsCellRange PrintRange, RectangleF PaintClipRect, TRepeatingRange RepeatingRange, int PagePercent, out int LastMinLevel)
		{
			CalcPrintedPageAndPageBreaks(StartCol, out EndCol, PrintRange.Left, PrintRange.Right, PaintClipRect.Left, PaintClipRect.Width, PaintClipRect.Right,
				RepeatingRange.FirstCol, RepeatingRange.MaxCol(StartCol), PagePercent, out LastMinLevel, false);
			
		}

		private real RealRowColHeight(int r, bool IsRows)
		{
			return IsRows? RealRowHeight(r): RealColWidth(r);
		}

		private bool IsFitToPage(bool IsRows)
		{
			return IsRows?
				FWorkbook.PrintNumberOfVerticalPages!=0:
				FWorkbook.PrintNumberOfHorizontalPages!=0;
		}

		private bool HasPageBreak(int r, bool IsRows)
		{
			return IsRows?
				FWorkbook.HasHPageBreak(r):
				FWorkbook.HasVPageBreak(r);
		}

        private void CalcPrintedPageAndPageBreaks(int StartRow, out int EndRow, int PageTop, int PageBottom, real ClipTop, real ClipHeight, real ClipBottom, int RepeatingTop, int RepeatingBottom,
            int PagePercent, out int LastMinLevel, bool IsRows)
        {
            LastMinLevel = -1;

            // Calc rows
            real Ch = ClipTop;

            //Headings
			if (FWorkbook.PrintHeadings) 
			{
				if (IsRows) Ch+= HeaderRowHeight(); else Ch += HeaderColWidth();
			}

            //Repeating rows
            for (int Row = RepeatingTop; Row <= RepeatingBottom; Row++)
            {
                Ch+= RealRowColHeight(Row, IsRows);
            }

            //Normal rows
            EndRow = PageBottom;
            int MinRowLevel = -1;
			bool BrokenPage = false;
            for (int Row = StartRow; Row <= PageBottom; Row++)
            {
                if (PagePercent < 100 && Ch > ClipTop + ClipHeight * PagePercent / 100.0)
                {
                    if (Row > StartRow)
                    {
						int KeepLevel = IsRows? FWorkbook.GetKeepRowsTogether(Row - 1) : FWorkbook.GetKeepColsTogether(Row - 1);
						if (MinRowLevel == -1 || KeepLevel <= MinRowLevel)
						{
							MinRowLevel = KeepLevel;
							LastMinLevel = Row - 1;
						}
					}
                }

				Ch+= RealRowColHeight(Row, IsRows);

                if ((Ch > ClipBottom) ||
                    ((!FWorkbook.PrintToFit || !IsFitToPage(IsRows)) && (Row > StartRow) && (HasPageBreak(Row - 1, IsRows)))
                    )
                {
                    EndRow = Row - 1;
					BrokenPage = true;
                    break;
                }
            }

			if (!BrokenPage)
			{
				LastMinLevel = -1; //If we arrived to the end without needing to break, there is no need to fit
			}

            if (EndRow < StartRow) EndRow = StartRow; //A cell taller than one page...
        }

        protected TXlsMargins GetMargins(RectangleF PageBounds)
        {
            TXlsMargins Margins = FWorkbook.GetPrintMargins();
            real InchesMulX = DispMul; //GraphicUnit.Display means 1/75 INCHES.
            real InchesMulY = DispMul;

            if (Margins.Left * InchesMulX < (MarginBounds.Left - PageBounds.Left)/100F*DispMul)
                Margins.Left = (MarginBounds.Left - PageBounds.Left)/100F*DispMul/ InchesMulX;
            if (Margins.Top * InchesMulY < (MarginBounds.Top - PageBounds.Top) /100F*DispMul)
                Margins.Top = (MarginBounds.Top - PageBounds.Top)/100F*DispMul/InchesMulY;

            if (Margins.Right * InchesMulX < (PageBounds.Right - MarginBounds.Right)/100F*DispMul)
                Margins.Right = MarginBounds.Left/100F*DispMul/ InchesMulX;
            if (Margins.Bottom * InchesMulY < (PageBounds.Bottom - MarginBounds.Bottom) /100F*DispMul)
                Margins.Bottom = (PageBounds.Bottom - MarginBounds.Bottom)/100F*DispMul/InchesMulY;

            if (Margins.Header * InchesMulY < (MarginBounds.Top - PageBounds.Top) /100F*DispMul)
                Margins.Header = (MarginBounds.Top - PageBounds.Top)/100F*DispMul/InchesMulY;
            if (Margins.Footer * InchesMulY < (PageBounds.Bottom - MarginBounds.Bottom) /100F*DispMul)
                Margins.Footer = (PageBounds.Bottom - MarginBounds.Bottom)/100F*DispMul/InchesMulY;

            return Margins;
        }

        private void CalcPrintParams(TXlsCellRange PrintRange, RectangleF PageBounds, out RectangleF PaintClipRect, TRepeatingRange RepeatingRange)
        {
            FMargins = GetMargins(PageBounds);
            real InchesMulX = DispMul; //GraphicUnit.Display means 1/75 INCHES.
            real InchesMulY = DispMul;

            real aZoom = 0;
            PaintClipRect = RectangleXY(PageBounds.Left / 100F * DispMul + (real)FMargins.Left * InchesMulX, PageBounds.Top / 100F * DispMul + (real)FMargins.Top * InchesMulY, (real)(PageBounds.Right / 100F * DispMul - FMargins.Right * InchesMulX), (real)(PageBounds.Bottom / 100F * DispMul - FMargins.Bottom * InchesMulY));

            if (FWorkbook.PrintToFit && (PaintClipRect.Right > PaintClipRect.Left) && (PaintClipRect.Bottom > PaintClipRect.Top)
                && PrintRange.ColCount > 0 && PrintRange.RowCount > 0)
            {
                Zoom100 = 1;
                real Cp = CalcAcumColWidth(PrintRange.Left, PrintRange.Right);
                real Rp = CalcAcumRowHeight(PrintRange.Top, PrintRange.Bottom);

                real fc = 0;
                real fr = 0;
                //To do this right, we must calculate how many sheets include repeatingrange, how many don't and how many include a part of it.
                //if (RepeatingRange.LastCol>=RepeatingRange.FirstCol) fc+=CalcAcumColWidth(RepeatingRange.FirstCol, RepeatingRange.LastCol+1);
                //if (RepeatingRange.LastRow>=RepeatingRange.FirstRow) fr+=CalcAcumRowHeight(RepeatingRange.FirstRow, RepeatingRange.LastRow+1);
                if (FWorkbook.PrintHeadings)
                {
                    fc += HeaderColWidth();
                    fr += HeaderRowHeight();
                }
                //First try
                real XZoom = 100;
                if (Cp > 0) XZoom = 100 * FWorkbook.PrintNumberOfHorizontalPages * (PaintClipRect.Right - PaintClipRect.Left - fc) / Cp;
                if (FWorkbook.PrintNumberOfHorizontalPages == 0) XZoom = 100;

                real YZoom = 100;
                if (Rp > 0) YZoom = 100 * FWorkbook.PrintNumberOfVerticalPages * (PaintClipRect.Bottom - PaintClipRect.Top - fr) / Rp;
                if (FWorkbook.PrintNumberOfVerticalPages == 0) YZoom = 100;
                aZoom = Math.Min(XZoom, YZoom);
                if (aZoom > 100) aZoom = 100;
                //Iterative try until is ok;
                bool PagesOk = false;
                if (aZoom > 10)
                    do
                    {
                        int XPageCount; int YPageCount;
                        Zoom100 = aZoom / 100;
                        InternalCalcNumberOfPrintingPages(PrintRange, PaintClipRect, out XPageCount, out YPageCount, RepeatingRange);
                        PagesOk = (XPageCount <= FWorkbook.PrintNumberOfHorizontalPages || FWorkbook.PrintNumberOfHorizontalPages == 0)
                            && (YPageCount <= FWorkbook.PrintNumberOfVerticalPages || FWorkbook.PrintNumberOfVerticalPages == 0);
                        if (!PagesOk) aZoom -= 1;
                    }
                    while ((aZoom > 10) && !PagesOk);
            }
            else
                aZoom = FWorkbook.PrintScale;

            if (aZoom < 10) aZoom = 10;
            if (aZoom > 400) aZoom = 400;
            Zoom100 = aZoom / 100;

        }

        private void InternalCalcNumberOfPrintingPages(TXlsCellRange PrintRange, RectangleF PaintClipRect, out int HCount, out int VCount, TRepeatingRange RepeatingRange )
        {
			if (FWorkbook.SheetType == TSheetType.Chart)
			{
				HCount = 1;
				VCount = 1;
				return;
			}

            VCount = 0;
            int Sr = PrintRange.Top;
            do
            {
				int Er; 
                CalcRowsInPage(Sr, out Er, PrintRange, PaintClipRect, RepeatingRange);
                VCount++;
                Sr = Er + 1;
            }
            while (Sr <= PrintRange.Bottom);

            HCount = 0;
			int Sc = PrintRange.Left;
            do
            {
				int Ec;
                CalcColsInPage(Sc, out Ec, PrintRange, PaintClipRect, RepeatingRange);
                HCount++;
                Sc = Ec + 1;
            }
            while (Sc <= PrintRange.Right);
        }

        #endregion

        #region Publics
        /// <summary>
        /// Always follow this call with a DisposeFontCache.
        /// </summary>
        public void CreateFontCache()
        {
            if (FontCache==null)
                FontCache = new TFontCache();
        }

        public void DisposeFontCache()
        {
            if (FontCache!=null) FontCache.Dispose();
            FontCache = null;
        }

        public TXlsCellRange[] InternalCalcPrintArea(TXlsCellRange FPrintRange)
        {
            bool IgnoreFormulaText = FWorkbook.IgnoreFormulaText;
            try
            {
                Workbook.IgnoreFormulaText = true; //even if printing formula text...
                if (FWorkbook.SheetType == TSheetType.Chart) return new TXlsCellRange[] { new TXlsCellRange(1, 1, 1, 1) };
                TXlsCellRange ResultRange = FPrintRange;
                if (ResultRange.Left <= 0 || ResultRange.Right <= 0 || ResultRange.Top <= 0 || ResultRange.Bottom <= 0)
                {
                    TXlsNamedRange Pr = Workbook.GetNamedRange(TXlsNamedRange.GetInternalName(InternalNameRange.Print_Area), -1, Workbook.ActiveSheet);
                    if (Pr != null)
                    {
                        TXlsCellRange[] Result = Pr.GetRanges();
                        if (Result == null || Result.Length == 0)
                        {
                            ResultRange = new TXlsCellRange(1, 1, MaxVisibleRow, MaxVisibleCol);
                        }
                        else
                        {
                            bool LeftToRight = (FWorkbook.PrintOptions & TPrintOptions.LeftToRight) != 0;
                            if (Result.Length > 0)
                            {
                                IComparer<TXlsCellRange> RangeDirection = new TPrintAreaSort(LeftToRight);
                                Array.Sort(Result, RangeDirection);
                            }

                            foreach (TXlsCellRange rng in Result)
                            {
                                if (rng.Top == 1 && rng.Bottom >= FlxConsts.Max_Rows + 1) rng.Bottom = MaxVisibleRow;
                                if (rng.Left == 1 && rng.Right >= FlxConsts.Max_Columns + 1) rng.Right = MaxVisibleCol;
                            }
                            return Result;
                        }
                    }
                    else
                        ResultRange = new TXlsCellRange(1, 1, MaxVisibleRow, MaxVisibleCol);
                }

                return new TXlsCellRange[] { ResultRange };
            }
            finally
            {
                FWorkbook.IgnoreFormulaText = IgnoreFormulaText;
            }
        }

        private void InitRender(TXlsCellRange PrintRange, int aPrintAreaRegion, int aPrintAreaCount)
        {
            ColMultDisplay = ExcelMetrics.ColMultDisplay(FWorkbook) * 100F / DispMul;
            RowMultDisplay = ExcelMetrics.RowMultDisplay(FWorkbook) * 100F / DispMul;
            FmlaMult = ExcelMetrics.FmlaMult(FWorkbook);

            CurrentPrintAreaRegion = aPrintAreaRegion;
            PageFormatCache.Clear();
            if (aPrintAreaRegion == 0)
            {
                CellMergedCache = new TCellMergedCache[aPrintAreaCount];
                ColumnWidthsCache = new int[0];
                ColumnStartCache = PrintRange.Left;
            }

            CacheColumns(PrintRange);
        }

		private bool HasKeepRowsTogether()
		{
			return Workbook.HasKeepRowsTogether();
		}

		private bool HasKeepColsTogether()
		{
			return Workbook.HasKeepColsTogether();
		}

        public void AddPageBreaks(int PagePercent,  RectangleF PageBounds, RectangleF aMarginBounds, int PageScale, 
            TXlsCellRange PrintRange, int PrintArea, int PrintAreaCount)
        {
            if (FWorkbook.SheetType == TSheetType.Chart) return;
            TRepeatingRange RepeatingRange = GetRepeatingRange(PrintRange);

            MarginBounds = aMarginBounds;
            InitRender(PrintRange, PrintArea, PrintAreaCount);
            CacheMergedCells(PrintRange, PrintArea);

            RectangleF PaintClipRect;
            CalcPrintParams(PrintRange, PageBounds, out PaintClipRect, RepeatingRange); //Will calc zoom100

            if (PageScale != 100)
            {
                PaintClipRect.Height *= PageScale / 100f;
                PaintClipRect.Width *= PageScale / 100f;
            }

			if (!FWorkbook.PrintToFit || !IsFitToPage(true)) //When fitting in a page, page breaks are ignored, so we will not add them.
			{
				if (HasKeepRowsTogether())
				{
					int Sr = PrintRange.Top;
					do
					{
						int Er;
						int LastMinLevel;
						CalcRowsInPage(Sr, out Er, PrintRange, PaintClipRect, RepeatingRange, PagePercent, out LastMinLevel);

						if (LastMinLevel >= 0)
						{
							if (LastMinLevel < PrintRange.Bottom) FWorkbook.InsertHPageBreak(LastMinLevel);
							Er = LastMinLevel;
						}

						Sr = Er + 1;
					}
					while (Sr <= PrintRange.Bottom);
				}
			}

			if (!FWorkbook.PrintToFit || !IsFitToPage(false))
			{
				if (HasKeepColsTogether())
				{
					int Sc = PrintRange.Left;
					do
					{
						int Ec;
						int LastMinLevel;
						CalcColsInPage(Sc, out Ec, PrintRange, PaintClipRect, RepeatingRange, PagePercent, out LastMinLevel);

						if (LastMinLevel >= 0)
						{
							if (LastMinLevel < PrintRange.Right) FWorkbook.InsertVPageBreak(LastMinLevel);
							Ec = LastMinLevel;
						}

						Sc = Ec + 1;
					}
					while (Sc <= PrintRange.Right);
				}
			}
        }

        public void InitializePrint(IFlxGraphics ACanvas, RectangleF PageBounds, RectangleF aMarginBounds, TXlsCellRange[] PrintRanges,
            out RectangleF[] PaintClipRect, out int TotalPages, out TXlsCellRange PagePrintRange)
        {
            TotalPages = 0;
            PagePrintRange = null;
            PaintClipRect = new RectangleF[PrintRanges.Length];
            CellMergedCache = new TCellMergedCache[PrintRanges.Length];
            ShapesCache = new TShapesCache[PrintRanges.Length];
            int i = 0;
            foreach (TXlsCellRange PrintRange in PrintRanges)
            {
                int PartialTotalPages;
                TXlsCellRange PartialPagePrintRange;
                RectangleF PartialPaintClipRect;
                InitializePrintOneRange(ACanvas, PageBounds, aMarginBounds, out PartialPaintClipRect, PrintRange, 
                    out PartialTotalPages, out PartialPagePrintRange, i, PrintRanges.Length);
                TotalPages += PartialTotalPages;
                if (i == 0) PagePrintRange = PartialPagePrintRange;
                PaintClipRect[i] = PartialPaintClipRect;
                i++;
            }
        }

        private void InitializePrintOneRange(IFlxGraphics ACanvas, RectangleF PageBounds, RectangleF aMarginBounds, 
            out RectangleF PaintClipRect, TXlsCellRange PrintRange, out int TotalPages, out TXlsCellRange PagePrintRange,
            int aPrintAreaRegion, int aPrintAreaCount)
        {
            TRepeatingRange RepeatingRange = GetRepeatingRange(PrintRange);

            MarginBounds = aMarginBounds;
            Canvas = ACanvas;
            InitRender(PrintRange, aPrintAreaRegion, aPrintAreaCount);
            CacheMergedCells(PrintRange, aPrintAreaRegion);

            CalcPrintParams(PrintRange, PageBounds, out PaintClipRect, RepeatingRange); //Will calc zoom100

            int HCount;
            int VCount;
            InternalCalcNumberOfPrintingPages(PrintRange, PaintClipRect, out HCount, out VCount, RepeatingRange);
            TotalPages = HCount * VCount;

            PagePrintRange = new TXlsCellRange();
            PagePrintRange.Left = PrintRange.Left;
            PagePrintRange.Top = PrintRange.Top;

            int Er; int Ec;
            CalcPrintedPage(PrintRange.Top, PrintRange.Left, out Er, out Ec, PrintRange, PaintClipRect, RepeatingRange);
            CacheShapes(aPrintAreaRegion, PrintRange, Er - PrintRange.Top + 1, Ec - PrintRange.Left + 1);
        }

        public void GenericPrint(IFlxGraphics ACanvas, RectangleF PageBounds, TXlsCellRange[] PrintRanges, int PageNumber, RectangleF[] PaintClipRect,
            int TotalPages, bool ReallyDraw, TXlsCellRange PagePrintRange, ref int PrintArea)
        {
            GenericPrint(ACanvas, PageBounds, PrintRanges, PageNumber, PaintClipRect[PrintArea], TotalPages, ReallyDraw, PagePrintRange, ref PrintArea);
        }

        private void GenericPrint(IFlxGraphics ACanvas, RectangleF PageBounds, TXlsCellRange[] PrintRanges, int PageNumber, RectangleF PaintClipRect, 
            int TotalPages, bool ReallyDraw, TXlsCellRange PagePrintRange, ref int PrintArea)
        {
            if (FWorkbook == null) return;
            CurrentPrintAreaRegion = PrintArea;
            TXlsCellRange PrintRange = PrintRanges[PrintArea];

            TReferenceStyle SaveRefStyle = Workbook.FormulaReferenceStyle;
            bool IgnoreFormulaText = Workbook.IgnoreFormulaText;
            try
            {
                if (Workbook.OptionsR1C1) Workbook.FormulaReferenceStyle = TReferenceStyle.R1C1; else Workbook.FormulaReferenceStyle = TReferenceStyle.A1;
                Workbook.IgnoreFormulaText = !FWorkbook.ShowFormulaText;

                TRepeatingRange RepeatingRange = GetRepeatingRange(PrintRange);
                Canvas = ACanvas;
                FDrawGridLines = FWorkbook.PrintGridLines;
                FPrintFormulas = FWorkbook.ShowFormulaText;
                GridLinesColor = FWorkbook.GridLinesColor.ToColor(Workbook, Colors.LightGray);

                Canvas.CreateSFormat(); //GenericTypographic returns a NEW instance.
                try
                {
                    RectangleF PaintClipRect1 = RectangleXY(
                        PaintClipRect.Left - 1 * DispMul / 100f,
                        PaintClipRect.Top - 1 * DispMul / 100f,
                        PaintClipRect.Right + 1 * DispMul / 100f,
                        PaintClipRect.Bottom + 1 * DispMul / 100f);
                    RectangleF PaintClipRect2 = new RectangleF(PageBounds.Left / 100F * DispMul, PageBounds.Top / 100F * DispMul, PageBounds.Width / 100F * DispMul, PageBounds.Height / 100F * DispMul);

                    TDrawObjects DrawObjects = new TDrawObjects(FWorkbook, Canvas, FontCache, Zoom100, FHidePrintObjects, ReverseRightToLeftStrings, this);

                    CalcPrintedPage(PagePrintRange.Top, PagePrintRange.Left, out PagePrintRange.Bottom, out PagePrintRange.Right, PrintRange, PaintClipRect, RepeatingRange);

                    if (ReallyDraw)
                    {
                        PageFormatCache.CreatePageCache(PagePrintRange);
                        try
                        {

                            DrawHeadersAndFooters(PaintClipRect2, PageNumber, TotalPages, PaintClipRect1, PaintClipRect2);  //This should be before drawing the cells.

                            if (FWorkbook.SheetType == TSheetType.Chart)
                            {
                                TShapeProperties Shp = new TShapeProperties();
                                Shp.ShapeOptions = new TShapeOptionList();
                                if (Workbook.ChartCount < 1) return;
                                DrawObjects.DrawEmbeddedChart(1, Canvas, Workbook, FontCache, Shp, PaintClipRect, new TShadowInfo(TShadowStyle.None, 0), TClippingStyle.None, 1, false);
                            }
                            else
                            {
                                real HHeading = 0;
                                real WHeading = 0;
                                if (FWorkbook.PrintHeadings)
                                {
                                    HHeading = HeaderRowHeight();
                                    WHeading = HeaderColWidth();
                                }

                                real HFixed = 0;
                                if (RepeatingRange.RowCount(PagePrintRange.Top) > 0)
                                    HFixed = CalcAcumRowHeight(RepeatingRange.FirstRow, RepeatingRange.MaxRow(PagePrintRange.Top) + 1);
                                real WFixed = 0;
                                if (RepeatingRange.ColCount(PagePrintRange.Left) > 0)
                                    WFixed = CalcAcumColWidth(RepeatingRange.FirstCol, RepeatingRange.MaxCol(PagePrintRange.Left) + 1);

                                real HOfs = 0;
                                if (FWorkbook.PrintVCentered)
                                {
                                    HOfs = HHeading + HFixed + CalcAcumRowHeight(PagePrintRange.Top, PagePrintRange.Bottom + 1);
                                    HOfs = (PaintClipRect.Bottom - PaintClipRect.Top - HOfs) / 2f;
                                }

                                if (MarginBounds.Top > 0) HOfs -= MarginBounds.Top / 100F * DispMul;

                                real WOfs = 0;
                                if (FWorkbook.PrintHCentered)
                                {
                                    WOfs = WHeading + WFixed + CalcAcumColWidth(PagePrintRange.Left, PagePrintRange.Right + 1);
                                    WOfs = (PaintClipRect.Right - PaintClipRect.Left - WOfs) / 2f;
                                }

                                if (MarginBounds.Left > 0) WOfs -= MarginBounds.Left / 100F * DispMul;


                                // Remember to Add room for double lines. (they go outside the margins)
                                real wt = PaintClipRect.Left + WHeading + WOfs;
                                real ht = PaintClipRect.Top + HHeading + HOfs;
                                RectangleF PaintClipRectRnC = RectangleXY(wt, ht, wt + WFixed + 2 * DispMul / 100f, ht + HFixed + 2 * DispMul / 100f);
                                RectangleF PaintClipRectDat = RectangleXY(wt + WFixed, ht + HFixed, PaintClipRect.Right + WHeading + WOfs, PaintClipRect.Bottom + HHeading + HOfs);
                                RectangleF PaintClipRectRow = RectangleXY(wt + WFixed, ht, PaintClipRect.Right + WHeading + WOfs, ht + HFixed + 2 * DispMul / 100f);
                                RectangleF PaintClipRectCol = RectangleXY(wt, ht + HFixed, wt + WFixed + 2 * DispMul / 100f, PaintClipRect.Bottom + HHeading + HOfs);

                                int rbottom = RepeatingRange.MaxRow(PagePrintRange.Top);
                                int rright = RepeatingRange.MaxCol(PagePrintRange.Left);
                                TXlsCellRange RowColRange = new TXlsCellRange(RepeatingRange.FirstRow, RepeatingRange.FirstCol, rbottom, rright);
                                TXlsCellRange RowRange = new TXlsCellRange(RepeatingRange.FirstRow, PagePrintRange.Left, rbottom, PagePrintRange.Right);
                                TXlsCellRange ColRange = new TXlsCellRange(PagePrintRange.Top, RepeatingRange.FirstCol, PagePrintRange.Bottom, rright);

                                TSpawnedCellList SpawnedCells = new TSpawnedCellList();

                                //First do the background and cells, so lines are not overwriten
                                if (RowColRange.RowCount > 0 && RowColRange.ColCount > 0) DrawPage(PaintClipRectRnC, RowColRange, true, SpawnedCells, DrawObjects, false, false);
                                if (RowRange.RowCount > 0) DrawPage(PaintClipRectRow, RowRange, true, SpawnedCells, DrawObjects, false, false);
                                if (ColRange.ColCount > 0) DrawPage(PaintClipRectCol, ColRange, true, SpawnedCells, DrawObjects, false, false);
                                //Draw Normal cells
                                DrawPage(PaintClipRectDat, PagePrintRange, true, SpawnedCells, DrawObjects, false, false);

                                //Now the rest of things.
                                if (RowColRange.RowCount > 0 && RowColRange.ColCount > 0) DrawPage(PaintClipRectRnC, RowColRange, false, SpawnedCells, DrawObjects, true, true);
                                if (RowRange.RowCount > 0) DrawPage(PaintClipRectRow, RowRange, false, SpawnedCells, DrawObjects, true, WFixed == 0);
                                if (ColRange.ColCount > 0) DrawPage(PaintClipRectCol, ColRange, false, SpawnedCells, DrawObjects, HFixed == 0, true);
                                //Draw Normal cells
                                DrawPage(PaintClipRectDat, PagePrintRange, false, SpawnedCells, DrawObjects, HFixed == 0, WFixed == 0);
                                DrawHeadings(PaintClipRect1, PaintClipRect2, WOfs, HOfs, PagePrintRange, RepeatingRange);


                            }
                        }
                        finally
                        {
                            PageFormatCache.DestroyPageCache(); //Remember to reset it so it is not used by any other method.
                        }

                    }

                    if ((FWorkbook.PrintOptions & TPrintOptions.LeftToRight) != 0)
                    {
                        PagePrintRange.Left = PagePrintRange.Right + 1;
                        if (PagePrintRange.Left > PrintRange.Right)
                        {
                            PagePrintRange.Left = PrintRange.Left;
                            PagePrintRange.Top = PagePrintRange.Bottom + 1;
                        }
                    }
                    else
                    {
                        PagePrintRange.Top = PagePrintRange.Bottom + 1;
                        if (PagePrintRange.Top > PrintRange.Bottom)
                        {
                            PagePrintRange.Top = PrintRange.Top;
                            PagePrintRange.Left = PagePrintRange.Right + 1;
                        }
                    }

                    if (PagePrintRange.Top > PrintRange.Bottom || PagePrintRange.Left > PrintRange.Right)
                    {
                        PrintArea++;
                        if (PrintArea < PrintRanges.Length)
                        {
                            PagePrintRange.Top = PrintRanges[PrintArea].Top;
                            PagePrintRange.Left = PrintRanges[PrintArea].Left;
                        }
                    }
                }
                finally
                {
                    if (Canvas != null) Canvas.DestroySFormat();
                }
            }
            finally
            {
                Workbook.IgnoreFormulaText = IgnoreFormulaText;
                Workbook.FormulaReferenceStyle = SaveRefStyle;
            }
        }
        #endregion

        #region DrawHeadings
        private void DrawHeadings(RectangleF PaintClipRect, RectangleF PaintClipRect2, real WOfs, real HOfs, TXlsCellRange PagePrintRange, TRepeatingRange RepeatingRange)
        {
            bool HasHeaders = FWorkbook.PrintHeadings;
            TFlxFont fx = FWorkbook.GetDefaultFontNormalStyle;
            Font AFont = FontCache.GetFont(fx, Zoom100);
            using (Pen APen = new Pen(Colors.Black))
            {
                APen.Width = 1 / 100F * DispMul;
                real dw = HasHeaders? HeaderColWidth() : 0;
                dw += PaintClipRect.Left + 1 * DispMul / 100f + WOfs;

                for (int c = PagePrintRange.Left - RepeatingRange.ColCount(PagePrintRange.Left); c <= PagePrintRange.Right; c++)
                {
                    int col = c;
                    if (c < PagePrintRange.Left) col = RepeatingRange.FirstCol + (c - PagePrintRange.Left + RepeatingRange.ColCount(PagePrintRange.Left));

                    real dw1 = dw + RealColWidth(col);

                    if (HasHeaders)
                    {
                        Canvas.SetClipReplace(RectangleXY(dw - 1 * DispMul / 100f, PaintClipRect.Top + HOfs, dw1 + 1 * DispMul / 100f, PaintClipRect.Top + HOfs + HeaderRowHeight() + 1 * DispMul / 100f));
                        Canvas.DrawLine(APen, dw, PaintClipRect.Top + HOfs, dw, PaintClipRect.Top + HOfs + HeaderRowHeight() + 1 * DispMul / 100f);

                        if (dw1 - dw > 0.1 * DispMul / 100f)
                        {
                            Canvas.SetClipReplace(RectangleXY(dw, PaintClipRect.Top + HOfs, dw1, PaintClipRect.Top + HOfs + HeaderRowHeight() + 1 * DispMul / 100f));
                            string st = Workbook.OptionsR1C1? col.ToString(CultureInfo.CurrentCulture): TCellAddress.EncodeColumn(col);
                            real ws = Canvas.MeasureString(st, AFont, new TPointF(0, 0)).Width;
                            Canvas.DrawString(st, AFont, Brushes.Black,
                                (dw + dw1 - ws) / 2, PaintClipRect.Top + HOfs + HeaderRowHeight());
                        }
                    }
                    dw = dw1;
                }

                real dh = HasHeaders ? HeaderRowHeight() : 0;
                dh += PaintClipRect.Top + 1 * DispMul / 100f + HOfs;
                for (int r = PagePrintRange.Top - RepeatingRange.RowCount(PagePrintRange.Top); r <= PagePrintRange.Bottom; r++)
                {
                    int row = r;
                    if (r < PagePrintRange.Top) row = RepeatingRange.FirstRow + (r - PagePrintRange.Top + RepeatingRange.RowCount(PagePrintRange.Top));

                    real dh1 = dh + RealRowHeight(row);

                    if (HasHeaders)
                    {
                        Canvas.SetClipReplace(RectangleXY(PaintClipRect.Left + WOfs, dh - 1 * DispMul / 100f, PaintClipRect.Left + WOfs + HeaderColWidth() + 1 * DispMul / 100f, dh1 + 1 * DispMul / 100f));
                        Canvas.DrawLine(APen, PaintClipRect.Left + WOfs, dh, PaintClipRect.Left + WOfs + HeaderColWidth() + 1 * DispMul / 100f, dh);

                        if (dh1 - dh > 0.1 * DispMul / 100f)
                        {
                            Canvas.SetClipReplace(RectangleXY(PaintClipRect.Left + WOfs, dh, PaintClipRect.Left + WOfs + HeaderColWidth() + 1 * DispMul / 100f, dh1));
                            string st = row.ToString(CultureInfo.CurrentCulture);
                            real ws = Canvas.MeasureString(st, AFont, new TPointF(0, 0)).Width;
                            Canvas.DrawString(row.ToString(CultureInfo.CurrentCulture), AFont, Brushes.Black,
                                PaintClipRect.Left + WOfs + (HeaderColWidth() - ws) / 2, dh1);
                        }
                    }
                    dh = dh1;
                }

                if (HasHeaders || FWorkbook.PrintGridLines)
                {
                    Canvas.SetClipReplace(PaintClipRect2);
                }

                if (HasHeaders)
                {
                    Canvas.DrawLine(APen, PaintClipRect.Left + WOfs + HeaderColWidth() + 1 * DispMul / 100f, PaintClipRect.Top + HOfs, PaintClipRect.Left + HeaderColWidth() + 1 * DispMul / 100f + WOfs, dh);
                    Canvas.DrawLine(APen, PaintClipRect.Left + WOfs, PaintClipRect.Top + HOfs + HeaderRowHeight() + 1 * DispMul / 100f, dw, PaintClipRect.Top + HeaderRowHeight() + 1 * DispMul / 100f + HOfs);
                }

                if (HasHeaders || FWorkbook.PrintGridLines)
                {
                    Canvas.DrawRectangle(APen, PaintClipRect.Left + WOfs, PaintClipRect.Top + HOfs, dw - PaintClipRect.Left - WOfs, dh - PaintClipRect.Top - HOfs);
                }
            }
        }

        #endregion

        #region  DrawPage
        protected void DrawPage(RectangleF PaintClipRect, TXlsCellRange PagePrintRange, bool FirstPass, TSpawnedCellList SpawnedCells, TDrawObjects DrawObjects, bool FirstRow, bool FirstCol)
        {
            RectangleF PaintClipRect1 = RectangleXY(
                PaintClipRect.Left - 1* DispMul / 100f,
                PaintClipRect.Top - 1* DispMul / 100f,
                PaintClipRect.Right + 1* DispMul / 100f,
                PaintClipRect.Bottom + 1* DispMul / 100f);

            RectangleF BorderRect = new RectangleF(PaintClipRect1.Left, PaintClipRect1.Top,
                Math.Min(PaintClipRect1.Width, CalcAcumColWidth(PagePrintRange.Left, PagePrintRange.Right + 1)),
                Math.Min(PaintClipRect1.Height, CalcAcumRowHeight(PagePrintRange.Top, PagePrintRange.Bottom + 1)));

            bool EmptyPage = (PagePrintRange.RowCount <= 1 && PagePrintRange.ColCount <= 1 && IsEmptyCell(PagePrintRange.Top, PagePrintRange.Left));
            Canvas.SetClipReplace(PaintClipRect1);
			RectangleF NewClipRect = PaintClipRect;
            if (!EmptyPage)
            {
                if (FirstPass)
                {
                    DrawBackground(PagePrintRange, PaintClipRect);
                    DrawCells(ref PaintClipRect, PagePrintRange, SpawnedCells, ref BorderRect);
                }
                else
                {
                    real FinalH = 0; real FinalW = 0;

                    Canvas.SetClipReplace(RectangleF.Inflate(PaintClipRect1, 2f* DispMul / 100f, 2f* DispMul / 100f));  //Add room for double lines. (they go outside the margins)
                    DrawLines(PagePrintRange, PaintClipRect, SpawnedCells, out FinalH, out FinalW, FirstRow, FirstCol);

					if (FDrawGridLines || FWorkbook.PrintHeadings)
					{
						NewClipRect = new RectangleF(PaintClipRect.Left, PaintClipRect.Top, FinalW, FinalH);
						Canvas.SetClipReplace(NewClipRect);
					}
                }
            }
            if (!FirstPass)
            {
                DrawAllImages(PagePrintRange, NewClipRect, DrawObjects);
                DrawHyperlinks(PagePrintRange, NewClipRect);
                DrawComments(PagePrintRange, NewClipRect);
            }
        }

		private void DrawAllImages(TXlsCellRange PagePrintRange, RectangleF PaintClipRect, TDrawObjects DrawObjects)
		{
			int ObjCount = FWorkbook.ObjectCount + 1;
			if (ObjCount > 1)
			{
				TShapePropertiesList ShapesInPage = null;
				if (ShapesCache != null)
				{
					ShapesInPage = ShapesCache[CurrentPrintAreaRegion].GetShapes(PagePrintRange.Top, PagePrintRange.Left, PagePrintRange.Bottom, PagePrintRange.Right);
				}
				else
				{
					ShapesInPage = new TShapePropertiesList();
					for (int i = 1; i <= FWorkbook.ObjectCount; i++)
					{
						TShapeProperties sp = FWorkbook.GetObjectProperties(i, true);
						sp.zOrder = i;
						ShapesInPage.Add(sp);
					}
				}
				DrawObjects.DrawImages(ShapesInPage, PagePrintRange, PaintClipRect, false);
			}
		}

        #endregion

        #region  DrawCells
        private void DrawCells(ref RectangleF PaintClipRect, TXlsCellRange PagePrintRange, TSpawnedCellList SpawnedCells, ref RectangleF BorderRect)
        {
            DrawOutsideCells(ref PaintClipRect, PagePrintRange, SpawnedCells, ref BorderRect);
            DrawInsideCells(ref PaintClipRect, PagePrintRange, SpawnedCells, ref BorderRect);
        }

        private void DrawOutsideCells(ref RectangleF PaintClipRect, TXlsCellRange PagePrintRange, TSpawnedCellList SpawnedCells, ref RectangleF BorderRect)
        {
            //Draw Merged and Spawned Cells that might come from the left.
            int Row = PagePrintRange.Top;
            int Col = PagePrintRange.Left;
            if (Col>1)
            {
                while (Row <= PagePrintRange.Bottom)
                {
                    bool Merged;
                    TXlsCellRange xl = XCellMergedBounds(Row, Col, out Merged);
                    if (xl != null && (xl.Top!=xl.Bottom || xl.Left!=xl.Right)) 
                    {
                        if(xl.Left < Col || xl.Top<Row)
                        {
                            RectangleF ARect = RectangleCell(xl.Top, xl.Left, ref PaintClipRect, ref PagePrintRange);
                            DrawCell(xl.Left, xl.Top, ARect, BorderRect, SpawnedCells, true, false, TSpanDirection.Left);
                        }
                        Row = xl.Bottom;
                    }
                    else
                        if (IsEmptyCell(Row, Col))
                    {
                        RectangleF ARect = RectangleCell(Row, Col, ref PaintClipRect, ref PagePrintRange);
                        DrawCell(Col, Row, ARect, BorderRect, SpawnedCells, true, true, TSpanDirection.Left);
                    }

                    Row++;
                }
            }

            //Draw Spawned Cells that might come from the right (Merged cells won't come from right).
            Row = PagePrintRange.Top;
            Col = PagePrintRange.Right;
            while (Row <= PagePrintRange.Bottom)
            {
                bool Merged;
                TXlsCellRange xl = XCellMergedBounds(Row, Col, out Merged);
                if (xl != null && (xl.Top!=xl.Bottom || xl.Left!=xl.Right)) 
                {
                    //No draw here. Merged cells will not come from the right.
                    Row = xl.Bottom;
                }
                else
                    if (IsEmptyCell(Row, Col))
                {
                    RectangleF ARect = RectangleCell(Row, Col, ref PaintClipRect, ref PagePrintRange);
                    DrawCell(Col, Row, ARect, BorderRect, SpawnedCells, true, true, TSpanDirection.Right);
                }

                Row++;
            }
            

            //Draw Merged Cells that might span from previous row on other page. Not Merged cells won't span) 
            Row = PagePrintRange.Top;
            Col = PagePrintRange.Left;
            if (Row>1)
            {
                while (Col <= PagePrintRange.Right)
                {
                    bool Merged;
                    TXlsCellRange xl = XCellMergedBounds(Row, Col, out Merged);
                    if (xl != null && xl.Top < Row)
                    {
                        if (xl.Left>=PagePrintRange.Left)  //If not this has been drawn when drawing cells coming from the left.
                        {
                            RectangleF ARect = RectangleCell(xl.Top, xl.Left, ref PaintClipRect, ref PagePrintRange);
                            DrawCell(xl.Left, xl.Top, ARect, BorderRect, SpawnedCells, true, false, TSpanDirection.Both);
                        }
                        Col = xl.Right;
                    }
                    Col++;
                }
            }
        }

        private void DrawInsideCells(ref RectangleF PaintClipRect, TXlsCellRange PagePrintRange, TSpawnedCellList SpawnedCells, ref RectangleF BorderRect)
        {
            real Ch = PaintClipRect.Top;
            int Row = PagePrintRange.Top;
			
            while (Row <= PagePrintRange.Bottom)
            {
                real Cw = PaintClipRect.Left;
                //Slow but trusted way. It will draw each cell and "clip" the area so it is not
                //drawn outside the cell.  For Merged and continued cells, makes much more sense to 
                //draw the string only once, and clip to the whole Merged/continued cell.
                /*
                                                        for (int Col=PagePrintRange.Left;Col<=PagePrintRange.Right;Col++)
                                                        {
                                                            RectangleF ARect= RectangleXY(Cw, Ch, (Cw+RealColWidth(Col)),(Ch+RealRowHeight(Row)));
                                                            DrawCell(Col, Row, Col, PagePrintRange.Left, ARect, ARect, TGridDrawState.Normal, true, false, false, Col, Row, PaintClipRect1, SpawnedCells, true);
                                                            Cw+=RealColWidth(Col);
                                                        }
                            */

                //Draw normal columns
                DrawColumns(PagePrintRange, SpawnedCells, ref BorderRect, Row, Ch, ref Cw);

                Ch += RealRowHeight(Row);
                Row++;
            }
        }

        private void DrawColumns(TXlsCellRange PagePrintRange, TSpawnedCellList 
            SpawnedCells, ref RectangleF BorderRect, int Row, real Ch, ref real Cw)
        {
            int MinColIndex = Workbook.ColToIndex(Row, PagePrintRange.Left);
            int MaxColIndex = Workbook.ColToIndex(Row, PagePrintRange.Right);

            int LastCol = PagePrintRange.Left;

            for (int ci = MinColIndex; ci <= MaxColIndex; ci++)
            {
                int Col = Workbook.ColFromIndex(Row, ci);
                if (Col > PagePrintRange.Right || Col < PagePrintRange.Left) continue;

                for (int i = LastCol; i < Col; i++)
                    Cw += RealColWidth(i);
                LastCol = Col;

                if (Col < PagePrintRange.Left || Col > PagePrintRange.Right) continue;
                if (IsEmptyCell(Row, Col)) continue;  //Merged cells or hidden columns might return indexes, but the cell is empty.
                RectangleF ARect = RectangleXY(Cw, Ch, (Cw + RealColWidth(Col)), (Ch + RealRowHeight(Row)));
                DrawCell(Col, Row, ARect, BorderRect, SpawnedCells, true, false, TSpanDirection.Both);

            }

            for (int i = LastCol; i <= PagePrintRange.Right; i++)
                Cw += RealColWidth(i);
        }

        #endregion

        #region Draw Lines

        private static DashStyle CalcDashStyle(TFlxBorderStyle aBorder)
        {
            switch (aBorder)
            {
                case TFlxBorderStyle.Dash_dot:  
                case TFlxBorderStyle.Medium_dash_dot:
                case TFlxBorderStyle.Slanted_dash_dot:
                    return DashStyles.DashDot;

                case TFlxBorderStyle.Dash_dot_dot:
                case TFlxBorderStyle.Medium_dash_dot_dot:
                    return DashStyles.DashDotDot;

                case TFlxBorderStyle.Dashed:
                case TFlxBorderStyle.Medium_dashed:
                    return DashStyles.Dash;

                case TFlxBorderStyle.Dotted:
                    return DashStyles.Dot;
            }
            return DashStyles.Solid;
        }


        private static TFlxOneBorder GetCorner(TFlxBorders Border, TCorner Corner)
        {
            switch (Corner)
            {
                case TCorner.Left:
                    return Border.Left;
                case TCorner.Top:
                    return Border.Top;
                case TCorner.Right:
                    return Border.Right;
                case TCorner.Bottom:
                    return Border.Bottom;
                default:
                    FlxMessages.ThrowException(FlxErr.ErrInternal);
                    return Border.Left; //just to compile
            }
        }

		private static int BorderWeight(TFlxBorderStyle b)
		{
            switch (b)
            {
                case TFlxBorderStyle.None:
                    return 0;
                case TFlxBorderStyle.Hair:
                    return 1;
                case TFlxBorderStyle.Thin:
                    return 2;
                case TFlxBorderStyle.Dotted:
                    return 3;
                case TFlxBorderStyle.Dashed:
                    return 4;
                case TFlxBorderStyle.Dash_dot:
                    return 5;
                case TFlxBorderStyle.Dash_dot_dot:
                    return 6;
                case TFlxBorderStyle.Medium:
                    return 7;
                case TFlxBorderStyle.Medium_dashed:
                    return 8;
                case TFlxBorderStyle.Medium_dash_dot:
                    return 9;
                case TFlxBorderStyle.Medium_dash_dot_dot:
                    return 10;
                case TFlxBorderStyle.Slanted_dash_dot:
                    return 11;

                case TFlxBorderStyle.Thick:
                    return 12;
                case TFlxBorderStyle.Double:
                    return 13;
            }

            return 0;
            
		}

        private TFlxOneBorder GetOneBorder(TFlxFormat fmt, int r1, int c1, TXlsCellRange PagePrintRange, TCorner Corner1, TCorner Corner2)
		{
			if (fmt == null)
			{
                if (r1 < PagePrintRange.Top || r1 > PagePrintRange.Bottom || c1 < PagePrintRange.Left || c1 > PagePrintRange.Right)
                    return new TFlxOneBorder(TFlxBorderStyle.None, TExcelColor.Automatic);
				TFlxFormat fmt3 = GetCellVisibleFormatDef(r1, c1);
				return GetCorner(fmt3.Borders, Corner2);
			}

			TFlxOneBorder Result1 = GetCorner(fmt.Borders, Corner1);

            if (r1 < PagePrintRange.Top || r1 > PagePrintRange.Bottom || c1 < PagePrintRange.Left || c1 > PagePrintRange.Right) return Result1;
			TFlxFormat fmt2 = GetCellVisibleFormatDef(r1, c1);
			TFlxOneBorder Result2 = GetCorner(fmt2.Borders, Corner2);
			if (BorderWeight(Result2.Style) > BorderWeight(Result1.Style)) return Result2; else return Result1;
		}

        private TFlxBorders GetBorder(int r, int c, TXlsCellRange PagePrintRange)
		{
			TFlxFormat fmt;
			return GetBorder(r, c, PagePrintRange, out fmt);
		}

		private TFlxBorders GetBorder(int r, int c, TXlsCellRange PagePrintRange, out TFlxFormat fmt)
		{
			TFlxBorders Result = new TFlxBorders();
			if (r < PagePrintRange.Top || r > PagePrintRange.Bottom || c < PagePrintRange.Left || c > PagePrintRange.Right) 
				fmt = null;
			 else fmt = GetCellVisibleFormatDef(r, c);

			Result.Top = GetOneBorder(fmt, r-1, c, PagePrintRange, TCorner.Top, TCorner.Bottom);
			Result.Left = GetOneBorder(fmt, r, c-1, PagePrintRange, TCorner.Left, TCorner.Right);
			Result.Bottom = GetOneBorder(fmt, r+1, c, PagePrintRange, TCorner.Bottom, TCorner.Top);
			Result.Right = GetOneBorder(fmt, r, c+1, PagePrintRange, TCorner.Right, TCorner.Left);

            if (fmt != null)
            {
                Result.DiagonalStyle = fmt.Borders.DiagonalStyle;
                Result.Diagonal = fmt.Borders.Diagonal;
            }

            CheckMerged(r, c, Result);
			return Result;
		}

        private void CheckMerged(int r, int c, TFlxBorders Result)
        {
            TXlsCellRange MergedRange = null;
            bool Merged = false;
            if (r <= FlxConsts.Max_Rows + 1 && c <= FlxConsts.Max_Columns + 1) MergedRange = XCellMergedBounds(r, c, out Merged);

            if (MergedRange != null && Merged)
            {
                if (r > MergedRange.Top) Result.Top = new TFlxOneBorder(TFlxBorderStyle.None, TExcelColor.Automatic);
                if (r < MergedRange.Bottom) Result.Bottom = new TFlxOneBorder(TFlxBorderStyle.None, TExcelColor.Automatic);
                if (c > MergedRange.Left) Result.Left = new TFlxOneBorder(TFlxBorderStyle.None, TExcelColor.Automatic);
                if (c < MergedRange.Right) Result.Right = new TFlxOneBorder(TFlxBorderStyle.None, TExcelColor.Automatic);

            }
        }

		private real GetBorderWidth(TFlxBorders FullBorder, TCorner Corner)
		{

            TFlxOneBorder Border = GetCorner(FullBorder, Corner);
			const real f = DispMul / 200f;
			if (Border.Style == TFlxBorderStyle.None) return 0;
			if (Border.Style == TFlxBorderStyle.Double)
			{
				return CalcLineWidth(TFlxBorderStyle.Thin)* f - LineJoinAdj * Zoom100  + DoubleLineSep; 
			}
			return CalcLineWidth(Border.Style) * f - LineJoinAdj * Zoom100;
		}

        private real CalcLineWidth(TFlxBorderStyle aBorder)
        {
            switch (aBorder)
            {
				case TFlxBorderStyle.Hair:
				case TFlxBorderStyle.None:  //gridlines
					return (1f/6f) * Zoom100;

                case TFlxBorderStyle.Medium_dash_dot:
                case TFlxBorderStyle.Slanted_dash_dot:
                case TFlxBorderStyle.Medium_dash_dot_dot:
                case TFlxBorderStyle.Medium_dashed:
                case TFlxBorderStyle.Medium:
                    return 2*Zoom100;

                case TFlxBorderStyle.Thick:
                    return 3*Zoom100;
            }
            return 1*Zoom100;
        }

        private void SelectPen(Pen APen, Color aColor, TFlxBorderStyle aBorder)
        {
            if (APen.Color != aColor) APen.Color = aColor;

            DashStyle ds =CalcDashStyle(aBorder);
            if (APen.DashStyle != ds) APen.DashStyle = ds;
            real dw = CalcLineWidth(aBorder);
            APen.Width = dw * DispMul / 100f;
        }

        private void CalcCorners(int Row, int Col, int ofsr, int ofsc, TXlsCellRange PagePrintRange, TCorner Corner, TCorner CornerLeft, 
			TFlxBorderStyle aBorder, bool DontIntersect,
            out TFlxBorders FmRDown, out real wRDown, out TFlxBorders FmRUp, out real wRUp, out real wR)
        {
            FmRDown = GetBorder(Row, Col, PagePrintRange);
            wRDown = GetBorderWidth(FmRDown, Corner);
            FmRUp = null; 
            if (Row + ofsr >= 1 && Row + ofsr <= FlxConsts.Max_Rows + 1 &&
                Col + ofsc >= 1 && Col + ofsc <= FlxConsts.Max_Columns + 1) FmRUp = GetBorder(Row + ofsr, Col + ofsc, PagePrintRange);
            
            wRUp = wRDown;
            if (FmRUp != null) wRUp = GetBorderWidth(FmRUp, Corner);

            wR = Math.Max(wRUp, wRDown);

			if (DontIntersect)
			{
				wR = - wR - 2* LineJoinAdj * Zoom100;
				wRUp = -wRUp - 2 * LineJoinAdj * Zoom100;
				wRDown = -wRDown - 2* LineJoinAdj * Zoom100;
				return;
			}

			bool HasDown = GetCorner(FmRDown, Corner).Style == TFlxBorderStyle.Double;
			bool HasUp = FmRUp != null && GetCorner(FmRUp, Corner).Style == TFlxBorderStyle.Double;
			bool HasLeft = false;
			if ((aBorder != TFlxBorderStyle.Double) && (HasDown ^ HasUp))
			{
                HasLeft = HasLeftCorner(Row, Col, PagePrintRange, Corner, CornerLeft);
			}

            if ((HasDown && HasUp) || (HasDown && HasLeft) || (HasUp && HasLeft))
            {
				real sep = 2 * DoubleLineSep;
                wR -= sep; //if a double line crosses, we need to stop at the left border.
                wRUp -= sep;
                wRDown -= sep;
            }
        }

        private bool HasLeftCorner(int Row, int Col, TXlsCellRange PagePrintRange, TCorner Corner, TCorner CornerLeft)
        {
            int o2r = 0; int o2c = 0;
            switch (Corner)
            {
                case TCorner.Left:
                    o2c = -1;
                    break;
                case TCorner.Right:
                    o2c = 1;
                    break;
                case TCorner.Top:
                    o2r = -1;
                    break;
                case TCorner.Bottom:
                    o2r = 1;
                    break;
            }
            if (Row + o2r >= 1 && Row + o2r <= FlxConsts.Max_Rows + 1 &&
                Col + o2c >= 1 && Col + o2c <= FlxConsts.Max_Columns + 1)
            {
                TFlxBorders FmLeft = GetBorder(Row + o2r, Col + o2c, PagePrintRange);
                return GetCorner(FmLeft, CornerLeft).Style == TFlxBorderStyle.Double;
            }
            return false;
        }

        private static void VerifyDoubleCorner(TFlxBorders FmRDown, TFlxBorders FmRUp, TCorner Corner, ref real wLDown, ref real wLUp)
        {
			bool HasDown = GetCorner(FmRDown, Corner).Style == TFlxBorderStyle.Double;
			bool HasUp = FmRUp != null && GetCorner(FmRUp, Corner).Style == TFlxBorderStyle.Double;
			if (HasDown && !HasUp)
			{
				real sep = 2 * DoubleLineSep;
				wLUp = wLDown;
				wLDown -= sep;
			}
			else
			if (HasUp && !HasDown) //We need to close the double line
			{
				real sep = 2 * DoubleLineSep;
				wLDown = wLUp;
				wLUp -= sep;
			}
        }

        private void VerifyDoubleDiags(int r, int c, TFlxBorders FmLDown, TFlxDiagonalBorder DiagStyle, TCorner CornerV, TCorner CornerH, ref real wL)
        {
			if (FmLDown == null) return;
            if ((FmLDown.DiagonalStyle == TFlxDiagonalBorder.Both || FmLDown.DiagonalStyle == DiagStyle) && FmLDown.Diagonal.Style == TFlxBorderStyle.Double)
            {
                TXlsCellRange Mb = null;
                bool Merged = false;
                if (r <= FlxConsts.Max_Rows + 1 && c <= FlxConsts.Max_Columns + 1) Mb = XCellMergedBounds(r, c, out Merged);
                if (Mb != null && Merged)
                {
                    switch (CornerH)
                    {
                        case TCorner.Left:
                            if (c != Mb.Left) return;
							break;
                        case TCorner.Right:
                            if (c != Mb.Right) return;
							break;
                        default:
                            FlxMessages.ThrowException(FlxErr.ErrInternal);
                            break;
                    }

                    switch (CornerV)
                    {
                        case TCorner.Top:
                            if (r != Mb.Top) return;
							break;
                        case TCorner.Bottom:
                            if (r != Mb.Bottom) return;
							break;
                        default:
                            FlxMessages.ThrowException(FlxErr.ErrInternal);
                            break;
                    }
                }

                wL -= DoubleLineSepDiag;
            }
        }

        private void DrawLine (int StartRow, int StartCol, int EndRow, int EndCol, int ofs, TXlsCellRange PagePrintRange, TCorner CornerAdj,
            Pen APen,  Color aColor, TFlxBorderStyle aBorder, real x1, real y1, real x2, real y2, bool Horizontal, bool ISectL, bool ISectR)
        {
			SelectPen(APen, aColor, aBorder);

			if (Horizontal)
			{
                TFlxBorders FmRDown; real wRDown; TFlxBorders FmRUp; real wRUp; real wR;
                CalcCorners(EndRow, EndCol, ofs, 0, PagePrintRange, TCorner.Right, CornerAdj, aBorder, ISectR, out FmRDown, out wRDown, out FmRUp, out wRUp, out wR);

                TFlxBorders FmLDown; real wLDown; TFlxBorders FmLUp; real wLUp; real wL;
                CalcCorners(StartRow, StartCol, ofs, 0, PagePrintRange, TCorner.Left, CornerAdj, aBorder, ISectL, out FmLDown, out wLDown, out FmLUp, out wLUp, out wL);

				if (aBorder == TFlxBorderStyle.Double)
				{
                    VerifyDoubleCorner(FmLDown, FmLUp, TCorner.Left, ref wLDown, ref wLUp);
                    VerifyDoubleCorner(FmRDown, FmRUp, TCorner.Right, ref wRDown, ref wRUp);

                    //EndRow & StartRow are the same.
                    VerifyDoubleDiags(EndRow, StartCol, FmLDown, TFlxDiagonalBorder.DiagDown, TCorner.Top, TCorner.Left, ref wLDown);
                    VerifyDoubleDiags(EndRow, EndCol, FmRDown, TFlxDiagonalBorder.DiagUp, TCorner.Top, TCorner.Right, ref wRDown);
                    VerifyDoubleDiags(EndRow - 1, StartCol, FmLUp, TFlxDiagonalBorder.DiagUp, TCorner.Bottom, TCorner.Left, ref wLUp);
                    VerifyDoubleDiags(EndRow - 1, EndCol, FmRUp, TFlxDiagonalBorder.DiagDown, TCorner.Bottom, TCorner.Right, ref wRUp);
                    
                    Canvas.DrawLine(APen, x1 - wLUp, y1 + ofs * DoubleLineSep, x2 + wRUp, y2 + ofs * DoubleLineSep);
					Canvas.DrawLine(APen, x1 - wLDown, y1-ofs*DoubleLineSep, x2 + wRDown, y2-ofs*DoubleLineSep);
				}
				else
				{
					Canvas.DrawLine(APen, x1 - wL, y1, x2 + wR, y2);
				}
			}
			
			else  //VERTICAL
			{
                TFlxBorders FmRDown; real wRDown; TFlxBorders FmRUp; real wRUp; real wR;
                CalcCorners(EndRow, EndCol, 0, ofs, PagePrintRange, TCorner.Bottom, CornerAdj, aBorder, ISectR, out FmRDown, out wRDown, out FmRUp, out wRUp, out wR);

                TFlxBorders FmLDown; real wLDown; TFlxBorders FmLUp; real wLUp; real wL;
                CalcCorners(StartRow, StartCol, 0, ofs, PagePrintRange, TCorner.Top, CornerAdj, aBorder, ISectL, out FmLDown, out wLDown, out FmLUp, out wLUp, out wL);

				if (aBorder == TFlxBorderStyle.Double)
				{
                    VerifyDoubleCorner(FmLDown, FmLUp, TCorner.Top, ref wLDown, ref wLUp);
                    VerifyDoubleCorner(FmRDown, FmRUp, TCorner.Bottom, ref wRDown, ref wRUp);

                    //EndCol & StartCol are the same.
                    VerifyDoubleDiags(StartRow, EndCol, FmLDown, TFlxDiagonalBorder.DiagDown, TCorner.Top, TCorner.Left, ref wLDown);
                    VerifyDoubleDiags(EndRow, EndCol, FmRDown, TFlxDiagonalBorder.DiagUp, TCorner.Bottom, TCorner.Left, ref wRDown);
                    VerifyDoubleDiags(StartRow, EndCol - 1, FmLUp, TFlxDiagonalBorder.DiagUp, TCorner.Top, TCorner.Right, ref wLUp);
                    VerifyDoubleDiags(EndRow, EndCol - 1, FmRUp, TFlxDiagonalBorder.DiagDown, TCorner.Bottom, TCorner.Right, ref wRUp);
                    
                    Canvas.DrawLine(APen, x1 - ofs * DoubleLineSep, y1 - wLDown, x2 - ofs * DoubleLineSep, y2 + wRDown);
					Canvas.DrawLine(APen, x1 + ofs * DoubleLineSep, y1 - wLUp, x2 + ofs*DoubleLineSep, y2 + wRUp);
				}
				else
				{
					Canvas.DrawLine(APen, x1, y1 - wL, x2, y2 + wR);
				}
				return;
			}
        }


		private void CalcDiagCoords(TXlsCellRange PagePrintRange, int RangeTop, int RangeLeft, int RangeBottom, int RangeRight, bool CalcDiagUp, out real ox1, out real ox2, out real oy1, out real oy2)
		{
			TFlxBorders BordersTopLeft =GetBorder(RangeTop, RangeLeft, PagePrintRange);
			TFlxBorders BordersBottomRight = GetBorder(RangeBottom, RangeRight, PagePrintRange);

			ox1 = 0; ox2 = 0; oy1 = 0; oy2 = 0;
			if (BordersTopLeft != null)
			{
				if (CalcDiagUp)
				{
					if (BordersTopLeft.Bottom.Style == TFlxBorderStyle.Double) oy1 = DoubleLineSep;
				}
				else
				{
					if (BordersTopLeft.Top.Style == TFlxBorderStyle.Double) oy1 = DoubleLineSep;
				}
				if (BordersTopLeft.Left.Style == TFlxBorderStyle.Double) ox1 = DoubleLineSep;
			}
			if (BordersBottomRight != null)
			{
				if (CalcDiagUp)
				{
					if (BordersBottomRight.Top.Style == TFlxBorderStyle.Double) oy2 = DoubleLineSep;
				}
				else
				{
					if (BordersBottomRight.Bottom.Style == TFlxBorderStyle.Double) oy2 = DoubleLineSep;
				}
				if (BordersBottomRight.Right.Style == TFlxBorderStyle.Double) ox2 = DoubleLineSep;
			}

		}

        private void DrawDoubleDiagLine(Pen APen, TFlxBorders Borders, real x1, real y1, real x2, real y2, TXlsCellRange PagePrintRange, TXlsCellRange range)
        {
			switch (Borders.DiagonalStyle)
			{
				case TFlxDiagonalBorder.DiagDown:
				{
					real ox1 = 0; real ox2 = 0; real oy1 = 0; real oy2 = 0;
					CalcDiagCoords(PagePrintRange, range.Top, range.Left, range.Bottom, range.Right, false, out ox1, out ox2, out oy1, out oy2);
					Canvas.DrawLine(APen, x1 + DoubleLineSepDiag + ox1, y1 + oy1, x2 - ox2, y2 - DoubleLineSepDiag - oy2);
					Canvas.DrawLine(APen, x1 + ox1, y1 + DoubleLineSepDiag + oy1, x2 - DoubleLineSepDiag - ox2, y2 - oy2);
                    break;
                }

                case TFlxDiagonalBorder.Both:
				{
                    real ox1 = 0; real ox2 = 0; real oy1 = 0; real oy2 = 0;
                    CalcDiagCoords(PagePrintRange, range.Top, range.Left, range.Bottom, range.Right, false, out ox1, out ox2, out oy1, out oy2);

                    TPointF TopLeft1 = new TPointF(x1 + DoubleLineSepDiag + ox1, y1 + oy1);
                    TPointF BottomRight1 = new TPointF(x2 - ox2, y2 - DoubleLineSepDiag - oy2);
                    TPointF TopLeft2 = new TPointF(x1 + ox1, y1 + DoubleLineSepDiag + oy1);
                    TPointF BottomRight2 = new TPointF(x2 - DoubleLineSepDiag - ox2, y2 - oy2);

                    CalcDiagCoords(PagePrintRange, range.Bottom, range.Left, range.Top, range.Right, true, out ox1, out ox2, out oy1, out oy2);

                    TPointF BottomLeft1 = new TPointF(x1 + DoubleLineSepDiag + ox1, y2 - oy1);
                    TPointF TopRight1 = new TPointF(x2 - ox2, y1 + DoubleLineSepDiag + oy2);
                    TPointF BottomLeft2 = new TPointF(x1 + ox1, y2 - DoubleLineSepDiag - oy1);
                    TPointF TopRight2 = new TPointF(x2 - DoubleLineSepDiag - ox2, y1 + oy2);

                    TPointF Middle = new TPointF((x1 + x2) / 2, (y1 + y2) / 2);
					real xofs = 0; real yofs = 0; real tgAlpha = 0;
					if (x2 - x1 > 0) tgAlpha = (y2 - y1) / (x2 - x1);
					if (tgAlpha != 0)
					{
						xofs = DoubleLineSepDiag * (1 + 1/tgAlpha) / 2;
						yofs = DoubleLineSepDiag * (1 + tgAlpha) / 2;
					}


					Canvas.DrawLines(APen, new TPointF[]{TopLeft1,  new TPointF(Middle.X, Middle.Y - yofs), TopRight2});
					Canvas.DrawLines(APen, new TPointF[]{TopLeft2,  new TPointF(Middle.X - xofs, Middle.Y), BottomLeft2});
					Canvas.DrawLines(APen, new TPointF[]{BottomLeft1,  new TPointF(Middle.X, Middle.Y + yofs), BottomRight2});
					Canvas.DrawLines(APen, new TPointF[]{BottomRight1,  new TPointF(Middle.X + xofs, Middle.Y), TopRight1});
					break;
				}

				case TFlxDiagonalBorder.DiagUp:
				{
					real ox1 = 0; real ox2 = 0; real oy1 = 0; real oy2 = 0;
					CalcDiagCoords(PagePrintRange, range.Bottom, range.Left, range.Top, range.Right, true, out ox1, out ox2, out oy1, out oy2);
					Canvas.DrawLine(APen, x1 + DoubleLineSepDiag + ox1, y2 - oy1, x2 - ox2, y1 + DoubleLineSepDiag + oy2);
					Canvas.DrawLine(APen, x1 + ox1, y2 - DoubleLineSepDiag - oy1, x2 - DoubleLineSepDiag - ox2, y1 + oy2);
					break;
                }

			}

		}

        private void DrawDiagonals(Pen APen, TFlxBorders Borders, real x1, real y1, int StartRow, int StartCol, TXlsCellRange PagePrintRange, TXlsCellRange MergedRange)
        {
            if (Borders == null || Borders.DiagonalStyle == TFlxDiagonalBorder.None) return;//This will be the most usual case

            if (StartRow != MergedRange.Top || StartCol != MergedRange.Left) return; //Only will draw in the main cell.

            float x2 = x1 + CalcAcumColWidth(MergedRange.Left, MergedRange.Right + 1);
            float y2 = y1 + CalcAcumRowHeight(MergedRange.Top, MergedRange.Bottom + 1);

            SelectPen(APen, GetFgColor(Borders.Diagonal.Color), Borders.Diagonal.Style);

			if (Borders.Diagonal.Style != TFlxBorderStyle.Double)
			{
				switch (Borders.DiagonalStyle)
				{
					case TFlxDiagonalBorder.DiagDown:
						Canvas.DrawLine(APen, x1, y1, x2, y2);
						break;
					case TFlxDiagonalBorder.Both:
						Canvas.DrawLine(APen, x1, y1, x2, y2);
						Canvas.DrawLine(APen, x1, y2, x2, y1);
						break;
					case TFlxDiagonalBorder.DiagUp:
						Canvas.DrawLine(APen, x1, y2, x2, y1);
						break;

				}
			}
			else
			{
				DrawDoubleDiagLine(APen, Borders, x1, y1, x2, y2, PagePrintRange, MergedRange);
			}


        }


        private static bool NeedsHTopBreak(TFlxBorders aBorders, TFlxBorders BordersUpLeft, bool DoLeft, ref bool LastWasDiagUp, out bool ISect)
        {
			ISect = false;
			bool aLastWasDiagUp = LastWasDiagUp;
            LastWasDiagUp = aBorders.DiagonalStyle != TFlxDiagonalBorder.None && aBorders.DiagonalStyle != TFlxDiagonalBorder.DiagDown && aBorders.Diagonal.Style == TFlxBorderStyle.Double; 
            if (aLastWasDiagUp) return true;

			//aBorders are the borders of the cell in the same row and one col to the right.
			if (aBorders.Left.Style == TFlxBorderStyle.Double) return true;

			if (BordersUpLeft.Right.Style == TFlxBorderStyle.Double) return true;

			if (aBorders.DiagonalStyle != TFlxDiagonalBorder.None && aBorders.DiagonalStyle != TFlxDiagonalBorder.DiagUp && aBorders.Diagonal.Style == TFlxBorderStyle.Double) 
			{
				return true;
			}
			
			int CurrentWeight = DoLeft? BorderWeight(BordersUpLeft.Bottom.Style): BorderWeight(aBorders.Top.Style);
			ISect = BorderWeight(aBorders.Left.Style) > CurrentWeight
				|| BorderWeight(BordersUpLeft.Right.Style) > CurrentWeight;

			return ISect;
        }

		private static bool NeedsVLeftBreak(TFlxBorders aBorders, TFlxBorders BordersUpLeft, bool DoTop, ref bool LastWasDiagUp, out bool ISect)
		{
			ISect = false;
            bool aLastWasDiagUp = LastWasDiagUp;
            LastWasDiagUp = aBorders.DiagonalStyle != TFlxDiagonalBorder.None && aBorders.DiagonalStyle != TFlxDiagonalBorder.DiagDown && aBorders.Diagonal.Style == TFlxBorderStyle.Double;

            //aBorders are the borders of the cell in the same row and one col to the right.
            if (aLastWasDiagUp) return true;
            if (aBorders.Top.Style == TFlxBorderStyle.Double) return true;

			if (BordersUpLeft.Bottom.Style == TFlxBorderStyle.Double) return true;

            if (aBorders.DiagonalStyle != TFlxDiagonalBorder.None && aBorders.DiagonalStyle != TFlxDiagonalBorder.DiagUp && aBorders.Diagonal.Style == TFlxBorderStyle.Double)
            {
                return true;
            }
			
			int CurrentWeight = DoTop? BorderWeight(BordersUpLeft.Right.Style): BorderWeight(aBorders.Left.Style);
			ISect = BorderWeight(aBorders.Top.Style) > CurrentWeight
				|| BorderWeight(BordersUpLeft.Bottom.Style) > CurrentWeight;

			return ISect;
		}

        internal void DrawLines(TXlsCellRange PagePrintRange, RectangleF PaintClipRect, TSpawnedCellList SpawnedCells, out real FinalH, out real FinalW, bool FirstRow, bool FirstCol)
        {
            using (Pen APen = new Pen(Colors.Black))
            {
                DrawHorizontalLines(PagePrintRange, PaintClipRect, APen, FirstRow);

                real Ch, Cw;
                DrawVerticalLines(PagePrintRange, PaintClipRect, APen, SpawnedCells, FirstCol, out Ch, out Cw);
				FinalH = Ch - PaintClipRect.Top;
				FinalW = Cw - PaintClipRect.Left;
            }
        }

        private void DrawHorizontalLines(TXlsCellRange PagePrintRange, RectangleF PaintClipRect, Pen APen, bool FirstRow)
        {
            int Row = PagePrintRange.Top;
            int Col = 0;
            real Ch = PaintClipRect.Top;
            real Cw = PaintClipRect.Left;
			bool[] LastCellWasFilled = new bool[PagePrintRange.ColCount + 1];
			
			while (Row <= PagePrintRange.Bottom + 1)
            {
                Col = PagePrintRange.Left;
                Cw = PaintClipRect.Left;
                real TopCw = Cw; 
                int StartTopCol = Col; 
				bool StartISect = false;
                Color TopColor0 = Colors.White;
                TFlxBorderStyle TopBorder0 = TFlxBorderStyle.None; 
				bool DTop0 = false;
				int LastCol = Math.Min(PagePrintRange.Right, FlxConsts.Max_Columns + 1);
				bool LastWasDiagUp = false;
				while (Col <= LastCol + 1)
                {
                    //Calc Borders
                    TXlsCellRange Mb = null;
                    bool Merged;
                    if (Row <= FlxConsts.Max_Rows + 1 && Col <= FlxConsts.Max_Columns + 1) Mb = XCellMergedBounds(Row, Col, out Merged);

                    Color TopColor = GridLinesColor;
                    TFlxBorderStyle TopBorder = TFlxBorderStyle.Hair; //for gridlines
                    bool DTop = false;
                    bool NeedsTopBreak = false;
					bool ISect = false;

					TFlxBorders LastBorders = null;
					TFlxBorders LastBordersUpLeft = null;

                    if (Col <= LastCol)
                    {
						TFlxFormat Fm1;
                        TFlxBorders Borders = GetBorder(Row, Col, PagePrintRange, out Fm1);
						LastBorders = Borders;
						TFlxBorders BordersUpLeft = GetBorder(Row - 1, Col - 1, PagePrintRange);
						LastBordersUpLeft = BordersUpLeft;

                        DrawDiagonals(APen, Borders, Cw, Ch, Row, Col, PagePrintRange, Mb);

						NeedsTopBreak = NeedsHTopBreak(Borders, BordersUpLeft, true, ref LastWasDiagUp, out ISect);				

                        bool CellGridLines = FDrawGridLines;
						bool NewCellWasFilled = (Fm1 !=null && (int)Fm1.FillPattern.Pattern > 1);
						if (NewCellWasFilled || LastCellWasFilled[Col - PagePrintRange.Left])CellGridLines = false;  //We are NOT going to draw them now, on the color of the background.
						LastCellWasFilled[Col - PagePrintRange.Left] = NewCellWasFilled;

                        DTop = CellGridLines && (Row > PagePrintRange.Top && Row <= PagePrintRange.Bottom); 
						FirstRow = false;

                        TopBorder = Borders.Top.Style;
                        if (TopBorder != TFlxBorderStyle.None)
                        {
                            TopColor = GetFgColor(Borders.Top.Color);
                            DTop = true;
                        }
                    }
                    else
                    {
                        DTop = false; 
                    }

                    if (Mb != null && Mb.Top < Row) DTop = false;

                    bool DrawnTop = (TopColor != TopColor0) || (TopBorder != TopBorder0) || NeedsTopBreak;
                    if (DTop0 && (!DTop || DrawnTop))
                    {
                        DrawLine(Row, StartTopCol, Row, Col - 1, -1, PagePrintRange, TCorner.Top, APen, TopColor0, TopBorder0, TopCw, Ch, Cw, Ch, true, StartISect, ISect);
                    }

                    if (DTop && (!DTop0 || DrawnTop)) 
                    { 
                        TopCw = Cw; StartTopCol = Col;
                        bool DiagUp = false;
                        if (LastBorders != null) NeedsHTopBreak(LastBorders, LastBordersUpLeft, false, ref DiagUp, out StartISect);
                    }

                    DTop0 = DTop;
                    TopColor0 = TopColor;
					TopBorder0 = TopBorder;

                    if (Col <= PagePrintRange.Right) Cw += RealColWidth(Col);
                    Col++;
                }
                if (Row <= PagePrintRange.Bottom) Ch += RealRowHeight(Row);
                Row++;
            }
        }

        //this is similar to drawhorizontallines, but for performance reasons it is unlooped here.
        private void DrawVerticalLines(TXlsCellRange PagePrintRange, RectangleF PaintClipRect, Pen APen, TSpawnedCellList SpawnedCells, bool FirstCol, out real Ch, out real Cw)
        {
			int Col = PagePrintRange.Left;
			int Row = 0;
			Cw = PaintClipRect.Left;
			Ch = PaintClipRect.Top;
		    bool[] LastCellWasFilled = new bool[PagePrintRange.RowCount + 1];
			while (Col <= PagePrintRange.Right + 1)
			{
				Row = PagePrintRange.Top;
				Ch = PaintClipRect.Top;
				real LeftCh = Ch; 
				int StartLeftRow = Row;  
				bool StartISect = false;
				Color LeftColor0 = Colors.White;
				TFlxBorderStyle LeftBorder0 = TFlxBorderStyle.None; 
				bool DLeft0 = false;
				int LastRow = Math.Min(PagePrintRange.Bottom, FlxConsts.Max_Rows + 1);
                bool LastWasDiagUp = false;
                while (Row <= LastRow + 1)
				{
					//Calc Borders
					TXlsCellRange Mb = null;
					bool Merged;
					if (Col <= FlxConsts.Max_Columns + 1 && Row <= FlxConsts.Max_Rows + 1) Mb = XCellMergedBounds(Row, Col, out Merged);

					Color LeftColor = GridLinesColor;
					TFlxBorderStyle LeftBorder = TFlxBorderStyle.Hair; //for gridlines
					bool DLeft = false;
					bool NeedsLeftBreak = false; 
					bool ISect = false;

					TFlxBorders LastBorders = null;
					TFlxBorders LastBordersUpLeft = null;

					if (Row <= LastRow)
					{
						TFlxFormat Fm1;
						TFlxBorders Borders = GetBorder(Row, Col, PagePrintRange, out Fm1);
						LastBorders = Borders;
						TFlxBorders BordersUpLeft = GetBorder(Row - 1, Col - 1, PagePrintRange);
						LastBordersUpLeft = BordersUpLeft;
						NeedsLeftBreak = NeedsVLeftBreak(Borders, BordersUpLeft, true, ref LastWasDiagUp, out ISect);

						bool CellGridLines = FDrawGridLines;
						bool NewCellWasFilled = (Fm1 !=null && (int)Fm1.FillPattern.Pattern > 1);
						if (NewCellWasFilled || LastCellWasFilled[Row - PagePrintRange.Top])CellGridLines = false;  //We are NOT going to draw them now, on the color of the background.
						LastCellWasFilled[Row - PagePrintRange.Top] = NewCellWasFilled;

						DLeft = CellGridLines && (Col > PagePrintRange.Left && Col <= PagePrintRange.Right);
						FirstCol = false;

						LeftBorder = Borders.Left.Style;
						if (LeftBorder != TFlxBorderStyle.None)
						{
							LeftColor = GetFgColor(Borders.Left.Color);
							DLeft = true;
						}

						if ((Col > 1) && DLeft && SpawnedCells.ContainsKey(FlxHash.MakeHash(Row, Col - 1))) DLeft = false;
					}
					else
					{
						DLeft = false; 
					}

					if (Mb != null && Mb.Left < Col) DLeft = false;

					bool DrawnLeft = (LeftColor != LeftColor0) || (LeftBorder != LeftBorder0) || NeedsLeftBreak;
					if (DLeft0 && (!DLeft || DrawnLeft))
					{
						DrawLine(StartLeftRow, Col, Row - 1, Col, -1, PagePrintRange, TCorner.Left, APen, LeftColor0, LeftBorder0, Cw, LeftCh, Cw, Ch, false, StartISect, ISect);
					}

					if (DLeft && (!DLeft0 || DrawnLeft)) 
                    { 
                        LeftCh = Ch; StartLeftRow = Row;
                        bool DiagUp = false;
                        if (LastBorders != null) NeedsVLeftBreak(LastBorders, LastBordersUpLeft, false, ref DiagUp, out StartISect);
                    }

					DLeft0 = DLeft;
					LeftColor0 = LeftColor;
					LeftBorder0 = LeftBorder;

					if (Row <= PagePrintRange.Bottom) Ch += RealRowHeight(Row);
					Row++;
				}
				if (Col <= PagePrintRange.Right) Cw += RealColWidth(Col);
				Col++;
			}
        }

        #endregion

        #region Draw Headers / Footers

        public static void DrawHeaderImage(ExcelFile Workbook, IFlxGraphics Canvas, THeaderAndFooterKind Kind, THeaderAndFooterPos Section, bool ReallyDraw, real x, real y, ref real dx, ref real dy, real ZoomHead)
        {
            try
            {
                THeaderOrFooterImageProperties ImgProp = Workbook.GetHeaderOrFooterImageProperties(Kind, Section);
                if (ImgProp==null || ImgProp.Anchor == null) return;
                real imgw = ImgProp.Anchor.Width / 96F * DispMul * ZoomHead;
                real imgh = ImgProp.Anchor.Height / 96F * DispMul * ZoomHead;
                RectangleF Coords= new RectangleF(x+dx, y-imgh, imgw, imgh);
                dx+= imgw;
                if (imgh> dy) dy = imgh;

                TXlsImgType ImageType = TXlsImgType.Unknown;

                using (MemoryStream ImgData = new MemoryStream())
                {
                    Workbook.GetHeaderOrFooterImage(Kind, Section, ref ImageType, ImgData);
                    if (ImgData!=null && ReallyDraw)
                    {
                        DrawShape.DrawOneImage(Canvas, ImgProp.CropArea, ImgProp.TransparentColor, ImgProp.Brightness, ImgProp.Contrast,
                            ImgProp.Gamma, ColorUtil.Empty, ImgProp.BiLevel, ImgProp.Grayscale, Coords, ImgData);
                    }
                }
            }
            catch (FileNotFoundException ex)
            {
				if (FlexCelTrace.HasListeners) FlexCelTrace.Write(new TRenderErrorDrawingImageError(ex.Message));
				
				//This will be raised if vj# is not installed. it is not a serious thing, we will just not draw the picture.
            }

        }

        private void DoWriteHeader(THeaderAndFooterKind Kind, THeaderAndFooterPos Section, int CurrentPage, int TotalPages, string Text, bool ReallyWrite, real x, real y, ref real[] dx, ref real[] dy, ref real[] LineMaxDescent, THFlxAlignment HAlign, real ZoomHead)
        {
			const int CacheLines = 10;

            const string  ItalicStr = "Italic";  //Do not localize
            const string BoldStr = "Bold";  //Do not localize)
			TFlxFont fnt = new TFlxFont();
			fnt.Name = "Arial"; fnt.Size20 = 200;
            Font AFont = FontCache.GetFont(fnt, ZoomHead);
            try
            {
                Color FontColor = Colors.Black;

                string aText = Text + "&";
                string TagValue = String.Empty;
                int p = 0;
                bool U = false;
                bool dU = false;
                bool SubScript=false;
                bool SuperScript=false;
				bool It = false;
				bool Bold = false;
				bool Striked = false;

				if (!ReallyWrite) 
				{
					dx = new real[CacheLines]; 
					dy = new real[CacheLines]; //when ReallyDraw, dy is already calculated, we will use this.
					LineMaxDescent = new real[CacheLines];
				}

				real AcumY = 0;
				real linedx = 0;
				real linedy = 0;

				int o = 0;
				int line = 0;
                do
                {
					real x1 = x;
					if (HAlign == THFlxAlignment.right) x1 = x - dx[line];
					else if (HAlign == THFlxAlignment.center) x1 = x - dx[line]/2;

					int q = aText.IndexOfAny(new char[]{'&','\n'}, p + o);
                    if (q < 0) return; //might happen on an unterminated string, f.i.
                    o = 0;
                    TRichString CurrentText = new TRichString(TagValue + aText.Substring(p, q - p), new TRTFRun[0], FWorkbook);
                    real MaxDescent = 0;
					
                    Font AFont2 = AFont;
					if (SuperScript || SubScript)
					{
						TFlxFont fnt2 = (TFlxFont)fnt.Clone(); fnt2.Size20 = (int)Math.Round(fnt2.Size20 / 1.5f);
						AFont2 = FontCache.GetFont(fnt2, ZoomHead);
					}
					
                    SizeF Sz = CalcTextExtent(AFont2, CurrentText, out MaxDescent);
					if (LineMaxDescent[line] < MaxDescent) LineMaxDescent[line] = MaxDescent;
                    if (ReallyWrite)
                    {
                        real ddy =0;
                        if (SubScript) 
                        {
                            SizeF Sz2 = Canvas.MeasureString("Mg", AFont2);
                            ddy = Sz2.Height/2.5F;
                        }
                        if (SuperScript) 
                        {
                            SizeF Sz2 = Canvas.MeasureString("Mg", AFont2);
                            ddy = -Sz2.Height/1.5F;
                        }
						real Desc = LineMaxDescent[line] - MaxDescent;
                        TextPainter.WriteText(FWorkbook, Canvas, FontCache, ZoomHead, AFont2, FontColor, 
                            x1 + linedx, y + AcumY - Desc, ddy, TextPainter.GetVisualString(CurrentText, ReverseRightToLeftStrings), 0, 0, null);  //descent in writetext does not work when there is no rich text.
                    }
                    
					TagValue = String.Empty;

                    linedx += Sz.Width;
					if (!ReallyWrite) dx[line] = linedx;

                    if (Sz.Height > linedy) linedy = Sz.Height;
					if (!ReallyWrite) dy[line] = linedy;
                    p = q + 1;

					if (aText[p - 1] == '\n') //new line.
					{
						if (line + 1 < dy.Length) AcumY += (float)(dy[line + 1] * Workbook.Linespacing);
						linedx = 0;
						linedy = 0;
						line++;
						if (line >= dx.Length)
						{
							real[] tmp = new real[dx.Length + CacheLines];
							Array.Copy(dx, 0, tmp, 0, dx.Length);
							dx = tmp;

							tmp = new real[dy.Length + CacheLines];
							Array.Copy(dy, 0, tmp, 0, dy.Length);
							dy = tmp;

							tmp = new real[LineMaxDescent.Length + CacheLines];
							Array.Copy(LineMaxDescent, 0, tmp, 0, LineMaxDescent.Length);
							LineMaxDescent = tmp;
						}
						continue;
					}


                    if (p < aText.Length)
                    {
						if (aText[p] == 'U')
						{
							if (U)
							{
								fnt.Underline = TFlxUnderline.None;
								AFont = FontCache.GetFont(fnt, ZoomHead);
							}
							else
							{
								fnt.Underline = TFlxUnderline.Single;
							}
							AFont = FontCache.GetFont(fnt, ZoomHead);
							dU = false;
							U = !U;
							p++;
						}
						else
							if (aText[p] == 'E')
						{
							if (dU && !U)
							{
								fnt.Underline = TFlxUnderline.None;
							}
							else
							{
								fnt.Underline = TFlxUnderline.Double;
							}
							AFont = FontCache.GetFont(fnt, ZoomHead);

							U = false;
							dU = !dU;
							p++;
						}
						else
							if (aText[p] == 'S')
						{
							if (Striked)
							{
								fnt.Style &= ~TFlxFontStyles.StrikeOut;
							}
							else
							{
								fnt.Style |= TFlxFontStyles.StrikeOut;
							}
							AFont = FontCache.GetFont(fnt, ZoomHead);
							Striked = ! Striked;
							p++;
						}
						else
							if (aText[p] == 'B' || aText[p] == 'b')
						{
							if (Bold)
							{
								fnt.Style &= ~TFlxFontStyles.Bold;
							}
							else
							{
								fnt.Style |= TFlxFontStyles.Bold;
							}
							AFont = FontCache.GetFont(fnt, ZoomHead);
							Bold = !Bold;
							p++;
						}
						else
							if (aText[p] == 'I' || aText[p] == 'i')
						{
							if (It)
							{
								fnt.Style &= ~TFlxFontStyles.Italic;
							}
							else
							{
								fnt.Style |= TFlxFontStyles.Italic;
							}
							AFont = FontCache.GetFont(fnt, ZoomHead);
							It = !It;
							p++;
						}
						else
							if (aText[p] == 'Y')
						{
							SubScript =!SubScript;
							SuperScript = false;
							p++;
						}
						else
							if (aText[p] == 'X')
						{
							SuperScript =!SuperScript;
							SubScript = false;
							p++;
						}
						else
							if (aText[p] == 'G')
						{
							DrawHeaderImage(FWorkbook, Canvas, Kind, Section, ReallyWrite, x1, y+AcumY, ref linedx, ref linedy, ZoomHead);
							if (!ReallyWrite) {dx[line] = linedx; dy[line] = linedy;}
							p++;
						}
						else
							if (aText[p] == '"')
						{
							p++;
							q = aText.IndexOf('"', p);
							if (q >= 0)
							{
								string FontName = String.Empty;
								string FontSt = String.Empty;
								int r = aText.IndexOf(',', p, q - p);
								if (r >= 0)
								{
									FontName = aText.Substring(p, r - p);
									FontSt = aText.Substring(r+1, q-r-1);
								}
								else
									FontName = aText.Substring(p, q - p);

								if (FontSt.IndexOf(BoldStr)>=0)
								{
									Bold = true;
									fnt.Style |= TFlxFontStyles.Bold;
								}
								else
								{
									fnt.Style &= ~TFlxFontStyles.Bold;
									Bold = false;
								}

								if (FontSt.IndexOf(ItalicStr)>=0)
								{
									fnt.Style |= TFlxFontStyles.Italic;
									It = true;
								}
								else
								{
									fnt.Style &= ~TFlxFontStyles.Italic;
									It = false;
								}

								fnt.Name = FontName;
								AFont = FontCache.GetFont(fnt, ZoomHead);

								p = q + 1;
							}
						}
						else
							if (aText[p] >= '0' && aText[p] <= '9')
						{
							int FSize = 0;
							do
							{
								FSize = FSize * 10 + (int)aText[p] - (int)'0';
								p++;
							}
							while (p < aText.Length && aText[p] >= '0' && aText[p] <= '9');

							fnt.Size20 = FSize * 20;
							AFont = FontCache.GetFont(fnt, ZoomHead);
						}
						else
							if (aText[p] == '&') //double & means a simple one.
						{
							o = 1; //to skip next & from search
						}
						else
							if (aText[p] == 'A') //SheetName
						{
							p++; //to skip next from search
							TagValue = Workbook.SheetName;
						}
						else
							if (aText[p] == 'D') //Date
						{
							p++; //to skip next from search
								TagValue = DateTime.Now.Date.ToShortDateString();
						}
						else
							if (aText[p] == 'T') //Time
						{
							p++; //to skip next from search
								TagValue = DateTime.Now.ToShortTimeString();
						}
						else
							if (aText[p] == 'P') //Page Number
						{
							p++; //to skip next from search
                            int FirstPage = FWorkbook.PrintFirstPageNumber.HasValue?  FWorkbook.PrintFirstPageNumber.Value -1: 0;
							TagValue = (FirstPage + CurrentPage).ToString();
						}
						else
							if (aText[p] == 'N') //PageCount
						{
							p++; //to skip next from search
							TagValue = TotalPages.ToString();
						}
						else
							if (aText[p] == 'F') //FileName
						{
							p++; //to skip next from search
							TagValue = Path.GetFileName(FWorkbook.ActiveFileName);
						}

						else
							if (aText[p] == 'Z') //FullFileName
						{
							p++; //to skip next from search
							TagValue = Path.GetFullPath(FWorkbook.ActiveFileName);
						}
						else  //unknown code
							p++;
                    }
                }
                while (p < aText.Length);
            }
            finally
            {
                //AFont.Dispose();
            }
        }

		private static real First(real[] value, bool AllButFirst)
		{
			if (value == null || value.Length == 0) return 0;
			if (AllButFirst)
			{
				real Result = 0;
				for (int i= 1; i< value.Length; i++) Result+=value[i];
				return -Result;
			}
			return value[0];
		}

        protected void DrawHeaderOrFooter(THeaderAndFooterKind Kind, int BaseSection, RectangleF ClipRect, real XOfs, string Text, int CurrentPage, int TotalPages, real y, bool Footer, real ZoomHead, real rx)
        {
            string Left = String.Empty;
            string Center = String.Empty;
            string Right = String.Empty;
            FWorkbook.FillPageHeaderOrFooter(Text, ref Left, ref Center, ref Right);
            real x = ClipRect.Left + rx - XOfs; real[] dx = null; real[] dy = null; real[] MaxDescent = null;
            DoWriteHeader(Kind, (THeaderAndFooterPos)(BaseSection), CurrentPage, TotalPages, Left, false, 0, 0, ref dx, ref dy, ref MaxDescent, THFlxAlignment.left, ZoomHead);
            DoWriteHeader(Kind, (THeaderAndFooterPos)(BaseSection), CurrentPage, TotalPages, Left, true, x, y + First(dy, Footer), ref dx, ref dy, ref MaxDescent, THFlxAlignment.left, ZoomHead);

            DoWriteHeader(Kind, (THeaderAndFooterPos)(BaseSection + 1), CurrentPage, TotalPages, Center, false, 0, 0, ref dx, ref dy, ref MaxDescent, THFlxAlignment.center, ZoomHead);
            DoWriteHeader(Kind, (THeaderAndFooterPos)(BaseSection + 1), CurrentPage, TotalPages, Center, true, (ClipRect.Right + ClipRect.Left) / 2 - XOfs, y + First(dy, Footer), ref dx, ref dy, ref MaxDescent, THFlxAlignment.center, ZoomHead);

            DoWriteHeader(Kind, (THeaderAndFooterPos)(BaseSection + 2), CurrentPage, TotalPages, Right, false, 0, 0, ref dx, ref dy, ref MaxDescent, THFlxAlignment.right, ZoomHead);
            DoWriteHeader(Kind, (THeaderAndFooterPos)(BaseSection + 2), CurrentPage, TotalPages, Right, true, ClipRect.Right - rx - XOfs, y + First(dy, Footer), ref dx, ref dy, ref MaxDescent, THFlxAlignment.right, ZoomHead);
        }

        private void DrawHeadersAndFooters(RectangleF PrintBounds,
            int CurrentPage, int TotalPages, RectangleF BorderRect, RectangleF PaintClipRect2)
        {
            try
            {
                THeaderAndFooter HeadFoot = FWorkbook.GetPageHeaderAndFooter();
                string Header = HeadFoot.GetHeader(CurrentPage);
                string Footer = HeadFoot.GetFooter(CurrentPage);
                THeaderAndFooterKind Kind = HeadFoot.GetHeaderAndFooterKind(CurrentPage);
                if (Footer.Length == 0 && Header.Length == 0) return;

                Canvas.SetClipReplace(PaintClipRect2);
                real XOfs = MarginBounds.Left>0? MarginBounds.Left/100F*DispMul: 0; 
                real YOfs = MarginBounds.Top>0? MarginBounds.Top/100F*DispMul: 0; 
                RectangleF ClipBounds = PrintBounds;

                if (HeadFoot.AlignMargins) { ClipBounds.X = BorderRect.X; ClipBounds.Width = BorderRect.Width; }

                real ZoomHead = HeadFoot.ScaleWithDoc ? Zoom100 : 1;
                real rx = HeadFoot.AlignMargins? 0: 0.75F * DispMul;  //This spacing is fixed in headers and footers


				if ((FHidePrintObjects & THidePrintObjects.Headers)==0)
				{
					DrawHeaderOrFooter(Kind, 0, ClipBounds, XOfs, Header, CurrentPage, TotalPages, PrintBounds.Top + (real)FMargins.Header * DispMul - YOfs, false, ZoomHead, rx);
				}
				if ((FHidePrintObjects & THidePrintObjects.Footers)==0)
				{
					DrawHeaderOrFooter(Kind, 3, ClipBounds, XOfs, Footer, CurrentPage, TotalPages, PrintBounds.Top + (real)(PrintBounds.Height - (FMargins.Footer) * DispMul - YOfs), true, ZoomHead, rx);
				}
            }
            finally
            {
                //Always reset for future processing.
                Canvas.SetClipReplace(BorderRect);
            } //finally
        }

        #endregion

        #region Draw Background
		private static bool SameFill(TFlxFormat Fm1, TFlxFormat Fm2)
		{
            if (Fm1.FillPattern.Pattern == TFlxPatternStyle.Gradient) return false; //gradients should be painted one by one except in merged cells.
			return
			Fm1.FillPattern.Pattern == Fm2.FillPattern.Pattern 
					&& Fm1.FillPattern.BgColor == Fm2.FillPattern.BgColor
					&& Fm1.FillPattern.FgColor == Fm2.FillPattern.FgColor;
		}

		private void ExpandRow(TUsedRangeList UsedRanges, int MaxCol, int Row, ref int Col, out TFlxFormat Fm0, out TXlsCellRange Mb)
		{
			bool Merged;
			Mb = XCellMergedBounds(Row, Col, out Merged);
			Fm0 = GetCellVisibleFormatDef(Mb.Top, Mb.Left);
            bool IsMergedWithGradient = IsCellMergedWithGradient(Fm0, Mb);
			
			Col = Mb.Right + 1;
			bool Similar = !IsMergedWithGradient;
			while (Col <= MaxCol && Similar)
			{
				if (UsedRanges.Find(Row, Col)) break;
				Mb = XCellMergedBounds(Row, Col, out Merged);
				TFlxFormat Fm = GetCellVisibleFormatDef(Mb.Top, Mb.Left);
                IsMergedWithGradient = IsCellMergedWithGradient(Fm, Mb);

				if (SameFill(Fm0, Fm))
				{
					Col= Mb.Right + 1;
				}
				else
				{
					Similar = false;
				}
			}

			if (Col > MaxCol + 1) Col = MaxCol + 1; //this might happen when the page ends at a merged cell.
		}

        private static bool IsCellMergedWithGradient(TFlxFormat Fm0, TXlsCellRange Mb)
        {
            return (Mb.ColCount > 1 || Mb.RowCount > 1) && Fm0.FillPattern.Pattern == TFlxPatternStyle.Gradient;
        }

		private void ExpandRange(TUsedRangeList UsedRanges, TXlsCellRange PagePrintRange, int Row, int Col, out TFlxFormat Fm0, out SizeF ResultingSize, out int dCol)
		{
			real TotalHeight = RealRowHeight(Row);

			int Col0 = Col;
            TXlsCellRange Mb;
			ExpandRow(UsedRanges, PagePrintRange.Right, Row, ref Col0, out Fm0, out Mb);
			int r;

            if (IsCellMergedWithGradient(Fm0, Mb))
            {
                Col0 = Mb.Right + 1;
                r = Mb.Bottom + 1;
				TotalHeight = CalcAcumRowHeight(Row, r);
            }
            else
            {
                r = Row + 1;
                while (r <= PagePrintRange.Bottom)
                {
                    int c = Col;
                    TFlxFormat Fm1;
                    ExpandRow(UsedRanges, Col0 - 1, r, ref c, out Fm1, out Mb);
                    if (SameFill(Fm0, Fm1) && c == Col0)
                    {
                        r++;
                        TotalHeight += RealRowHeight(r - 1);
                    }
                    else
                    {
                        break;
                    }
                }
            }

			real TotalWidth = CalcAcumColWidth(Col, Col0); 

			UsedRanges.Add(Row, Col, r, Col0);
			dCol = Col0 - Col;
			ResultingSize = new SizeF(TotalWidth, TotalHeight);
		}

		private void DrawBackground(TXlsCellRange PagePrintRange, RectangleF PaintClipRect)
		{
			//Fill the cells
			//We do this first, so lines are drawn on the background
			int Row = PagePrintRange.Top;
			real Ch = PaintClipRect.Top;
			TUsedRangeList UsedRanges = new TUsedRangeList();

			while (Row <= PagePrintRange.Bottom)
			{
				int Col = PagePrintRange.Left;
				real Cw = PaintClipRect.Left;

				while (Col <= PagePrintRange.Right)
				{
					int Index;
					if (UsedRanges.Find(Row, Col, out Index))
					{
						TXlsCellRange cr = UsedRanges[Index];
                        Cw += CalcAcumColWidth(Col, cr.Right + 1);
                        Col = cr.Right + 1;
						continue;
					}

					TFlxFormat Fm;
					SizeF ResultingSize;
					int dCol;

					ExpandRange(UsedRanges, PagePrintRange, Row, Col, out Fm, out ResultingSize, out dCol);
					if (Fm.FillPattern.Pattern != TFlxPatternStyle.None)
					{
						TExcelColor ColorBg = Fm.FillPattern.BgColor;
						TExcelColor ColorFg = Fm.FillPattern.FgColor;

						Color FgColor = Colors.Black;
						if (Fm.FillPattern.Pattern == TFlxPatternStyle.Solid) FgColor = Colors.White; //in this case, FgColor will be used as background, so automatic must be white.
						Color ABrushFg = ColorFg.ToColor(Workbook, FgColor);
						Color ABrushBg = ColorBg.ToColor(Workbook, Colors.White);

                        RectangleF CellRect = new RectangleF(Cw, Ch, ResultingSize.Width, ResultingSize.Height);
						using (Brush ABrush = FlgConsts.CreatePattern(Fm.FillPattern.Pattern, ABrushFg, ABrushBg, Fm.FillPattern.Gradient, CellRect, Workbook))
						{
							Canvas.FillRectangle(ABrush, CellRect);
						}
					}

					Col += dCol;
					Cw += ResultingSize.Width;
				}

				Ch += RealRowHeight(Row);
				Row++;
				UsedRanges.CleanUpUsed(Row);
			}
		}

        #endregion

        #region Hyperlinks
        private void DrawHyperlinks(TXlsCellRange PagePrintRange, RectangleF PaintClipRect)
        {
            if ((FHidePrintObjects & THidePrintObjects.Hyperlynks)!=0) return;
            int aCount = FWorkbook.HyperLinkCount;
            for (int i = 1; i <= aCount; i++)
            {
                THyperLink hl = FWorkbook.GetHyperLink(i);
                TXlsCellRange Range= FWorkbook.GetHyperLinkCellRange(i);

                if (hl.LinkType == THyperLinkType.URL)
                {
                    int l = Math.Max(Range.Left, PagePrintRange.Left);
                    int r = Math.Min(Range.Right, PagePrintRange.Right);
                    int t = Math.Max(Range.Top, PagePrintRange.Top);
                    int b = Math.Min(Range.Bottom, PagePrintRange.Bottom);

                    if (l<=r && t<=b)
                        Canvas.AddHyperlink(
                            PaintClipRect.Left + CalcAcumColWidth(PagePrintRange.Left, l),
                            PaintClipRect.Top + CalcAcumRowHeight(PagePrintRange.Top, t),
                            CalcAcumColWidth(l,r+1),
                            CalcAcumRowHeight(t,b+1),
                            hl.Text);
                }
            }
        }

        #endregion

        #region Comments
        private void DrawComments(TXlsCellRange PagePrintRange, RectangleF PaintClipRect)
        {
            if ((FHidePrintObjects & THidePrintObjects.Comments)!=0) return;
            int rLast = Math.Min(PagePrintRange.Bottom, FWorkbook.CommentRowCount());
            for (int r = PagePrintRange.Top; r <= rLast; r++)
                for (int i=1; i<= FWorkbook.CommentCountRow(r); i++)				
                {
                    TRichString text = FWorkbook.GetCommentRow(r, i);
                    int c= FWorkbook.GetCommentRowCol(r, i);
                    if (c<PagePrintRange.Left || c>PagePrintRange.Right) continue;

                    TXlsCellRange Range = FWorkbook.CellMergedBounds(r,c);
                    int l = Math.Max(Range.Left, PagePrintRange.Left);
                    int right = Math.Min(Range.Right, PagePrintRange.Right);
                    int t = Math.Max(Range.Top, PagePrintRange.Top);
                    int b = Math.Min(Range.Bottom, PagePrintRange.Bottom);

                    if (l<=right && t<=b)

                        Canvas.AddComment(
                            PaintClipRect.Left + CalcAcumColWidth(PagePrintRange.Left, l),
                            PaintClipRect.Top + CalcAcumRowHeight(PagePrintRange.Top, t),
                            CalcAcumColWidth(l,right+1),
                            CalcAcumRowHeight(t,b+1),
                            text.ToString(CultureInfo.CurrentCulture));
                }
        }

        #endregion

		#region RenderCell
        private Image RenderCellsInternal(ExcelFile xls, int row1, int col1, int row2, int col2, bool aDrawBackground, real dpi, SmoothingMode aSmoothingMode, InterpolationMode aInterpolationMode, bool antiAliased)
        {
            Workbook = xls;
            TXlsCellRange PagePrintRange = new TXlsCellRange(row1, col1, row2, col2);
            InitRender(PagePrintRange, 0, 1);
			CacheMergedCells(PagePrintRange, 0);

            FDrawGridLines = xls.PrintGridLines;
            FPrintFormulas = xls.ShowFormulaText;
            GridLinesColor = xls.GridLinesColor.ToColor(Workbook, Colors.LightGray);

			RectangleF PaintClipRect = new RectangleF(0, 0, CalcAcumColWidth(col1, col2 + 1), CalcAcumRowHeight(row1, row2 + 1));
			
			int wPix = (int)Math.Round(PaintClipRect.Width * dpi / FlexCelRender.DispMul);
			int hPix = (int)Math.Round(PaintClipRect.Height * dpi / FlexCelRender.DispMul);

			Bitmap Result = BitmapConstructor.CreateBitmap(wPix, hPix, PixelFormat.Format32bppArgb);
			Result.SetResolution(dpi, dpi);

            using (Graphics ImageGraphics = Graphics.FromImage(Result))
            {
				ImageGraphics.SmoothingMode = aSmoothingMode;
				ImageGraphics.InterpolationMode = aInterpolationMode;
				if (antiAliased)ImageGraphics.TextRenderingHint = TextRenderingHint.AntiAlias;
				ImageGraphics.PageUnit = GraphicsUnit.Point;

                Canvas = new GdiPlusGraphics(ImageGraphics);
                

                Canvas.CreateSFormat(); //GenericTypographic returns a NEW instance.
                try
                {
					RectangleF PaintClipBig = new RectangleF(0, 0, PaintClipRect.Width + 1, PaintClipRect.Height + 1);
                    PageFormatCache.CreatePageCache(PagePrintRange);
                    try
                    {
                        RectangleF ARect = new RectangleF(0, 0, 0, 0);

						if (aDrawBackground) 
						{
							ImageGraphics.SmoothingMode = SmoothingMode.None;
							try
							{
								Canvas.FillRectangle(Brushes.White, PaintClipBig);
								DrawBackground(PagePrintRange, PaintClipBig);
							}
							finally
							{
								ImageGraphics.SmoothingMode = aSmoothingMode;
							}
						}

                        for (int row = row1; row <= row2; row++)
                        {
                            ARect.Height = RealRowHeight(row);
                            for (int col = col1; col <= col2; col++)
                            {
                                ARect.Width = RealColWidth(col);
                                if (!IsEmptyCell(row, col))  //Merged cells or hidden columns might return indexes, but the cell is empty.
                                {
                                    DrawCell(col, row, ARect, PaintClipRect, null, true, false, TSpanDirection.Both);
                                }
                                ARect.X += ARect.Width;
                            }
							ARect.X = 0;
                            ARect.Y += ARect.Height;
                        }
                    }
                    finally
                    {
                        PageFormatCache.DestroyPageCache(); //Remember to reset it so it is not used by any other method.
                    }
                }
                finally
                {
                    Canvas.DestroySFormat();
                }
            }

			return Result;
        }

        public static Image RenderCells(ExcelFile xls, int row1, int col1, int row2, int col2, bool drawBackground, real dpi, SmoothingMode aSmoothingMode, InterpolationMode aInterpolationMode, bool antiAliased)
        {
            FlexCelRender Renderer = new FlexCelRender();
            Renderer.CreateFontCache();
            try
            {
                return Renderer.RenderCellsInternal(xls, row1, col1, row2, col2, drawBackground, dpi, aSmoothingMode, aInterpolationMode, antiAliased);
            }
            finally
            {
                Renderer.DisposeFontCache();
            }
        }

        public static RectangleF CalcCellRangeSize(ExcelFile aXls, int row1, int col1, int row2, int col2, bool IncludeMargins)
        {
            FlexCelRender Renderer = new FlexCelRender();
            using (TGraphicCanvas GrCanvas = new TGraphicCanvas())
            {
                Renderer.SetCanvas(GrCanvas.Canvas);
                Renderer.Workbook = aXls;
                Renderer.CreateFontCache();
                try
                {
                    TXlsCellRange PagePrintRange = new TXlsCellRange(row1, col1, row2, col2);
                    TXlsCellRange FinalRange = Renderer.InternalCalcPrintArea(PagePrintRange)[0];

                    Renderer.InitRender(FinalRange, 0, 1);
                    Renderer.CacheMergedCells(FinalRange, 0);
                    TRepeatingRange RepeatingRange = Renderer.GetRepeatingRange(FinalRange);
                    RectangleF PaintClipRect;
                    RectangleF PageBounds = FlexCelRender.GetPageSize(aXls);
                    Renderer.CalcPrintParams(FinalRange, PageBounds, out PaintClipRect, RepeatingRange); //Will calc zoom100

                    double mh = 0;
                    double mw = 0;
                    if (IncludeMargins)
                    {
                        TXlsMargins Margins = aXls.GetPrintMargins();
                        mh = (Margins.Left + Margins.Right) * DispMul;
                        mw = (Margins.Top + Margins.Bottom) * DispMul;
                    }


                    return new RectangleF(0, 0, (float)mw + Renderer.CalcAcumColWidth(FinalRange.Left, FinalRange.Right + 1), (float)mh + Renderer.CalcAcumRowHeight(FinalRange.Top, FinalRange.Bottom + 1));
                }
                finally
                {
                    Renderer.DisposeFontCache();
                }
            }
        }

		#endregion

        #region Auto Page Breaks
        internal static void AutoPageBreaks(ExcelFile aXls, int PercentOfUsedSheet, RectangleF PageBounds, RectangleF MarginBounds, int PageScale)
        {
            FlexCelRender Renderer = new FlexCelRender();

            using (TGraphicCanvas GrCanvas = new TGraphicCanvas())
            {
                Renderer.SetCanvas(GrCanvas.Canvas);
                Renderer.Workbook = aXls;
                Renderer.CreateFontCache();
                try
                {
                    TXlsCellRange[] PrintRanges =  Renderer.InternalCalcPrintArea(new TXlsCellRange(-1, -1, -1, -1));
                    if (PageBounds.Width <= 0 || PageBounds.Height <= 0)
                    {
                        PageBounds = GetPageSize(aXls);
                    }

                    if (MarginBounds.Width <= 0 || MarginBounds.Height <= 0)
                    {
                        MarginBounds = PageBounds;
                    }

                    int i = 0;
                    foreach (TXlsCellRange PrintRange in PrintRanges)
                    {
                        Renderer.AddPageBreaks(PercentOfUsedSheet, PageBounds, MarginBounds, PageScale, PrintRange, i, PrintRanges.Length);                        
                        i++;
                    }
                }
                finally
                {
                    Renderer.DisposeFontCache();
                }
            }
        }

        internal static RectangleF GetPageSize(ExcelFile Workbook)
        {
            TPaperDimensions pd = Workbook.PrintPaperDimensions;
            if ((Workbook.PrintOptions & TPrintOptions.Orientation) == 0)
            {
                real w = pd.Width;
                pd.Width = pd.Height;
                pd.Height = w;
            }

            return new RectangleF(0, 0, pd.Width, pd.Height);
        }

        #endregion
    }

    #region Auxiliary classes
    /// <summary>
    /// Directions on where the cell can span.
    /// </summary>
    internal enum TSpanDirection
    {
        Left,
        Right,
        Both
    }

    internal enum TCorner
    {
        Left,
        Top,
        Right,
        Bottom       
    }

    internal class TRepeatingRange
    {
        internal int FirstRow;
        internal int LastRow;
        internal int FirstCol;
        internal int LastCol;

        internal TRepeatingRange(int aFirstRow, int aLastRow, int aFirstCol, int aLastCol)
        {
            FirstRow = aFirstRow;
            LastRow = aLastRow;
            FirstCol = aFirstCol;
            LastCol = aLastCol;
        }

        internal int MaxRow(int StartRow)
        {
            return Math.Min(LastRow, StartRow-1);
        }

        internal int MaxCol(int StartCol)
        {
            return Math.Min(LastCol, StartCol-1);
        }

        internal int RowCount(int TopRow)
        {
            int lr = MaxRow(TopRow);
            if (lr>=FirstRow) return lr -FirstRow+1;
            return 0;
        }

        internal int ColCount(int TopCol)
        {
            int lc = MaxCol(TopCol);
            if (lc>=FirstCol) return lc-FirstCol+1;
            return 0;
        }
    }

    class TPrintAreaSort : IComparer<TXlsCellRange>
    {
        bool LeftToRight;

        public TPrintAreaSort(bool aLeftToRight)
        {
            LeftToRight = aLeftToRight;
        }

        #region IComparer<TXlsCellRange> Members

        public int Compare(TXlsCellRange x, TXlsCellRange y)
        {
            int Result;

            if (LeftToRight)
            {
                Result = x.Left.CompareTo(y.Left);
                if (Result != 0) return Result;
            }

            Result = x.Top.CompareTo(y.Top);
            if (Result != 0) return Result;

            return x.Left.CompareTo(y.Left);
        }

        #endregion
    }

   #endregion


}
