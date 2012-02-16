using System;
using FlexCel.Core;
using System.Resources;
using System.IO;
using System.Reflection;
using System.Diagnostics;
using System.Text;
using System.Collections.Generic;

using System.Globalization;
using real = System.Single;

#if (MONOTOUCH)
using System.Drawing;
using Color = MonoTouch.UIKit.UIColor;
#else
#if (WPF)
using System.Windows.Media;
using Colors = FlexCel.Core.Colors;
#else
using System.Drawing;
using Colors = System.Drawing.Color;
#endif
#endif


namespace FlexCel.XlsAdapter
{
    #region Sheets

    /// <summary>
	/// Base Excel Sheet. It can be a worksheet, a Chart, a macro sheet, etc.
	/// </summary>
	internal abstract class TSheet: TBaseSection
    {
        #region Variables
        internal TWorkbookGlobals FWorkbookGlobals;

        internal TSheetGlobals SheetGlobals;
        internal TPageSetup PageSetup;
        internal TBgPicRecord BgPic;
        internal TMiscRecordList BigNames;
        internal TSheetProtection SheetProtection;

        internal TColInfoList Columns;
        internal TScenarios Scenarios;
        internal TSortAndFilter SortAndFilter;
        internal TDimensionsRecord OriginalDimensions;

        internal TCells Cells;
        internal TDrawing Drawing;

        internal TDrawing HeaderImages;
        internal TNoteList Notes;

        internal TPivotViewList PivotView;
        internal TDConn Connections;

        internal TWindow Window;

        internal TCustomViewList CustomViews;
        internal TMiscRecordList RRSort;

        internal TRangeList<TMergedCells> MergedCells;
        internal TLRngRecord LRng;

        internal TQueryTableList QueryTable;
        internal TPhoneticRecord Phonetic;

        internal TRangeList<TCondFmt> ConditionalFormats;
        internal THLinkList HLinks;
        internal TDataValidationList DataValidation;

        internal TCodeNameRecord CodeNameRecord;

        internal TMiscRecordList WebPub;
        internal TMiscRecordList CellWatches;

        internal TSheetExtRecord SheetExt;

        internal TMiscRecordList Feat;
        internal TMiscRecordList Feat11;

        internal TMiscRecordList FutureRecords;

        internal TChartDef Chart;

#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
        internal TXlsxPivotTableList XlsxPivotTables;
#endif
        #endregion

        #region Constructors And Initializers
        internal TSheet(TWorkbookGlobals aWorkbookGlobals): base()
		{
			FWorkbookGlobals=aWorkbookGlobals;

            SheetGlobals = new TSheetGlobals();
            PageSetup = new TPageSetup();
            BigNames = new TMiscRecordList();
            SheetProtection = new TSheetProtection();

            Columns = new TColInfoList(aWorkbookGlobals, false); //will be modified later.
            Scenarios = new TScenarios();
            SortAndFilter = new TSortAndFilter();

            Cells = new TCells(aWorkbookGlobals, Columns, SheetGlobals);
            Drawing = new TDrawing(FWorkbookGlobals.DrawingGroup, xlr.MSODRAWING, 0);

            HeaderImages = new TDrawing(FWorkbookGlobals.HeaderImages, xlr.HEADERIMG, 14);
            Notes = new TNoteList();

            PivotView = new TPivotViewList();
            Connections = new TDConn();

            Window = new TWindow();

            CustomViews = new TCustomViewList();
            RRSort = new TMiscRecordList();

            MergedCells = new TRangeList<TMergedCells>();

            QueryTable = new TQueryTableList();

            ConditionalFormats = new TRangeList<TCondFmt>();
            HLinks = new THLinkList();
            DataValidation = new TDataValidationList();

            WebPub = new TMiscRecordList();
            CellWatches = new TMiscRecordList();

            Feat = new TMiscRecordList();
            Feat11 = new TMiscRecordList();

            FutureRecords = new TMiscRecordList();

            Chart = new TChartDef(FWorkbookGlobals);
#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
            XlsxPivotTables = new TXlsxPivotTableList();
#endif
        }

        protected TMiscRecordList CopyMiscList(TMiscRecordList a, TSheetInfo SheetInfo)
        {
            TMiscRecordList Result = new TMiscRecordList();
            Result.CopyFrom(a, SheetInfo);
            return Result;
        }

        private void DoRealCopyTo(TSheet Result, TSheetInfo SheetInfo, bool CopyData)
        {
            Result.sBOF = (TBOFRecord)TBOFRecord.Clone(sBOF, SheetInfo);
            Result.sEOF = (TEOFRecord)TEOFRecord.Clone(sEOF, SheetInfo);

            Result.SheetGlobals = SheetGlobals.Clone(SheetInfo);
            Result.PageSetup = PageSetup.Clone(SheetInfo);
            Result.BgPic = (TBgPicRecord)TBgPicRecord.Clone(BgPic, SheetInfo);
            Result.BigNames = CopyMiscList(BigNames, SheetInfo);

            if (CopyData)
            {
                Result.SheetProtection = TSheetProtection.Clone(SheetProtection, SheetInfo);

                Result.Columns = new TColInfoList(Result.FWorkbookGlobals, Columns.AllowStandardWidth);
                Result.Columns.CopyFrom(Columns);
            }

            Result.Scenarios = Scenarios.Clone(SheetInfo);
            Result.SortAndFilter = SortAndFilter.Clone(SheetInfo);

            if (CopyData)
            {
                Result.OriginalDimensions = (TDimensionsRecord)TBaseRecord.Clone(OriginalDimensions, SheetInfo);

                Result.Cells = new TCells(SheetInfo.DestGlobals, Result.Columns, Result.SheetGlobals);
                Result.Cells.CopyFrom(Cells, SheetInfo);

                Result.Drawing = new TDrawing(SheetInfo.DestGlobals.DrawingGroup, xlr.MSODRAWING, 0);
                Result.Drawing.CopyFrom(0, 0, Drawing, SheetInfo);
                Result.HeaderImages = new TDrawing(SheetInfo.DestGlobals.HeaderImages, xlr.HEADERIMG, 14);
                Result.HeaderImages.CopyFrom(0, 0, HeaderImages, SheetInfo);

                Result.Notes = new TNoteList();
                Result.Notes.CopyFrom(Notes, SheetInfo);
                Result.PivotView = PivotView.Clone(SheetInfo);
#if(FRAMEWORK30 && !COMPACTFRAMEWORK)
                Result.XlsxPivotTables = new TXlsxPivotTableList();
                XlsxPivotTables.CopyTo(Result.XlsxPivotTables, SheetInfo.DestGlobals);
#endif
            }
            Result.Connections = Connections.Clone(SheetInfo);

            Result.Window = Window.Clone(SheetInfo);

            if (CopyData)
            {
                Result.CustomViews = CustomViews.Clone(SheetInfo);
                Result.RRSort = CopyMiscList(RRSort, SheetInfo);

                Result.MergedCells = new TRangeList<TMergedCells>();
                Result.MergedCells.CopyFrom(MergedCells, SheetInfo);
            }
            Result.LRng = (TLRngRecord)TLRngRecord.Clone(LRng, SheetInfo);

            Result.QueryTable = QueryTable.Clone(SheetInfo);
            Result.Phonetic = (TPhoneticRecord)TPhoneticRecord.Clone(Phonetic, SheetInfo);

            if (CopyData)
            {
                Result.ConditionalFormats = new TRangeList<TCondFmt>();
                Result.ConditionalFormats.CopyFrom(ConditionalFormats, SheetInfo);
                Result.HLinks = new THLinkList();
                Result.HLinks.CopyFrom(HLinks, SheetInfo);
                Result.DataValidation = new TDataValidationList();
                Result.DataValidation.CopyFrom(DataValidation, Result.Drawing, SheetInfo);
            }

            Result.CodeNameRecord = null;

            Result.WebPub = CopyMiscList(WebPub, SheetInfo);
            if (CopyData)
            {
                Result.CellWatches = CopyMiscList(CellWatches, SheetInfo);
            }

            Result.SheetExt = (TSheetExtRecord)TSheetExtRecord.Clone(SheetExt, SheetInfo);

            Result.Feat = CopyMiscList(Feat, SheetInfo);
            Result.Feat11 = CopyMiscList(Feat11, SheetInfo);

            Result.FutureRecords = CopyMiscList(FutureRecords, SheetInfo);
            if (FutureStorage != null) Result.FutureStorage = FutureStorage.Clone();

            Result.Chart = Chart.Clone(SheetInfo);

            DoExtraCopy(Result, SheetInfo, CopyData);
        }

        protected virtual void DoExtraCopy(TSheet DestSheet, TSheetInfo SheetInfo, bool CopyData)
        {
        }

        protected abstract TSheet CreateSheet(TWorkbookGlobals DestGlobals);

        protected virtual TSheet DoCopyTo(TSheetInfo SheetInfo)
        {
            TWorkbookGlobals DestGlobals = SheetInfo.DestGlobals == null ? FWorkbookGlobals : SheetInfo.DestGlobals;
            TSheet Result = CreateSheet(DestGlobals);

            if (SheetInfo.DestSheet == null) SheetInfo.DestSheet = Result;

            DoRealCopyTo(Result, SheetInfo, true);

            Result.Notes.FixDwgIds(Result.Drawing, this, false, SheetInfo.CopiedGen); //After copying Drawing.

            return Result;
        }

        internal TSheet CopyMiscData(TSheetInfo SheetInfo)
        {
            SheetInfo.DestSheet = CreateSheet(SheetInfo.DestGlobals);
            DoRealCopyTo(SheetInfo.DestSheet, SheetInfo, false);
            SheetInfo.DestSheet.Columns.DefColWidthChars = Columns.DefColWidthChars;
            SheetInfo.DestSheet.Columns.DefColWidthChars256 = Columns.DefColWidthChars256;
            return SheetInfo.DestSheet;
        }

        internal static TSheet Clone(TSheet Self, TSheetInfo SheetInfo) //This method can't be virtual
        {
            if (Self == null) return null; else return Self.DoCopyTo(SheetInfo);
        }

        internal void Destroy()
        {
            Cells.Clear();
            Notes.Clear();
            Drawing.Destroy();
            HeaderImages.Destroy();
        }

        internal TSheet ClearValues()
        {
            Destroy();
            TSheet Result = CreateSheet(FWorkbookGlobals);
            Result.EnsureRequiredRecords();
            return Result;
        }
        #endregion

        #region Load
		internal override void LoadFromStream(TBaseRecordLoader RecordLoader, TBOFRecord First)
		{
			TLoaderInfo Loader = new TLoaderInfo();
			do
			{
				int rRow;
				TBaseRecord R=RecordLoader.LoadRecord(out rRow, false);
      
				if (R!= null)  //Null will be ignored
					R.LoadIntoSheet(this, rRow, RecordLoader, ref Loader);
			}
			while (!Loader.Eof);

            EnsureRequiredRecords();

			Notes.FixDwgIds(Drawing, this, true, new TCopiedGen(0));
			Cells.CellList.Sort();

			Drawing.SaveObjectCoords(this);

			//this must be the last statement, so if there is an exception, we dont take First
			sBOF= First;

		}        
        #endregion

        #region Save
        internal virtual void EnsureRequiredRecords()
        {
            if (sBOF == null) sBOF = TBOFRecord.CreateEmptyWorksheet(SheetType);
            if (sEOF == null) sEOF = new TEOFRecord();

            SheetGlobals.EnsureRequiredRecords();
            PageSetup.EnsureRequiredRecords(false);
            SortAndFilter.EnsureRequiredRecords();
            Window.EnsureRequiredRecords(SheetType);
        }

        internal void SaveGenericSheet(IDataStream DataStream, TSaveData SaveData, TXlsCellRange CellRange)
        {
            if ((sBOF == null) || (sEOF == null)) XlsMessages.ThrowException(XlsErr.ErrSectionNotLoaded);
            sBOF.SaveToStream(DataStream, SaveData, 0);
            if (International) TInternationalRecord.SaveRecord(DataStream);
            FWorkbookGlobals.CalcOptions.SaveToStream(DataStream);
            SheetGlobals.SaveToStream(Cells, Columns, DataStream, SaveData, CellRange);
            PageSetup.SaveToStream(DataStream, SaveData, Guid.Empty);

            if (SheetType != TSheetType.Dialog)
            {
                if (BgPic != null) BgPic.SaveToStream(DataStream, SaveData, 0);
            }

            BigNames.SaveToStream(DataStream, SaveData, 0);

            if (SheetType != TSheetType.Dialog)
            {
                SheetProtection.SaveFirstPart(DataStream, SaveData);
                Columns.SaveToStream(DataStream, SaveData, CellRange);
            }
            else
            {
                SheetProtection.SaveDialogSheetProtect(DataStream, SaveData);
                Columns.SaveDefColWidth(DataStream, SaveData);
            }

            if (SheetType == TSheetType.Worksheet)
            {
                Scenarios.SaveToStream(DataStream, SaveData);
            }

            if (SheetType != TSheetType.Dialog)
            {
                SortAndFilter.SaveToStream(DataStream, SaveData, SheetType == TSheetType.Macro);
            }
            

            Cells.SaveToStream(DataStream, SaveData, CellRange); //This includes dimensions

            //Excel doesn't save drawings to the clipboard
            if (CellRange == null)
            {
                Drawing.SaveToStream(DataStream, SaveData);
                HeaderImages.SaveToStream(DataStream, SaveData);
                Notes.SaveToStream(DataStream, SaveData, CellRange);
            }

            if (SheetType == TSheetType.Worksheet)
            {
                PivotView.SaveToStream(DataStream, SaveData);
            }

            if (SheetType != TSheetType.Dialog)
            {
                Connections.SaveToStream(DataStream, SaveData);
            }
            
            Window.SaveToStream(DataStream, SaveData, SheetType == TSheetType.Dialog);
            CustomViews.SaveToStream(DataStream, SaveData, CellRange);

            if (SheetType != TSheetType.Dialog)
            {
                RRSort.SaveToStream(DataStream, SaveData, 0);
                Columns.SaveStandardWidth(DataStream);


                if (SheetType == TSheetType.Worksheet)
                {
                    MergedCells.SaveToStream(DataStream, SaveData, CellRange);
                    if (LRng != null) LRng.SaveToStream(DataStream, SaveData, 0);
                    QueryTable.SaveToStream(DataStream, SaveData);
                }

                if (Phonetic != null) Phonetic.SaveToStream(DataStream, SaveData, 0);

                if (SheetType == TSheetType.Worksheet)
                {
                    ConditionalFormats.SaveToStream(DataStream, SaveData, CellRange);
                    HLinks.SaveToStream(DataStream, SaveData, CellRange);
                    DataValidation.SaveToStream(DataStream, SaveData, CellRange);
                }
            }

            if (CodeNameRecord != null) CodeNameRecord.SaveToStream(DataStream, SaveData, 0);

            if (SheetType == TSheetType.Worksheet)
            {
                WebPub.SaveToStream(DataStream, SaveData, 0);
            }

            if (SheetType != TSheetType.Dialog)
            {
                CellWatches.SaveToStream(DataStream, SaveData, 0);
            }

            if (SheetExt != null) SheetExt.SaveToStream(DataStream, SaveData, 0);
            SheetProtection.SaveSecondPart(DataStream, SaveData);
            Feat.SaveToStream(DataStream, SaveData, 0);
            Feat11.SaveToStream(DataStream, SaveData, 0);
            FutureRecords.SaveToStream(DataStream, SaveData, 0);

            sEOF.SaveToStream(DataStream, SaveData, 0);
        }

        protected virtual long TotalSheetSize(TXlsCellRange CellRange)
        {
            long Result = 0;
            if (International) Result += TInternationalRecord.StandardSize();
            Result += TCalcOptions.TotalSize();
            Result += SheetGlobals.TotalSize(CellRange);
            Result += PageSetup.TotalSize();

            if (SheetType != TSheetType.Dialog)
            {
                if (BgPic != null) Result += BgPic.TotalSize();
            }

            Result += BigNames.TotalSize;

            if (SheetType != TSheetType.Dialog)
            {
                Result += SheetProtection.TotalSizeFirst();
                Result += Columns.TotalSizeNoStdWidth(CellRange);
            }
            else
            {
                Result += SheetProtection.TotalSizeDialogSheet();
                Result += TDefColWidthRecord.StandardSize();
            }

            if (SheetType == TSheetType.Worksheet)
            {
                Result += Scenarios.TotalSize();
            }

            if (SheetType != TSheetType.Dialog)
            {
                Result += SortAndFilter.TotalSize(SheetType == TSheetType.Macro);
            }


            Result += Cells.TotalSize(CellRange); //This includes dimensions

            //Excel doesn't save drawings to the clipboard
            if (CellRange == null)
            {
                Result += Drawing.TotalSize();
                Result += HeaderImages.TotalSize();
            }

            Result += Notes.TotalSize(CellRange);

            if (SheetType == TSheetType.Worksheet)
            {
                Result += PivotView.TotalSize();
            }

            if (SheetType != TSheetType.Dialog)
            {
                Result += Connections.TotalSize();
            }
            
            Result += Window.TotalSize(SheetType == TSheetType.Dialog);
            Result += CustomViews.TotalSize(CellRange);

            if (SheetType != TSheetType.Dialog)
            {
                Result += RRSort.TotalSize;
                Result += Columns.StandardWidthSize();


                if (SheetType == TSheetType.Worksheet)
                {
                    Result += MergedCells.TotalSize(CellRange);
                    if (LRng != null) Result += LRng.TotalSize();
                    Result += QueryTable.TotalSize();
                }

                if (Phonetic != null) Result += Phonetic.TotalSize();

                if (SheetType == TSheetType.Worksheet)
                {
                    Result += ConditionalFormats.TotalSize(CellRange);
                    Result += HLinks.TotalSize(CellRange);
                    Result += DataValidation.TotalSize(CellRange);
                }
            }

            if (CodeNameRecord != null) Result += CodeNameRecord.TotalSize();

            if (SheetType == TSheetType.Worksheet)
            {
                Result += WebPub.TotalSize;
            }

            if (SheetType != TSheetType.Dialog)
            {
                Result += CellWatches.TotalSize;
            }

            if (SheetExt != null) Result += SheetExt.TotalSize();
            Result += SheetProtection.TotalSizeSecond();
            Result += Feat.TotalSize;
            Result += Feat11.TotalSize;
            Result += FutureRecords.TotalSize;

            return Result;
        }

        internal override long TotalRangeSize(int SheetIndex, TXlsCellRange CellRange, TEncryptionData Encryption, bool Repeatable)
        {
            long Result = base.TotalRangeSize(SheetIndex, CellRange, Encryption, Repeatable) +
                TotalSheetSize(CellRange);

            return Result;
        }

        internal override long TotalSize(TEncryptionData Encryption, bool Repeatable)
        {
            long Result = base.TotalSize(Encryption, Repeatable) +
                TotalSheetSize(null);

            return Result;
        }

        #endregion

        #region Properties
        internal abstract TSheetType SheetType { get; }
        internal virtual bool International { get { return false; } set { } }

        internal string CodeName
        {
            get
            {
                if (CodeNameRecord == null) return String.Empty; else return CodeNameRecord.SheetName;
            }
            set
            {
                if (value == null || value.Length == 0) CodeNameRecord = null;
                else CodeNameRecord = new TCodeNameRecord(value);
            }
        }

        internal bool Selected 
		{
			get
			{
				return Window.Window2.Selected;
			}
			set
			{
				Window.Window2.Selected=value;
			}
        }

        internal int DefRowHeight
        {
            get
            {
                if (SheetGlobals.DefRowHeight == null) return 0xFF; //might happen when loading xlsx, if there is not a defRowHeight present (before we call ensure records).
                return SheetGlobals.DefRowHeight.Height;
            }
            set
            {
                SheetGlobals.DefRowHeight.Height = value;
            }
        }

        internal int DefRowFlags
        {
            get
            {
                return SheetGlobals.DefRowHeight.Flags;
            }
            set
            {
                SheetGlobals.DefRowHeight.Flags = value;
            }
        }

        internal int VisualDefRowHeight { get { if ((DefRowFlags & 0x2) != 0) return 0; return DefRowHeight; } }

        /// <summary>
        /// Margins in Inches
        /// </summary>
        internal TXlsMargins Margins
        {
            get
            {
                byte[] l0 = { 0xFC, 0xFD, 0x7E, 0xBF, 0xDF, 0xEF, 0xE7, 0x3F };
                byte[] t0 = { 0xE0, 0xEF, 0xF7, 0xFB, 0xFD, 0x7E, 0xEF, 0x3F };
                TXlsMargins Result = new TXlsMargins();

                if (PageSetup.LeftMargin >= 0) Result.Left = PageSetup.LeftMargin; else Result.Left = BitConverter.ToDouble(l0, 0);
                if (PageSetup.RightMargin >= 0) Result.Right = PageSetup.RightMargin; else Result.Right = BitConverter.ToDouble(l0, 0);
                if (PageSetup.TopMargin >= 0) Result.Top = PageSetup.TopMargin; else Result.Top = BitConverter.ToDouble(t0, 0);
                if (PageSetup.BottomMargin >= 0) Result.Bottom = PageSetup.BottomMargin; else Result.Bottom = BitConverter.ToDouble(t0, 0);

                Result.Header = PageSetup.Setup.HeaderMargin;
                Result.Footer = PageSetup.Setup.FooterMargin;

                return Result;
            }
            set
            {
                PageSetup.LeftMargin = value.Left;
                PageSetup.RightMargin = value.Right;
                PageSetup.TopMargin = value.Top;
                PageSetup.BottomMargin = value.Bottom;

                PageSetup.Setup.HeaderMargin = value.Header;
                PageSetup.Setup.FooterMargin = value.Footer;
            }
        }

        internal int Zoom
        {
            get
            {
                if (Window.Scl != null) return Window.Scl.Zoom; else return 100;
            }
            set
            {
                if (Window.Scl == null) Window.Scl = new TSCLRecord(100);
                Window.Scl.Zoom = value;
            }
        }

        internal TExcelColor GetSheetTabColor()
        {
            if (SheetExt != null) return SheetExt.GetTabColor(FWorkbookGlobals.Workbook); else return TExcelColor.Automatic;
        }

        internal void SetSheetTabColor(TExcelColor aColor)
        {
            if (aColor.IsAutomatic)
            {
                SheetExt = null;
                return;
            }
            if (SheetExt == null) SheetExt = new TSheetExtRecord(aColor); else SheetExt.SetTabColor(aColor);
        }

        internal TCommentProperties GetCommentProperties(ExcelFile Workbook, int Row, int Index)
        {
            TNoteRecord Note = Notes[Row][Index];
            if (Note.Dwg == null) return null; // TCommentProperties.GetDefaultProps(Row + 1, Note.Col, Workbook); //this might happen when pasting from clipboard./
            TMsObj o = (TMsObj)Note.Dwg.ClientData;
            TEscherClientTextBoxRecord t = Note.GetClientTextBox();
            TEscherOPTRecord Opt = Note.GetOpt();
            TShapeOptionList ShapeOptions = Opt.ShapeOptions();

            string AltText = ShapeOptions.AsUnicodeString(TShapeOption.wzDescription, null);
            long ft = ShapeOptions.AsLong(TShapeOption.fFitTextToShape, 0);
            bool AutoSize = ((ft & 0x2) != 0 && (ft & 0x20000) != 0);

            long lar = ShapeOptions.AsLong(TShapeOption.fLockAgainstGrouping, 0);
            bool LockAspectRatio = (lar & 0x80) != 0 && (lar & 0x800000) != 0;

            TFillStyle FillColor = GetObjectBackground(ShapeOptions, TShapeOption.fillColor, ColorUtil.BgrToRgb(TCommentProperties.DefaultFillColorRGB), TCommentProperties.DefaultFillColorSystem);
            TFillStyle BorderColor = GetObjectBackground(ShapeOptions, TShapeOption.lineColor, -1, TSystemColor.None);
            TTextRotation TextRotation = t.TextRotation;
            bool HasFill = ShapeOptions.AsBool(TShapeOption.fNoFillHitTest, true, 4);
            bool HasBorder = ShapeOptions.AsBool(TShapeOption.fNoLineDrawDash, true, 3);
            bool aHidden = !Opt.Visible;
            bool aIs3D = false; //doesn't matter in comments.

            return new TCommentProperties(Note.GetAnchor(Row, this), String.Empty, null, new TCropArea(),
                FlxConsts.NoTransparentColor, FlxConsts.DefaultBrightness, FlxConsts.DefaultContrast, FlxConsts.DefaultGamma,
                o.IsLocked, o.IsPrintable, o.IsPublished, o.IsDisabled, o.IsDefaultSize, o.IsAutoFill, o.IsAutoLine,
                AltText, null, LockAspectRatio, new TObjectTextProperties(t.LockText, t.HAlign, t.VAlign, TextRotation),
                AutoSize, new TShapeFill(HasFill, FillColor), 
                new TShapeLine(HasBorder, new TLineStyle(BorderColor)), aHidden, aIs3D, true);
        }

        internal static TFillStyle GetObjectBackground(TShapeOptionList ShapeOptions, TShapeOption ShProp, long DefaultColor, TSystemColor DefaultSysColor)
        {
            TFillStyle FillColor = null;
            unchecked
            {
                long clr = ShapeOptions.AsLong(ShProp, DefaultColor);
                if (clr == -1) return null;

                if ((clr & 0x8000000) == 0) //not indexed color
                {
                    FillColor = new TSolidFill(ColorUtil.FromArgb(ColorUtil.BgrToRgb(clr)));
                }
                else
                {
                    TSystemColor SysColor = ColorUtil.GetSystemColor((clr & 0xFF) - 56);
                    if (SysColor == TSystemColor.None)
                    {
                        if (DefaultSysColor != TSystemColor.None) FillColor = new TSolidFill(TDrawingColor.FromSystem(DefaultSysColor));
                    }
                    else
                    {
                        FillColor = new TSolidFill(TDrawingColor.FromSystem(SysColor));
                    }
                }
            }
            return FillColor;
        }

        internal virtual int ChartCount
        {
            get
            {
                return 0;
            }
        }

        internal bool HasExternRefs()
        {
            return Chart.HasExternRefs() || Drawing.HasExternRefs();
        }


        #endregion

        #region Row & Col
        internal bool HasCol(int aCol)
        {
            return Columns[aCol] != null;
        }

        internal int GetRowOptions(int aRow)
        {
            if (!Cells.CellList.HasRow(aRow)) return 0;
            else
                return Cells.CellList.RowOptions(aRow);
        }

        internal int GetColOptions(int aCol)
        {
            TColInfo ci = Columns[aCol];
            if (ci == null) return 0; else return ci.Options;
        }

        internal void SetRowOptions(int aRow, int Value)
        {
            Cells.CellList.SetRowOptions(aRow, Value);
        }

        internal void CollapseRows(int aRow, int Level, TCollapseChildrenMode CollapseChildren, bool IsNode)
        {
            Cells.CellList.CollapseRows(aRow, Level, CollapseChildren, IsNode);
        }

        internal void CollapseCols(int aCol, int Level, TCollapseChildrenMode CollapseChildren, bool IsNode)
        {
            int ColLevel = GetColOutlineLevel(aCol);
            if (ColLevel == 0) return;
            bool NeedsToHide = ColLevel >= Level;
            const byte Mask = 0x01;
            const int CMask = 0x1000;

            int Options = GetColOptions(aCol);
            int OrigOptions = Options;

            if (NeedsToHide)
            {
                if (IsNode)
                {
                    if (CollapseChildren == TCollapseChildrenMode.Collapsed) Options = (Options | CMask);
                    if (CollapseChildren == TCollapseChildrenMode.Expanded) Options = (Options & ~CMask);
                }

                Options = Options | Mask;
            }
            else
            {
                if (IsNode && ColLevel == Level - 1)
                    Options = Options & ~Mask | CMask; //The parent node has one level less than the hidden children.
                else
                    Options = Options & ~Mask & ~CMask;
            }

            if (Options != OrigOptions) SetColOptions(aCol, Options);
        }

        internal void SetColOptions(int aCol, int Value)
        {
            TColInfo ci = Columns[aCol];
            if (ci != null)
                ci.Options = Value;
            else
                Columns[aCol] = new TColInfo(Columns.DefColWidth, FlxConsts.DefaultFormatId, Value, true);
        }

        internal int GetRowHeight(int aRow, bool HiddenIsZero)
        {
            if (!Cells.CellList.HasRow(aRow)) return VisualDefRowHeight;
            else
            {
                if (HiddenIsZero && Cells.CellList[aRow].RowRecord.IsHidden()) return 0;
                return Cells.CellList.RowHeight(aRow);
            }
        }

        internal int GetColWidth(int aCol, bool HiddenIsZero)
        {
            TColInfo ci = Columns[aCol];
            if (ci == null)
                return Columns.DefColWidth;
            else
            {
                if (HiddenIsZero && ((ci.Options & 0x1) == 0x1)) return 0;
                return ci.Width;
            }
        }

        internal void SetRowHeight(int aRow, int Value)
        {
            Cells.CellList.SetRowHeight(aRow, Value);
        }

        internal void SetColWidth(int aCol, int Value)
        {
            TColInfo ci = Columns[aCol];
            if (ci != null)
                ci.Width = Value;
            else
                Columns[aCol] = new TColInfo(Value, FlxConsts.DefaultFormatId, 0, false);
        }


        internal bool GetRowHidden(int aRow)
        {
            if (!Cells.CellList.HasRow(aRow)) return false;
            else
                return Cells.CellList[aRow].RowRecord.IsHidden();
        }

        internal bool GetColHidden(int aCol)
        {
            TColInfo ci = Columns[aCol];
            if (ci == null)
                return false;
            else
                return (ci.Options & 0x1) == 0x1;
        }

        internal void SetRowHidden(int aRow, bool Value)
        {
            Cells.CellList.AddRow(aRow);
            Cells.CellList[aRow].RowRecord.Hide(Value);
        }

        internal void SetColHidden(int aCol, bool Value)
        {
            TColInfo ci = Columns[aCol];
            if (ci != null)
            {
                if (Value)
                {
                    ci.Options = (ci.Options | 0x1);
                }
                else
                {
                    if (ci.Width == 0) ci.Width = Columns.DefColWidth;
                    ci.Options = (ci.Options & (~1));
                }
            }
            else
                if (Value)
                    Columns[aCol] = new TColInfo(Columns.DefColWidth, FlxConsts.DefaultFormatId, 0x1, true);
        }

        internal int GetRowFormat(int aRow)
        {
            if (!Cells.CellList.HasRow(aRow) || !Cells.CellList[aRow].RowRecord.IsFormatted()) return -1;
            else
                return Cells.CellList[aRow].RowRecord.XF;
        }

        internal int GetColFormat(int aCol)
        {
            TColInfo ci = Columns[aCol];
            if (ci == null)
                return -1;
            else return ci.XF;
        }

        private void DoResetRow(int aRow, int aXF)
        {
            //Reset all cells in row to format XF
            if ((aRow >= 0) && (aRow < Cells.CellList.Count))
                for (int i = 0; i < Cells.CellList[aRow].Count; i++)
                    Cells.CellList[aRow][i].XF = aXF;
        }

        internal void SetRowFormat(int aRow, int Value, bool ResetRow)
        {
            if (Value < 0)
            {
                if (Cells.CellList[aRow].RowRecord == null) return;
                Value = FlxConsts.DefaultFormatId;
            }

            Cells.CellList.AddRow(aRow);
            Cells.CellList[aRow].RowRecord.XF = Value;

            if (ResetRow) DoResetRow(aRow, Value);
        }

        internal void SetColFormat(int aCol, int Value, bool ResetColumn)
        {
            TColInfo ci = Columns[aCol];
            if (ci != null)
            {
                if (Value < 0) Value = FlxConsts.DefaultFormatId;
                ci.XF = Value;
            }
            else
            {
                if (Value < 0) return;
                Columns[aCol] = new TColInfo(Columns.DefColWidth, Value, 0, true);
            }

            //Reset all cells in column to format XF
            if (ResetColumn)
            {
                int Index = -1;
                for (int i = 0; i < Cells.CellList.Count; i++)
                    if (Cells.CellList[i].Find(aCol, ref Index)) Cells.CellList[i][Index].XF = Value;
            }
        }

        #endregion

        #region Outline
        internal void SetRowOutlineLevel(int aRow, int level)
        {
            if (level <= 0 && !Cells.CellList.HasRow(aRow)) return;
            if (SheetGlobals.Guts != null) SheetGlobals.Guts = new TGutsRecord(); //CreateFromData
            SheetGlobals.Guts.RecalcNeeded = true;
            Cells.CellList.AddRow(aRow);
            Cells.CellList[aRow].RowRecord.SetRowOutlineLevel(level);
        }

        internal int GetRowOutlineLevel(int row)
        {
            return GetRowOptions(row) & 0x07;
        }

        internal void SetColOutlineLevel(int aCol, int level)
        {
            if (level <= 0 && !HasCol(aCol)) return;
            if (SheetGlobals.Guts != null) SheetGlobals.Guts = new TGutsRecord(); //CreateFromData
            SheetGlobals.Guts.RecalcNeeded = true;
            TColInfo ci = Columns[aCol];
            if (ci == null)
            {
                ci = new TColInfo(Columns.DefColWidth, FlxConsts.DefaultFormatId, 0, true);
                Columns[aCol] = ci;
            }

            ci.SetColOutlineLevel(level);

        }

        internal int GetColOutlineLevel(int aCol)
        {
            return (GetColOptions(aCol) >> 8) & 0x07;
        }

        #endregion

        #region InsertAndCopy

        internal void InsertAndCopyRange(TXlsCellRange SourceRange, int DestRow, int DestCol, int aRowCount, int aColCount, TRangeCopyMode CopyMode, TFlxInsertMode InsertMode, TSheetInfo SheetInfo)
        {
            Debug.Assert((aRowCount == 0) || (aColCount == 0));

            Cells.InsertAndCopyRange(SourceRange, InsertMode, DestRow, DestCol, aRowCount, aColCount, CopyMode, SheetInfo);
            MergedCells.InsertAndCopyRange(SourceRange, DestRow, DestCol, aRowCount, aColCount, CopyMode, InsertMode, SheetInfo);
            ConditionalFormats.InsertAndCopyRange(SourceRange, DestRow, DestCol, aRowCount, aColCount, CopyMode, InsertMode, SheetInfo);
            DataValidation.InsertAndCopyRange(SourceRange, DestRow, DestCol, aRowCount, aColCount, CopyMode, InsertMode, SheetInfo);
            Notes.InsertAndCopyRange(SourceRange, InsertMode, DestRow, DestCol, aRowCount, aColCount, CopyMode, SheetInfo);
            Drawing.InsertAndCopyRange(CopyMode == TRangeCopyMode.AllIncludingDontMoveAndSizeObjects, SourceRange, InsertMode, DestRow, DestCol, aRowCount, aColCount, CopyMode, SheetInfo);
            HLinks.InsertAndCopyRange(SourceRange, InsertMode, DestRow, DestCol, aRowCount, aColCount, CopyMode, SheetInfo);
            if (Window.Pane != null) Window.Pane.ArrangeInsertRange(SourceRange.OffsetForIns(DestRow, DestCol, InsertMode), aRowCount, aColCount, Window.Window2.IsFrozen);

#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
#endif

            if ((aRowCount > 0) && (SourceRange.Left <= 0) && (SourceRange.Right >= FlxConsts.Max_Columns))
            {
                SheetGlobals.HPageBreaks.InsertRows(DestRow, SourceRange.RowCount * aRowCount);
            }
            if ((aColCount > 0) && (SourceRange.Top <= 0) && (SourceRange.Bottom >= FlxConsts.Max_Rows))
            {
                SheetGlobals.VPageBreaks.InsertCols(DestCol, SourceRange.ColCount * aColCount);
            }
        }

        internal void ClearRange(TXlsCellRange CellRange)
        {
            Cells.ClearRange(CellRange);
            Drawing.ClearRange(CellRange);
            Notes.ClearRange(CellRange);
            HLinks.ClearRange(CellRange);
        }

        internal void ClearFormats(TXlsCellRange CellRange)
        {
            Cells.ClearFormats(CellRange);
        }

        internal void DeleteRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
        {
            Cells.DeleteRange(CellRange, aRowCount, aColCount, SheetInfo);
            Drawing.DeleteRange(CellRange, aRowCount, aColCount, SheetInfo);

#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
            XlsxPivotTables.ArrangeInsertRange(CellRange, -aRowCount, -aColCount, SheetInfo);
#endif

            MergedCells.DeleteRange(CellRange, aRowCount, aColCount, SheetInfo);
            ConditionalFormats.DeleteRange(CellRange, aRowCount, aColCount, SheetInfo);
            DataValidation.DeleteRange(CellRange, aRowCount, aColCount, SheetInfo);
            Notes.DeleteRange(CellRange, aRowCount, aColCount, SheetInfo);
            HLinks.DeleteRange(CellRange, aRowCount, aColCount, SheetInfo);
            if (Window.Pane != null) Window.Pane.ArrangeInsertRange(CellRange, -aRowCount, -aColCount, Window.Window2.IsFrozen);

            if ((aRowCount > 0) && (CellRange.Left <= 0) && (CellRange.Right >= FlxConsts.Max_Columns))
            {
                SheetGlobals.HPageBreaks.DeleteRows(CellRange.Top, CellRange.RowCount * aRowCount);
            }
            if ((aColCount > 0) && (CellRange.Top <= 0) && (CellRange.Bottom >= FlxConsts.Max_Rows))
            {
                SheetGlobals.VPageBreaks.DeleteCols(CellRange.Left, CellRange.ColCount * aColCount);
            }
        }

        internal void MoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
        {
            Cells.MoveRange(CellRange, NewRow, NewCol, SheetInfo);
            Drawing.MoveRange(CellRange, NewRow, NewCol, SheetInfo);
            MergedCells.MoveRange(CellRange, NewRow, NewCol, SheetInfo);
            ConditionalFormats.MoveRange(CellRange, NewRow, NewCol, SheetInfo);
            DataValidation.MoveRange(CellRange, NewRow, NewCol, SheetInfo);
            Notes.MoveRange(CellRange, NewRow, NewCol, SheetInfo);
            HLinks.MoveRange(CellRange, NewRow, NewCol, SheetInfo);

#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
#endif

        }

        internal void ArrangeInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
        {
            Cells.ArrangeInsertRange(CellRange, aRowCount, aColCount, SheetInfo);
            Drawing.ArrangeInsertRange(CellRange, aRowCount, aColCount, SheetInfo);

#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
            XlsxPivotTables.ArrangeInsertRange(CellRange, aRowCount, aColCount, SheetInfo);
#endif
            Chart.ArrangeInsertRange(CellRange, aRowCount, aColCount, SheetInfo);
        }

        internal void ArrangeMoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
        {
            Cells.ArrangeMoveRange(CellRange, NewRow, NewCol, SheetInfo);
            Drawing.ArrangeMoveRange(CellRange, NewRow, NewCol, SheetInfo);
            Chart.ArrangeMoveRange(CellRange, NewRow, NewCol, SheetInfo);
#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
            XlsxPivotTables.ArrangeMoveRange(CellRange, NewRow, NewCol, SheetInfo);
#endif

        }

        internal void ArrangeCopySheet(TSheetInfo SheetInfo)
        {
            Cells.ArrangeInsertSheet(SheetInfo);
            Drawing.ArrangeCopySheet(SheetInfo);
            Chart.ArrangeCopySheet(SheetInfo);
        }

        internal void UpdateDeletedRanges(TDeletedRanges DeletedRanges)
        {
            Cells.CellList.UpdateDeletedRanges(DeletedRanges);
            DataValidation.UpdateDeletedRanges(DeletedRanges);
            ConditionalFormats.UpdateDeletedRanges(DeletedRanges);
            Drawing.UpdateDeletedRanges(DeletedRanges);
            Chart.UpdateDeletedRanges(DeletedRanges);
        }

        #endregion

        #region Fix
        internal void FixGuts()
        {
            SheetGlobals.PrepareToSave(Cells, Columns);
        }

        internal void DeleteEmptyRowRecords()
        {
            Cells.DeleteEmptyRowRecords(DefRowHeight, DefRowFlags, Notes);
        }

        internal void FixNotes()
        {
            Notes.FixDwgOfs(this);
        }
        
        internal void FixRadioButtons(ExcelFile Workbook, int ActiveSheet)
        {
            Drawing.FixRadioButtons(Workbook, ActiveSheet);
        }

        #endregion

        #region Pxl
        internal void MergeFromPxlSheet(TSheet Source)
        {
            SheetGlobals.MergeFromPxl(Source.SheetGlobals);
            Columns.Clear();
            Columns.CopyFrom(Source.Columns);
            Cells.MergeFromPxlCells(Source.Cells);

            Window.MergeFromPxl(Source.Window);

            Window.Selection = Source.Window.Selection;
        }
        #endregion

        #region Images
        internal void RestoreObjectCoords()
        {
            Drawing.RestoreObjectCoords(this);
        }

        #endregion

        #region Merged Cells
        internal TXlsCellRange CellMergedBounds(int aRow, int aCol)
        {
            //Find the cell into the MergedCells array
            TXlsCellRange Result = new TXlsCellRange();
            Result.Left = aCol;
            Result.Right = aCol;
            Result.Top = aRow;
            Result.Bottom = aRow;
            for (int i = 0; i < MergedCells.Count; i++)
            {
                TMergedCells mc = MergedCells[i] as TMergedCells;
                if (mc != null)
                    if (mc.CheckCell(aRow, aCol, Result)) return Result;
            }

            return Result;
        }

        internal TXlsCellRange CellMergedList(int index)
        {
            //Find the cell into the MergedCells array
            TXlsCellRange Result = new TXlsCellRange();
            Result.Left = 0;
            Result.Right = 0;
            Result.Top = 0;
            Result.Bottom = 0;

            if (index < 0) return Result;
            int p = 0;
            for (int i = 0; i < MergedCells.Count; i++)
            {
                TMergedCells mc = (MergedCells[i] as TMergedCells);

                int k = mc.MergedCount();
                if (index < p + k)
                {
                    Result = mc.MergedCell(index - p);
                    return Result;
                }
                p += k;
            }
            return Result;
        }

        internal int CellMergedListCount()
        {
            int Result = 0;
            for (int i = 0; i < MergedCells.Count; i++)
                Result += ((TMergedCells)MergedCells[i]).MergedCount();
            return Result;
        }

        internal void MergeCells(int aRow1, int aCol1, int aRow2, int aCol2)
        {
            if (aRow1 > aRow2) { int x = aRow2; aRow2 = aRow1; aRow1 = x; }
            if (aCol1 > aCol2) { int y = aCol2; aCol2 = aCol1; aCol1 = y; }

            //We have to take all existing included merged cells

            TMergedCells Mc = null;
            int bRow1 = aRow1; int bRow2 = aRow2; int bCol1 = aCol1; int bCol2 = aCol2;
            do
            {
                bRow1 = aRow1; bRow2 = aRow2; bCol1 = aCol1; bCol2 = aCol2;
                for (int i = 0; i < MergedCells.Count; i++)
                {
                    Mc = ((TMergedCells)MergedCells[i]);
                    Mc.PreMerge(ref aRow1, ref aCol1, ref aRow2, ref aCol2);
                }
            }
            while ((aRow1 != bRow1) || (aRow2 != bRow2) || (aCol1 != bCol1) || (aCol2 != bCol2));

            if (Mc == null)
            {
                Mc = new TMergedCells();
                MergedCells.Add(Mc);
            }
            Mc.MergeCells(aRow1, aCol1, aRow2, aCol2);

            //We are not copying the value here as this is not what Excel does anyway.
            //Arranging the format of the merged cells is left to the user.
            //FixOneMergedCell(aRow1, aCol1, aRow2, aCol2);
        }

        internal void UnMergeCells(int aRow1, int aCol1, int aRow2, int aCol2)
        {
            if (aRow1 > aRow2) { int x = aRow2; aRow2 = aRow1; aRow1 = x; }
            if (aCol1 > aCol2) { int y = aCol2; aCol2 = aCol1; aCol1 = y; }

            for (int i = 0; i < MergedCells.Count; i++)
            {
                ((TMergedCells)MergedCells[i]).UnMergeCells(aRow1, aCol1, aRow2, aCol2);
            }
        }
        #endregion

        #region Autofilter

        /// <summary>
        /// Changes the AutoFilter record to fit the "FilterDatabase" named range. Failing to do this will cause AV in the arrows outside the columns in AutoFilter
        /// </summary>
        internal void FixAutoFilter(int Sheet)
        {
            TXlsCellRange R = GetAutoFilterRange(Sheet);
            if (R == null || R.Top > R.Bottom || R.Left > R.Right) return;

            if (SortAndFilter.AutoFilter.AutoFilterInfo < 0) return;
            SortAndFilter.AutoFilter.AutoFilterInfo = R.Right - R.Left + 1;
        }


        internal TXlsCellRange GetAutoFilterRange(int Sheet)
        {
            if (SortAndFilter.AutoFilter.AutoFilterInfo < 0) return null;
            TNameRecord Name = FWorkbookGlobals.Names.GetName(Sheet, TXlsNamedRange.GetInternalName(InternalNameRange.Filter_DataBase));
            if (Name == null) return null;

            return new TXlsCellRange(Name.R1, Name.C1, Name.R2, Name.C2);

        }

        private int[][] GetColumnsForAutoFilter(int Row, int Col1, int Col2)
        {
            List<int[]> al = new List<int[]>();
            int ci = Col1;
            while (ci <= Col2)
            {
                TXlsCellRange R = CellMergedBounds(Row, ci);
                if (R.Left == ci)
                {
                    al.Add(new int[] { R.Left, R.Right });
                }
                ci = R.Right + 1;
            }

            int[][] Cols = al.ToArray();
            return Cols;
        }

        internal void SetAutoFilter(int SheetIndex, int Row, int Col1, int Col2)
        {
            int[][] Cols = GetColumnsForAutoFilter(Row, Col1, Col2);

            if (Cols.Length == 0)
            {
                RemoveAutoFilter();
                return;
            }

            SortAndFilter.AutoFilter.AutoFilterInfo = Math.Abs(Col2 - Col1 + 1); //here we need *all* columns, not just the merged ones. The last autofilters will AV when selected if that is not the case.

            Drawing.RemoveAutoFilter();
            Drawing.AddAutoFilter(Row, Cols, this);

            //AutoFilters refer to an internal name, and we have to add it.
            FWorkbookGlobals.Names.AddName(new TXlsNamedRange(TXlsNamedRange.GetInternalName(InternalNameRange.Filter_DataBase), SheetIndex,
                SheetIndex, Row, Col1, Row, Col2, 0x21, null),
                FWorkbookGlobals, Cells.CellList);

        }



        internal void AddNewAutoFilters(int Sheet, int Row1, int Row2, int DestCol1, int DestCol2)
        {
            TXlsCellRange R = GetAutoFilterRange(Sheet);
            if (R == null) return;

            if (R.Left > DestCol2 || R.Right < DestCol1 || R.Top > Row2 || R.Top < Row1) return;

            SetAutoFilter(Sheet, R.Top, R.Left, R.Right);

            //While this would be the "correct" thing to do instead of calling SetAutoFilter above, it will leave the blue
            //arrows in the wrong position when a range is filtered. (you can see this with autof3.xls test file).
            //Anyway, there is no way to choose many disjunct ranges for an autofilter, so there is no risk of setting it 
            //in cells it wasn't before.
            //int[][] Cols = GetColumnsForAutoFilter(R.Top, DestCol, DestCol + ColumnsInserted - 1);
            //FDrawing.AddAutoFilter(R.Top, Cols, this);
        }

        internal void RemoveAutoFilter()
        {
            SortAndFilter.AutoFilter.AutoFilterInfo = -1;
            SortAndFilter.AutoFilter.Filters.Clear();
            SortAndFilter.AutoFilter.FutureStorage = null;
            SortAndFilter.AutoFilter.Sort12.Clear();
            SortAndFilter.FilterMode = false;

            Drawing.RemoveAutoFilter();
        }

        internal bool HasAutoFilter()
        {
            return SortAndFilter.AutoFilter.AutoFilterInfo >= 0;
        }

        internal bool HasAutoFilter(int sheet, int row, int col)
        {
            if (!HasAutoFilter()) return false;

            TXlsCellRange R = GetAutoFilterRange(sheet);
            if (R == null) return false;

            return row >= R.Top && row <= R.Bottom && col >= R.Left && col <= R.Right;
        }

        #endregion

        #region Window
        internal void ScrollWindow(TPanePosition Pane, int row, int col)
        {
            if (Pane == TPanePosition.UpperLeft || Pane == TPanePosition.UpperRight)
                Window.Window2.FirstRow = row;

            if (Pane == TPanePosition.UpperLeft || Pane == TPanePosition.LowerLeft)
                Window.Window2.FirstCol = col;

            if (Window.Pane != null)
            {
                if (Pane == TPanePosition.LowerLeft || Pane == TPanePosition.LowerRight)
                    Window.Pane.FirstVisibleRow = row;

                if (Pane == TPanePosition.UpperRight || Pane == TPanePosition.LowerRight)
                    Window.Pane.FirstVisibleCol = col;
            }
        }

        internal TCellAddress GetWindowScroll(TPanePosition Pane)
        {
            int r = 0;
            int c = 0;

            if (Pane == TPanePosition.UpperLeft || Pane == TPanePosition.UpperRight)
                r = Window.Window2.FirstRow;

            if (Pane == TPanePosition.UpperLeft || Pane == TPanePosition.LowerLeft)
                c = Window.Window2.FirstCol;

            if (Window.Pane != null)
            {
                if (Pane == TPanePosition.LowerLeft || Pane == TPanePosition.LowerRight)
                    r = Window.Pane.FirstVisibleRow;

                if (Pane == TPanePosition.UpperRight || Pane == TPanePosition.LowerRight)
                    c = Window.Pane.FirstVisibleCol;
            }
            return new TCellAddress(r, c);
        }

        internal void FreezePanes(int row, int col)
        {
            bool Frost = row > 0 || col > 0;
            Window.Window2.IsFrozen = Frost;
            Window.Window2.IsFrozenButNoSplit = Frost;

            AddOrRemovePane(Frost);
            if (Window.Pane == null)
                return;

            int row1 = row < 1 ? 1 : row;
            int col1 = col < 1 ? 1 : col;

            Window.Pane.RowSplit = row;
            if (Window.Pane.FirstVisibleRow < row1) Window.Pane.FirstVisibleRow = row1;
            Window.Pane.ColSplit = col;
            if (Window.Pane.FirstVisibleCol < col1) Window.Pane.FirstVisibleCol = col1;

            Window.Pane.EnsureSelectedVisible();
        }

        internal void SplitWindow(int xOffset, int yOffset)
        {
            bool Frost = xOffset > 0 || yOffset > 0;
            Window.Window2.IsFrozen = false;
            Window.Window2.IsFrozenButNoSplit = false;

            AddOrRemovePane(Frost);
            if (Window.Pane == null)
                return;

            Window.Pane.RowSplit = yOffset;
            Window.Pane.ColSplit = xOffset;

            if (Window.Pane.FirstVisibleRow < 1) Window.Pane.FirstVisibleRow = 1;
            if (Window.Pane.FirstVisibleCol < 1) Window.Pane.FirstVisibleCol = 1;

            Window.Pane.EnsureSelectedVisible();
        }

        internal TCellAddress GetFrozenPanes()
        {
            if (Window.Pane != null && Window.Window2.IsFrozen)
            {
                return new TCellAddress(Window.Pane.RowSplit, Window.Pane.ColSplit);
            }
            return new TCellAddress(0, 0);
        }

        internal TPanePosition GetActivePaneForSelection()
        {
            if (Window.Pane == null || !Window.Window2.IsFrozen) return TPanePosition.UpperLeft;
            return Window.Pane.ActivePaneForSelection();
        }

        internal TPoint GetSplitWindow()
        {
            if (Window.Pane != null && !Window.Window2.IsFrozen)
            {
                return new TPoint(Window.Pane.ColSplit, Window.Pane.RowSplit);
            }
            return new TPoint(0, 0);
        }

        protected void AddOrRemovePane(bool Add)
        {
            if (Add)
            {
                if (Window.Pane == null) Window.Pane = new TPaneRecord((int)xlr.PANE, new byte[10]);
            }
            else
            {
                Window.Pane = null;
            }
        }

        #endregion

        #region Conditional Formats
        internal TFlxFormat ConditionallyModifyFormat(ExcelFile Xls, int SheetIndex, TFlxFormat Format, int aRow, int aCol)
        {
            if (ConditionalFormats.Modified != -1)
            {
                if (ConditionalFormats.Modified != 0) Cells.CellList.CleanRowCF();
                ConditionalFormats.Modified = -1;
                for (int i = ConditionalFormats.Count - 1; i >= 0; i--)  //loop must go descending.
                {
                    ((TCondFmt)ConditionalFormats[i]).UpdateCFRows(Cells.CellList);
                }
            }

            TCellCondFmt[] cf = Cells.CellList.GetRowCondFmt(aRow);
            if (cf == null) return null;

            int p = Array.BinarySearch(cf, new TCellCondFmt(aCol, aCol, null));

            if (p < 0) p = ~p - 1; //binary search didn't found exact match.

            if (p < 0 || p >= cf.Length) return null; //the smallest element is bigger than the col we are searching for.
            if (cf[p].C1 <= aCol && cf[p].C2 >= aCol)
            {
                TCondFmt fmt = cf[p].Fmt;
                if (fmt != null)
                {
                    TFlxFormat Result = (TFlxFormat)Format.Clone();
                    fmt.ConditionallyModifyFormat(Result, Xls, SheetIndex, aRow, aCol);
                    return Result;
                }
            }
            return null;
        }


        internal void SetConditionalFormat(int aRow1, int aCol1, int aRow2, int aCol2, TConditionalFormatRule[] ConditionalFormat)
        {
            if (aRow1 > aRow2) { int x = aRow2; aRow2 = aRow1; aRow1 = x; }
            if (aCol1 > aCol2) { int y = aCol2; aCol2 = aCol1; aCol1 = y; }

            /*This code would not work on Excel2007
            //Find the correct conditional format if there is one.

            bool Added = false;
            for (int i = FConditionalFormats.Count - 1; i >= 0;i--)
            {
                TCondFmt Fmt=((TCondFmt)FConditionalFormats[i]);
                if (Fmt.EqualsDef(ConditionalFormat, FWorkbookGlobals.Names, Cells.CellList))
                {
                    Fmt.AddRange(aRow1, aCol1, aRow2, aCol2);
                    Added = true;
                }
                else
                {
                    Fmt.ClearRange(aRow1, aCol1, aRow2, aCol2);
                }
            }

            if (!Added && ConditionalFormat != null && ConditionalFormat.Length > 0)
            {
                TCondFmt Fmt = new TCondFmt(aRow1, aCol1, aRow2, aCol2, ConditionalFormat, FWorkbookGlobals.Names, Cells.CellList);
                FConditionalFormats.Add(Fmt);
            }
            */

            TCondFmt Fmt = new TCondFmt(aRow1, aCol1, aRow2, aCol2, ConditionalFormat, Cells.CellList);
            ConditionalFormats.Add(Fmt);

        }
        #endregion

        #region Printing
        internal TPrinterDriverSettings PrinterDriverSettings
        {
            get
            {
                if (PageSetup.Pls == null) return null;
                else
                    return new TPrinterDriverSettings(PageSetup.Pls.Data);
            }

            set
            {
                if (value == null || value.GetData() == null)
                {
                    PageSetup.Pls = null;
                }
                else
                {
                    byte[] aData = new byte[value.GetData().Length];
                    value.GetData().CopyTo(aData, 0);
                    PageSetup.Pls = new TPlsRecord((int)xlr.PLS, aData);
                    PrintOptions = (byte)(PrintOptions & ~0x4);

                    //Not needed. If you want to copy the PaperSize, just use PaperSize property. 
                    //PaperSize=(TPaperSize)BitOps.GetWord(aData, 80);
                }
            }
        }

        internal int PrintOptions
        {
            get
            {
                  return PageSetup.Setup.GetPrintOptions(SheetType != TSheetType.Chart);
            }
            set
            {
                if ((value < 0) || (value > 0xFFFF))
                    XlsMessages.ThrowException(XlsErr.ErrXlsIndexOutBounds, value, "PrintOptions", 0, 0xFFFF);
                PageSetup.Setup.SetPrintOptions(value);
            }
        }

        internal TPaperSize PaperSize
        {
            get
            {
                return (TPaperSize)PageSetup.Setup.PaperSize;
            }
            set
            {
                if (((int)value < 0) || ((int)value > 0xFFFF))
                    XlsMessages.ThrowException(XlsErr.ErrXlsIndexOutBounds, value, "PrintPageSize", 0, 0xFFFF);
                PageSetup.Setup.PaperSize = (int)value;
            }
        }
        #endregion

        #region Custom Views
        internal TPageSetup GetCustomViewSetup(Guid CustomView)
        {
            if (CustomView.Equals(Guid.Empty)) return PageSetup;
            TCustomView cv;
            if (!CustomViews.Find(CustomView, out cv)) return null;
            return cv.Setup;
        }

        internal TAutoFilter GetCustomViewAutoFilter(Guid CustomView)
        {
            if (CustomView.Equals(Guid.Empty)) return SortAndFilter.AutoFilter;
            TCustomView cv;
            if (!CustomViews.Find(CustomView, out cv)) return null;
            return cv.AutoFilter;
        }

        #endregion


        /// <summary>
        /// An empty xlsx file has different defaults than an empty xls one.
        /// </summary>
        internal void InitXlsx()
        {
            PageSetup.HeaderAndFooter.AlignMargins = true;
            PageSetup.Setup.PaperSize = 1;
            PageSetup.Setup.HPrintRes = 600;
            PageSetup.Setup.VPrintRes = 600;
        }

        #region Protection
        internal TSheetProtectionOptions GetSheetProtectionOptions()
        {
            TSheetProtectionOptions Result = new TSheetProtectionOptions();
            TSheetProtection Sp = SheetProtection;
            Result.Contents = Sp.Protect != null && Sp.Protect.Protected;
            Result.Objects = Sp.ObjProtect == null || Sp.ObjProtect.Protected;
            Result.Scenarios = Sp.ScenProtect == null || Sp.ScenProtect.Protected;

            if (Sp.SheetProtect != null)
            {
                Result.CellFormatting = Sp.SheetProtect.GetProtect(2);
                Result.ColumnFormatting = Sp.SheetProtect.GetProtect(3);
                Result.RowFormatting = Sp.SheetProtect.GetProtect(4);
                Result.InsertColumns = Sp.SheetProtect.GetProtect(5);
                Result.InsertRows = Sp.SheetProtect.GetProtect(6);
                Result.InsertHyperlinks = Sp.SheetProtect.GetProtect(7);
                Result.DeleteColumns = Sp.SheetProtect.GetProtect(8);
                Result.DeleteRows = Sp.SheetProtect.GetProtect(9);
                Result.SelectLockedCells = Sp.SheetProtect.GetProtect(10);
                Result.SortCellRange = Sp.SheetProtect.GetProtect(11);
                Result.EditAutoFilters = Sp.SheetProtect.GetProtect(12);
                Result.EditPivotTables = Sp.SheetProtect.GetProtect(13);
                Result.SelectUnlockedCells = Sp.SheetProtect.GetProtect(14);
            }
            else
            {
                Result.SelectLockedCells = true;
                Result.SelectUnlockedCells = true;
            }

            return Result;
        }

        internal void SetSheetProtectionOptions(TSheetProtectionOptions value)
        {
            TSheetProtection Sp = SheetProtection;

            if (value == null) value = new TSheetProtectionOptions(false);
            if (Sp.Protect == null) Sp.Protect = new TProtectRecord();
            Sp.Protect.Protected = value.Contents;

            if (Sp.ScenProtect == null) Sp.ScenProtect = new TScenProtectRecord();
            Sp.ScenProtect.Protected = value.Scenarios;

            if (Sp.ObjProtect == null) Sp.ObjProtect = new TObjProtectRecord();
            Sp.ObjProtect.Protected = value.Objects;

            if (Sp.SheetProtect == null) Sp.SheetProtect = new TSheetProtectRecord();
            Sp.SheetProtect.SetProtect(0, value.Objects);
            Sp.SheetProtect.SetProtect(1, value.Scenarios);
            Sp.SheetProtect.SetProtect(2, value.CellFormatting);
            Sp.SheetProtect.SetProtect(3, value.ColumnFormatting);
            Sp.SheetProtect.SetProtect(4, value.RowFormatting);
            Sp.SheetProtect.SetProtect(5, value.InsertColumns);
            Sp.SheetProtect.SetProtect(6, value.InsertRows);
            Sp.SheetProtect.SetProtect(7, value.InsertHyperlinks);
            Sp.SheetProtect.SetProtect(8, value.DeleteColumns);
            Sp.SheetProtect.SetProtect(9, value.DeleteRows);
            Sp.SheetProtect.SetProtect(10, value.SelectLockedCells);
            Sp.SheetProtect.SetProtect(11, value.SortCellRange);
            Sp.SheetProtect.SetProtect(12, value.EditAutoFilters);
            Sp.SheetProtect.SetProtect(13, value.EditPivotTables);
            Sp.SheetProtect.SetProtect(14, value.SelectUnlockedCells);
        }
        #endregion
    }

    internal class TFlxChart: TSheet
    {
        #region Variables
        internal TFrtInfoRecord FrtInfo;
        internal TClrtClientRecord ClrtClient;
        internal TPaletteRecord Palette;
        internal TSxViewLinkRecord ViewLink;
        internal TPivotChartBitsRecord PivotChartBits;
        internal TChartSBaseRefRecord SBaseRef;
        internal TSeriesData SeriesData;
        internal int Units;
        internal TMiscRecordList CrtMlFrt;
        internal bool Embedded;

        internal TBaseRecord RemainingData;
        #endregion

        #region Constructor
        public TFlxChart(TWorkbookGlobals aWorkbookGlobals, bool aEmbedded)
            : base(aWorkbookGlobals)
        {
            SeriesData = new TSeriesData(null);
            CrtMlFrt = new TMiscRecordList();
            Embedded = aEmbedded;
        }

        protected override TSheet CreateSheet(TWorkbookGlobals DestGlobals)
        {
            return new TFlxChart(DestGlobals, Embedded);
        }

        protected override void DoExtraCopy(TSheet DestSheet, TSheetInfo SheetInfo, bool CopyData)
        {
            TFlxChart Result = (TFlxChart)DestSheet;
            Result.ClrtClient = (TClrtClientRecord)TClrtClientRecord.Clone(ClrtClient, SheetInfo);
            Result.Palette = (TPaletteRecord)TPaletteRecord.Clone(Palette, SheetInfo);
            Result.FrtInfo = (TFrtInfoRecord)TFrtInfoRecord.Clone(FrtInfo, SheetInfo);
            Result.ViewLink = (TSxViewLinkRecord)TSxViewLinkRecord.Clone(ViewLink, SheetInfo);
            Result.PivotChartBits = (TPivotChartBitsRecord)TPivotChartBitsRecord.Clone(PivotChartBits, SheetInfo);
            Result.SBaseRef = (TChartSBaseRefRecord)TChartSBaseRefRecord.Clone(SBaseRef, SheetInfo);
            Result.SeriesData = SeriesData.Clone(SheetInfo);
            Result.CrtMlFrt = CopyMiscList(CrtMlFrt, SheetInfo);
            Result.Embedded = Embedded;
        }
        #endregion

        #region Sheet Type
        internal override TSheetType SheetType
        {
            get { return TSheetType.Chart; }
        }
        #endregion

        #region Save
        internal override void EnsureRequiredRecords()
        {
            base.EnsureRequiredRecords();
            if (PageSetup.ChartPrintSize == TChartPrintSize.NotDefined) PageSetup.ChartPrintSize = TChartPrintSize.DefinedInChart;
            SeriesData.EnsureRequiredRecords();
        }

        internal void Save(IDataStream DataStream, TSaveData SaveData, TXlsCellRange CellRange)
        {
            if ((sBOF == null) || (sEOF == null)) XlsMessages.ThrowException(XlsErr.ErrSectionNotLoaded);
            sBOF.SaveToStream(DataStream, SaveData, 0);
            if (FWorkbookGlobals.FileEncryption.WriteProt != null) FWorkbookGlobals.FileEncryption.WriteProt.SaveToStream(DataStream, SaveData, 0);
            if (SheetExt != null) SheetExt.SaveToStream(DataStream, SaveData, 0);
            WebPub.SaveToStream(DataStream, SaveData, 0);
            HeaderImages.SaveToStream(DataStream, SaveData);

            if (FrtInfo != null) FrtInfo.SaveToStream(DataStream, SaveData, 0); //Actually it should save here if CrtMlFrt.Count > 0, or before the first future record. But since this is too difficult to implement, we wil just save it here. It will crash excel if there are crtfrt and this record is not here.

            PageSetup.SaveToStream(DataStream, SaveData, Guid.Empty);

            if (BgPic != null) BgPic.SaveToStream(DataStream, SaveData, 0);
            //no fbi or fbi2

            if (ClrtClient != null) ClrtClient.SaveToStream(DataStream, SaveData, 0);
            SheetProtection.SaveFirstPart(DataStream, SaveData);
            if (Palette != null) Palette.SaveToStream(DataStream, SaveData, 0);

            if (ViewLink != null) ViewLink.SaveToStream(DataStream, SaveData, 0);
            if (PivotChartBits != null) PivotChartBits.SaveToStream(DataStream, SaveData, 0);
            if (SBaseRef != null) SBaseRef.SaveToStream(DataStream, SaveData, 0);

            Drawing.SaveToStream(DataStream, SaveData);
            TUnitsRecord.SaveRecord(DataStream, Units);
            Chart.SaveToStream(DataStream, SaveData);

            SeriesData.SaveToStream(DataStream, SaveData);
            if (!Embedded)
            {
                Window.SaveToStream(DataStream, SaveData, false);
                CustomViews.SaveToStream(DataStream, SaveData, CellRange);
            }

            if (CodeNameRecord != null) CodeNameRecord.SaveToStream(DataStream, SaveData, 0);
            CrtMlFrt.SaveToStream(DataStream, SaveData, 0);
            
            Feat.SaveToStream(DataStream, SaveData, 0);
            Feat11.SaveToStream(DataStream, SaveData, 0);
            FutureRecords.SaveToStream(DataStream, SaveData, 0);
            sEOF.SaveToStream(DataStream, SaveData, 0);
        }

        internal override void SaveToStream(IDataStream DataStream, TSaveData SaveData)
        {
            Save(DataStream, SaveData, null);
        }

        internal override void SaveRangeToStream(IDataStream DataStream, TSaveData SaveData, int SheetIndex, TXlsCellRange CellRange)
        {
            //Can't save a chart range
            Save(DataStream, SaveData, CellRange);
        }

        internal override void SaveToPxl(TPxlStream PxlStream, TPxlSaveData SaveData)
        {
            // Nothing, pxl has no charts.
        }

        protected override long TotalSheetSize(TXlsCellRange CellRange)
        {
            long Result = 0;
            if (FWorkbookGlobals.FileEncryption.WriteProt != null) Result += FWorkbookGlobals.FileEncryption.WriteProt.TotalSize();
            if (SheetExt != null) Result += SheetExt.TotalSize();
            Result += WebPub.TotalSize;
            Result += HeaderImages.TotalSize();
            if (FrtInfo != null) Result += FrtInfo.TotalSize();
            Result += PageSetup.TotalSize();

            if (BgPic != null) Result += BgPic.TotalSize();

            if (ClrtClient != null) Result += ClrtClient.TotalSize();
            Result += SheetProtection.TotalSizeFirst();
            if (Palette != null) Result += Palette.TotalSize();
            
            if (ViewLink != null) Result += ViewLink.TotalSize();
            if (PivotChartBits != null) Result += PivotChartBits.TotalSize();
            if (SBaseRef != null) Result += SBaseRef.TotalSize();


            Result += Drawing.TotalSize();
            Result += TUnitsRecord.StandardSize();

            Result += Chart.TotalSize();

            Result += SeriesData.TotalSize();

            if (!Embedded)
            {
                Result += Window.TotalSize(false);
                Result += CustomViews.TotalSize(CellRange);
            }

            if (CodeNameRecord != null) Result += CodeNameRecord.TotalSize();

            Result += CrtMlFrt.TotalSize;
            Result += Feat.TotalSize;
            Result += Feat11.TotalSize;
            Result += FutureRecords.TotalSize;

            return Result;
        }

        #endregion

        internal override void LoadFromStream(TBaseRecordLoader RecordLoader, TBOFRecord First)
        {
            base.LoadFromStream(RecordLoader, First);
            if (sEOF.Continue != null)
            {
                RemainingData = sEOF.Continue;
                sEOF.Continue = null;
            }

            if (OriginalDimensions != null) SeriesData.Dimensions = new TXlsCellRange((int)OriginalDimensions.FirstRow(),
                (int)OriginalDimensions.FirstCol(), (int)OriginalDimensions.LastRow() - 1, (int)OriginalDimensions.LastCol() - 1);
        }

        internal override int ChartCount
        {
            get
            {
                return 1;
            }
        }

    }


	/// <summary>
	/// An Excel Worksheet or a dialog sheet.
	/// </summary>
	internal class TWorkSheet: TSheet
    {

        #region Constructor
        internal TWorkSheet(TWorkbookGlobals aWorkbookGlobals)
            : base(aWorkbookGlobals)
        {
            Columns.AllowStandardWidth = true; //If it is a dialog, it will be reset by the loadintoworksheet.
        }

        protected override TSheet CreateSheet(TWorkbookGlobals DestGlobals)
        {
            return new TWorkSheet(DestGlobals);
        }
        #endregion

        #region Sheet Type
        internal override TSheetType SheetType
        {
            get { if (SheetGlobals.WsBool.Dialog) return TSheetType.Dialog; return TSheetType.Worksheet; }
        }
        #endregion

        #region Save
        internal override void EnsureRequiredRecords()
        {
            base.EnsureRequiredRecords();
        }

        internal override void SaveToStream(IDataStream DataStream, TSaveData SaveData)
        {
            SaveGenericSheet(DataStream, SaveData, null);
        }

        internal override void SaveRangeToStream(IDataStream DataStream, TSaveData SaveData, int SheetIndex, TXlsCellRange CellRange)
        {
            SaveGenericSheet(DataStream, SaveData, CellRange);
        }

        internal override void SaveToPxl(TPxlStream PxlStream, TPxlSaveData SaveData)
        {
            if ((sBOF == null) || (sEOF == null)) XlsMessages.ThrowException(XlsErr.ErrSectionNotLoaded);
            sBOF.SaveToPxl(PxlStream, 0, SaveData);
            SheetGlobals.SaveToPxl(PxlStream, SaveData);

            Columns.SaveToPxl(PxlStream, SaveData);
            Cells.SaveToPxl(PxlStream, SaveData);

            Window.SaveToPxl(PxlStream, SaveData);

            Window.Selection.SaveToPxl(PxlStream, SaveData);

            sEOF.SaveToPxl(PxlStream, 0, SaveData);
        }
        #endregion

        #region New Worksheet

        internal static Stream GetEmptyWorkbook(TExcelFileFormat excelFileFormat)
        {
            switch (excelFileFormat)
            {
                case TExcelFileFormat.v2007:
                    return Assembly.GetExecutingAssembly().GetManifestResourceStream("FlexCel.XlsAdapter.EmptyWorkbook2007.xls");

                case TExcelFileFormat.v2010:
                    return Assembly.GetExecutingAssembly().GetManifestResourceStream("FlexCel.XlsAdapter.EmptyWorkbook2010.xls");

                default:
                    return Assembly.GetExecutingAssembly().GetManifestResourceStream("FlexCel.XlsAdapter.EmptyWorkbook.xls");
            }
        }

        internal static TWorkSheet CreateFromData(TWorkbookGlobals aWorkbookGlobals, TXlsBiffVersion aXlsBiffVersion, TExcelFileFormat aFileFormat)
        {
            TWorkSheet Result = new TWorkSheet(aWorkbookGlobals);
            Result.DoCreateFromData(aWorkbookGlobals, aXlsBiffVersion, aFileFormat);
            return Result;
        }

        private void DoCreateFromData(TWorkbookGlobals aWorkbookGlobals, TXlsBiffVersion aXlsBiffVersion, TExcelFileFormat aFileFormat)
        {
            TEncryptionData Encryption = new TEncryptionData(String.Empty, null, null);  //Resource is not encrypted.

            using (Stream MemStream = GetEmptyWorkbook(aFileFormat))
            {
                using (TOle2File DataStream = new TOle2File(MemStream))
                {
                    DataStream.SelectStream(XlsConsts.WorkbookString);

                    TBaseRecordLoader RecordLoader = new TXlsRecordLoader(DataStream, aWorkbookGlobals.Biff8XF, aWorkbookGlobals.SST, aWorkbookGlobals.Workbook,
                            aWorkbookGlobals.Borders, aWorkbookGlobals.Patterns, Encryption, aXlsBiffVersion, aWorkbookGlobals.Names, null);
                    RecordLoader.ReadHeader();
                    TBaseRecord R = null;

                    do
                    {
                        R = RecordLoader.LoadRecord(true);
                    }
                    while (!(R is TEOFRecord));

                    TBOFRecord BOF = (TBOFRecord)RecordLoader.LoadRecord(false);
                    LoadFromStream(RecordLoader, BOF);

                }
            }
        }
        #endregion

    }


    /// <summary>
    /// Excel 4.0 macro sheet.
    /// </summary>
    internal class TMacroSheet : TSheet
    {
        #region Variables
        private bool FInternational;
        #endregion

        #region Constructors
        public TMacroSheet(TWorkbookGlobals aWorkbookGlobals) : base(aWorkbookGlobals) 
        {
            Columns.AllowStandardWidth = true;
        }

        protected override TSheet CreateSheet(TWorkbookGlobals DestGlobals)
        {
            return new TMacroSheet(DestGlobals);
        }
        #endregion

        #region Sheet Type
        internal override TSheetType SheetType
        {
            get { return TSheetType.Macro; }
        }

        internal override bool International
        {
            get
            {
                return FInternational;
            }
            set
            {
                FInternational = value;
            }
        }
        #endregion

        #region Save
        internal override void EnsureRequiredRecords()
        {
            base.EnsureRequiredRecords();
        }

        internal override void SaveToStream(IDataStream DataStream, TSaveData SaveData)
        {
            SaveGenericSheet(DataStream, SaveData, null);
        }

        internal override void SaveRangeToStream(IDataStream DataStream, TSaveData SaveData, int SheetIndex, TXlsCellRange CellRange)
        {
            SaveGenericSheet(DataStream, SaveData, CellRange);
        }

        internal override void SaveToPxl(TPxlStream PxlStream, TPxlSaveData SaveData)
        {
            // Nothing, pxl has no macro 4 sheets.
        }
        #endregion
    }

    /// <summary>
    /// Something we don't support... shouldn't really happen
    /// </summary>
    internal class TFlxUnsupportedSheet : TSheet
    {
        #region Constructors
        public TFlxUnsupportedSheet(TWorkbookGlobals aWorkbookGlobals) : base(aWorkbookGlobals) { }

        protected override TSheet CreateSheet(TWorkbookGlobals DestGlobals)
        {
            return new TFlxUnsupportedSheet(DestGlobals);
        }
        #endregion

        #region Sheet Type
        internal override TSheetType SheetType
        {
            get { return TSheetType.Other; }
        }
        #endregion

        #region Load/ Save
		internal override void LoadFromStream(TBaseRecordLoader RecordLoader, TBOFRecord First)
		{
            int Level = 1;
			do
			{
				TBaseRecord R=RecordLoader.LoadUnsupportdedRecord();
      
				if (R!= null)  //Null will be ignored
				    FutureRecords.Add(R);

                if (R is TBOFRecord) Level++;
                if (R is TEOFRecord) Level--;
			}
			while (Level > 0);

            EnsureRequiredRecords();

            //this must be the last statement, so if there is an exception, we dont take First
			sBOF= First;

		}               

        internal override void EnsureRequiredRecords()
        {
            base.EnsureRequiredRecords();
        }

        internal override void SaveToStream(IDataStream DataStream, TSaveData SaveData)
        {
            if ((sBOF == null)) XlsMessages.ThrowException(XlsErr.ErrSectionNotLoaded);
            sBOF.SaveToStream(DataStream, SaveData, 0);
            FutureRecords.SaveToStream(DataStream, SaveData, 0);
        }

        internal override void SaveRangeToStream(IDataStream DataStream, TSaveData SaveData, int SheetIndex, TXlsCellRange CellRange)
        {
            SaveToStream(DataStream, SaveData);
        }

        internal override void SaveToPxl(TPxlStream PxlStream, TPxlSaveData SaveData)
        {
            //Nothing
        }

        internal override long TotalSize(TEncryptionData Encryption, bool Repeatable)
        {
            return sBOF.TotalSize() + FutureRecords.TotalSize;
        }

        internal override long TotalRangeSize(int SheetIndex, TXlsCellRange CellRange, TEncryptionData Encryption, bool Repeatable)
        {
            return TotalSize(Encryption, Repeatable);
        }
        #endregion
    }

    internal struct TLoaderInfo
    {
        internal TFormulaRecord LastFormula;
        internal int LastFormulaRow;
        private TShrFmlaRecordList FShrFmlas;
        internal bool Eof;
        internal TCustomView CustomView;

        internal TShrFmlaRecordList ShrFmlas { get { if (FShrFmlas == null) FShrFmlas = new TShrFmlaRecordList(); return FShrFmlas; } }
    }
    #endregion

    #region Shared sections

    internal class TSheetGlobals
    {
        #region Variables
        public bool PrintHeaders;
        public bool PrintGridLines;
        public bool GridSet;
        public TGutsRecord Guts;
        public TDefaultRowHeightRecord DefRowHeight;
        public TWsBool WsBool;
        public TSyncRecord Sync; //optional
        public TLprRecord Lpr; //optional
        public THPageBreakList HPageBreaks; //optional
        public TVPageBreakList VPageBreaks; //optional
        public TProtectedRangeList ProtectedRanges;
        #endregion

        internal TSheetGlobals()
        {
            GridSet = true;

            WsBool.Init();
            HPageBreaks = new THPageBreakList();
            VPageBreaks = new TVPageBreakList();
            ProtectedRanges = new TProtectedRangeList();
        }

        internal TSheetGlobals Clone(TSheetInfo SheetInfo)
        {
            TSheetGlobals Result = (TSheetGlobals)MemberwiseClone();

            Result.Guts = (TGutsRecord)TGutsRecord.Clone(Guts, SheetInfo);
            Result.DefRowHeight = (TDefaultRowHeightRecord)TDefaultRowHeightRecord.Clone(DefRowHeight, SheetInfo);
            Sync = (TSyncRecord)TSyncRecord.Clone(Sync, SheetInfo);
            Lpr = (TLprRecord)TLprRecord.Clone(Lpr, SheetInfo);

            Result.HPageBreaks = new THPageBreakList();
            Result.HPageBreaks.CopyFrom(HPageBreaks);
            Result.VPageBreaks = new TVPageBreakList();
            Result.VPageBreaks.CopyFrom(VPageBreaks);
            Result.ProtectedRanges = ProtectedRanges.Clone();
            return Result;            
        }

        internal void PrepareToSave(TCells CellList, TColInfoList Columns)
        {
            if (Guts.RecalcNeeded)
            {
                CellList.CellList.CalcRowGuts(Guts);
                Columns.CalcGuts(Guts);
                Guts.RecalcNeeded = false;
            }
        }

        internal void SaveToStream(TCells CellList, TColInfoList Columns, IDataStream DataStream, TSaveData SaveData, TXlsCellRange CellRange)
        {
            PrepareToSave(CellList, Columns);

            TPrintHeadersRecord.SaveRecord(DataStream, PrintHeaders);
            TPrintGridLinesRecord.SaveRecord(DataStream, PrintGridLines);
            TGridSetRecord.SaveRecord(DataStream, GridSet);
            Guts.SaveToStream(DataStream, SaveData, 0);
            DefRowHeight.SaveToStream(DataStream, SaveData, 0);
            TWsBoolRecord.SaveWsBool(DataStream, WsBool);
            if (Sync != null) Sync.SaveToStream(DataStream, SaveData, 0);
            if (Lpr != null) Lpr.SaveToStream(DataStream, SaveData, 0);

            
            HPageBreaks.SaveToStream(DataStream, SaveData, CellRange);
            VPageBreaks.SaveToStream(DataStream, SaveData, CellRange);
        }

        internal int TotalSize(TXlsCellRange CellRange)
        {
            return
            TPrintHeadersRecord.StandardSize() +
            TPrintGridLinesRecord.StandardSize() +
            TGridSetRecord.StandardSize() +
            Guts.TotalSize() +
            DefRowHeight.TotalSize() +
            (int)TWsBoolRecord.StandardSize() +
            (Sync != null? Sync.TotalSize(): 0) +
            (Lpr != null ? Lpr.TotalSize(): 0) +
            (int)HPageBreaks.TotalSize(CellRange) +
            (int)VPageBreaks.TotalSize(CellRange);

        }

        internal void EnsureRequiredRecords()
        {
            if (Guts == null) Guts = new TGutsRecord();
            if (DefRowHeight == null) DefRowHeight = new TDefaultRowHeightRecord();
        }

        internal void SaveToPxl(TPxlStream PxlStream, TPxlSaveData SaveData)
        {
            DefRowHeight.SaveToPxl(PxlStream, 0, SaveData);
        }

        internal void MergeFromPxl(TSheetGlobals SheetGlobals)
        {
            DefRowHeight = SheetGlobals.DefRowHeight;
        }

    }

    internal class TPageSetup
    {
        #region Variables
        public THeaderAndFooter HeaderAndFooter;//not optional in root, but optional in custom views.

        public bool HCenter;//not optional in root, but optional in custom views.
        public bool VCenter;//not optional in root, but optional in custom views.

        public double LeftMargin; //optional
        public double RightMargin;//optional
        public double TopMargin;//optional
        public double BottomMargin;//optional
        public TPlsRecord Pls;//optional
        public TSetupRecord Setup;  //not optional in root, but optional in custom views.

        public TChartPrintSize ChartPrintSize; //required in charts, not allowed anywhere else.
        #endregion

        public TPageSetup()
        {
            LeftMargin = -1;
            RightMargin = -1;
            TopMargin = -1;
            BottomMargin = -1;

            ChartPrintSize = TChartPrintSize.NotDefined;

            HeaderAndFooter.ScaleWithDoc = true;
            Setup = new TSetupRecord();
        }

        internal TPageSetup Clone(TSheetInfo SheetInfo)
        {
            TPageSetup Result = (TPageSetup)MemberwiseClone();
            Result.Pls = (TPlsRecord)TPlsRecord.Clone(Pls, SheetInfo);
            Setup = (TSetupRecord)TSetupRecord.Clone(Setup, SheetInfo);
            return Result;
        }

        internal void SaveToStream(IDataStream DataStream, TSaveData SaveData, Guid CustomView)
        {
            TPageHeaderRecord.SaveRecord(DataStream, xlr.HEADER, HeaderAndFooter.DefaultHeader);
            TPageFooterRecord.SaveRecord(DataStream, xlr.FOOTER, HeaderAndFooter.DefaultFooter);
            THCenterRecord.SaveRecord(DataStream, HCenter);
            TVCenterRecord.SaveRecord(DataStream, VCenter);

            TLeftMarginRecord.SaveRecord(DataStream, LeftMargin);
            TRightMarginRecord.SaveRecord(DataStream, RightMargin);
            TTopMarginRecord.SaveRecord(DataStream, TopMargin);
            TBottomMarginRecord.SaveRecord(DataStream, BottomMargin);
            if (Pls != null) Pls.SaveToStream(DataStream, SaveData, 0);
            if (Setup != null) Setup.SaveToStream(DataStream, SaveData, 0);

            if (ChartPrintSize != TChartPrintSize.NotDefined) TPrintSizeRecord.SaveRecord(DataStream, ChartPrintSize);
            THeaderFooterExtRecord.SaveRecord(DataStream, CustomView, HeaderAndFooter);
        }

        internal int TotalSize()
        {
            return
            TPageHeaderRecord.StandardSize(HeaderAndFooter.DefaultHeader) +
            TPageFooterRecord.StandardSize(HeaderAndFooter.DefaultFooter) +
            THCenterRecord.StandardSize() +
            TVCenterRecord.StandardSize() +

            TLeftMarginRecord.StandardSize(LeftMargin) +
            TRightMarginRecord.StandardSize(RightMargin) +
            TTopMarginRecord.StandardSize(TopMargin) +
            TBottomMarginRecord.StandardSize(BottomMargin) +
            (Pls != null ? Pls.TotalSize() : 0) +
            (Setup == null? 0: Setup.TotalSize())+
            TPrintSizeRecord.StandardSize(ChartPrintSize) +
            THeaderFooterExtRecord.StandardSize(HeaderAndFooter);
        }

        internal void EnsureRequiredRecords(bool IsChart)
        {
            if (Setup == null) Setup = new TSetupRecord();
            if (IsChart && ChartPrintSize == TChartPrintSize.NotDefined) ChartPrintSize = TChartPrintSize.DefinedInChart;
        }
    }

    internal class TWindow
    {
        #region Variables
        internal TWindow2Record Window2;
        internal TPlvRecord Plv; //optional
        internal TSCLRecord Scl; //optional
        internal TPaneRecord Pane; //optional
        internal TSheetSelection Selection;

        internal TFutureStorage FutureStorage;
        #endregion

        public TWindow()
        {
            Selection = new TSheetSelection();
        }

        public TWindow Clone(TSheetInfo SheetInfo)
        {
            TWindow Result = new TWindow();
            Result.Window2 = (TWindow2Record)TWindow2Record.Clone(Window2, SheetInfo);
            Result.Plv = (TPlvRecord)TPlvRecord.Clone(Plv, SheetInfo);
            Result.Scl = (TSCLRecord)TSCLRecord.Clone(Scl, SheetInfo);
            Result.Pane = (TPaneRecord)TPaneRecord.Clone(Pane, SheetInfo);
            Result.Selection = TSheetSelection.Clone(Selection);
            Result.FutureStorage = TFutureStorage.Clone(FutureStorage);
            return Result;
        }

        internal void SaveToStream(IDataStream DataStream, TSaveData SaveData, bool IsDialogSheet)
        {
            Window2.SaveToStream(DataStream, SaveData, 0);
            if (!IsDialogSheet)
            {
                if (Plv != null) Plv.SaveToStream(DataStream, SaveData, 0);
                if (Scl != null) Scl.SaveToStream(DataStream, SaveData, 0);
            }
            if (Pane != null) Pane.SaveToStream(DataStream, SaveData, 0);
            Selection.SaveToStream(DataStream, SaveData, this);
        }

        internal int TotalSize(bool IsDialogSheet)
        {
            int ExtraRecords =             
                (Plv != null ? Plv.TotalSize() : 0) +
                (Scl != null ? Scl.TotalSize() : 0);
            
            if (IsDialogSheet) ExtraRecords = 0;

            return
            Window2.TotalSize() +
            ExtraRecords +
            (Pane != null ? Pane.TotalSize() : 0) +
            (int)Selection.TotalSize(this);
        }

        internal void EnsureRequiredRecords(TSheetType SheetType)
        {
            if (Window2 == null) Window2 = new TWindow2Record(SheetType == TSheetType.Chart);
        }


        internal void SaveToPxl(TPxlStream PxlStream, TPxlSaveData SaveData)
        {
            Window2.SaveToPxl(PxlStream, 0, SaveData);
            if (Pane != null) Pane.SaveToPxl(PxlStream, 0, SaveData);
        }

        internal void MergeFromPxl(TWindow Source)
        {
            Window2 = Source.Window2;
            Pane = Source.Pane;
        }

        internal void AddFutureStorage(TFutureStorageRecord R)
        {
            TFutureStorage.Add(ref FutureStorage, R);
        }
    }

    internal class TScenarios
    {
        #region Variables
        internal TScenManRecord ScenMan;
        internal TMiscRecordList Scenarios;
        #endregion

        public TScenarios()
        {
            Scenarios = new TMiscRecordList();
        }

        internal TScenarios Clone(TSheetInfo SheetInfo)
        {
            TScenarios Result = new TScenarios();
            Result.ScenMan = (TScenManRecord)TScenManRecord.Clone(ScenMan, SheetInfo);
            Result.Scenarios.CopyFrom(Scenarios, SheetInfo);

            return Result;
        }

        internal void SaveToStream(IDataStream DataStream, TSaveData SaveData)
        {
            if (ScenMan != null)ScenMan.SaveToStream(DataStream, SaveData, 0);
            Scenarios.SaveToStream(DataStream, SaveData, 0);
        }

        internal int TotalSize()
        {
            return
                (ScenMan == null? 0: ScenMan.TotalSize()) +
            (int)Scenarios.TotalSize;
        }

        internal void EnsureRequiredRecords()
        {
        }
    }

    internal class TSortAndFilter
    {
        #region Variables
        internal TSortRecord Sort;
        internal TSortDataRecord SortData;
        internal bool FilterMode;
        internal TDropDownObjIdsRecord DropDownObjIds;
        internal TAutoFilter AutoFilter;

        #endregion

        public TSortAndFilter()
        {
            AutoFilter = new TAutoFilter();
        }

        internal TSortAndFilter Clone(TSheetInfo SheetInfo)
        {
            TSortAndFilter Result = (TSortAndFilter)MemberwiseClone();
            Result.Sort = (TSortRecord)TSortRecord.Clone(Sort, SheetInfo);
            Result.SortData = (TSortDataRecord)TSortDataRecord.Clone(SortData, SheetInfo);
            Result.DropDownObjIds = (TDropDownObjIdsRecord)TDropDownObjIdsRecord.Clone(DropDownObjIds, SheetInfo);
            Result.AutoFilter = AutoFilter.Clone(SheetInfo);

            return Result;
        }

        internal void SaveToStream(IDataStream DataStream, TSaveData SaveData, bool IsMacroSheet)
        {
            if (Sort != null) Sort.SaveToStream(DataStream, SaveData, 0);
            if (SortData != null) SortData.SaveToStream(DataStream, SaveData, 0);
            if (FilterMode && !IsMacroSheet) TFilterModeRecord.SaveRecord(DataStream);
            if (DropDownObjIds != null) DropDownObjIds.SaveToStream(DataStream, SaveData, 0);
            AutoFilter.SaveToStream(DataStream, SaveData);
        }

        internal int TotalSize(bool IsMacroSheet)
        {
            return
                (Sort == null ? 0 : Sort.TotalSize()) +
                (SortData == null ? 0 : SortData.TotalSize()) +
                (FilterMode && !IsMacroSheet? TFilterModeRecord.StandardSize() : 0) +
                (DropDownObjIds == null ? 0 : DropDownObjIds.TotalSize()) +
                AutoFilter.TotalSize();
        }

        internal void EnsureRequiredRecords()
        {
        }
    }

    internal class TAutoFilter
    {
        #region Variables
        internal int AutoFilterInfo;
        internal TMiscRecordList Filters;  //Define which filter is applied to column n
        internal TMiscRecordList Sort12;
        internal TFutureStorage FutureStorage;
        #endregion

        public TAutoFilter()
        {
            AutoFilterInfo = -1;
            Filters = new TMiscRecordList();
            Sort12 = new TMiscRecordList();
        }

        internal TAutoFilter Clone(TSheetInfo SheetInfo)
        {
            TAutoFilter Result = (TAutoFilter)MemberwiseClone();
            Result.Filters = new TMiscRecordList();
            Result.Filters.CopyFrom(Filters, SheetInfo);
            Result.Sort12 = new TMiscRecordList();
            Result.Sort12.CopyFrom(Sort12, SheetInfo);
            Result.FutureStorage = TFutureStorage.Clone(FutureStorage);

            return Result;
        }

        internal void SaveToStream(IDataStream DataStream, TSaveData SaveData)
        {
            if (AutoFilterInfo > 0)TAutoFilterInfoRecord.SaveRecord(DataStream, AutoFilterInfo);
            Filters.SaveToStream(DataStream, SaveData, 0);
            Sort12.SaveToStream(DataStream, SaveData, 0);
        }

        internal int TotalSize()
        {
            return
                (AutoFilterInfo > 0? TAutoFilterInfoRecord.StandardSize() : 0) +
                (int)Filters.TotalSize +
                (int)Sort12.TotalSize;
        }
    }

    internal class TPivotViewList
    {
        #region Variables
        internal TMiscRecordList PivotItems;
        #endregion

        public TPivotViewList()
        {
            PivotItems = new TMiscRecordList();
        }

        internal TPivotViewList Clone(TSheetInfo SheetInfo)
        {
            TPivotViewList Result = new TPivotViewList();
            Result.PivotItems.CopyFrom(PivotItems, SheetInfo);
            return Result;
        }

        internal void SaveToStream(IDataStream DataStream, TSaveData SaveData)
        {
            PivotItems.SaveToStream(DataStream, SaveData, 0);
        }

        internal long TotalSize()
        {
            return PivotItems.TotalSize;
        }
    }

    internal class TDConn
    {
        #region Variables
        internal TDConRecord DCon;
        internal TMiscRecordList DConList;
        #endregion

        public TDConn()
        {
            DConList = new TMiscRecordList();
        }

        internal TDConn Clone(TSheetInfo SheetInfo)
        {
            TDConn Result = new TDConn();
            Result.DCon = (TDConRecord)TDConRecord.Clone(DCon, SheetInfo);
            Result.DConList.CopyFrom(DConList, SheetInfo);

            return Result;
        }

        internal void SaveToStream(IDataStream DataStream, TSaveData SaveData)
        {
            if (DCon != null)
            {
                DCon.SaveToStream(DataStream, SaveData, 0);
                DConList.SaveToStream(DataStream, SaveData, 0);
            }
        }

        internal long TotalSize()
        {
            if (DCon == null) return 0;
            return DCon.TotalSize() + DConList.TotalSize;
        }

    }

    internal class TCustomViewList
    {
        private List<TCustomView> FList;
        private Dictionary<Guid, TCustomView> ViewCache;

        internal TCustomViewList()
        {
            FList = new List<TCustomView>();
            ViewCache = new Dictionary<Guid,TCustomView>();
        }

        internal TCustomViewList Clone(TSheetInfo SheetInfo)
        {
            TCustomViewList Result = new TCustomViewList();
            for (int i = 0; i < Result.FList.Count; i++)
            {
                Result.FList.Add(FList[i].Clone(SheetInfo));
                Result.ViewCache[FList[i].View.CustomView] = Result.FList[i];
            }

            return Result;
        }

        internal void SaveToStream(IDataStream DataStream, TSaveData SaveData, TXlsCellRange CellRange)
        {
            for (int i = 0; i < FList.Count; i++)
            {
                FList[i].SaveToStream(DataStream, SaveData, CellRange);
            }
        }

        internal long TotalSize(TXlsCellRange CellRange)
        {
            int Result = 0;
            for (int i = 0; i < FList.Count; i++)
            {
                Result += FList[i].TotalSize(CellRange);
            }
            return Result;
        }


        internal bool Find(Guid CustomView, out TCustomView ResultValue)
        {
            return ViewCache.TryGetValue(CustomView, out ResultValue);
        }

        internal TCustomView Add(TUserSViewBeginRecord UserView)
        {
            TCustomView ncv = new TCustomView(UserView);
            FList.Add(ncv);
            ViewCache[UserView.CustomView] = ncv;
            return ncv;
        }
    }

    internal class TCustomView
    {
        #region Variables
        internal TUserSViewBeginRecord View;

        internal TSheetSelection Selection;
        internal THPageBreakList HPageBreaks; //optional
        internal TVPageBreakList VPageBreaks; //optional
        internal TPageSetup Setup;
        internal TAutoFilter AutoFilter;
        #endregion

        public TCustomView(TUserSViewBeginRecord aView)
        {
            View = aView;
            Selection = new TSheetSelection();
            HPageBreaks = new THPageBreakList();
            VPageBreaks = new TVPageBreakList();
            Setup = new TPageSetup();
            AutoFilter = new TAutoFilter();
        }

        internal TCustomView Clone(TSheetInfo SheetInfo)
        {
            TCustomView Result = new TCustomView((TUserSViewBeginRecord) TUserSViewBeginRecord.Clone(View, SheetInfo));
            Result.Selection = TSheetSelection.Clone(Selection);
            Result.HPageBreaks.CopyFrom(HPageBreaks);
            Result.VPageBreaks.CopyFrom(VPageBreaks);
            Result.Setup = Setup.Clone(SheetInfo);
            Result.AutoFilter = AutoFilter.Clone(SheetInfo);
            return Result;
        }

        internal void SaveToStream(IDataStream DataStream, TSaveData SaveData, TXlsCellRange CellRange)
        {
            View.SaveToStream(DataStream, SaveData, 0);
            Selection.SaveToStream(DataStream, SaveData, null);
            HPageBreaks.SaveToStream(DataStream, SaveData, CellRange);
            VPageBreaks.SaveToStream(DataStream, SaveData, CellRange);
            Setup.SaveToStream(DataStream, SaveData, View.CustomView);
            AutoFilter.SaveToStream(DataStream, SaveData);
            TUserSViewEndRecord.SaveRecord(DataStream);
        }

        internal int TotalSize(TXlsCellRange CellRange)
        {
            return
                View.TotalSize() +
                (int)Selection.TotalSize(null) +
                (int)HPageBreaks.TotalSize(CellRange) +
                (int)VPageBreaks.TotalSize(CellRange) +
                Setup.TotalSize() +
                AutoFilter.TotalSize() +
                TUserSViewEndRecord.StandardSize();
        }

    }

    internal class TQueryTableList
    {
        #region Variables
        internal TMiscRecordList QueryItems;
        #endregion

        public TQueryTableList()
        {
            QueryItems = new TMiscRecordList();
        }

        internal TQueryTableList Clone(TSheetInfo SheetInfo)
        {
            TQueryTableList Result = new TQueryTableList();
            Result.QueryItems.CopyFrom(QueryItems, SheetInfo);
            return Result;
        }

        internal void SaveToStream(IDataStream DataStream, TSaveData SaveData)
        {
            QueryItems.SaveToStream(DataStream, SaveData, 0);
        }

        internal int TotalSize()
        {
            return
                (int)QueryItems.TotalSize;
        }


    }

    #endregion

    #region Chart

    internal class TChartDef
    {
        #region Variables
        private TWorkbookGlobals FWorkbookGlobals;
        internal TChartRecordList ChartRecords;
		private TChartChartRecord ChartCache;
        #endregion

        #region Constructor and load
        internal TChartDef(TWorkbookGlobals aWorkbookGlobals)
        {
            FWorkbookGlobals = aWorkbookGlobals;
            ChartRecords = new TChartRecordList();
        }

        internal TChartDef Clone(TSheetInfo SheetInfo)
        {
            TChartDef Result = new TChartDef(SheetInfo.DestGlobals);
            Result.ChartRecords = new TChartRecordList();
            Result.ChartRecords.CopyFrom(ChartRecords, SheetInfo);
            Result.FixCachePointers();
            return Result;
        }

        private void FixCachePointers()
        {
            ChartCache = null;

            for (int i = 0; i < ChartRecords.Count; i++)
            {
                TBaseRecord R = ChartRecords[i] as TBaseRecord;
                if (R != null) //CachePointers are only on chart sheets, so they are on first level. No need to recurse.
                {
                    TChartChartRecord Cr = R as TChartChartRecord;
                    if (Cr != null)
                    {
                        ChartCache = Cr;
                        return;
                    }
                }
            }
        }

        internal void LoadFromStream(TBaseRecordLoader RecordLoader, TChartChartRecord First)
        {
            int Level = 0;
            ChartRecords.Add(First);
            ChartCache = First;

            TChartRecordList CurrentChart = ChartRecords;
            Stack<TChartRecordList> RecordStack = new Stack<TChartRecordList>();
            TChartCache MasterCache = ChartRecords.GetCache();
            int RecordId = 0;
            do
            {
                RecordId = RecordLoader.RecordHeader.Id;
                int Row;
                TBaseRecord R = RecordLoader.LoadRecord(out Row, false);

                switch ((xlr)RecordId)
                {
                    case xlr.BEGIN:
                        TxChartBaseRecord RC = CurrentChart[CurrentChart.Count - 1] as TxChartBaseRecord;
                        if (RC == null)
                        {
                            XlsMessages.ThrowException(XlsErr.ErrInvalidChart);
                        }
                        RC.CreateChildren(MasterCache);
                        RC.Children.Add(R);
                        RecordStack.Push(CurrentChart);
                        CurrentChart = RC.Children;
                        Level++;
                        break;

                    case xlr.END:
                        Level--;
                        CurrentChart.Add(R);
                        if (Level < 0) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
                        CurrentChart = RecordStack.Pop();
                        break;

                    case xlr.ChartFbi:
                        TChartFBIRecord Fbi = R as TChartFBIRecord;
                        int FontId = Fbi.FontId;
                        if (FontId > 4) FontId--;
                        if (FontId > 0 && FontId < FWorkbookGlobals.Fonts.Count)
                        FWorkbookGlobals.Fonts[FontId].Reuse = false;
                        Fbi.Globals = FWorkbookGlobals;
                        CurrentChart.Add(R);
                        break; 


                    default:
                        CurrentChart.Add(R);
                        break;
                }

            }
            while (RecordId != (int)xlr.END || Level > 0);
        } 

 		internal void Clear()
		{
			ChartRecords.Clear();
        }
        #endregion

        #region Save
        internal void SaveToStream(IDataStream DataStream, TSaveData SaveData)
        {
            ChartRecords.SaveToStream(DataStream, SaveData, 0);
        }

        internal long TotalSize()
        {
            return ChartRecords.TotalSize;
        }

        #endregion

        #region InsertAndCopy
        internal void ArrangeCopyRange(int RowOfs, int ColOfs, TSheetInfo SheetInfo)
		{
			ChartRecords.ArrangeCopyRange(RowOfs, ColOfs, SheetInfo);
		}


		internal void ArrangeCopySheet(TSheetInfo SheetInfo)
		{
			ChartRecords.ArrangeCopySheet(SheetInfo);
		}

		internal void ArrangeInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
		{
			ChartRecords.ArrangeInsertRange(CellRange, aRowCount, aColCount, SheetInfo);
		}

		internal void ArrangeMoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
		{
			ChartRecords.ArrangeMoveRange(CellRange, NewRow, NewCol, SheetInfo);
		}

		
		internal bool HasExternRefs()
		{
			return ChartRecords.HasExternRefs();
        }

        #endregion

        #region Chart properties
        internal TChartCache Cache
		{
			get
			{
				return ChartRecords.GetCache();
			}
		}


		internal TChartChartRecord GetChartCache
		{
			get
			{
				return ChartCache;
			}
		}

		internal void Add(TxChartBaseRecord Record)
		{
			ChartRecords.Add(Record);
		}

		internal void DeleteSeries(int index)
		{
            int Count = 0;

            int i = 0;
            while (i < ChartCache.Children.Count)
            {
                if (ChartCache.Children[i] is TChartSeriesRecord)
                {
                    if (Count == index)
                    {
                        ChartCache.Children.Delete(i);
                        Count++;
                        continue; //do not inc i.
                    }
                    Count++;
                }

                TChartTextRecord CText = ChartCache.Children[i] as TChartTextRecord;
                if (CText != null && CText.SeriesIndex() == index)
                {
                    ChartCache.Children.Delete(i);
                    continue; // do not inc i.
                }

                i++;
            }
        }

		internal void AddSeries(TChartSeriesRecord Series)
		{
			int position = 0;

			for (int i = ChartCache.Children.Count - 1; i >= 0; i--)
			{
				if (ChartCache.Children[i] is TChartSeriesRecord)
				{
					position = i;
					break;
				}
			}

			ChartCache.Children.Insert(position + 1, Series);
		}

		internal T FindRec<T>() where T:TChartBaseRecord
		{
			T Result;
			if (ChartRecords.FindRec(out Result)) return Result;
			return null;
		}

		internal TChartAxis[] GetChartAxis(XlsFile Workbook, TSheet CurrentSheet, TCellList CellList, double FontScaling)
		{
			List<TBaseRecord> Result1 = ChartCache.FindAllRec(typeof(TChartAxisParentRecord));
			if (Result1.Count <= 0) return null;

			TChartAxis[] Result = new TChartAxis[Result1.Count];
			for (int i = 0; i < Result.Length; i++)
			{
				TChartAxisParentRecord AR = (TChartAxisParentRecord)Result1[i];

				TBaseAxis Categories = null;
				TValueAxis Values = null;
				TDataLabel ValueAxisCaption = null;
				TDataLabel CatAxisCaption = null;

				for (int z = 0; z < AR.Children.Count; z++)
				{
					TChartTextRecord TR = AR.Children[z] as TChartTextRecord;
					if (TR != null)
					{
						int SheetIndex = Workbook.InternalWorkbook.Sheets.IndexOf(CurrentSheet);
						if (SheetIndex >= 0)
						{
							TDataLabel AxisCaption = TR.GetDataLabel(Workbook, CellList, SheetIndex, true, true, FontScaling);
							if (AxisCaption.LinkedTo == TLinkOption.XAxisTitle) CatAxisCaption = AxisCaption;
							if (AxisCaption.LinkedTo == TLinkOption.YAxisTitle) ValueAxisCaption = AxisCaption;
						}
						continue;
					}

					TChartAxisRecord Axis = AR.Children[z] as TChartAxisRecord;
					if (Axis != null)
					{
						TFlxChartFont AxisFont = null;
						string AxisNumberFormat = null;
						TChartAxcExtRecord AxEx = null;
						TChartValueRangeRecord AxVa = null; 
						TAxisLineOptions AxisLineOptions = new TAxisLineOptions();
						TAxisTickOptions AxisTickOptions = new TAxisTickOptions(TTickType.Outside, TTickType.None, TAxisLabelPosition.NextToAxis, TBackgroundMode.Transparent, Colors.Black, 0);
						TAxisRangeOptions AxisRangeOptions = new TAxisRangeOptions(1, 1, false, false, false);
						int CatCrossValue = 1;

						for (int k = 0; k < Axis.Children.Count; k++)
						{
							TxChartBaseRecord Rec = Axis.Children[k] as TxChartBaseRecord;
							if (Rec == null) continue;

							switch ((xlr)Rec.Id)
							{
								case xlr.ChartFontx:
								{
									TChartFontXRecord FontX = (TChartFontXRecord)Rec;
									AxisFont = FontX.GetFont(FWorkbookGlobals, FontScaling);
									break;
								}

								case xlr.ChartIfmt:
								{
									TChartIFmtRecord IFmt = (TChartIFmtRecord)Rec;
									int FormatIndex = IFmt.FormatIndex;
									AxisNumberFormat = FWorkbookGlobals.Formats.Format(FormatIndex);
									break;
								}
									
								case xlr.ChartAxcext:
								{
									AxEx = (TChartAxcExtRecord)Rec;
									break;
								}

								case xlr.ChartValuerange:
								{
									AxVa = (TChartValueRangeRecord)Rec;
									break;
								}

								case xlr.ChartTick:
								{
									TChartTickRecord Tick = (TChartTickRecord)Rec;
									AxisTickOptions.BackgroundMode = (TBackgroundMode) Tick.BackgroundMode;
									long cl = Tick.LabelColor;
									AxisTickOptions.LabelColor = ColorUtil.FromArgb((int)(cl & 0xFF), (int)((cl & 0xFF00) >>8), (int)((cl & 0xFF0000) >>16));
									AxisTickOptions.LabelPosition = (TAxisLabelPosition) Tick.LabelPosition;
									AxisTickOptions.MajorTickType = (TTickType) Tick.MajorType;
									AxisTickOptions.MinorTickType = (TTickType) Tick.MinorType;
									AxisTickOptions.Rotation = Tick.Rotation;
									break;
								}

								case xlr.ChartCatserrange:
								{
									TChartCatSerRangeRecord CatSer = (TChartCatSerRangeRecord)Rec;
									CatCrossValue = CatSer.CatCross;
									AxisRangeOptions.LabelFrequency = CatSer.LabelFrequency;
									AxisRangeOptions.TickFrequency = CatSer.TickFrequency;
									AxisRangeOptions.ValueAxisBetweenCategories = (CatSer.Flags & 0x01) != 0;
									AxisRangeOptions.ValueAxisAtMaxCategory = (CatSer.Flags & 0x02) != 0;
									AxisRangeOptions.ReverseCategories = (CatSer.Flags & 0x04) != 0;
									break;
								}

								case xlr.ChartAxislineformat:
								{
									if ((k + 1) < Axis.Children.Count)
									{
										TChartLineFormatRecord Lf = Axis.Children[k+1] as TChartLineFormatRecord;
										if (Lf != null)
										{
											TChartAxisLineFormatRecord AxLineFmt = (TChartAxisLineFormatRecord)Rec;
											switch (AxLineFmt.AxisType)
											{
												case 0: 
													AxisLineOptions.MainAxis = Lf.GetLineFormat();
													AxisLineOptions.DoNotDrawLabelsIfNotDrawingAxis = (Lf.Flags & 0x04) == 0;
													break;
												case 1: AxisLineOptions.MajorGridLines = Lf.GetLineFormat();break;
												case 2: AxisLineOptions.MinorGridLines = Lf.GetLineFormat();break;
												case 3: AxisLineOptions.WallLines = Lf.GetLineFormat();break;
											}
										}
									}
									break;
								}
							}
						}

						switch ((TAxisType)Axis.AxisType)
						{
							case TAxisType.Category:
								if (AxEx != null)
								{
									Categories = new TCategoryAxis(AxEx.Min, AxEx.Max, AxEx.MajorValue, AxEx.MajorUnits, AxEx.MinorValue, AxEx.MinorUnits, AxEx.BaseUnits,
										CatCrossValue, (TCategoryAxisOptions) AxEx.AxisOptions, AxisFont, AxisNumberFormat, AxisLineOptions, AxisTickOptions, AxisRangeOptions, null);
								}
								else
								{
									if (AxVa != null)
									{
										Categories = new TValueAxis(AxVa.Min, AxVa.Max, AxVa.Major, AxVa.Minor, AxVa.CrossValue, (TValueAxisOptions) AxVa.AxisOptions,
											AxisFont, AxisNumberFormat, AxisLineOptions, AxisTickOptions, AxisRangeOptions, null);
									}
								}
								break;

							case TAxisType.Value:
								if (AxVa != null)
								{
									Values = new TValueAxis(AxVa.Min, AxVa.Max, AxVa.Major, AxVa.Minor, AxVa.CrossValue, (TValueAxisOptions) AxVa.AxisOptions,
										AxisFont, AxisNumberFormat, AxisLineOptions, AxisTickOptions, AxisRangeOptions, null);
								}
								break;
						}
					}

					//Only now we can load AxisCaption
					if (Categories != null) Categories.Caption = CatAxisCaption;
					if (Values != null) Values.Caption = ValueAxisCaption;
				}

				Result[i] = new TChartAxis(AR.Index, AR.Rect, Categories, Values);
			}

			return Result;
		}

		#endregion

		#region Chart Legend
		internal TChartLegend GetChartLegend(double FontScale)
		{
			List<TBaseRecord> Result1 = ChartCache.FindAllRec(typeof(TChartAxisParentRecord));
			if (Result1.Count <= 0) return null;

			for (int i = 0; i < Result1.Count; i++)
			{
				TChartAxisParentRecord AR = (TChartAxisParentRecord)Result1[i];

				List<TBaseRecord> Result2 = AR.FindAllRec(typeof(TChartChartFormatRecord));
				for (int k = 0; k < Result2.Count; k++)
				{
					TChartChartFormatRecord ChartFormat = (TChartChartFormatRecord) Result2[k];
					TChartLegendRecord Legend = (TChartLegendRecord) ChartFormat.FindRec<TChartLegendRecord>();
					if (Legend == null) continue;
					TChartLegend Result = new TChartLegend(Legend.xPos, Legend.yPos, Legend.xSize, Legend.ySize, (TChartLegendPos) Legend.LegendType, null, null);

					for (int z = 0; z < Legend.Children.Count; z++)
					{
						TxChartBaseRecord Rec = Legend.Children[z] as TxChartBaseRecord;
						if (Rec == null) continue;

						switch ((xlr)Rec.Id)
						{
							case xlr.ChartText:
							{
								TChartTextRecord Text = (TChartTextRecord)Rec;
								Result.TextOptions = Text.GetTextOptions(FWorkbookGlobals, FontScale);
								break;
							}

							case xlr.ChartFrame:
							{
								Result.Frame = ((TChartFrameRecord) Rec).GetFrameOptions();
								break;
							}

						}
					}

					return Result;
				}
			}
			return null;
		}

		#endregion

		#region Named ranges
		internal void UpdateDeletedRanges(TDeletedRanges DeletedRanges)
		{
			ChartRecords.UpdateDeletedRanges(DeletedRanges);
		}

    }
		#endregion

    internal class TSeriesData
    {
        List<TChartSIIndexRecord> SIIndex;
        internal TXlsCellRange Dimensions;

        internal TSeriesData(TXlsCellRange aDimensions)
        {
            SIIndex = new List<TChartSIIndexRecord>();
            Dimensions = aDimensions;
        }

        internal void Add(TChartSIIndexRecord SIIndexRecord)
        {
            SIIndex.Add(SIIndexRecord);
        }

        internal TSeriesData Clone(TSheetInfo SheetInfo)
        {
            TXlsCellRange NewDims = Dimensions == null ? new TXlsCellRange() : (TXlsCellRange)Dimensions.Clone();
            TSeriesData Result = new TSeriesData(NewDims);
            foreach (TChartSIIndexRecord r in SIIndex)
            {
                Result.SIIndex.Add((TChartSIIndexRecord)TChartSIIndexRecord.Clone(r, SheetInfo));
            }

            return Result;
        }

        internal void SaveToStream(IDataStream DataStream, TSaveData SaveData)
        {
            TCells.WriteDimensions(DataStream, SaveData, Dimensions);
            for (int i = 0; i < SIIndex.Count; i++)
            {
                SIIndex[i].SaveToStream(DataStream, SaveData, 0);
            }
        }

        internal long TotalSize()
        {
            long Result = TCells.DimensionsSize;
            for (int i = 0; i < SIIndex.Count; i++)
            {
                Result += SIIndex[i].TotalSize();
            }

            return Result;
        }


        internal void EnsureRequiredRecords()
        {
            if (Dimensions == null) Dimensions = new TXlsCellRange();
            if (SIIndex.Count >= 3) return;

            bool[] Used = new bool[3];
            for (int i = 0; i < SIIndex.Count; i++)
			{
                int num = (int)SIIndex[i].NumIndex - 1;
                if (num < 0 || num >= Used.Length) continue;
                Used[num] = true;
			}

            for (int i = 0; i < Used.Length; i++)
            {
                if (!Used[i]) SIIndex.Add(new TChartSIIndexRecord((TChartSIIndexType)(i + 1)));
            }
        }


        internal void DeleteSeries(int index)
        {
            for (int i = 0; i < SIIndex.Count; i++)
            {
                SIIndex[i].DeleteSeries(index - 1);
            }
        }
    }
    #endregion

    internal class RowColSize : IRowColSize
    {
        #region Variables
        private TSheet Sheet;
        real FHeightCorrection;
        real FWidthCorrection;
        #endregion

        public RowColSize(float aHeightCorrection, float aWidthCorrection, TSheet aSheet)
        {
            FHeightCorrection = aHeightCorrection;
            FWidthCorrection = aWidthCorrection;
            Sheet = aSheet;

        }

        #region IRowColSize Members

        public int DefaultColWidth
        {
            get
            {
                return Sheet.Columns.DefColWidth;
            }
            set
            {
                Sheet.Columns.DefColWidth = value;
            }
        }

        public int DefaultRowHeight
        {
            get
            {
                return Sheet.DefRowHeight;
            }
            set
            {
                Sheet.DefRowHeight = value;
            }
        }

        public bool IsEmptyRow(int row)
        {
            return !(Sheet.Cells.CellList.HasRow(row - 1));
        }

        public int GetRowHeight(int row, bool HiddenIsZero)
        {
            return Sheet.GetRowHeight(row - 1, HiddenIsZero);
        }

        public int GetColWidth(int col, bool HiddenIsZero)
        {
            return Sheet.GetColWidth(col - 1, HiddenIsZero);
        }

        public bool ShowFormulaText
        {
            get
            {
                return Sheet.Window.Window2.ShowFormulaText;
            }
            set
            {
                Sheet.Window.Window2.ShowFormulaText = value;
            }
        }

        public TFlxFont GetDefaultFont
        {
            get
            {
                int FontIndex = Sheet.FWorkbookGlobals.CellXF[0].FontIndex;
                return Sheet.FWorkbookGlobals.Fonts.GetFont(FontIndex);
            }
        }

        public real WidthCorrection
        {
            get
            {
                return FWidthCorrection;
            }
            set
            {
                FWidthCorrection = value;
            }
        }

        public real HeightCorrection
        {
            get
            {
                return FHeightCorrection;
            }
            set
            {
                FHeightCorrection = value;
            }
        }

        #endregion
    }
}
