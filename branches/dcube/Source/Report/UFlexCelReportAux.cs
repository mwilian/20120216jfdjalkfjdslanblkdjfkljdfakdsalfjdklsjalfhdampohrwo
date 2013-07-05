using System;
using System.Data;
using System.IO;
using FlexCel.Core;
using System.Collections.Generic;
using System.Collections;

namespace FlexCel.Report
{
    #region Progress indicator
    /// <summary>
    /// Phase of the report we are in.
    /// </summary>
    public enum FlexCelReportProgressPhase
    {
        /// <summary>
        /// Report is inactive.
        /// </summary>
        NotRunning,

		/// <summary>
        /// Reading and parsing the template.
        /// </summary>
        ReadTemplate,
        
		/// <summary>
        /// Organizing the data on the template.
        /// </summary>
        OrganizeData,
        
		/// <summary>
        /// Inserting the needed rows/ranges for the report.
        /// </summary>
        CopyStructure,
        
		/// <summary>
        /// Replacing the tags with the new values.
        /// </summary>
        FillData,
        
		/// <summary>
        /// Fixing pagebreak or delete rows tags.
        /// </summary>
        FinalCleanup,
        
		/// <summary>
        /// Report has finished.
        /// </summary>
        Done
    }

    /// <summary>
    /// Indicates how much of the report has been generated.
    /// </summary>
    public class FlexCelReportProgress
    {
        private volatile FlexCelReportProgressPhase FPhase;
        private volatile int FCounter;
        private volatile int FSheet;

        internal FlexCelReportProgress()
        {
            Clear();
        }

        internal void Clear()
        {
            FPhase = FlexCelReportProgressPhase.NotRunning;
            FCounter = 0;
            FSheet = 0;
        }

        internal void SetPhase(FlexCelReportProgressPhase aPhase)
        {
            FPhase = aPhase;
        }

        internal void SetCounter(int aCounter)
        {
            FCounter = aCounter;
        }

        internal void SetSheet(int aSheet)
        {
            FSheet = aSheet;
        }

        /// <summary>
        /// Phase of the report.
        /// </summary>
        public FlexCelReportProgressPhase Phase { get { return FPhase; } }

        /// <summary>
        /// A meaningless counter that is increased from time to time. It is not possible to know what is the final count.
        /// </summary>
        public int Counter { get { return FCounter; } }

        /// <summary>
        /// The Sheet we are working on.
        /// </summary>
        public int Sheet { get { return FSheet; } }
    }
    #endregion

    #region SQL Parameters
    /// <summary>
    /// How the parameters for Direct SQL queries are.
    /// Change it only if your database uses positional parameters and it is
    /// not ODBC or OLEDB.
    /// </summary>
    public enum TSqlParametersType
    {
        /// <summary>
        /// FlexCel will try to guess the correct type.
        /// Currently it will return Positional for ODBC and OLEDB Parameters
        /// and Named for everything else.
        /// </summary>
        Automatic,

        /// <summary>
        /// Parameter is named. For example, "@param1" or ":Param2"
        /// </summary>
        Named,

        /// <summary>
        /// Name of parameter is not important, and we only care about position.
        /// For example "?"
        /// </summary>
        Positional
    }
    #endregion

    #region Event Handlers

    #region Generate
    /// <summary>
    /// Arguments passed on <see cref="FlexCel.Report.FlexCelReport.BeforeGenerateWorkbook"/>, <see cref="FlexCel.Report.FlexCelReport.AfterGenerateWorkbook"/>, <see cref="FlexCel.Report.FlexCelReport.BeforeGenerateSheet"/> and <see cref="FlexCel.Report.FlexCelReport.AfterGenerateSheet"/>
    /// </summary>
    public class GenerateEventArgs : EventArgs
    {
        private readonly ExcelFile FExcelFile;

        /// <summary>
        /// Creates a new Argument.
        /// </summary>
        /// <param name="aExcelFile">The file we are processing.</param>
        public GenerateEventArgs(ExcelFile aExcelFile)
        {
            FExcelFile = aExcelFile;
        }

        /// <summary>
        /// The file with the report data.
        /// </summary>
        public ExcelFile File
        {
            get { return FExcelFile; }
        }

    }

    /// <summary>
    /// Generic delegate for After/Before generate events.
    /// </summary>
    public delegate void GenerateEventHandler(object sender, GenerateEventArgs e);
    #endregion

    #region ImageData
    /// <summary>
    /// Arguments passed on <see cref="FlexCel.Report.FlexCelReport.GetImageData"/>
    /// </summary>
    public class GetImageDataEventArgs : EventArgs
    {
        private readonly ExcelFile FExcelFile;
        private string FImageName;
        private byte[] FImageData;
        private double FHeight;
        private double FWidth;

        /// <summary>
        /// Creates a new Argument.
        /// </summary>
        /// <param name="aExcelFile">The file we are processing.</param>
        /// <param name="aImageName">The name of the image on the Excel sheet. Use it to identify it.</param>
        /// <param name="aImageData">The image data.</param>
        /// <param name="aHeight">The height of the image in pixels. Change it to resize the image.</param>
        /// <param name="aWidth">The width of the image in pixels. Change it to resize the image.</param>
        public GetImageDataEventArgs(ExcelFile aExcelFile, string aImageName, byte[] aImageData, double aHeight, double aWidth)
        {
            FExcelFile = aExcelFile;
            FImageName = aImageName;
            FImageData = aImageData;
            FHeight = aHeight;
            FWidth = aWidth;
        }

        /// <summary>
        /// The file with the report data.
        /// </summary>
        public ExcelFile File
        {
            get { return FExcelFile; }
        }


        /// <summary>
        /// The name of the image on the Excel sheet. Use it to identify it.
        /// </summary>
        public string ImageName { get { return FImageName; } set { FImageName = value; } }

        /// <summary>
        /// The data of the image. You can modify it to return another image format.
        /// </summary>
        public byte[] ImageData { get { return FImageData; } set { FImageData = value; } }

        /// <summary>
        /// The height of the image in pixels. Change it to resize the image.
        /// </summary>
        public double Height { get { return FHeight; } set { FHeight = value; } }

        /// <summary>
        /// The width of the image in pixels. Change it to resize the image.
        /// </summary>
        public double Width { get { return FWidth; } set { FWidth = value; } }
    }

    /// <summary>
    /// Delegate for GetImageData event.
    /// </summary>
    public delegate void GetImageDataEventHandler(object sender, GetImageDataEventArgs e);
    #endregion

    #region GetInclude
    /// <summary>
    /// Arguments passed on <see cref="FlexCel.Report.FlexCelReport.GetInclude"/>
    /// </summary>
    public class GetIncludeEventArgs : EventArgs
    {
        private readonly ExcelFile FExcelFile;
        private byte[] FIncludeData;
        private string FFileName;

        /// <summary>
        /// Creates a new Argument.
        /// </summary>
        /// <param name="aExcelFile">The ExcelFile that has the report doing the include.</param>
        /// <param name="aFileName">File that we are trying to include.</param>
        /// <param name="aIncludeData">The included file as an array of bytes. If you return null, the file will be searched on disk.</param>
        public GetIncludeEventArgs(ExcelFile aExcelFile, string aFileName, byte[] aIncludeData)
        {
            FExcelFile = aExcelFile;
            FIncludeData = aIncludeData;
            FFileName = aFileName;
        }

        /// <summary>
        /// The file with the report.
        /// </summary>
        public ExcelFile File
        {
            get { return FExcelFile; }
        }

        /// <summary>
        /// File we are trying to include. you can modify it to point to other place. 
        /// If the including file is a real file (not an stream) and FileName is relative, it will be relative to the
        /// including file path.
        /// </summary>
        public string FileName { get { return FFileName; } set { FFileName = value; } }

        /// <summary>
        /// Here you can return the included file as an array of bytes.
        /// If you return null, the filename will be used to search for a file on the disk. 
        /// If the including file is a real file (not an stream) and FileName is relative, it will be relative to the
        /// including file path.
        /// </summary>
        public byte[] IncludeData { get { return FIncludeData; } set { FIncludeData = value; } }

    }

    /// <summary>
    /// Delegate for GetInclude event.
    /// </summary>
    public delegate void GetIncludeEventHandler(object sender, GetIncludeEventArgs e);
    #endregion

    #region UserTable
    /// <summary>
    /// Arguments passed on <see cref="FlexCel.Report.FlexCelReport.UserTable"/>
    /// </summary>
    public class UserTableEventArgs : EventArgs
    {
        private readonly string FTableName;
        private readonly string FParameters;

        /// <summary>
        /// Creates a new Argument.
        /// </summary>
        /// <param name="aTableName">The value written on the cell &quot;Table name&quot; on the config sheet. You can use it as an extra parameter.</param>
        /// <param name="aParameters">The parameters passed on the &lt;#User Table(parameters)&gt; tag.</param>
        public UserTableEventArgs(string aTableName, string aParameters)
        {
            FTableName = aTableName;
            FParameters = aParameters;
        }

        /// <summary>
        /// The value written on the cell &quot;Table name&quot; on the config sheet. You can use it as an extra parameter.
        /// </summary>
        public string TableName
        {
            get { return FTableName; }
        }

        /// <summary>
        /// The parameters on the &lt;#User Table(parameters)&gt; tag.
        /// </summary>
        public string Parameters
        {
            get { return FParameters; }
        }
    }

    /// <summary>
    /// Delegate for UserTable event.
    /// </summary>
    public delegate void UserTableEventHandler(object sender, UserTableEventArgs e);
    #endregion

    #region Table
    /// <summary>
    /// Arguments passed on <see cref="FlexCel.Report.FlexCelReport.LoadTable"/>
    /// </summary>
    public class LoadTableEventArgs : EventArgs
    {
        private readonly string FTableName;

        /// <summary>
        /// Creates a new Argument.
        /// </summary>
        /// <param name="aTableName">The table that needs to be loaded on demand.</param>
        public LoadTableEventArgs(string aTableName)
        {
            FTableName = aTableName;
        }

        /// <summary>
        /// The table that needs to be loaded on demand.
        /// </summary>
        public string TableName
        {
            get { return FTableName; }
        }
    }

    /// <summary>
    /// Delegate for LoadTable event.
    /// </summary>
    public delegate void LoadTableEventHandler(object sender, LoadTableEventArgs e);
    #endregion

    #endregion

    #region Rectangles

    internal enum TRectPos
    {
        Inside,
        Outside,
        Intersect,
        Separated,
        Equal
    }

    internal sealed class RectUtils
    {
        private RectUtils() { }

        internal static TRectPos TestInRect(TXlsCellRange R1, TXlsCellRange R2)
        {
            if (R1.Left == R2.Left && R1.Top == R2.Top && R1.Right == R2.Right && R1.Bottom == R2.Bottom) return TRectPos.Equal;
            if (R1.Right < R2.Left || R1.Left > R2.Right || R1.Bottom < R2.Top || R1.Top > R2.Bottom) return TRectPos.Separated;
            if (R1.Left <= R2.Left && R1.Top <= R2.Top && R1.Right >= R2.Right && R1.Bottom >= R2.Bottom) return TRectPos.Outside;
            if (R2.Left <= R1.Left && R2.Top <= R1.Top && R2.Right >= R1.Right && R2.Bottom >= R1.Bottom) return TRectPos.Inside;
            return TRectPos.Intersect;
        }
    }

    #endregion

    #region TInclude

    internal class TInclude : IDisposable
    {
        #region privates
        private byte[] FData;
        private string FTagText;
        private TBandType FBandType;
        private TBand MainBand;
        private string FRangeName;
        private bool StaticInclude;

        private List<TKeepTogether> KeepRows;
        private List<TKeepTogether> KeepCols;

        private FlexCelReport Report;
        #endregion

        internal TInclude(byte[] aData, string aRange, TBandType aBandType, TBand aParentBand,
            string aTagText, int aNestedLevel, TDataSourceInfoList aDsInfoList, string aFileName, bool aStaticInclude, FlexCelReport aParentReport)
        {
            try
            {
                FData = new byte[aData.Length];
                Array.Copy(aData, 0, FData, 0, FData.Length);
                FBandType = aBandType;
                FTagText = aTagText;
                StaticInclude = aStaticInclude;

                //Preloading at read time has the advantage of fully checking the template on load, 
                //so errors will be detected sooner. (If an include is conditional, might not be detected until much later)
                //For this same reason it is a little slower than preloading on demand, but it is worth.
                Preload(aParentBand, aRange, aFileName, aDsInfoList, aNestedLevel, aParentReport);
            }
            catch
            {
                Dispose();
                throw;
            }
        }

        private static TBand CreateStartingBand(TXlsCellRange XlsRange, TBand aParentBand, string aRange)
        {
            return new TBand(null, aParentBand, XlsRange, aRange, TBandType.Static, false, String.Empty);
        }

        private void DoPreload(TBand aParentBand, string aRange, string aFileName, TDataSourceInfoList aDsInfoList, int aNestedLevel, ExcelFile Result, MemoryStream MStream, FlexCelReport aParentReport)
        {
            Result.Open(MStream);
            Result.ActiveFileName = aFileName;

            TXlsNamedRange XlsRange = Result.GetNamedRange(aRange, -1);
            if (XlsRange == null) FlxMessages.ThrowException(FlxErr.ErrCantFindNamedRange, aRange);
            FRangeName = aRange;
            MainBand = CreateStartingBand(XlsRange, aParentBand, aRange);
            Report = new FlexCelReport(aNestedLevel, FTagText, aDsInfoList, aParentReport);
            if (!StaticInclude)
                Report.PreLoad(Result, ref MainBand, XlsRange.SheetIndex, ref FData, out KeepRows, out KeepCols);
            Result.ActiveSheet = XlsRange.SheetIndex;
        }

        private void Preload(TBand aParentBand, string aRange, string aFileName, TDataSourceInfoList aDsInfoList, int aNestedLevel, FlexCelReport aParentReport)
        {
            ExcelFile Result = new XlsAdapter.XlsFile();
            using (MemoryStream MStream = new MemoryStream(FData))
            {
                MStream.Position = 0;
                if (aNestedLevel > 1)
                    DoPreload(aParentBand, aRange, aFileName, aDsInfoList, aNestedLevel, Result, MStream, aParentReport);
                else
                    try //We only catch a level 1 include. if not, we would end up with a nested message.
                    {
                        DoPreload(aParentBand, aRange, aFileName, aDsInfoList, aNestedLevel, Result, MStream, aParentReport);
                    }
                    catch (Exception e)
                    {
                        FlxMessages.ThrowException(e, FlxErr.ErrOnIncludeReport, FTagText, e.Message);
                    }
            }
        }

        internal ExcelFile Run()
        {
            ExcelFile Result = new XlsAdapter.XlsFile();
            using (MemoryStream MStream = new MemoryStream(FData))
            {
                MStream.Position = 0;
                try
                {
                    Result.Open(MStream);
                    Result.ActiveSheet = Result.ActiveSheet; //Just in case...
					if (!StaticInclude) 
					{
						Report.RunPreloaded(Result, MainBand, KeepRows, KeepCols);
					}
                }
                catch (Exception e)
                {
                    FlxMessages.ThrowException(e, FlxErr.ErrOnIncludeReport, FTagText, e.Message);
                }
            }

            return Result;
        }

        public string RangeName { get { return FRangeName; } set { FRangeName = value; } }
        public TBandType BandType { get { return FBandType; } set { FBandType = value; } }

        public override string ToString()
        {
            return String.Empty;
        }

        internal static byte[] OpenInclude(GetIncludeEventArgs ea, ExcelFile Workbook)
        {
            string iPath = Workbook.ActiveFileName;
            try
            {
                if (iPath.Length > 0) iPath = Path.GetDirectoryName(iPath) + Path.DirectorySeparatorChar;
            }
            catch (ArgumentException)
            {
                iPath = String.Empty;
            }
            if (Path.IsPathRooted(ea.FileName)) iPath = String.Empty;

            ea.FileName = iPath + ea.FileName;
            using (FileStream f = new FileStream(ea.FileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) //FileShare.ReadWrite is the way we have to open a file even if it is being used by excel.
            {
                byte[] Result = new byte[f.Length];
                Sh.Read(f, Result, 0, Result.Length);
                return Result;
            }
        }
        #region IDisposable Members

        public void Dispose()
        {
            if (MainBand != null) MainBand.Dispose();
            if (Report != null)
            {
                Report.Unload();
#if (!COMPACTFRAMEWORK)
                ((IDisposable)Report).Dispose();
#endif
            }
            Report = null;
            MainBand = null;

            GC.SuppressFinalize(this);

        }

        #endregion
    }

    #endregion

    #region TWaitingRange
    internal class TWaitingCoords
    {
        internal int RowOfs;
        internal int ColOfs;
        internal int InsRowOfs;
        internal int InsColOfs;
        internal int LastRangeRow;
        internal int LastRangeCol;

        internal TWaitingCoords(int aRowOfs, int aColOfs, int aInsRowOfs, int aInsColOfs, int aLastRangeRow, int aLastRangeCol)
        {
            RowOfs = aRowOfs;
            ColOfs = aColOfs;
            InsRowOfs = aInsRowOfs;
            InsColOfs = aInsColOfs;
            LastRangeRow = aLastRangeRow;
            LastRangeCol = aLastRangeCol;
        }

    }

    /// <summary>
    /// Holds a generic range that will be inserted/deleted after a band has been replaced.
    /// </summary>
    internal abstract class TWaitingRange : ITopLeft
    {
        internal TBandType BandType;
        protected int FTop;
        protected int FLeft;

        internal TWaitingRange(TBandType aBandType, int aTop, int aLeft)
        {
            BandType = aBandType;
            FTop = aTop;
            FLeft = aLeft;
        }

        internal TFlxInsertMode InsertMode
        {
            get
            {
                switch (BandType)
                {
                    case TBandType.ColFull: return TFlxInsertMode.ShiftColRight;
                    case TBandType.RowRange: return TFlxInsertMode.ShiftRangeDown;
                    case TBandType.ColRange: return TFlxInsertMode.ShiftRangeRight;
                    case TBandType.FixedCol: return TFlxInsertMode.NoneRight;
                    case TBandType.FixedRow: return TFlxInsertMode.NoneDown;
                }
                return TFlxInsertMode.ShiftRowDown;

            }
        }

        internal abstract void Execute(ExcelFile Workbook, TWaitingCoords Coords, TBand Band);
        #region ITopLeft Members

        public int Top
        {
            get
            {
                return FTop;
            }
        }

        public int Left
        {
            get
            {
                return FLeft;
            }
        }

        #endregion
    }

    internal class TIncludeWaitingRange : TWaitingRange
    {
        private TInclude FInclude;
        private bool CopyColFormats;
        private bool CopyRowFormats;

        internal TIncludeWaitingRange(TInclude aInclude, int aRow, int aCol, bool aCopyRowFormats, bool aCopyColFormats)
            : base(aInclude.BandType, aRow, aCol)
        {
            FInclude = aInclude;
            CopyRowFormats = aCopyRowFormats;
            CopyColFormats = aCopyColFormats;
        }

        internal override void Execute(ExcelFile Workbook, TWaitingCoords Coords, TBand Band)
        {
            ExcelFile IncludedReport = FInclude.Run();
            //Workbook.InsertAndCopyRange(IncludedReport.GetNamedRange(FInclude.RangeName, IncludedReport.ActiveSheet), 
            //    Top+RowOfs, Left+ColOfs, 1, InsertMode , TRangeCopyMode.All,  IncludedReport, IncludedReport.ActiveSheet);

            //This is to avoid inserting one row more on the include.
            TXlsNamedRange range = IncludedReport.GetNamedRange(FInclude.RangeName, IncludedReport.ActiveSheet);

			//We don't want to copy he full range even if using "__". Just *insert* the full range
			//if (InsertMode == TFlxInsertMode.ShiftRowDown) {range.Left = 1; range.Right = FlxConsts.Max_Columns + 1;}
			//if (InsertMode == TFlxInsertMode.ShiftColRight) {range.Top = 1; range.Bottom = FlxConsts.Max_Rows + 1;}
            
            if (InsertMode == TFlxInsertMode.ShiftColRight || InsertMode == TFlxInsertMode.ShiftRangeRight)
            {
                TXlsCellRange rangeToInsert= new TXlsCellRange(range.Top, range.Left, range.Bottom, range.Right - 1);
                if (InsertMode == TFlxInsertMode.ShiftColRight) {rangeToInsert.Top = 1; rangeToInsert.Bottom = FlxConsts.Max_Rows + 1;}

                if (range.Left > range.Right)
                { }// not possible. Workbook.DeleteRange();
                else if (range.Left < range.Right)
                {
                    int CopyTop = Top + Coords.RowOfs;
                    if (InsertMode == TFlxInsertMode.ShiftColRight) CopyTop = 1;

                    Workbook.InsertAndCopyRange(rangeToInsert, CopyTop, Left + Coords.ColOfs, 1, InsertMode, TRangeCopyMode.None);
                    if (Band != null)
                    {
                        Band.AddTmpExpandedCols(rangeToInsert.ColCount, Top + Coords.RowOfs, Top + Coords.RowOfs + rangeToInsert.RowCount);
                        Band.TmpPartialCols += rangeToInsert.ColCount;
                    }
                }

				CopyRowAndColFormat(Workbook, Coords, IncludedReport, range);

                if (range.Left <= range.Right)
                    Workbook.InsertAndCopyRange(range, Top + Coords.RowOfs, Left + Coords.ColOfs, 1, TFlxInsertMode.NoneRight, TRangeCopyMode.AllIncludingDontMoveAndSizeObjects, IncludedReport, IncludedReport.ActiveSheet);

            }
            else
            {
                TXlsCellRange rangeToInsert = new TXlsCellRange(range.Top, range.Left, range.Bottom - 1, range.Right);
                if (InsertMode == TFlxInsertMode.ShiftRowDown) {rangeToInsert.Left = 1; rangeToInsert.Right = FlxConsts.Max_Columns + 1;}

                if (range.Top > range.Bottom)
                { }// not possible. Workbook.DeleteRange();
                else if (range.Top < range.Bottom)
                {
                    int CopyLeft = Left + Coords.ColOfs;
                    if (InsertMode == TFlxInsertMode.ShiftRowDown) CopyLeft = 1;

                    Workbook.InsertAndCopyRange(rangeToInsert, Top + Coords.RowOfs, CopyLeft, 1, InsertMode, TRangeCopyMode.None);
                    if (Band != null)
                    {
                        Band.AddTmpExpandedRows(rangeToInsert.RowCount, Left + Coords.ColOfs, Left + Coords.ColOfs + rangeToInsert.ColCount);
                        Band.TmpPartialRows += rangeToInsert.RowCount;
                    }
                }

                CopyRowAndColFormat(Workbook, Coords, IncludedReport, range);

                if (range.Top <= range.Bottom)
                    Workbook.InsertAndCopyRange(range, Top + Coords.RowOfs, Left + Coords.ColOfs, 1, TFlxInsertMode.NoneDown, TRangeCopyMode.AllIncludingDontMoveAndSizeObjects, IncludedReport, IncludedReport.ActiveSheet);
            }

        }

        private void CopyRowAndColFormat(ExcelFile Workbook, TWaitingCoords Coords, ExcelFile IncludedReport, TXlsCellRange range)
        {
			//Columns go before rows.
			if (CopyColFormats || (InsertMode == TFlxInsertMode.ShiftColRight && (range.Top >1 || range.Bottom < FlxConsts.Max_Rows + 1)))
			{
				for (int c = range.Left; c < range.Right; c++)
				{
					if (!IncludedReport.IsNotFormattedCol(c))
					{
						int c1 = c - range.Left + Left + Coords.ColOfs;
                        int cw = IncludedReport.GetColWidth(c);
						Workbook.SetColWidth(c1, cw);

                        int co = IncludedReport.GetColOptions(c);
                        if (cw != IncludedReport.DefaultColWidth || cw != Workbook.DefaultColWidth) co |= 0x02; //the column has no standard width. 

						Workbook.SetColOptions(c1, co);
						TFlxFormat fmt = IncludedReport.GetFormat(IncludedReport.GetColFormat(c));
						fmt.LinkedStyle.AutomaticChoose = false;
						Workbook.SetColFormat(c1, Workbook.AddFormat(fmt));
						if (c1 + 1 <= FlxConsts.Max_Columns) Workbook.KeepColsTogether(c1, c1 + 1, IncludedReport.GetKeepColsTogether(c), true);
					}
				}
			}

			if (CopyRowFormats || (InsertMode == TFlxInsertMode.ShiftRowDown && (range.Left >1 || range.Right < FlxConsts.Max_Columns + 1)))
            {
                for (int r = range.Top; r < range.Bottom; r++)
                {
					if (!IncludedReport.IsEmptyRow(r))
					{
						int r1 = r - range.Top + Top + Coords.RowOfs;
						Workbook.SetRowHeight(r1, IncludedReport.GetRowHeight(r));
						Workbook.SetRowOptions(r1, IncludedReport.GetRowOptions(r));
						TFlxFormat fmt = IncludedReport.GetFormat(IncludedReport.GetRowFormat(r));
						fmt.LinkedStyle.AutomaticChoose = false;
						Workbook.SetRowFormat(r1, Workbook.AddFormat(fmt));
						if (r1 + 1 <= FlxConsts.Max_Rows) Workbook.KeepRowsTogether(r1, r1 + 1, IncludedReport.GetKeepRowsTogether(r), true);
					}
                }
            }
        }
    }

    internal class TDeleteRowWaitingRange : TWaitingRange
    {
        internal int LastCol;

        internal TDeleteRowWaitingRange(int aRow, int aFirstCol, int aLastCol)
            : base(TBandType.RowRange, aRow, aFirstCol)
        {
            LastCol = aLastCol;
        }

        internal override void Execute(ExcelFile Workbook, TWaitingCoords Coords, TBand Band)
        {
            Workbook.DeleteRange(new TXlsCellRange(Top + Coords.RowOfs, Left + Coords.ColOfs, Top + Coords.RowOfs, LastCol + Coords.ColOfs), TFlxInsertMode.ShiftRangeDown);
            Band.AddTmpExpandedRows(-1, Left + Coords.ColOfs, LastCol + Coords.ColOfs);
            Band.TmpPartialRows -= 1;
        }
    }

    internal class TDeleteColWaitingRange : TWaitingRange
    {
        internal int LastRow;

        internal TDeleteColWaitingRange(int aCol, int aFirstRow, int aLastRow)
            : base(TBandType.ColRange, aFirstRow, aCol)
        {
            LastRow = aLastRow;
        }

        internal override void Execute(ExcelFile Workbook, TWaitingCoords Coords, TBand Band)
        {
            Workbook.DeleteRange(new TXlsCellRange(Top + Coords.RowOfs, Left + Coords.ColOfs, LastRow + Coords.RowOfs, Left + Coords.ColOfs), TFlxInsertMode.ShiftRangeRight);
            Band.AddTmpExpandedCols(-1, Top + Coords.RowOfs, LastRow + Coords.RowOfs);
            Band.TmpPartialCols -= 1;
        }
    }

	internal abstract class TRangeWaitingRange: TWaitingRange, IDisposable
	{
		protected TOneCellValue Range;
		protected int Bottom;
		protected int Right;
		protected int Sheet1;
		protected int Sheet2;

		protected bool RowAbs1;
		protected bool ColAbs1;
		protected bool RowAbs2;
		protected bool ColAbs2;

		internal TRangeWaitingRange(TOneCellValue aRange, TBandType aBandType)
			: base(aBandType, 0, 0)
		{
			if (aRange != null)
			{
				if (aRange.Count == 1 && aRange[0].ValueType == TValueType.Const) //optimize most common case
				{
					CalcBounds(aRange);
				}
				else
				{
					Range = aRange;
				}
			}
		}


		protected void CalcBounds(TOneCellValue aRange)
		{
			TValueAndXF val = new TValueAndXF();
			val.Workbook = aRange.Workbook;
			aRange.Evaluate(0, 0, 0, 0, val);

			TXlsNamedRange XlsRange = val.Workbook.GetNamedRange(FlxConvert.ToString(val.Value), -1, val.Workbook.ActiveSheet);
			if (XlsRange == null)
				XlsRange = val.Workbook.GetNamedRange(FlxConvert.ToString(val.Value), -1, 0);

			if (XlsRange != null)
			{
				FTop = XlsRange.Top;
				FLeft = XlsRange.Left;
				Bottom = XlsRange.Bottom;
				Right = XlsRange.Right;
				Sheet1 = XlsRange.SheetIndex;
				Sheet2 = XlsRange.SheetIndex;
				RowAbs1 = false; ColAbs1 = false; RowAbs2 = false; ColAbs2 = false;
				return;
			}


			string[] Addresses = FlxConvert.ToString(val.Value).Split(TFormulaMessages.TokenChar(TFormulaToken.fmRangeSep));
			if (Addresses == null || (Addresses.Length != 2 && Addresses.Length != 1))
				FlxMessages.ThrowException(FlxErr.ErrInvalidRef, FlxConvert.ToString(val.Value));
			TCellAddress FirstCell = new TCellAddress(Addresses[0]);
			if (FirstCell.Sheet == null || FirstCell.Sheet.Length == 0) Sheet1 = -1; else Sheet1 = aRange.Workbook.GetSheetIndex(FirstCell.Sheet);
			FTop = FirstCell.Row;
			FLeft = FirstCell.Col;

			RowAbs1 = FirstCell.RowAbsolute;
			ColAbs1 =FirstCell.ColAbsolute;

			if (Addresses.Length > 1) FirstCell = new TCellAddress(Addresses[1]);
			if (FirstCell.Sheet == null || FirstCell.Sheet.Length == 0) Sheet2 = Sheet1; else Sheet2 = aRange.Workbook.GetSheetIndex(FirstCell.Sheet);
			Bottom = FirstCell.Row;
			Right = FirstCell.Col;
			RowAbs2 = FirstCell.RowAbsolute;
			ColAbs2 =FirstCell.ColAbsolute;
		}

		protected void GetBounds(TWaitingCoords Coords, out int t, out int l, out int b, out int r)
		{
			if (Range != null)
			{
				CalcBounds(Range);
			}

			t = Top; if (!RowAbs1) {t += Coords.RowOfs; if (t > Coords.LastRangeRow) t += Coords.InsRowOfs;}
			b = Bottom; if (!RowAbs2) {b += Coords.RowOfs; if (b > Coords.LastRangeRow) b += Coords.InsRowOfs;}
			l = Left; if (!ColAbs1) {l += Coords.ColOfs; if (l > Coords.LastRangeCol) l += Coords.InsColOfs;}
			r = Right; if (!ColAbs2) {r += Coords.ColOfs; if (r > Coords.LastRangeCol) r += Coords.InsColOfs;}
		}

		#region IDisposable Members

		protected virtual void Dispose(bool Disposing)
		{
			if (Disposing)
			{
				if (Range != null) Range.Dispose();
                Range = null;
			}
		}

		public void Dispose()
		{
			Dispose(true);
			GC.SuppressFinalize(this);
		}

		#endregion


	}

    internal class TDeleteRangeWaitingRange : TRangeWaitingRange
    {
        internal TDeleteRangeWaitingRange(TOneCellValue aRange, TBandType aBandType)
            : base(aRange, aBandType)
        {
        }

        internal override void Execute(ExcelFile Workbook, TWaitingCoords Coords, TBand Band)
        {
            int t, l, b, r;
            GetBounds(Coords, out t, out l, out b, out r);
            TXlsCellRange rangeToDelete = new TXlsCellRange(t, l, b, r);

            if (Sheet1 <= 0 || Sheet2 <= 0)
            {
                Workbook.DeleteRange(rangeToDelete, InsertMode);
            }
            else
            {
                Workbook.DeleteRange(Sheet1, Sheet2, rangeToDelete, InsertMode);
            }

            if (Sheet1 <= 0 || Sheet2 <= 0 || (Sheet1 <= Workbook.ActiveSheet && Sheet2 >= Workbook.ActiveSheet))
            {
                if (Band != null)
                {
                    if (InsertMode == TFlxInsertMode.ShiftRangeDown || InsertMode == TFlxInsertMode.ShiftRowDown)
                    {
                        Band.AddTmpExpandedRows(-rangeToDelete.RowCount, Left + Coords.ColOfs, Left + Coords.ColOfs + rangeToDelete.ColCount);
                        Band.TmpPartialRows += -rangeToDelete.RowCount;
                    }
                    else if (InsertMode == TFlxInsertMode.ShiftColRight || InsertMode == TFlxInsertMode.ShiftRangeRight)
                    {
                        Band.AddTmpExpandedCols(-rangeToDelete.ColCount, Top + Coords.RowOfs, Top + Coords.RowOfs + rangeToDelete.RowCount);
                        Band.TmpPartialCols += -rangeToDelete.ColCount;
                    }
                }

            }
        }

    }

    /// <summary>
    /// Holds all includes/delete row/columns that will be inserted.
    /// </summary>
    internal class TWaitingRangeList
    {
        private List<TWaitingRange> FList;

        internal TWaitingRangeList()
        {
            FList = new List<TWaitingRange>();
        }

        internal void Add(TWaitingRange wr)
        {
            FList.Add(wr);
        }

        internal int Count
        {
            get
            {
                return FList.Count;
            }
        }

        internal TWaitingRange this[int index]
        {
            get
            {
                return FList[index];
            }
        }
    }
    #endregion

    #region TFormatRange
    /// <summary>
    /// Holds an XF format for a range of cells.
    /// </summary>
    internal class TFormatRange : TRangeWaitingRange
    {
        private TFormatList FormatList;
		private TOneCellValue XFDef;

        internal TFormatRange(TFormatList aFormatList, TOneCellValue aRange, TOneCellValue aXFDef):
			base(aRange, TBandType.Static)
        {
            FormatList = aFormatList;
			XFDef = aXFDef;
        }

		internal override void Execute(ExcelFile Workbook, TWaitingCoords Coords, TBand Band)
		{
			int t, l, b, r;
			GetBounds(Coords, out t, out l, out b, out r);

			int fr = Math.Min(t, b);
			int fc = Math.Min(l, r);
			int lr = Math.Max(t, b);
			int lc = Math.Max(l, r);

			TValueAndXF val = new TValueAndXF();
			val.Workbook = Workbook;
			XFDef.Evaluate(0, 0, 0, 0, val);
			TConfigFormat fmt = FormatList.GetValue(FlxConvert.ToString(val.Value));

			if (fmt.ApplyFmt == null)
			{
				Workbook.SetCellFormat(fr, fc, lr, lc, fmt.XF);
			}
			else
			{
				Workbook.SetCellFormat(fr, fc, lr, lc, Workbook.GetFormat(fmt.XF), fmt.ApplyFmt, fmt.ExteriorBorders);
			}

		}

		internal TFormatRange ShallowClone()
		{
			return (TFormatRange)MemberwiseClone();
		}

		internal void SetCell(int aRow, int aCol)
		{
			FTop = aRow;
			Bottom = aRow;
			FLeft = aCol;
			Right = aCol;
		}

		internal void ApplyFormat(ExcelFile Workbook, int RowOfs, int ColOfs)
		{
			Execute(Workbook, new TWaitingCoords(RowOfs, ColOfs, 0, 0, 0, 0), null);
		}
		#region IDisposable Members

		protected override void Dispose(bool Disposing)
		{
			if (Disposing)
			{
				if (XFDef != null) XFDef.Dispose();
                XFDef = null;
			}
			base.Dispose (Disposing);
		}

		#endregion
	}

    /// <summary>
    /// Holds a list of Format Ranges.
    /// </summary>
    internal class TFormatRangeList : IEnumerable, IEnumerable<TFormatRange>
    {
        private List<TFormatRange> FList;
        private int LastCellSet;

        internal TFormatRangeList()
        {
            FList = new List<TFormatRange>();
            LastCellSet = 0;
        }

        internal void Add(TFormatRange fr)
        {
            FList.Add(fr.ShallowClone());
        }

        internal void SetCurrentCell(int Row, int Col)
        {
            int aCount = FList.Count;
            for (int i = LastCellSet; i < aCount; i++)
            {
                FList[i].SetCell(Row, Col);
            }
            LastCellSet = aCount;

        }

        #region IEnumerable Members

        public System.Collections.IEnumerator GetEnumerator()
        {
            return FList.GetEnumerator();
        }

        #endregion

        #region IEnumerable<TFormatRange> Members

        IEnumerator<TFormatRange> IEnumerable<TFormatRange>.GetEnumerator()
        {
            return FList.GetEnumerator();
        }

        #endregion
    }
    #endregion

	#region Exceptions Options
		/// <summary>
		/// Enumerates what to do on different FlexCel error situations.
		/// </summary>
		[Flags]
			public enum TErrorActions
		{
			/// <summary>
			/// FlexCel will try to recover from most errors.
			/// </summary>
			None=0,

			/// <summary>
			/// When true and the number of manual pagebreaks is bigger than the maximum Excel allows,
			/// an Exception will be raised. When false, the page break will be silently ommited.
			/// </summary>
			ErrorOnTooManyPageBreaks=1
		}
		#endregion

    #region Image Params
	internal class TImageSizeParams
	{
		private double FZoom;
		private double FAspectRatio;
		private bool FBoundImage;

		internal TImageSizeParams(double aZoom, double aAspectRatio, bool aBoundImage)
		{
			Zoom = aZoom;
			AspectRatio = aAspectRatio;
			BoundImage = aBoundImage;
		}

		public double Zoom { get { return FZoom; } set { if (value < 0) FZoom = 0; else FZoom = value; } }

		/// <summary>
		/// Negative aspect ratios mean height is preserved. Positives will preserve width.
		/// </summary>
		public double AspectRatio { get { return FAspectRatio; } set { FAspectRatio = value; } }

		public bool BoundImage {get {return FBoundImage;} set{FBoundImage = value;}}


	}

	internal enum TImageHAlign  
	{
		None, Left, Center, Right
	}

	internal enum TImageVAlign  
	{
		None, Top, Center, Bottom
	}
	
	internal class TImagePosParams
	{
		internal TImageVAlign RowAlign;
		internal TImageHAlign ColAlign;
		internal TOneCellValue RowOffs;
		internal TOneCellValue ColOffs;

		internal TImagePosParams(TImageVAlign aRowAlign, TImageHAlign aColAlign, TOneCellValue aRowOffs, TOneCellValue aColOffs)
		{
			RowAlign = aRowAlign;
			ColAlign = aColAlign;
			RowOffs = aRowOffs;
			ColOffs = aColOffs;
		}

        internal static double Eval(ExcelFile Workbook, TOneCellValue v)
        {
            return TImageFitParams.Eval(Workbook, v);
        }

	}

	internal class TImageFitParams
	{
		internal TAutofitGrow FitInRows;
		internal TAutofitGrow FitInCols;
		internal TOneCellValue RowMargin;
		internal TOneCellValue ColMargin;

		internal TImageFitParams(TAutofitGrow aFitInRows, TAutofitGrow aFitInCols, TOneCellValue aRowMargin, TOneCellValue aColMargin)
		{
			FitInRows = aFitInRows;
			FitInCols = aFitInCols;
			RowMargin = aRowMargin;
			ColMargin = aColMargin;
		}

        internal static double Eval(ExcelFile Workbook, TOneCellValue v)
        {
            if (v == null) return 0;

            TValueAndXF val = new TValueAndXF();
            val.Workbook = Workbook;
            v.Evaluate(0, 0, 0, 0, val);
            return Convert.ToDouble(val.Value);
        }
	}

	#endregion

    #region TExpression
    internal class TExpression
    {
        internal string[] Parameters;
        internal object Value;

        public TExpression(string[] aParameters, object aValue)
        {
            Parameters = aParameters;
            Value = aValue;
        }
    }
    #endregion

    #region String utils
    internal sealed class StrUtils
    {
        private StrUtils() { }

        internal static string[] Split(string s, string separator)
        {
            if (s == null || separator == null || separator.Length == 0) return new string[0];

            List<string> Result = new List<string>();
            int pos = 0;
            do
            {
                int newpos = s.IndexOf(separator, pos);
                if (newpos < 0) newpos = s.Length;
                Result.Add(s.Substring(pos, newpos - pos));
                pos = newpos + separator.Length;
            }
            while (pos < s.Length);

            return Result.ToArray();
        }
    }
    #endregion

    #region TIncludeHtml
    internal enum TIncludeHtml
    {
        Undefined = 0,
        Yes,
        No
    }
    #endregion

    #region ImageData
    internal struct TCopiedImageData
    {
        internal int RecordPos;
        internal TExcelObjectList OrigObjects;

        internal TCopiedImageData(int aRecordPos)
        {
            RecordPos = 0;
            OrigObjects = new TExcelObjectList(true);
        }

        internal long[] GetObjects()
        {
            return OrigObjects.GetObjects(RecordPos);
        }
    }
    #endregion

    #region TAutofitInfo
    internal class TAutofitInfo
    {
        internal TAutofitType AutofitType;
		internal float GlobalAdjustment;
		internal int GlobalAdjustmentFixed;
        internal bool KeepAutofit;
		internal TAutofitMerged MergedMode;

        internal TAutofitInfo()
        {
            AutofitType = TAutofitType.None;
            GlobalAdjustment = 1;
            KeepAutofit = true;
			MergedMode = TAutofitMerged.OnLastCell;
        }

    }

    internal enum TAutofitType
    {
        //Do not autofit anything.
        None,

        //Autofit all rows on the sheet.
        Sheet,

        //Autofit only the rows and columns marked for autofit.
        OnlyMarked
    }

	internal enum TAutofitGrow
	{
		//Normal
		All,

		//Only grow the rows, never shrink them.
		DontShrink,

		//Only shrink, never grow.
		DontGrow,

		//dont do autofit
		None

	}
    #endregion

    #region TKeepTogether

    internal class TKeepTogether
    {
        internal int R1;
        internal int R2;
        internal int Level;

        internal TKeepTogether(int aR1, int aR2, int aLevel)
        {
            R1 = aR1;
            R2 = aR2;
            Level = aLevel;
        }
    }
    #endregion

	#region SheetState
	internal class TSheetState
	{
		internal int AutoPageBreaksPercent;
		internal int AutoPageBreaksPageScale;

		internal TAutofitInfo AutofitInfo;

		internal TSheetState()
		{
			AutoPageBreaksPercent = -1;
			AutoPageBreaksPageScale = -1;
			AutofitInfo = new TAutofitInfo();
		}
	}
	#endregion


}
