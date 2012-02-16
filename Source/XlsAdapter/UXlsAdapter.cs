using System;
using System.IO;
using System.Reflection;
using System.Resources;
using System.Globalization;
using System.Text;
using FlexCel.Core;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Collections;

#if (MONOTOUCH)
    using Color = MonoTouch.UIKit.UIColor;
    using System.Drawing;
    using real = System.Single;
#else
#if (WPF)
using RectangleF = System.Windows.Rect;
  using PointF = System.Windows.Point;
  using System.Windows.Media;
using real = System.Double;
  #else
  using System.Drawing;
  using System.Drawing.Drawing2D;
  using real = System.Single;
  #endif
#endif

    // Note: Excel uses 1-Based arrays, and that's the interface we present in this API.
    // But, TExcelWorkbook uses 0-Based arrays, to be consistent with the file format (made in C)
    // So here we have to add and substract 1 everywere to be consistent.

namespace FlexCel.XlsAdapter
{
    /// <summary>
    /// This is the FlexCel Native Engine. Use this class to natively read or write an Excel 97 or newer file. Note that to read xlsx files
    /// (Excel 2007 or newer) you need FlexCel for .NET Framework 3.5 or newer.
    /// </summary>
    /// <remarks>
    /// Note that most arrays here are 1-based and not 0-based, to follow the convention used on Excel Automation.
    /// So, if you are using C# or C++, your loops should look like:
    /// <code>for (int i=1;i&lt;=SomeProperty.Count;i++)</code> and not
    /// <code>for (int i=0;i&lt;SomeProperty.Count;i++)</code>
    /// </remarks>
    [ClassInterface(ClassInterfaceType.None), ComVisible(false)]
    public class XlsFile : ExcelFile
    {
        #region Private variables
        private TWorkbook FWorkbook; //You can access it internally with internalworkbook.
        private byte[] OtherStreams;
        private TUnsupportedFormulaList FUnsupportedFormulaList;
#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
        private byte[] OtherXlsxParts;
        private byte[] MacroData;
#endif

        private int FActiveSheet;

        private TSheet FActiveSheetObject;
        private string FActiveFileName = String.Empty;
        private TRecalcMode FRecalcMode;
        private bool FNeedsRecalc = false;
        private bool FRecalcForced = true;
        private bool FRecalculating = false;

        private real FHeightCorrection = 1;
        private real FWidthCorrection = 1;
        private double FLinespacing = 1;
        private bool FIgnoreFormulaText = false;

        internal TFileProps FileProps = new TFileProps();
        private TOle2Properties Ole2StdProperties = null;
        private TOle2Properties Ole2ExtProperties = null;

        private TFileFormats FFileFormatWhenOpened;

        private TCustomFormulaList LocalFormulaFunctions = new TCustomFormulaList();
        private static TCustomFormulaList GlobalFormulaFunctions = new TCustomFormulaList(); //STATIC**
        private static TCustomFormulaList BuiltInFormulaFunctions = GetBuiltInUserDefFunctions(); //STATIC**
        private static object GlobalFormulaFunctionAccess = new object();

        #endregion

        /// <summary>
        /// Creates a new XlsFile.
        /// </summary>
        public XlsFile()
            : base()
        {
            FActiveFileName = String.Empty;
            FWorkbook = new TWorkbook(this, FileProps);
            FRecalcMode = TRecalcMode.Smart;
            FNeedsRecalc = false;
            FRecalculating = false;
            FFileFormatWhenOpened = TFileFormats.Xls;
        }

        /// <summary>
        /// Creates a new XlsFile and sets the desired Overwriting mode for files.
        /// </summary>
        /// <param name="aAllowOverwritingFiles">When true calling "Save" will overwrite existing files. See <see cref="FlexCel.Core.ExcelFile.AllowOverwritingFiles"/></param>
        public XlsFile(bool aAllowOverwritingFiles)
            : this()
        {
            AllowOverwritingFiles = aAllowOverwritingFiles;
        }

        /// <summary>
        /// Creates a new XlsFile and opens the desired file.
        /// </summary>
        /// <param name="aFileName">Name of the file to open.</param>
        public XlsFile(string aFileName)
            : base()
        {
            FActiveFileName = String.Empty;
            FWorkbook = new TWorkbook(this, FileProps);
            FRecalcMode = TRecalcMode.Smart;
            FNeedsRecalc = false;
            FRecalculating = false;
            Open(aFileName);
        }

        /// <summary>
        /// Creates a new XlsFile and opens the desired file. Sets the desired Overwriting mode for files.
        /// </summary>
        /// <param name="aFileName">Name of the file to open.</param>
        /// <param name="aAllowOverwritingFiles">When true calling "Save" will overwrite existing files. See <see cref="FlexCel.Core.ExcelFile.AllowOverwritingFiles"/></param>
        public XlsFile(string aFileName, bool aAllowOverwritingFiles)
            : this(aFileName)
        {
            AllowOverwritingFiles = aAllowOverwritingFiles;
        }

        ///<inheritdoc />
        public override bool IsXltTemplate
        {
            get
            {
                CheckWkConnected();
                return FWorkbook.IsXltTemplate;
            }
            set
            {
                CheckWkConnected();
                FWorkbook.IsXltTemplate = value;
            }
        }

        #region Internal User Defined Functions
        private static TCustomFormulaList GetBuiltInUserDefFunctions()
        {
            TCustomFormulaList Result = new TCustomFormulaList();
            AddFn(Result, new FlexCel.AddinFunctions.DurationImpl(false));
            AddFn(Result, new FlexCel.AddinFunctions.DurationImpl(true));
            AddFn(Result, new FlexCel.AddinFunctions.EDateImpl());
            AddFn(Result, new FlexCel.AddinFunctions.CoupDaysBSImpl());
            AddFn(Result, new FlexCel.AddinFunctions.CoupDaysImpl());
            AddFn(Result, new FlexCel.AddinFunctions.CoupDaysNCImpl());
            AddFn(Result, new FlexCel.AddinFunctions.CoupNCDImpl());
            AddFn(Result, new FlexCel.AddinFunctions.CoupNumImpl());
            AddFn(Result, new FlexCel.AddinFunctions.CoupPCDImpl());

            AddFn(Result, new FlexCel.AddinFunctions.YearFracImpl());

            AddFn(Result, new FlexCel.AddinFunctions.Bin2DecImpl());
            AddFn(Result, new FlexCel.AddinFunctions.Bin2OctImpl());
            AddFn(Result, new FlexCel.AddinFunctions.Bin2HexImpl());

            AddFn(Result, new FlexCel.AddinFunctions.Dec2BinImpl());
            AddFn(Result, new FlexCel.AddinFunctions.Dec2OctImpl());
            AddFn(Result, new FlexCel.AddinFunctions.Dec2HexImpl());

            AddFn(Result, new FlexCel.AddinFunctions.Oct2BinImpl());
            AddFn(Result, new FlexCel.AddinFunctions.Oct2DecImpl());
            AddFn(Result, new FlexCel.AddinFunctions.Oct2HexImpl());

            AddFn(Result, new FlexCel.AddinFunctions.Hex2BinImpl());
            AddFn(Result, new FlexCel.AddinFunctions.Hex2DecImpl());
            AddFn(Result, new FlexCel.AddinFunctions.Hex2OctImpl());

            AddFn(Result, new FlexCel.AddinFunctions.IsOddImpl());
            AddFn(Result, new FlexCel.AddinFunctions.IsEvenImpl());
            AddFn(Result, new FlexCel.AddinFunctions.DeltaImpl());

            AddFn(Result, new FlexCel.AddinFunctions.DollarDeImpl());
            AddFn(Result, new FlexCel.AddinFunctions.DollarFrImpl());

            AddFn(Result, new FlexCel.AddinFunctions.EffectImpl());
            AddFn(Result, new FlexCel.AddinFunctions.EOMonthImpl());
            AddFn(Result, new FlexCel.AddinFunctions.FactDoubleImpl());
            AddFn(Result, new FlexCel.AddinFunctions.GcdImpl());
            
            AddFn(Result, new FlexCel.AddinFunctions.GeStepImpl());

            AddFn(Result, new FlexCel.AddinFunctions.LcmImpl());
            AddFn(Result, new FlexCel.AddinFunctions.MRoundImpl());

            AddFn(Result, new FlexCel.AddinFunctions.MultinomialImpl());

            AddFn(Result, new FlexCel.AddinFunctions.NetWorkDaysImpl(false));
            AddFn(Result, new FlexCel.AddinFunctions.NetWorkDaysImpl(true), TUserDefinedFunctionLocation.Internal); //_xlfn., it's internal.
            AddFn(Result, new FlexCel.AddinFunctions.WorkDayImpl(false));
            AddFn(Result, new FlexCel.AddinFunctions.WorkDayImpl(true), TUserDefinedFunctionLocation.Internal); //_xlfn., it's internal.

            AddFn(Result, new FlexCel.AddinFunctions.NominalImpl());
            AddFn(Result, new FlexCel.AddinFunctions.QuotientImpl());
            AddFn(Result, new FlexCel.AddinFunctions.RandBetweenImpl());

            AddFn(Result, new FlexCel.AddinFunctions.SeriesSumImpl());
            AddFn(Result, new FlexCel.AddinFunctions.SqrtPiImpl());

            AddFn(Result, new FlexCel.AddinFunctions.WeekNumImpl());

            AddFn(Result, new FlexCel.AddinFunctions.ConvertImpl());

            return Result;
        }

        private static void AddFn(TCustomFormulaList Result, TUserDefinedFunction fn)
        {
            AddFn(Result, fn, TUserDefinedFunctionLocation.External);
        }

        private static void AddFn(TCustomFormulaList Result, TUserDefinedFunction fn, TUserDefinedFunctionLocation Location)
        {
            TUserDefinedFunctionContainer fnc;
            fnc = new TUserDefinedFunctionContainer(Location, fn);
            Result.Add(fnc);
        }

        #endregion

        #region Utilities
        private void CheckConnected()
        {
            if (FActiveSheet < 1) FlxMessages.ThrowException(FlxErr.ErrNotConnected);
        }

        private void CheckWkConnected()
        {
            if (FWorkbook == null || !FWorkbook.Loaded) FlxMessages.ThrowException(FlxErr.ErrNotConnected);
        }

        internal TSheet ActiveSheetObject
        {
            get
            {
                return FActiveSheetObject;
            }
        }

        private static void CheckRowAndCol(int row, int col)
        {
            if ((row < 1) || (row > FlxConsts.Max_Rows + 1)) FlxMessages.ThrowException(FlxErr.ErrInvalidRow, row);
            if ((col < 1) || (col > FlxConsts.Max_Columns + 1)) FlxMessages.ThrowException(FlxErr.ErrInvalidColumn, col);
        }

        private void CheckSheet(int sheet)
        {
            if (sheet < 1 || sheet > FWorkbook.Sheets.Count)
                XlsMessages.ThrowException(XlsErr.ErrInvalidSheetNo, sheet, 1, FWorkbook.Sheets.Count);
        }

        internal static void CheckRange(int val, int lowest, int highest, FlxParam paramName)
        {
            if (lowest > highest) FlxMessages.ThrowException(FlxErr.ErrInvalidValue2, paramName.ToString(), val);
            if ((val < lowest) || (val > highest)) FlxMessages.ThrowException(FlxErr.ErrInvalidValue, paramName.ToString(), val, lowest, highest);
        }

        internal static void CheckRangeObjPath(string ObjPath, int val, int lowest, int highest, FlxParam paramName)
        {
            if (ObjPath != null && 
                (ObjPath.StartsWith(FlxConsts.ObjectPathAbsolute, StringComparison.InvariantCulture)
                || ObjPath.StartsWith(FlxConsts.ObjectPathObjName, StringComparison.InvariantCulture)
                || ObjPath.StartsWith(FlxConsts.ObjectPathSpId, StringComparison.InvariantCulture)
                )) return;
            CheckRange(val, lowest, highest, paramName);
        }

        private static void Swap(ref int a1, ref int a2)
        {
            int tmp = a1;
            a1 = a2;
            a2 = tmp;
        }

        private void RestoreObjectSizes()
        {
            FActiveSheetObject.RestoreObjectCoords();
        }

        #endregion

        #region Internal Methods
        internal TWorkbook InternalWorkbook
        {
            get
            {
                return FWorkbook;
            }
        }
        #endregion

        #region IExcelFile Members

        #region Globals

        #endregion

        #region File Managment

        ///<inheritdoc />
        public override TExcelFileFormat ExcelFileFormat
        {
            get 
            {
                CheckWkConnected();
                return FWorkbook.Globals.sBOF.BiffFileFormat();
            }
        } 

        ///<inheritdoc />
        public override string ActiveFileName
        {
            get
            {
                return FActiveFileName;
            }
            set
            {
                FActiveFileName = value;
            }
        }

        ///<inheritdoc />
        public override void NewFile(int aSheetCount, TExcelFileFormat fileFormat)
        {
            CheckRange(aSheetCount, 1, FlxConsts.Max_Sheets + 1, FlxParam.SheetCount);
            FActiveFileName = String.Empty;

            using (Stream MemStream = TWorkSheet.GetEmptyWorkbook(fileFormat))
            {
                OpenXls(MemStream, false);
            }

            if (aSheetCount > 1) InsertAndCopySheets(0, 2, aSheetCount - 1);

        }

        ///<inheritdoc />
        public override void Open(string fileName, TFileFormats fileFormat, char delimiter, int firstRow, int firstCol, ColumnImportType[] columnFormats, string[] dateFormats, Encoding fileEncoding, bool detectEncodingFromByteOrderMarks)
        {
            FActiveFileName = String.Empty;
            using (FileStream f = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) //FileShare.ReadWrite is the way we have to open a file even if it is being used by excel.
            {
                Open(f, fileFormat, delimiter, firstRow, firstCol, columnFormats, dateFormats, fileEncoding, detectEncodingFromByteOrderMarks);
            }
            FActiveFileName = fileName;
        }

        ///<inheritdoc />
        public override void Open(Stream aStream, TFileFormats fileFormat, char delimiter, int firstRow, int firstCol, ColumnImportType[] columnFormats, string[] dateFormats, Encoding fileEncoding, bool detectEncodingFromByteOrderMarks)
        {
            FActiveFileName = String.Empty;
            //CheckConnected();

            switch (fileFormat)
            {
                case TFileFormats.Automatic:
                    //Automatic
                    if (!OpenXls(aStream, true))
                        if (!OpenXlsx(aStream, true))
                            if (!OpenPxl(aStream, true))
                                XlsMessages.ThrowException(XlsErr.ErrFileIsNotSupported);
                    break;

                case TFileFormats.Xls:
                    OpenXls(aStream, false);
                    break;

                case TFileFormats.Text:
                    using (StreamReader sr = new StreamReader(new TUndisposableStream(aStream), fileEncoding, detectEncodingFromByteOrderMarks))
                    {
                        OpenText(sr, delimiter, firstRow, firstCol, columnFormats, dateFormats);
                    }
                    break;

                case TFileFormats.Pxl:
                    OpenPxl(aStream, false);
                    break;

                case TFileFormats.Xlsx:
                case TFileFormats.Xlsm:
                    OpenXlsx(aStream, false);
                    break;

                default:
                    XlsMessages.ThrowException(XlsErr.ErrFileIsNotSupported);
                    break;
            }


            FileStream fStream = aStream as FileStream;
            if (fStream != null) FActiveFileName = fStream.Name;
            NeedsRecalc = false;
            FRecalculating = false;
            if (RecalcMode == TRecalcMode.Manual)
                FWorkbook.ClearFormulaResults();  //Here we provide a "Clean" sheet, so if we later assign formula values
            //they will stay. If recalc is not manual, we do not want to clear them,
            //so we can save a file without changes and preserve formula results.

            OnVirtualCellEndReading(this, new VirtualCellEndReadingEventArgs());
        }


        private bool OpenXls(Stream aStream, bool AvoidExceptions)
        {
            long StreamPosition = aStream.Position;
            using (MemoryStream MemFile = new MemoryStream())
            {
                using (TOle2File DataStream = new TOle2File(aStream, AvoidExceptions))
                {
                    if (DataStream.NotXls97) return false;
                    if (!DataStream.SelectStream(XlsConsts.WorkbookString, AvoidExceptions))
                    {
                        aStream.Position = StreamPosition;
                        return false;
                    }

                    using (FWorkbook.Globals.GetBiff8XFGuard())
                    {
                        TSST SST = GetSST(FWorkbook);
                        TVirtualReader VirtualReader = CreateVirtualReader();
                        
                        TXlsRecordLoader RecordLoader = new TXlsRecordLoader(DataStream, FWorkbook.Globals.Biff8XF,
                            SST, this,
                            FWorkbook.Globals.Borders, FWorkbook.Globals.Patterns,
                            new TEncryptionData(Protection.OpenPassword, Protection.OnPassword, Protection.Xls),
                            XlsBiffVersion, FWorkbook.Globals.Names, VirtualReader);
                        FWorkbook.LoadFromStream(RecordLoader, Protection);  //Saves the workbook stream.
                    }
                    DataStream.PrepareForWrite(MemFile, XlsConsts.WorkbookString, new string[0]);  //Saves all the other streams.
                }
                OtherStreams = MemFile.ToArray();//do this after the using.
#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
                MacroData = null;
#endif
                Ole2StdProperties = null;
                Ole2ExtProperties = null;
                FileProps.Clear();
            }
            FActiveSheet = FWorkbook.ActiveSheet + 1;
            FActiveSheetObject = FWorkbook.Sheets[FActiveSheet - 1] as TSheet;
            FFileFormatWhenOpened = TFileFormats.Xls;
            return true;
        }

        private TSST GetSST(TWorkbook aWorkbook)
        {
            if (VirtualMode) return new TSST();
            return aWorkbook.Globals.SST;
        }

        private TVirtualReader CreateVirtualReader()
        {
            if (!VirtualMode) return null;
            TVirtualReader VirtualReader = new TVirtualReader(this, FWorkbook);
            return VirtualReader;
        }

        private bool OpenXlsx(Stream aStream, bool AvoidExceptions)
        {
#if (!FRAMEWORK30 || COMPACTFRAMEWORK)
			if (AvoidExceptions) return false;
            XlsMessages.ThrowException(XlsErr.ErrFileIsNotSupported);
            return false;
#else
            bool MacroEnabled;
            MacroData = null;
            using (TOpenXmlReader DataStream = new TOpenXmlReader(aStream, AvoidExceptions, 
                new TEncryptionData(Protection.OpenPassword, Protection.OnPassword, Protection.Xls), ActiveFileName, ErrorActions))
            {
                if (DataStream.NotXlsx) return false;
                FileProps.Clear();
                TSST SST = GetSST(FWorkbook);
                TVirtualReader VirtualReader = CreateVirtualReader();
                TXlsxRecordLoader RecordLoader = new TXlsxRecordLoader(DataStream, this, SST, VirtualReader);
                FWorkbook.LoadFromStream(RecordLoader, Protection, out MacroEnabled);
                OtherXlsxParts = DataStream.ReadOtherParts();
            }
            Ole2StdProperties = null;
            Ole2ExtProperties = null;
            OtherStreams = GetGenericOtherStreams();


            FActiveSheet = FWorkbook.ActiveSheet + 1;
            FActiveSheetObject = FWorkbook.Sheets[FActiveSheet - 1];
            if (MacroEnabled) FFileFormatWhenOpened = TFileFormats.Xlsm; else FFileFormatWhenOpened = TFileFormats.Xlsx;
            return true;
#endif
        }

        private static byte[] GetGenericOtherStreams()
        {
            using (Stream InStream = TWorkSheet.GetEmptyWorkbook(TExcelFileFormat.v2007))
            {
                using (MemoryStream OutStream = new MemoryStream())
                {
                    using (TOle2File DataStream = new TOle2File(InStream))
                    {
                        DataStream.PrepareForWrite(OutStream, XlsConsts.WorkbookString, new string[0]);
                    }
                    return OutStream.ToArray();//do this after the using.
                }
            }
        }

        private bool OpenPxl(Stream aStream, bool AvoidExceptions)
        {
            TExternSheetList ExternSheetList = new TExternSheetList();
            TWorkbook PxlWorkbook = new TWorkbook(this, FileProps);
            using (PxlWorkbook.Globals.GetBiff8XFGuard())
            {
                NewFile(1);  //Create a new empty file where to merge the old one. We need to do this before loading the pxl

                TSST SST = GetSST(PxlWorkbook);
                TVirtualReader VirtualReader = CreateVirtualReader();

                TPxlRecordLoader RecordLoader = new TPxlRecordLoader(aStream, ExternSheetList,
                new TEncryptionData(Protection.OpenPassword, Protection.OnPassword, Protection.Xls),
                SST, this,
                PxlWorkbook.Globals.Borders, PxlWorkbook.Globals.Patterns, this, PxlWorkbook.Globals.Biff8XF, FWorkbook.Globals.CellXF.Count,
                FWorkbook.Globals.Names, VirtualReader);
                if (!RecordLoader.CheckHeader()) return false;

                PxlWorkbook.LoadFromStream(RecordLoader, Protection);  //Saves the workbook stream.

                if (PxlWorkbook.Sheets.Count > 1) InsertAndCopySheets(0, 2, PxlWorkbook.Sheets.Count - 1);
                FWorkbook.MergeFromPxlWorkbook(PxlWorkbook);
            }

            for (int i = 0; i < ExternSheetList.Count; i++)
            {
                TExternSheetEntry Es = (TExternSheetEntry)ExternSheetList[i];
                int Index = FWorkbook.Globals.References.AddSheet(FWorkbook.Globals.SheetCount, Es.FirstSheet, Es.LastSheet);
                Debug.Assert(Index == i, "Index Should be equal to i");
            }

            FActiveSheet = FWorkbook.ActiveSheet + 1;
            FActiveSheetObject = FWorkbook.Sheets[FActiveSheet - 1];
            FFileFormatWhenOpened = TFileFormats.Pxl;
            return true;
        }

        private void OpenText(StreamReader aStreamReader, char delimiter, int FirstRow, int FirstCol, ColumnImportType[] ColumnFormats, string[] DateFormats)
        {
            NewFile(1);
            TextDelim.Read(aStreamReader, this, delimiter, FirstRow, FirstCol, ColumnFormats, DateFormats); //CF!
            FFileFormatWhenOpened = TFileFormats.Xls; //Text doesn't make a good default for saving.
        }


        private TFileFormats GetFileFormat(string FileName, TFileFormats FileFormat)
        {
            if (FileFormat != TFileFormats.Automatic) return FileFormat;
            if (FileName != null)
            {
                switch (Path.GetExtension(FileName).ToLower(CultureInfo.InvariantCulture))
                {
                    case ".xlsx":
                    case ".xltx":
                        return TFileFormats.Xlsx;

                    case ".xlsm":
                    case ".xltm":
                        return TFileFormats.Xlsm;

                    case ".xls":
                    case ".xlt": 
                        return TFileFormats.Xls;

                    case ".csv": return TFileFormats.Text;

                    case ".txt": return TFileFormats.Text;

                    case ".pxl": return TFileFormats.Pxl;
                }
            }

            return DefaultFileFormat;
        }

        ///<inheritdoc />
        public override void Import(TextReader aTextReader, int firstRow, int firstCol, char delimiter, ColumnImportType[] columnFormats, string[] dateFormats)
        {
            TextDelim.Read(aTextReader, this, delimiter, firstRow, firstCol, columnFormats, dateFormats);
        }

        ///<inheritdoc />
        public override void Import(TextReader aTextReader, int firstRow, int firstCol, int[] columnWidths, ColumnImportType[] columnFormats, string[] dateFormats)
        {
            TextFixedWidth.Read(aTextReader, this, columnWidths, firstRow, firstCol, columnFormats, dateFormats);
        }

        ///<inheritdoc />
        public override void Save(string fileName, TFileFormats fileFormat, char delimiter, Encoding fileEncoding)
        {
            try
            {
                fileFormat = GetFinalFileFormat(fileName, fileFormat);
                FileMode fm = FileMode.CreateNew;
                if (AllowOverwritingFiles) fm = FileMode.Create;
                FileAccess fa = FileAccess.Write;
                if (fileFormat == TFileFormats.Xlsx || fileFormat == TFileFormats.Xlsm) fa = FileAccess.ReadWrite;
                
                using (FileStream f = new FileStream(fileName, fm, fa))
                {
                    Save(f, fileFormat, delimiter, fileEncoding);
                }
            }

            catch (IOException)
            {
                //Don't delete the file in an io exception. It might be because allowoverwritefiles was false, and the file existed.
                throw;
            }
            catch (FlexCelException)
            {
                File.Delete(fileName);
                throw;
            }
            FActiveFileName = fileName;
        }

        private bool PendingRecalc()
        {
            return
                (RecalcMode == TRecalcMode.Smart && NeedsRecalc) ||
                RecalcMode == TRecalcMode.Forced ||
                RecalcMode == TRecalcMode.OnEveryChange;
        }

        ///<inheritdoc />
        public override void Save(Stream aStream, TFileFormats fileFormat, char delimiter, Encoding fileEncoding)
        {
            CheckConnected();

            FileStream fs = aStream as FileStream;
            string fileName = fs != null? fs.Name: null;
            fileFormat = GetFinalFileFormat(fileName, fileFormat);

            if (OtherStreams == null) FlxMessages.ThrowException(FlxErr.ErrNotConnected);

            for (int i = 0; i < SheetCount; i++)
                if (i == FActiveSheet - 1) FWorkbook.Sheets[i].Selected = true; else FWorkbook.Sheets[i].Selected = false;

            Recalc(false);

            switch (fileFormat)
            {
                case TFileFormats.Automatic:
                    XlsMessages.ThrowException(XlsErr.ErrInternal);
                    break;
                case TFileFormats.Xls:
                    {
                        byte[] RestOfStreams = OtherStreams;


                        using (MemoryStream MemFile = new MemoryStream(RestOfStreams))
                        {
                            using (TOle2File DataStream = new TOle2File(MemFile))
                            {
                                string Pass = Protection.OpenPassword;
                                if (Pass.Length == 0 && Protection.HasWorkbookPassword) Pass = XlsConsts.EmptyExcelPassword;
                                if (Pass.Length > 0)
                                {
                                    if (Protection.EncryptionType == TEncryptionType.Xor)
                                        DataStream.Encryption.Engine = new TXorEncryption(Pass);
                                    else
                                        if (Protection.EncryptionType == TEncryptionType.Standard)
                                            DataStream.Encryption.Engine = new TStandardEncryption(Pass, FTesting);
                                        else
                                            XlsMessages.ThrowException(XlsErr.ErrNotSupportedEncryption);
                                }
                                DataStream.PrepareForWrite(aStream, XlsConsts.WorkbookString, new string[0]);
                                FWorkbook.SaveToStream(DataStream, new TSaveData(this, FWorkbook.Globals.Borders, FWorkbook.Globals, XlsBiffVersion, ErrorActions));
                            }
                        }
                    }
                    break;

                case TFileFormats.Text:
                    using (StreamWriter sw = new StreamWriter(new TUndisposableStream(aStream), fileEncoding))
                    {
                        TextDelim.Write(sw, this, delimiter, true);
                    }
                    break;

                case TFileFormats.Pxl:
                    FWorkbook.SaveToPxl(new TPxlStream(aStream));
                    break;

                case TFileFormats.Xlsx:
                    SaveAsXlsx(aStream, false);
                    break;
                
                case TFileFormats.Xlsm:
                    SaveAsXlsx(aStream, true);
                    break;
                default:
                    XlsMessages.ThrowException(XlsErr.ErrFileIsNotSupported);
                    break;
            }
        }

        private TFileFormats GetFinalFileFormat(string fileName, TFileFormats fileFormat)
        {
            //Check a normal case.
            if (fileFormat == TFileFormats.Automatic)
            {
                fileFormat = GetFileFormat(fileName, fileFormat);
            }

            if (fileFormat == TFileFormats.Automatic)
            {
                fileFormat = FFileFormatWhenOpened;
            }
            return fileFormat;
        }

        private void SaveAsXlsx(Stream aStream, bool MacroEnabled)
        {
#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
            Stream RealStream = aStream;
            try
            {
                if (!String.IsNullOrEmpty(Protection.OpenPassword))
                {
                    RealStream = new TXlsxCryptoStreamWriter(aStream, Protection);
                }

                using (TOpenXmlWriter DataStream = new TOpenXmlWriter(RealStream, AllowOverwritingFiles))
                {
                    TXlsxRecordWriter RecordWriter = new TXlsxRecordWriter(DataStream, this, FWorkbook, FileProps);
                    RecordWriter.Save(Protection, OtherXlsxParts, MacroEnabled);
                }
            }
            finally
            {
                if (RealStream != aStream) RealStream.Dispose();
            }
#else
            XlsMessages.ThrowException(XlsErr.ErrFileIsNotSupported);
#endif
        }


        private byte[] RepeatableOtherStreams()
        {
            using (MemoryStream MemFile = new MemoryStream(OtherStreams))
            {
                using (MemoryStream NewMemFile = new MemoryStream())
                {
                    using (TOle2File DataStream = new TOle2File(MemFile))
                    {
                        string[] Dirs = DataStream.ListStreams();
                        Array.Sort(Dirs);
                        foreach (string s in Dirs)
                        {
                            string[] fullname = s.Split((char)0);
                            if (fullname.Length < 1) continue;
                            string sn = fullname[fullname.Length - 1];
                            if (sn == XlsConsts.WorkbookString || sn == XlsConsts.DocumentPropertiesStringExtended ||
                                sn == XlsConsts.DocumentPropertiesString || sn == XlsConsts.ProjectString) continue; //extended props have timestamps.

                            DataStream.SelectStream(s);
                            byte[] b = Encoding.ASCII.GetBytes("*FLX" + s + "*");
                            NewMemFile.Write(b, 0, b.Length);
                            byte[] buff = new byte[DataStream.Length];
                            DataStream.Read(buff, buff.Length);
                            NewMemFile.Write(buff, 0, buff.Length);
                        }
                    }
                    return NewMemFile.ToArray();  //do this after the using
                }
            }
        }

        ///<inheritdoc />
        public override void SaveForHashing(Stream aStream, TExcludedRecords excludedRecords)
        {
            CheckConnected();
            byte[] b = Encoding.ASCII.GetBytes("FlexCel Repeatable File Format 1.0");
            aStream.Write(b, 0, b.Length);

            if (OtherStreams != null)
            {
                byte[] os = RepeatableOtherStreams(); //Document settings have timestamps that will mess the file.
                aStream.Write(os, 0, os.Length);
            }


#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
            if (MacroData != null) aStream.Write(MacroData, 0, MacroData.Length);
            if (OtherXlsxParts != null) aStream.Write(OtherXlsxParts, 0, OtherXlsxParts.Length);
#endif

            b = Encoding.ASCII.GetBytes("*FLXWORKBOOK*");
            aStream.Write(b, 0, b.Length);

            using (MemOle2 DataStream = new MemOle2(aStream))
            {
                FWorkbook.SaveToStream(DataStream, new TSaveData(this, FWorkbook.Globals.Borders, FWorkbook.Globals, XlsBiffVersion, ErrorActions, excludedRecords, true));
            }
        }


        ///<inheritdoc />
        public override void Export(TextWriter aTextWriter, TXlsCellRange range, char delimiter, bool exportHiddenRowsOrColumns)
        {
            TextDelim.Write(aTextWriter, this, delimiter, range, exportHiddenRowsOrColumns);
        }

        ///<inheritdoc />
        public override void Export(TextWriter aTextWriter, TXlsCellRange range, int charactersForFirstColumn, int[] columnWidths, bool exportHiddenRowsOrColumns, bool exportTextOutsideCells)
        {
            TextFixedWidth.Write(aTextWriter, this, range, columnWidths, charactersForFirstColumn, exportHiddenRowsOrColumns, exportTextOutsideCells);
        }
        #endregion

        #region Sheet Operations

        ///<inheritdoc />
        public override int ActiveSheet
        {
            get
            {
                return FActiveSheet;
            }
            set
            {
                CheckConnected();
                CheckRange(value, 1, SheetCount, FlxParam.ActiveSheet);
                FWorkbook.ActiveSheet = value - 1;
                FActiveSheet = value;
                FActiveSheetObject = FWorkbook.Sheets[FActiveSheet - 1];
            }
        }

        internal void FastActiveSheet(int aSheet)
        {
            FActiveSheet = aSheet;
            FActiveSheetObject = FWorkbook.Sheets[FActiveSheet - 1];
        }


        ///<inheritdoc />
        public override string ActiveSheetByName
        {
            get
            {
                return SheetName;
            }
            set
            {
                ActiveSheet = GetSheetIndex(value);
            }
        }

        ///<inheritdoc />
        public override int GetSheetIndex(string sheetName, bool throwException)
        {
            CheckWkConnected();

            int idx;
            if (FWorkbook.FindSheet(sheetName, out idx)) return idx + 1;
            
            if (!throwException)
                return -1;
            FlxMessages.ThrowException(FlxErr.ErrInvalidSheet, sheetName);
            return 0; //just to compile
        }

        ///<inheritdoc />
        public override string GetSheetName(int sheetIndex)
        {
            CheckRange(sheetIndex, 1, SheetCount, FlxParam.ActiveSheet);
            return FWorkbook.Globals.GetSheetName(sheetIndex - 1);
        }

        internal override void GetSheetsFromExternSheet(int externSheet, out int Sheet1, out int Sheet2, out bool ExternalSheets, out string ExternBookName)
        {
            CheckWkConnected();
            FWorkbook.Globals.References.GetSheetsFromExternSheet(externSheet, out Sheet1, out Sheet2, out ExternalSheets, out ExternBookName);
        }

        ///<inheritdoc />
        public override int SheetCount
        {
            get
            {
                return FWorkbook.Globals.SheetCount;
            }
        }


        ///<inheritdoc />
        public override string SheetName
        {
            get
            {
                CheckWkConnected();
                return FWorkbook.Globals.GetSheetName(FActiveSheet - 1);
            }
            set
            {
                CheckWkConnected();
                FWorkbook.Globals.SetSheetName(FActiveSheet - 1, value);
            }
        }


        ///<inheritdoc />
        public override string SheetCodeName
        {
            get
            {
                CheckConnected();
                return FWorkbook.Sheets[FActiveSheet - 1].CodeName;
            }
        }


        ///<inheritdoc />
        public override TXlsSheetVisible SheetVisible
        {
            get
            {
                CheckConnected();
                return FWorkbook.Globals.GetSheetVisible(FActiveSheet - 1);
            }
            set
            {
                CheckConnected();
                FWorkbook.Globals.SetSheetVisible(FActiveSheet - 1, value);
            }
        }

        ///<inheritdoc />
        public override int SheetZoom
        {
            get
            {
                CheckConnected();
                return FActiveSheetObject.Zoom;
            }
            set
            {
                CheckConnected();
                FActiveSheetObject.Zoom = value;
            }
        }

        ///<inheritdoc />
        public override int FirstSheetVisible
        {
            get
            {
                CheckWkConnected();
                return FWorkbook.Globals.GetFirstSheetVisible() + 1;
            }
            set
            {
                CheckWkConnected();
                CheckRange(value, 1, SheetCount, FlxParam.FirstSheetVisible);
                FWorkbook.Globals.SetFirstSheetVisible(value - 1);
            }
        }


        ///<inheritdoc />
        public override TExcelColor SheetTabColor
        {
            get
            {
                CheckConnected();
                return ActiveSheetObject.GetSheetTabColor();
            }
            set
            {
                CheckConnected();
                //This doesn't have to be a worksheet
                FWorkbook.Sheets[FActiveSheet - 1].SetSheetTabColor(value);
            }
        }


        ///<inheritdoc />
        public override void InsertAndCopySheets(int copyFrom, int insertBefore, int aSheetCount, ExcelFile sourceWorkbook)
        {
            CheckConnected();
            if (sourceWorkbook == this) sourceWorkbook = null;

            if (sourceWorkbook == null)
                CheckRange(copyFrom, 0, SheetCount, FlxParam.SheetFrom);
            else
                CheckRange(copyFrom, 0, sourceWorkbook.SheetCount, FlxParam.SheetFrom);

            CheckRange(insertBefore, 1, SheetCount + 1, FlxParam.SheetDest);
            CheckRange(aSheetCount, 1, FlxConsts.Max_Sheets + 1 - SheetCount, FlxParam.SheetCount);

            NeedsRecalc = true;
            if (sourceWorkbook == null)
                FWorkbook.InsertSheets(copyFrom - 1, insertBefore - 1, aSheetCount, GetWorkbook(sourceWorkbook));
            else
            {
                XlsFile xls = (sourceWorkbook as XlsFile);
                for (int i = 0; i < aSheetCount; i++)
                {
                    FWorkbook.InsertSheets(-1, insertBefore - 1 + i, 1, GetWorkbook(xls));
                    int SaveActiveSheet = ActiveSheet;
                    try
                    {
                        ActiveSheet = insertBefore + i;
                        if (xls != null)
                        {
                            TSheetInfo SheetInfo = new TSheetInfo(copyFrom - 1, copyFrom - 1, ActiveSheet - 1, xls.FWorkbook.Globals, FWorkbook.Globals, xls.FWorkbook.Sheets[copyFrom - 1], FWorkbook.Sheets[ActiveSheet - 1], SemiAbsoluteReferences);
                            FWorkbook.Globals.Names.InsertSheets(copyFrom - 1, insertBefore - 1 + i, 1, SheetInfo, false); //before copying the sheet, so formulas are already copied.
                            FWorkbook.Sheets[ActiveSheet - 1] = TWorkbook.CopySheetMisc(SheetInfo);
                            CopyHeaderImages(xls.FWorkbook, copyFrom - 1);
                        }

                        //We must do this after copy misc, so default widths (not column width, default widths) are already copied. If they weren't images would be changed.
                        //InsAndcopyRange will clear the cell range, but this does not currently clear the page breaks. If it did, page breaks should be copied after this.
                        InsertAndCopyRange(TXlsCellRange.FullRange(), 1, 1, 1, TFlxInsertMode.NoneDown, TRangeCopyMode.AllIncludingDontMoveAndSizeObjects, sourceWorkbook, copyFrom);
                    }
                    finally
                    {
                        ActiveSheet = SaveActiveSheet;
                    }
                }
            }
        }

        private static TWorkbook GetWorkbook(ExcelFile xls)
        {
            XlsFile x = xls as XlsFile;
            if (x == null) return null;
            return x.InternalWorkbook;
        }

        ///<inheritdoc />
        public override void InsertAndCopySheets(int[] copyFrom, int insertBefore, ExcelFile sourceWorkbook)
        {
            CheckConnected();
            if (sourceWorkbook == this || sourceWorkbook == null)
                FlxMessages.ThrowException(FlxErr.ErrWorkbookNull);

            if (copyFrom == null || copyFrom.Length <= 0) return;
            for (int i = 0; i < copyFrom.Length; i++)
                CheckRange(copyFrom[i], 1, sourceWorkbook.SheetCount, FlxParam.SheetFrom);
            CheckRange(insertBefore, 1, SheetCount + 1, FlxParam.SheetDest);
            CheckRange(copyFrom.Length + SheetCount, 1, FlxConsts.Max_Columns + 1, FlxParam.SheetCount);

            NeedsRecalc = true;
            XlsFile xls = (sourceWorkbook as XlsFile);

            int SaveActiveSheet = ActiveSheet;
            try
            {
                int SaveSourceActiveSheet = sourceWorkbook.ActiveSheet;
                try
                {
                    //First rename the sheets so they don't crash when copied.
                    for (int i = 0; i < copyFrom.Length; i++)
                    {
                        FWorkbook.InsertSheets(-1, insertBefore - 1 + i, 1, GetWorkbook(xls));
                        
                        ActiveSheet = insertBefore + i;
                        SheetName = "#&#&__#$$$$$_" + FlxConvert.ToString(i); //CF :-(,CultureInfo.InvariantCulture);
                    }

                    //Alter, copy and rename the sheet structure, so formulas will refer to the correct names.
                    for (int i = 0; i < copyFrom.Length; i++)
                    {
                        ActiveSheet = insertBefore + i;
                        sourceWorkbook.ActiveSheet = copyFrom[i];
                        SheetName = sourceWorkbook.SheetName;
                    }

                    //Now copy the actual sheet contents.
                    for (int i = 0; i < copyFrom.Length; i++)
                    {
                        ActiveSheet = insertBefore + i;
                        if (xls != null)
                        {
                            TSheetInfo SheetInfo = new TSheetInfo(copyFrom[i] - 1, copyFrom[i] - 1, ActiveSheet - 1, xls.FWorkbook.Globals, FWorkbook.Globals, xls.FWorkbook.Sheets[copyFrom[i] - 1],
                                FWorkbook.Sheets[ActiveSheet - 1], SemiAbsoluteReferences);
                            FWorkbook.Globals.Names.InsertSheets(copyFrom[i] - 1, insertBefore - 1 + i, 1, SheetInfo, false);
                            FWorkbook.Sheets[ActiveSheet - 1] = TWorkbook.CopySheetMisc(SheetInfo);

                        }
                        InsertAndCopyRange(TXlsCellRange.FullRange(), 1, 1, 1, TFlxInsertMode.NoneDown, TRangeCopyMode.AllIncludingDontMoveAndSizeObjects, sourceWorkbook, copyFrom[i]);
                    }
                }
                finally
                {
                    sourceWorkbook.ActiveSheet = SaveSourceActiveSheet;
                }
            }
            finally
            {
                ActiveSheet = SaveActiveSheet;
            }

        }

        ///<inheritdoc />
        public override void ClearSheet()
        {
            CheckConnected();
            NeedsRecalc = true;
            FWorkbook.Sheets[ActiveSheet - 1] = ActiveSheetObject.ClearValues();
            FActiveSheetObject = FWorkbook.Sheets[FActiveSheet - 1];
        }

        ///<inheritdoc />
        public override void DeleteSheet(int aSheetCount)
        {
            CheckConnected();
            if ((SheetCount <= aSheetCount) || (SheetCount < 0)) XlsMessages.ThrowException(XlsErr.ErrNoSheetVisible);
            NeedsRecalc = true;

            FWorkbook.DeleteSheets(FActiveSheet - 1, aSheetCount);
            if (FActiveSheet > SheetCount) ActiveSheet = SheetCount;  //Guarantee that ActiveSheet remains valid.
            else
            {
                FActiveSheetObject = FWorkbook.Sheets[FActiveSheet - 1]; //FActiveSheetObject changes when deleting the activesheet.
            }
        }

        ///<inheritdoc />
        public override bool ShowGridLines
        {
            get
            {
                CheckConnected();
                return ActiveSheetObject.Window.Window2.ShowGridLines;
            }
            set
            {
                CheckConnected();
                ActiveSheetObject.Window.Window2.ShowGridLines = value;
            }
        }

        ///<inheritdoc />
        public override bool ShowFormulaText
        {
            get
            {
                CheckConnected();
                return ActiveSheetObject.Window.Window2.ShowFormulaText;
            }
            set
            {
                CheckConnected();
                ActiveSheetObject.Window.Window2.ShowFormulaText = value;
            }
        }

        ///<inheritdoc />
        public override TExcelColor GridLinesColor
        {
            get
            {
                CheckConnected();
                return ActiveSheetObject.Window.Window2.GetGridLinesColor(this);
            }
            set
            {
                CheckConnected();
                ActiveSheetObject.Window.Window2.SetGridLinesColor(value);
            }
        }

        ///<inheritdoc />
        public override bool HideZeroValues
        {
            get
            {
                CheckConnected();
                return ActiveSheetObject.Window.Window2.HideZeroValues;
            }
            set
            {
                ActiveSheetObject.Window.Window2.HideZeroValues = value;
            }
        }


        ///<inheritdoc />
        public override TSheetType SheetType
        {
            get
            {
                CheckConnected();
                return FWorkbook.SheetType(FActiveSheet - 1);
            }
        }


        ///<inheritdoc />
        public override TSheetOptions SheetOptions
        {
            get
            {
                CheckConnected();
                return FWorkbook.Sheets[FActiveSheet - 1].Window.Window2.Options;
            }
            set
            {
                CheckConnected();
                FWorkbook.Sheets[FActiveSheet - 1].Window.Window2.Options = value;
            }
        }

        ///<inheritdoc />
        public override TSheetWindowOptions SheetWindowOptions
        {
            get
            {
                CheckWkConnected();
                return FWorkbook.Globals.WindowOptions;
            }
            set
            {
                CheckWkConnected();
                FWorkbook.Globals.WindowOptions = value;
            }
        }


        #endregion

        #region Page Breaks
        ///<inheritdoc />
        public override bool HasHPageBreak(int row)
        {
            CheckConnected();
            CheckRowAndCol(row, 1);
            return ActiveSheetObject.SheetGlobals.HPageBreaks.HasPageBreak(row); //Page Break arrays are 1-based
        }

        ///<inheritdoc />
        public override bool HasVPageBreak(int col)
        {
            CheckConnected();
            CheckRowAndCol(1, col);
            return ActiveSheetObject.SheetGlobals.VPageBreaks.HasPageBreak(col); //Page Break arrays are 1-based
        }

        ///<inheritdoc />
        internal override void InsertHPageBreak(int row, bool aGoesAfter)
        {
            CheckConnected();
            CheckRowAndCol(row, 1);
            ActiveSheetObject.SheetGlobals.HPageBreaks.AddBreak(row, aGoesAfter);
        }

        ///<inheritdoc />
        internal override void InsertVPageBreak(int col, bool aGoesAfter)
        {
            CheckConnected();
            CheckRowAndCol(1, col);
            ActiveSheetObject.SheetGlobals.VPageBreaks.AddBreak(col, aGoesAfter);
        }

        ///<inheritdoc />
        public override void DeleteHPageBreak(int row)
        {
            CheckConnected();
            CheckRowAndCol(row, 1);
            ActiveSheetObject.SheetGlobals.HPageBreaks.DeleteBreak(row);
        }

        ///<inheritdoc />
        public override void DeleteVPageBreak(int col)
        {
            CheckConnected();
            CheckRowAndCol(1, col);
            ActiveSheetObject.SheetGlobals.VPageBreaks.DeleteBreak(col);
        }

        ///<inheritdoc />
        public override void ClearPageBreaks()
        {
            CheckConnected();
            ActiveSheetObject.SheetGlobals.HPageBreaks.Clear();
            ActiveSheetObject.SheetGlobals.VPageBreaks.Clear();
        }

        ///<inheritdoc />
        public override void KeepRowsTogether(int row1, int row2, int level, bool replaceLowerLevels)
        {
            CheckConnected();
            CheckRowAndCol(row1, 1);
            CheckRowAndCol(row2, 1);
            CheckRange(level, 0, Int32.MaxValue, FlxParam.Level);

            FActiveSheetObject.Cells.KeepRowsTogeher(row1 - 1, row2 - 1, level, replaceLowerLevels);
        }

        ///<inheritdoc />
        public override void KeepColsTogether(int col1, int col2, int level, bool replaceLowerLevels)
        {
            CheckConnected();
            CheckRowAndCol(1, col1);
            CheckRowAndCol(1, col2);
            CheckRange(level, 0, Int32.MaxValue, FlxParam.Level);

            FActiveSheetObject.Columns.KeepColsTogether(col1 - 1, col2 - 1, level, replaceLowerLevels);
        }

        ///<inheritdoc />
        public override int GetKeepRowsTogether(int row)
        {
            CheckConnected();
            CheckRowAndCol(row, 1);

            return FActiveSheetObject.Cells.GetKeepRowsTogeher(row - 1);
        }

        ///<inheritdoc />
        public override int GetKeepColsTogether(int col)
        {
            CheckConnected();
            CheckRowAndCol(1, col);

            return FActiveSheetObject.Columns.GetKeepColsTogether(col - 1);
        }

        ///<inheritdoc />
        public override bool HasKeepRowsTogether()
        {
            CheckConnected();
            return FActiveSheetObject.Cells.HasKeepRowsTogether();
        }

        ///<inheritdoc />
        public override bool HasKeepColsTogether()
        {
            CheckConnected();
            return FActiveSheetObject.Columns.HasKeepColsTogether();
        }


        ///<inheritdoc />
        public override void AutoPageBreaks(int PercentOfUsedSheet, int PageScale)
        {
#if (!COMPACTFRAMEWORK && !MONOTOUCH && !SILVERLIGHT)
            CheckConnected();
            CheckRange(PercentOfUsedSheet, 0, 100, FlxParam.PercentOfUsedSheet);
            CheckRange(PageScale, 50, 100, FlxParam.PageScale);
            FlexCel.Render.FlexCelRender.AutoPageBreaks(this, PercentOfUsedSheet, new RectangleF(), new RectangleF(), PageScale);
#endif
        }

        ///<inheritdoc />
        public override void AutoPageBreaks(int PercentOfUsedSheet, RectangleF PageBounds)
        {
#if (!COMPACTFRAMEWORK && !MONOTOUCH && !SILVERLIGHT)
            CheckConnected();
            CheckRange(PercentOfUsedSheet, 0, 100, FlxParam.PercentOfUsedSheet);
            FlexCel.Render.FlexCelRender.AutoPageBreaks(this, PercentOfUsedSheet, PageBounds, PageBounds, 100);
#endif
        }
        #endregion

        #region Cell Value
        ///<inheritdoc />
        public override object GetCellValue(int row, int col, ref int XF)
        {
            CheckConnected();
            CheckRowAndCol(row, col);
            object value = null;
            ActiveSheetObject.Cells.CellList.GetValue(this, row - 1, col - 1, -1, ref value, ref XF);
            return value;
        }

        ///<inheritdoc />
        public override object GetCellValue(int sheet, int row, int col, ref int XF)
        {
            CheckConnected();
            CheckRowAndCol(row, col);
            CheckSheet(sheet);
            TSheet Ws = FWorkbook.Sheets[sheet - 1];
            object v = null;
            Ws.Cells.CellList.GetValue(this, row - 1, col - 1, -1, ref v, ref XF);
            return v;
        }

        ///<inheritdoc />
        public override object GetCellValueIndexed(int sheet, int row, int colIndex, ref int XF)
        {
            CheckConnected();
            CheckRowAndCol(row, 1);
            CheckSheet(sheet);
            object value = null;
            TSheet Ws = FWorkbook.Sheets[sheet - 1];
            Ws.Cells.CellList.GetValue(this, row - 1, -1, colIndex - 1, ref value, ref XF);
            return value;
        }

        internal override int PartialSheetCount()
        {
            return FWorkbook.Sheets.Count;
        }

        internal override object GetCellValueAndRecalc(int sheet, int row, int col, TCalcState CalcState, TCalcStack CalcStack)
        {
            TSheet Ws = FWorkbook.Sheets[sheet - 1];
            return Ws.Cells.CellList.GetValueAndRecalc(sheet - 1, row - 1, col - 1, this, CalcState, CalcStack);
        }

        ///<inheritdoc />
        public override void SetCellValue(int row, int col, object value, int XF)
        {
            CheckConnected();
            CheckRowAndCol(row, col);
            NeedsRecalc = true;

            ActiveSheetObject.Cells.CellList.SetValueWithArrays(row - 1, col - 1, value, XF);
        }

        ///<inheritdoc />
        public override void SetCellValue(int sheet, int row, int col, object value, int XF)
        {
            CheckConnected();
            CheckRowAndCol(row, col);
            CheckSheet(sheet);
            NeedsRecalc = true;

            TSheet Ws = FWorkbook.Sheets[sheet - 1];
            Ws.Cells.CellList.SetValueWithArrays(row - 1, col - 1, value, XF);
        }

        /// <summary>
        /// Will try to parse a date time string and find out if:
        ///   1)it has a date
        ///   2)it has a time
        ///   This method does not intend to be foolproof, jut to give a good aproximation.
        ///   There are too many issues with dates and internationalization.
        /// </summary>
        private static int CalcInternalFormatId(string value)
        {
            int i = 0;
            int dates = 0;
            string am = DateTimeFormatInfo.CurrentInfo.AMDesignator;
            if (am == null || am.Length == 0) am = FlxMessages.GetString(FlxMessage.TxtDefaultTimeAMString);
            am = am.ToUpper(CultureInfo.InvariantCulture);
            string pm = DateTimeFormatInfo.CurrentInfo.PMDesignator;
            if (pm == null || pm.Length == 0) pm = FlxMessages.GetString(FlxMessage.TxtDefaultTimePMString);
            pm = pm.ToUpper(CultureInfo.InvariantCulture);

            while ((i = value.IndexOf(DateTimeFormatInfo.CurrentInfo.DateSeparator, i) + 1) > 0 && dates < 3)
            {
                dates++;
            }
            int times = 0;
            while ((i = value.IndexOf(DateTimeFormatInfo.CurrentInfo.TimeSeparator, i) + 1) > 0 && times < 3)
            {
                times++;
            }

            if (dates >= 2)
            {
                if (times >= 1)  //excel does not convert to dd/mm/yyyy hh:mm:ss, always to dd/mm/yyyy hh:mm
                    return 0x16;  //dd/mm/yyyy hh:mm in international format
                else
                    return 0x0E;  //dd/mm/yyyy
            }
            else
                if (dates == 1)
                {
                    if (times >= 1)  //excel does not convert to mm/yyyy hh:mm, always to dd/mm/yyyy hh:mm
                        return 0x16;  //dd/mm/yyyy hh:mm in international format
                    else
                        return 0x11;  //mm/yyyy
                }
                else  //only time.
                {
                    if (times >= 2)
                    {
                        string valueUp = value.ToUpper(CultureInfo.InvariantCulture);
                        if (valueUp.IndexOf(am) > 0 || valueUp.IndexOf(pm) > 0)
                            return 0x13; //hh:mm:ss AM/PM
                        return 0x15; //hh:mm:ss
                    }
                    else
                        if (times == 1)
                        {
                            string valueUp = value.ToUpper(CultureInfo.InvariantCulture);
                            if (valueUp.IndexOf(am) > 0 || valueUp.IndexOf(pm) > 0)
                                return 0x12; //hh:mm AM/PM

                            return 0x14;  //hh:mm
                        }
                        else return 0x16; //Error, it is not a recognized format. return all.
                }
        }

        ///<inheritdoc />
        public override object ConvertString(TRichString value, ref int XF, string[] dateFormats)
        {
            if (value == null) return null;
            string sValue = value.ToString();
            if (sValue == null || sValue.Length == 0) return null;

            //See if it is a formula.
            if ((sValue != null) && (value.Length > 1) &&
                (
                (sValue.Substring(0, 1) == TFormulaMessages.TokenString(TFormulaToken.fmStartFormula)) ||
                (sValue.Substring(0, 1) == TFormulaMessages.TokenString(TFormulaToken.fmOpenArray))
                )
                )
            {
                TFormula Fmla = new TFormula();
                bool FormulaOk = false;
                try
                {
                    //if r1c1 it doesn't matter as we don't use the result from Ps.
                    TFormulaConvertTextToInternal Ps = new TFormulaConvertTextToInternal(this, ActiveSheet, false, sValue, false);
                    Ps.Parse();
                    Fmla.Text = sValue;
                    FormulaOk = !Ps.HasErrors;
                }
                catch (FlexCelException)
                {
                    FormulaOk = false;
                }
                if (FormulaOk)
                {
                    return Fmla;
                }
            }

            //try to convert to number
            double Result = 0;
            TNumberFormat nf;
            if (TCompactFramework.ConvertToNumber(sValue, CultureInfo.CurrentCulture, out Result, out nf))
            {
                if (nf.HasExp)
                {
                    XF = ChangeCellFormat(XF, TFormatRecordList.GetInternalFormat(11));
                }

                else
                if (nf.HasPercent)
                {
                    int PercentIndex = 9;
                    if (sValue.IndexOf(CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator) >= 0) PercentIndex = 10;
                    XF = ChangeCellFormat(XF, TFormatRecordList.GetInternalFormat(PercentIndex));
                }
                else
                    if (nf.HasCurrency)
                    {
                        //This has to be entered with the current currency locale, not what is in the file (in a NewFile we will see Euro)
                        NumberFormatInfo nfi = CultureInfo.CurrentCulture.NumberFormat;
                        string CurrFmt = string.Empty;
                        string n = "###,###";
                        if (sValue.IndexOf(nfi.CurrencyDecimalSeparator) >= 0)
                        {
                            n += "." + new string('0', nfi.CurrencyDecimalDigits);
                        }


                        string Sp = "\\ ";
                        switch (nfi.CurrencyPositivePattern)
                        {
                            case 1: CurrFmt = n + nfi.CurrencySymbol; break;
                            case 2: CurrFmt = nfi.CurrencySymbol + Sp + n; break;
                            case 3: CurrFmt = n + Sp + nfi.CurrencySymbol; break;
                            default:
                                CurrFmt = nfi.CurrencySymbol + n;
                                break;
                        }
                        
                        string NegSep = ";[Red]";
                        string NegSep2 = "_)" + NegSep;
                        string Op = "\\(";
                        string Cp = "\\)";

                        switch (nfi.CurrencyNegativePattern)
                        {
                            case 1: CurrFmt += NegSep + nfi.NegativeSign + nfi.CurrencySymbol + n; break;
                            case 2: CurrFmt += NegSep + nfi.CurrencySymbol + nfi.NegativeSign + n; break;
                            case 3: CurrFmt += NegSep + nfi.CurrencySymbol + n + nfi.NegativeSign; break;
                            case 4: CurrFmt += NegSep2 + Op + n + nfi.CurrencySymbol + Cp; break;
                            case 5: CurrFmt += NegSep + nfi.NegativeSign + n + nfi.CurrencySymbol; break;
                            case 6: CurrFmt += NegSep + n + nfi.NegativeSign + nfi.CurrencySymbol; break;
                            case 7: CurrFmt += NegSep + n + nfi.CurrencySymbol + nfi.NegativeSign; break;

                            case 8: CurrFmt += NegSep + nfi.NegativeSign + n + Sp + nfi.CurrencySymbol; break;
                            case 9: CurrFmt += NegSep + nfi.NegativeSign + nfi.CurrencySymbol + Sp + n; break;
                            case 10: CurrFmt += NegSep + n + Sp + nfi.CurrencySymbol + nfi.NegativeSign; break;
                            case 11: CurrFmt += NegSep + nfi.CurrencySymbol + Sp + n + nfi.NegativeSign; break;
                            case 12: CurrFmt += NegSep + nfi.CurrencySymbol + Sp + nfi.NegativeSign + n; break;
                            case 13: CurrFmt += NegSep + n + nfi.NegativeSign + Sp + nfi.CurrencySymbol; break;
                            case 14: CurrFmt += NegSep2 + Op + nfi.CurrencySymbol + Sp + n + Cp; break;
                            case 15: CurrFmt += NegSep2 + Op + n + Sp + nfi.CurrencySymbol + Cp; break;
 
                            default:
                                CurrFmt += NegSep2 + Op + nfi.CurrencySymbol + n + Cp; break;
                        }


                        XF = ChangeCellFormat(XF, CurrFmt);
                    }


                return Result;
            }

            //Try to convert to boolean.
            if (String.Equals(sValue, TFormulaMessages.TokenString(TFormulaToken.fmTrue), StringComparison.CurrentCultureIgnoreCase))
            {
                return true;
            }
            if (String.Equals(sValue, TFormulaMessages.TokenString(TFormulaToken.fmFalse), StringComparison.CurrentCultureIgnoreCase))
            {
                return false;
            }

            //try to convert to a date. This one is tricky, as we have to enter it as number, and format the cell as date.
            DateTime DateResult;
            if (TCompactFramework.ConvertDateToNumber(sValue, dateFormats, out DateResult))
            {
                string fmString = TFormatRecordList.GetInternalFormat(CalcInternalFormatId(sValue));
                XF = ChangeCellFormat(XF, fmString);
                return DateResult;
            }

            //Finally, if it wasn't anything else, enter the string.
            //If there are cr, convert the format to Wraptext.
            TRichString newValue = value.Replace("\r", String.Empty);
            if (newValue.ToString().IndexOf("\n") >= 0)
            {
                TFlxFormat fm = GetFormat(XF);
                if (!fm.WrapText)
                {
                    fm.WrapText = true;
                    XF = AddFormat(fm);
                }
            }
            return value;
        }

        private int ChangeCellFormat(int XF, string fmString)
        {
            TFlxFormat fm = GetFormat(XF);

            if (fm.Format != fmString)
            {
                fm.Format = fmString;
                XF = AddFormat(fm);
            }
            return XF;
        }

        ///<inheritdoc />
        public override void SetCellFromString(int row, int col, TRichString value, int XF, string[] dateFormats)
        {
            CheckConnected();
            CheckRowAndCol(row, col);

            SetCellValue(row, col, ConvertString(value, ref XF, dateFormats), XF);
        }

        ///<inheritdoc />
        public override TRichString GetStringFromCell(int row, int col, ref int XF, ref Color aColor)
        {
            CheckConnected();
            CheckRowAndCol(row, col);
            object value = GetCellValue(row, col, ref XF);
            if ((XF < 0) || (XF >= FWorkbook.Globals.CellXF.Count)) XF = 0;

            return TFlxNumberFormat.FormatValue(value, FWorkbook.Globals.Formats.Format(FWorkbook.Globals.CellXF[XF].FormatIndex), ref aColor, this);
        }

        ///<inheritdoc />
        public override void SetCellFromHtml(int row, int col, string htmlText, int XF)
        {
            CheckConnected();
            CheckRowAndCol(row, col);
            TFlxFormat fmt = null;
            if (XF >= 0)
            {
                fmt = GetFormat(XF);
            }
            else
            {
                fmt = GetCellVisibleFormatDef(row, col);
            }
            TRichString Rs = TRichString.FromHtml(htmlText, fmt, this);
            SetCellValue(row, col, Rs, XF);

        }

        ///<inheritdoc />
        public override string GetHtmlFromCell(int row, int col, THtmlVersion htmlVersion, THtmlStyle htmlStyle, Encoding encoding)
        {
            CheckConnected();
            CheckRowAndCol(row, col);

            int XF = -1;
            object value = GetCellValue(row, col, ref XF);
            if ((XF < 0) || (XF > FWorkbook.Globals.CellXF.Count)) XF = 0;

            TFormula Fmla = value as TFormula;
            if (Fmla != null)
            {
                if (ShowFormulaText) return THtmlEntities.EncodeAsHtml(Fmla.Text, htmlVersion, encoding);
                value = Fmla.Result;
            }

            TRichString Rs = value as TRichString;
            if (Rs != null)
            {
                TFlxFormat fmt = GetCellVisibleFormatDef(row, col);
                /*TFlxFormat CFFmt = ConditionallyModifyFormat(fmt, row, col);
                if (CFFmt != null) fmt = CFFmt;*/
                return Rs.ToHtml(this, fmt, htmlVersion, htmlStyle, encoding);
            }

            Color aColor = ColorUtil.Empty;
            TRichString Result = TFlxNumberFormat.FormatValue(value, FWorkbook.Globals.Formats.Format(FWorkbook.Globals.CellXF[XF].FormatIndex), ref aColor, this);
            if (Result == null) return String.Empty;

            string FinalString = THtmlEntities.EncodeAsHtml(Result.ToString(), htmlVersion, encoding);
            if (!aColor.Equals(ColorUtil.Empty))
            {
                return THtmlTagCreator.StartFontColor(aColor, htmlStyle) + FinalString + THtmlTagCreator.EndFontColor(htmlStyle);
            }
            return FinalString;
        }


        ///<inheritdoc />
        public override void CopyCell(ExcelFile sourceWorkbook, int sourceSheet, int destSheet, int sourceRow, int sourceCol, int destRow, int destCol, TRangeCopyMode copyMode)
        {
            CheckConnected();
            CheckRowAndCol(sourceRow, sourceCol);
            CheckRowAndCol(destRow, destCol);

            XlsFile sourceXls = sourceWorkbook as XlsFile;
            if (sourceXls == null) return;
            TSheet source = sourceXls.FWorkbook.Sheets[sourceSheet - 1];
            TSheet dest = FWorkbook.Sheets[destSheet - 1];

            dest.Cells.CellList.CopyCell(sourceSheet - 1, destSheet - 1,
                source.Cells.CellList, sourceRow - 1, sourceCol - 1, destRow - 1, destCol - 1, copyMode, source, dest);

        }



        #endregion

        #region Cell Format
        #region XF
        ///<inheritdoc />
        public override TFlxFormat GetFormat(int XF)
        {
            //CheckWkConnected();
            return FWorkbook.Globals.GetCellFormat(XF);
        }

        ///<inheritdoc />
        public override int FormatCount
        {
            get
            {
                CheckWkConnected();
                return FWorkbook.Globals.CellXF.Count;
            }
        }

        ///<inheritdoc />
        public override int AddFormat(TFlxFormat format)
        {
            CheckWkConnected();
            int Result = -1;
            if (format.IsStyle) FlxMessages.ThrowException(FlxErr.ErrCantAddStyleFormats);
            TXFRecord XF = new TXFRecord(format, false, FWorkbook.Globals, true);
            if (FWorkbook.Globals.CellXF.FindFormat(XF, ref Result))
                return Result;

            FWorkbook.Globals.CellXF.Add(XF);
            return FWorkbook.Globals.CellXF.Count - 1;
        }

        ///<inheritdoc />
        public override void SetFormat(int formatIndex, TFlxFormat aFormat)
        {
            CheckWkConnected();
            TXFRecordList XFList = aFormat.IsStyle ? FWorkbook.Globals.StyleXF : FWorkbook.Globals.CellXF;

            CheckRange(formatIndex, 0, XFList.Count, FlxParam.FormatIndex);
            TXFRecord OldXFRecord = XFList[formatIndex];
            if (OldXFRecord.IsStyle != aFormat.IsStyle) FlxMessages.ThrowException(FlxErr.ErrCantMixCellAndStyleFormats);
            TXFRecord XFRecord = new TXFRecord(aFormat, aFormat.IsStyle && formatIndex == 0, FWorkbook.Globals, true);
            XFList[formatIndex] = XFRecord;
            if (aFormat.IsStyle) FWorkbook.Globals.CellXF.UpdateChangedStyleInCellXF(formatIndex, XFRecord, false);
        }

        #endregion

        #region Font
        ///<inheritdoc />
        public override TFlxFont GetFont(int fontIndex)
        {
            return FWorkbook.Globals.Fonts.GetFont(fontIndex);

        }

        ///<inheritdoc />
        public override TFlxFont GetDefaultFont
        {
            get
            {
                int FontIndex = FWorkbook.Globals.CellXF[0].FontIndex;
                return FWorkbook.Globals.Fonts.GetFont(FontIndex);
            }
        }

        ///<inheritdoc />
        public override TFlxFont GetDefaultFontNormalStyle
        {
            get
            {
                CheckWkConnected();
                int FontIndex = FWorkbook.Globals.StyleXF[0].FontIndex;
                return FWorkbook.Globals.Fonts.GetFont(FontIndex);
            }
        }


        ///<inheritdoc />
        public override void SetFont(int fontIndex, TFlxFont aFont)
        {
            CheckWkConnected();
            FWorkbook.Globals.Fonts.SetFont(fontIndex, aFont);
        }


        ///<inheritdoc />
        public override int FontCount
        {
            get
            {
                return FWorkbook.Globals.Fonts.Count;
            }
        }

        ///<inheritdoc />
        public override int AddFont(TFlxFont font)
        {
            return FWorkbook.Globals.Fonts.AddFont(font);
        }


        #endregion

        ///<inheritdoc />
        public override void SetCellFormat(int row, int col, int XF)
        {
            CheckConnected();
            CheckRowAndCol(row, col);
            ActiveSheetObject.Cells.CellList.SetFormat(row - 1, col - 1, XF);
        }

        ///<inheritdoc />
        public override void SetCellFormat(int row1, int col1, int row2, int col2, int XF)
        {
            CheckConnected();
            CheckRowAndCol(row1, col1);
            CheckRowAndCol(row2, col2);
            for (int r = row1 - 1; r < row2; r++)
                for (int c = col1 - 1; c < col2; c++)
                {
                    ActiveSheetObject.Cells.CellList.SetFormat(r, c, XF);
                }
        }

        ///<inheritdoc />
        public override void SetCellFormat(int row1, int col1, int row2, int col2, TFlxFormat newFormat, TFlxApplyFormat applyNewFormat, bool exteriorBorders)
        {
            CheckConnected();
            CheckRowAndCol(row1, col1);
            CheckRowAndCol(row2, col2);

            int CacheXF = -2;
            int CacheNewXF = -1;
            bool CacheApply = false;

            TFlxApplyFormat applyNew = exteriorBorders ? (TFlxApplyFormat)applyNewFormat.Clone() : applyNewFormat; //If doing exterior borders, we will need to modify this thing.
            if (exteriorBorders && applyNewFormat.HasOnlyBorders)  //Optimization to draw only borders and not loop over all cells.
            {
                applyNew.Borders.SetAllMembers(false);
                for (int c = col1 - 1; c < col2; c++)
                {
                    applyNew.Borders.Bottom = false;
                    applyNew.Borders.Top = applyNewFormat.Borders.Top;
                    CacheXF = -2; FormatOneCell(newFormat, applyNew, ref CacheXF, ref CacheNewXF, ref CacheApply, row1 - 1, c);
                    applyNew.Borders.Top = false;
                    applyNew.Borders.Bottom = applyNewFormat.Borders.Bottom;
                    CacheXF = -2; FormatOneCell(newFormat, applyNew, ref CacheXF, ref CacheNewXF, ref CacheApply, row2 - 1, c);
                }

                applyNew.Borders.SetAllMembers(false);
                for (int r = row1 - 1; r < row2; r++)
                {
                    applyNew.Borders.Right = false;
                    applyNew.Borders.Left = applyNewFormat.Borders.Left;
                    CacheXF = -2; FormatOneCell(newFormat, applyNew, ref CacheXF, ref CacheNewXF, ref CacheApply, r, col1 - 1);
                    applyNew.Borders.Left = false;
                    applyNew.Borders.Right = applyNewFormat.Borders.Right;
                    CacheXF = -2; FormatOneCell(newFormat, applyNew, ref CacheXF, ref CacheNewXF, ref CacheApply, r, col2 - 1);
                }
            }
            else
            {
                if (exteriorBorders) applyNew.Borders.SetAllMembers(false);
                for (int r = row1 - 1; r < row2; r++)
                {
                    for (int c = col1 - 1; c < col2; c++)
                    {
                        if (exteriorBorders)
                        {
                            applyNew.Borders.Top = (r == row1 - 1);
                            applyNew.Borders.Bottom = (r == row2 - 1);
                            applyNew.Borders.Left = (c == col1 - 1);
                            applyNew.Borders.Right = (c == col2 - 1);
                            CacheXF = -2;
                        }
                        FormatOneCell(newFormat, applyNew, ref CacheXF, ref CacheNewXF, ref CacheApply, r, c);
                    }
                }
            }
        }

        private void FormatOneCell(TFlxFormat newFormat, TFlxApplyFormat applyNewFormat, ref int CacheXF, ref int CacheNewXF, ref bool CacheApply, int r, int c)
        {
            int XF = ActiveSheetObject.Cells.CellList.GetXFFormat(r, c, -1);
            if (CacheXF < -1 || XF != CacheXF)
            {
                CacheXF = XF;
                TFlxFormat Fmt = GetFormat(XF);
                CacheApply = applyNewFormat.Apply(Fmt, newFormat);
                if (CacheApply)
                {
                    CacheNewXF = AddFormat(Fmt);
                }
            }
            if (CacheApply) ActiveSheetObject.Cells.CellList.SetFormat(r, c, CacheNewXF);

        }

        ///<inheritdoc />
        public override int GetCellFormat(int row, int col)
        {
            CheckConnected();
            CheckRowAndCol(row, col);
            return ActiveSheetObject.Cells.CellList.GetXFFormat(row - 1, col - 1, -1);
        }

        ///<inheritdoc />
        public override int GetCellFormat(int sheet, int row, int col)
        {
            CheckConnected();
            CheckRowAndCol(row, col);
            CheckSheet(sheet);
            TSheet Ws = FWorkbook.Sheets[sheet - 1];
            return Ws.Cells.CellList.GetXFFormat(row - 1, col - 1, -1);
        }

        ///<inheritdoc />
        public override TFlxFormat ConditionallyModifyFormat(TFlxFormat format, int row, int col)
        {
            CheckConnected();
            CheckRowAndCol(row, col);
            return ActiveSheetObject.ConditionallyModifyFormat(this, FActiveSheet, format, row - 1, col - 1);
        }

        /*///<inheritdoc />
        public override int ConditionalFormatCount
        {
            get
            {
                return 0;
                CheckConnected();
                return ActiveSheetObject.ConditionalFormats.Count;
            }
        }

        ///<inheritdoc />
        public override TConditionalFormatRule[] GetConditionalFormat(int index, out TXlsCellRange range)
        {
            CheckConnected();
            CheckRange(index, 1, ConditionalFormatCount, FlxParam.ConditionalFormatIndex);

            TCondFmt Fmt = (TCondFmt)ActiveSheetObject.ConditionalFormats[index - 1];
            range = Fmt.CellRange1Based();
            return Fmt.GetCFDef(FWorkbook.Globals.Names, ActiveSheetObject.Cells.CellList);
        }

        ///<inheritdoc />
        public override void SetConditionalFormat(int firstRow, int firstCol, int lastRow, int lastCol, TConditionalFormatRule[] conditionalFormat)
        {
            CheckConnected();
            CheckRowAndCol(firstRow, firstCol);
            CheckRowAndCol(lastRow, lastCol);
            ActiveSheetObject.SetConditionalFormat(firstRow-1, firstCol-1, lastRow-1, lastCol-1, conditionalFormat);
        }
*/



        #endregion

        #region Styles
        ///<inheritdoc />
        public override int StyleCount
        {
            get
            {
                return FWorkbook.Globals.Styles.Count;
            }
        }

        ///<inheritdoc />
        public override string GetStyleName(int index)
        {
            return FWorkbook.Globals.Styles.GetStyleName(index - 1);
        }

        ///<inheritdoc />
        public override TFlxFormat GetStyle(int index)
        {
            return FWorkbook.Globals.GetStyleFormat(FWorkbook.Globals.Styles.GetStyle(index - 1));
        }

        ///<inheritdoc />
        public override TFlxFormat GetStyle(string name, bool convertToCellStyle)
        {
            int fmt = FWorkbook.Globals.Styles.GetStyle(name);
            if (fmt < 0)
            {
                int Level;
                int BuiltInId = TBuiltInStyles.GetIdAndLevel(name, out Level);
                return ConvertFormat(TBuiltInStyles.GetDefaultStyle(BuiltInId, Level), name, convertToCellStyle);
            }
            
            return ConvertFormat(FWorkbook.Globals.GetStyleFormat(fmt), name, convertToCellStyle);
        }

        private static TFlxFormat ConvertFormat(TFlxFormat aFormat, string name, bool convertToCellStyle)
        {
            if (convertToCellStyle)
            {
                aFormat.IsStyle = !convertToCellStyle;
                aFormat.ParentStyle = name;
            }
            return aFormat;
        }

        ///<inheritdoc />
        public override void RenameStyle(string oldName, string newName)
        {
            FWorkbook.Globals.Styles.RenameStyle(oldName, newName);
        }

        ///<inheritdoc />
        public override void SetStyle(string name, TFlxFormat fmt)
        {
            FWorkbook.Globals.Styles.SetStyle(name, FWorkbook.Globals.AddStyleFormat(fmt, name));
        }

        ///<inheritdoc />
        public override void DeleteStyle(string name)
        {
            FWorkbook.Globals.Styles.DeleteStyle(name, FWorkbook.Globals.CellXF);
        }

        ///<inheritdoc />
        public override string GetBuiltInStyleName(TBuiltInStyle style, int level)
        {
            return TBuiltInStyles.GetName((byte)style, (level - 1));
        }

        ///<inheritdoc />
        public override bool TryGetBuiltInStyleType(string styleName, out TBuiltInStyle style, out int level)
        {
            style = TBuiltInStyle.Normal;
            int index = TBuiltInStyles.GetIdAndLevel(styleName, out level);
            if (index < 0) return false;

            style = (TBuiltInStyle)index;
            if (style == TBuiltInStyle.ColLevel || style == TBuiltInStyle.RowLevel) level++; //only for outline styles.
            return true;
        }


        #endregion

        #region Merged Cells
        ///<inheritdoc />
        public override TXlsCellRange CellMergedBounds(int row, int col)
        {
            CheckConnected();
            CheckRowAndCol(row, col);
            TXlsCellRange Result = ActiveSheetObject.CellMergedBounds(row - 1, col - 1);
            Result.Left++;
            Result.Top++;
            Result.Right++;
            Result.Bottom++;
            return Result;
        }

        ///<inheritdoc />
        public override void MergeCells(int firstRow, int firstCol, int lastRow, int lastCol)
        {
            CheckConnected();
            CheckRowAndCol(firstRow, firstCol);
            CheckRowAndCol(lastRow, lastCol);
            ActiveSheetObject.MergeCells(firstRow - 1, firstCol - 1, lastRow - 1, lastCol - 1);
        }

        ///<inheritdoc />
        public override void UnMergeCells(int firstRow, int firstCol, int lastRow, int lastCol)
        {
            CheckConnected();
            CheckRowAndCol(firstRow, firstCol);
            CheckRowAndCol(lastRow, lastCol);
            ActiveSheetObject.UnMergeCells(firstRow - 1, firstCol - 1, lastRow - 1, lastCol - 1);
        }

        ///<inheritdoc />
        public override int CellMergedListCount
        {
            get
            {
                CheckConnected();
                return ActiveSheetObject.CellMergedListCount();
            }
        }

        ///<inheritdoc />
        public override TXlsCellRange CellMergedList(int index)
        {
            CheckConnected();
            CheckRange(index, 1, CellMergedListCount, FlxParam.CellMergedIndex);
            return ActiveSheetObject.CellMergedList(index - 1).Inc();
        }
        #endregion

        #region Rows and Cols
        ///<inheritdoc />
        public override int RowCount
        {
            get
            {
                CheckConnected();
                return ActiveSheetObject.Cells.CellList.Count;
            }
        }

        ///<inheritdoc />
        public override int GetRowCount(int sheet)
        {
            TSheet ws = FWorkbook.Sheets[sheet - 1];
            return ws.Cells.CellList.Count;
        }


        ///<inheritdoc />
        public override int ColCount
        {
            get
            {
                CheckConnected();
                return Math.Max(ActiveSheetObject.Cells.ColCount, ActiveSheetObject.Columns.ColCount);
            }
        }

        ///<inheritdoc />
        public override int GetColCount(int sheet, bool includeFormattedColumns)
        {
            TSheet ws = FWorkbook.Sheets[sheet - 1];
            if (includeFormattedColumns)
            {
                return Math.Max(ws.Cells.ColCount, ws.Columns.ColCount);
            }
            else
            {
                return ws.Cells.ColCount;
            }
        }

        ///<inheritdoc />
        public override bool IsEmptyRow(int row)
        {
            CheckConnected();
            CheckRowAndCol(row, 1);
            return !(ActiveSheetObject.Cells.CellList.HasRow(row - 1));
        }

        ///<inheritdoc />
        public override bool IsNotFormattedCol(int col)
        {
            CheckConnected();
            CheckRowAndCol(1, col);
            return !ActiveSheetObject.HasCol(col - 1);
        }

        ///<inheritdoc />
        public override int GetRowFormat(int row)
        {
            CheckConnected();
            CheckRowAndCol(row, 1);
            return ActiveSheetObject.GetRowFormat(row - 1);
        }

        ///<inheritdoc />
        public override int GetRowFormat(int sheet, int row)
        {
            CheckConnected();
            CheckRowAndCol(row, 1);
            CheckSheet(sheet);
            TSheet Ws = FWorkbook.Sheets[sheet - 1];
            return Ws.GetRowFormat(row - 1);
        }


        ///<inheritdoc />
        public override void SetRowFormat(int row, int XF, bool resetRow)
        {
            CheckConnected();
            CheckRowAndCol(row, 1);
            ActiveSheetObject.SetRowFormat(row - 1, XF, resetRow);
        }

        ///<inheritdoc />
        public override void SetRowFormat(int row, TFlxFormat newFormat, TFlxApplyFormat applyNewFormat, bool resetRow)
        {
            CheckConnected();
            CheckRowAndCol(row, 1);

            TFlxFormat existingFmt = GetFormat(GetRowFormat(row));
            applyNewFormat.Apply(existingFmt, newFormat);
            ActiveSheetObject.SetRowFormat(row - 1, AddFormat(existingFmt), false);

            if (resetRow)
            {
                int CacheXF = -2;
                int CacheNewXF = -1;
                bool CacheApply = false;

                if ((row - 1 >= 0) && (row - 1 < ActiveSheetObject.Cells.CellList.Count))
                {
                    for (int i = 0; i < ActiveSheetObject.Cells.CellList[row - 1].Count; i++)
                    {
                        FormatOneCell(newFormat, applyNewFormat, ref CacheXF, ref CacheNewXF, ref CacheApply, row - 1, ActiveSheetObject.Cells.CellList[row - 1][i].Col);
                    }
                }
            }
        }


        ///<inheritdoc />
        public override int GetColFormat(int col)
        {
            CheckConnected();
            CheckRowAndCol(1, col);
            return ActiveSheetObject.GetColFormat(col - 1);
        }

        ///<inheritdoc />
        public override int GetColFormat(int sheet, int col)
        {
            CheckConnected();
            CheckRowAndCol(1, col);
            CheckSheet(sheet);
            TSheet Ws = FWorkbook.Sheets[sheet - 1];
            return Ws.GetColFormat(col - 1);
        }


        ///<inheritdoc />
        public override void SetColFormat(int col, int XF, bool resetColumn)
        {
            CheckConnected();
            CheckRowAndCol(1, col);
            ActiveSheetObject.SetColFormat(col - 1, XF, resetColumn);
        }

        ///<inheritdoc />
        public override void SetColFormat(int col, TFlxFormat newFormat, TFlxApplyFormat applyNewFormat, bool resetColumn)
        {
            CheckConnected();
            CheckRowAndCol(1, col);

            TFlxFormat existingFmt = GetFormat(GetColFormat(col));
            applyNewFormat.Apply(existingFmt, newFormat);
            ActiveSheetObject.SetColFormat(col - 1, AddFormat(existingFmt), false);

            if (resetColumn)
            {
                int CacheXF = -2;
                int CacheNewXF = -1;
                bool CacheApply = false;

                int Index = -1;
                for (int i = 0; i < ActiveSheetObject.Cells.CellList.Count; i++)
                    if (ActiveSheetObject.Cells.CellList[i].Find(col - 1, ref Index))
                    {
                        FormatOneCell(newFormat, applyNewFormat, ref CacheXF, ref CacheNewXF, ref CacheApply, i, col - 1);
                    }
            }
        }


        ///<inheritdoc />
        public override int GetRowOptions(int row)
        {
            CheckConnected();
            CheckRowAndCol(row, 1);
            return ActiveSheetObject.GetRowOptions(row - 1);
        }

        ///<inheritdoc />
        public override void SetRowOptions(int row, int options)
        {
            CheckConnected();
            CheckRowAndCol(row, 1);
            ActiveSheetObject.SetRowOptions(row - 1, options);
        }

        ///<inheritdoc />
        public override int GetColOptions(int col)
        {
            CheckConnected();
            CheckRowAndCol(1, col);
            return ActiveSheetObject.GetColOptions(col - 1);
        }

        ///<inheritdoc />
        public override void SetColOptions(int col, int options)
        {
            CheckConnected();
            CheckRowAndCol(1, col);
            ActiveSheetObject.SetColOptions(col - 1, options);
        }

        ///<inheritdoc />
        public override int GetRowHeight(int row)
        {
            CheckConnected();
            CheckRowAndCol(row, 1);
            return ActiveSheetObject.GetRowHeight(row - 1, false);
        }

        ///<inheritdoc />
        public override int GetRowHeight(int sheet, int row, bool HiddenIsZero)
        {
            CheckConnected();
            CheckRowAndCol(row, 1);
            CheckSheet(sheet);
            TSheet Ws = FWorkbook.Sheets[sheet - 1];
            return Ws.GetRowHeight(row - 1, HiddenIsZero);
        }


        ///<inheritdoc />
        public override void SetRowHeight(int row, int height)
        {
            CheckConnected();
            CheckRowAndCol(row, 1);
            ActiveSheetObject.SetRowHeight(row - 1, height);
            RestoreObjectSizes();
        }

        ///<inheritdoc />
        public override int GetColWidth(int col)
        {
            CheckConnected();
            CheckRowAndCol(1, col);
            return ActiveSheetObject.GetColWidth(col - 1, false);
        }

        ///<inheritdoc />
        public override int GetColWidth(int col, bool HiddenIsZero)
        {
            CheckConnected();
            CheckRowAndCol(1, col);
            return ActiveSheetObject.GetColWidth(col - 1, HiddenIsZero);
        }

        ///<inheritdoc />
        public override int GetColWidth(int sheet, int col, bool HiddenIsZero)
        {
            CheckConnected();
            CheckRowAndCol(1, col);
            CheckSheet(sheet);
            TSheet Ws = FWorkbook.Sheets[sheet - 1];
            return Ws.GetColWidth(col - 1, HiddenIsZero);
        }

        ///<inheritdoc />
        public override void SetColWidth(int col, int width)
        {
            CheckConnected();
            CheckRowAndCol(1, col);
            ActiveSheetObject.SetColWidth(col - 1, width);
            RestoreObjectSizes();
        }

        internal override void SetColWidthInternal(int col, int width)
        {
            ActiveSheetObject.SetColWidth(col - 1, width);            
        } 

        ///<inheritdoc />
        public override int DefaultRowHeight
        {
            get
            {
                CheckConnected();
                return ActiveSheetObject.DefRowHeight;
            }
            set
            {
                CheckConnected();
                ActiveSheetObject.DefRowHeight = value;
            }
        }

        ///<inheritdoc />
        public override bool DefaultRowHeightAutomatic
        {
            get
            {
                return (ActiveSheetObject.DefRowFlags & 0x01) == 0;
            }
            set
            {
                if (value) ActiveSheetObject.DefRowFlags &= ~0x01; else ActiveSheetObject.DefRowFlags |= 0x01;
            }
        }

        ///<inheritdoc />
        public override int DefaultColWidth
        {
            get
            {
                CheckConnected();
                return ActiveSheetObject.Columns.DefColWidth;
            }
            set
            {
                CheckConnected();
                ActiveSheetObject.Columns.DefColWidth = value;
            }
        }

        ///<inheritdoc />
        public override bool GetRowHidden(int row)
        {
            CheckConnected();
            CheckRowAndCol(row, 1);
            return ActiveSheetObject.GetRowHidden(row - 1);
        }

        ///<inheritdoc />
        public override bool GetRowHidden(int sheet, int row)
        {
            CheckConnected();
            CheckRowAndCol(row, 1);
            CheckSheet(sheet);

            TSheet Ws = FWorkbook.Sheets[sheet - 1];
            return Ws.GetRowHidden(row - 1);
        }

        ///<inheritdoc />
        public override void SetRowHidden(int row, bool hide)
        {
            CheckConnected();
            CheckRowAndCol(row, 1);
            ActiveSheetObject.SetRowHidden(row - 1, hide);
            RestoreObjectSizes();
        }

        ///<inheritdoc />
        public override bool GetColHidden(int col)
        {
            CheckConnected();
            CheckRowAndCol(1, col);
            return ActiveSheetObject.GetColHidden(col - 1);
        }

        ///<inheritdoc />
        public override void SetColHidden(int col, bool hide)
        {
            CheckConnected();
            CheckRowAndCol(1, col);
            ActiveSheetObject.SetColHidden(col - 1, hide);
            RestoreObjectSizes();
        }

        ///<inheritdoc />
        public override bool GetAutoRowHeight(int row)
        {
            CheckConnected();
            CheckRowAndCol(row, 1);
            return ActiveSheetObject.Cells.CellList.IsAutoRowHeight(row - 1);
        }

        ///<inheritdoc />
        public override void SetAutoRowHeight(int row, bool autoRowHeight)
        {
            CheckConnected();
            CheckRowAndCol(row, 1);
            ActiveSheetObject.Cells.CellList.AutoRowHeight(row - 1, autoRowHeight);
        }

        ///<inheritdoc />
        public override void AutofitRow(int row1, int row2, bool autofitNotAutofittingRows, bool keepHeightAutomatic, real adjustment, int adjustmentFixed, int minHeight, int maxHeight, TAutofitMerged autofitMerged)
        {
#if (!COMPACTFRAMEWORK && !MONOTOUCH && !SILVERLIGHT)
            CheckConnected();
            CheckRowAndCol(row1, 1);
            CheckRowAndCol(row2, 1);
            ActiveSheetObject.Cells.CellList.RecalcRowHeights(this, row1 - 1, row2 - 1, autofitNotAutofittingRows, keepHeightAutomatic, false, adjustment, adjustmentFixed, minHeight, maxHeight, autofitMerged);
            RestoreObjectSizes();
#endif
        }

        ///<inheritdoc />
        public override void AutofitCol(int col1, int col2, bool ignoreStrings, real adjustment, int adjustmentFixed, int minWidth, int maxWidth, TAutofitMerged autofitMerged)
        {
#if (!COMPACTFRAMEWORK && !MONOTOUCH && !SILVERLIGHT)
            CheckConnected();
            CheckRowAndCol(1, col1);
            CheckRowAndCol(1, col2);
            ActiveSheetObject.Cells.CellList.RecalcColWidths(this, col1 - 1, col2 - 1, ignoreStrings, false, adjustment, adjustmentFixed, minWidth, maxWidth, autofitMerged);
            RestoreObjectSizes();
#endif
        }

        ///<inheritdoc />
        public override void AutofitRowsOnWorkbook(bool autofitNotAutofittingRows, bool keepSizesAutomatic, real adjustment, int adjustmentFixed, int minHeight, int maxHeight, TAutofitMerged autofitMerged)
        {
#if (!COMPACTFRAMEWORK && !MONOTOUCH)
            CheckConnected();
            FWorkbook.RecalcRowHeights(this, autofitNotAutofittingRows, keepSizesAutomatic, adjustment, adjustmentFixed, minHeight, maxHeight, autofitMerged);
            //RestoreObjectSizes();  NOT HERE, or it will only restore objects on the current sheet. Must be done inside FWorkbook.RecalcRowHeights.
#endif
        }

        ///<inheritdoc />
        public override void MarkRowForAutofit(int row, bool autofit, real adjustment, int adjustmentFixed, int minHeight, int maxHeight, bool isMerged)
        {
            CheckConnected();
            CheckRowAndCol(row, 1);
            ActiveSheetObject.Cells.MarkRowForAutofit(row - 1, autofit, adjustment, adjustmentFixed, minHeight, maxHeight, isMerged);
        }


        ///<inheritdoc />
        public override void MarkColForAutofit(int col, bool autofit, real adjustment, int adjustmentFixed, int minWidth, int maxWidth, bool isMerged)
        {
            CheckConnected();
            CheckRowAndCol(1, col);
            ActiveSheetObject.Columns.MarkColForAutofit(col - 1, autofit, adjustment, adjustmentFixed, minWidth, maxWidth, isMerged);
        }

        ///<inheritdoc />
        public override void AutofitMarkedRowsAndCols(bool keepSizesAutomatic, bool ignoreStringsOnColumnFit, real adjustment, int adjustmentFixed, int minHeight, int maxHeight, int minWidth, int maxWidth, TAutofitMerged autofitMerged)
        {
#if (!COMPACTFRAMEWORK && !MONOTOUCH && !SILVERLIGHT)
            CheckConnected();
            ActiveSheetObject.Cells.CellList.RecalcRowHeights(this, 0, RowCount - 1, true, keepSizesAutomatic, true, adjustment, adjustmentFixed, minHeight, maxHeight, autofitMerged);
            ActiveSheetObject.Cells.CellList.RecalcColWidths(this, 0, ColCount - 1, ignoreStringsOnColumnFit, true, adjustment, adjustmentFixed, minWidth, maxWidth, autofitMerged);
            RestoreObjectSizes();
#endif
        }


        #region Cols by index.
        ///<inheritdoc />
        public override int ColCountInRow(int row)
        {
            CheckConnected();
            CheckRowAndCol(row, 1);
            if (row - 1 < 0 || row - 1 >= ActiveSheetObject.Cells.CellList.Count) return 0;
            return ActiveSheetObject.Cells.CellList[row - 1].Count;
        }

        ///<inheritdoc />
        public override int ColCountInRow(int sheet, int row)
        {
            CheckRowAndCol(row, 1);
            CheckSheet(sheet);
            TSheet ws = FWorkbook.Sheets[sheet - 1];
            if (row - 1 < 0 || row - 1 >= ws.Cells.CellList.Count) return 0;
            return ws.Cells.CellList[row - 1].Count;
        }

        ///<inheritdoc />
        public override int ColFromIndex(int row, int colIndex)
        {
            CheckConnected();
            CheckRowAndCol(row, 1);
            if (IsEmptyRow(row)) return 0;
            if (colIndex <= 0 || colIndex > ActiveSheetObject.Cells.CellList[row - 1].Count) return 0;
            return ActiveSheetObject.Cells.CellList[row - 1][colIndex - 1].Col + 1;
        }

        ///<inheritdoc />
        public override int ColFromIndex(int sheet, int row, int colIndex)
        {
            CheckRowAndCol(row, 1);
            CheckSheet(sheet);
            TSheet ws = FWorkbook.Sheets[sheet - 1];
            if (!ws.Cells.CellList.HasRow(row - 1)) return 0;
            if (colIndex <= 0 || colIndex > ws.Cells.CellList[row - 1].Count) return 0;
            return ws.Cells.CellList[row - 1][colIndex - 1].Col + 1;
        }

        ///<inheritdoc />
        public override int ColToIndex(int row, int col)
        {
            CheckConnected();
            CheckRowAndCol(row, col);
            if (IsEmptyRow(row) || row - 1 >= ActiveSheetObject.Cells.CellList.Count) return 0;
            int Result = 0;
            ActiveSheetObject.Cells.CellList[row - 1].Find(col - 1, ref Result);
            return Result + 1;
        }

        ///<inheritdoc />
        public override int ColToIndex(int sheet, int row, int col)
        {
            CheckRowAndCol(row, col);
            CheckSheet(sheet);
            TSheet ws = FWorkbook.Sheets[sheet - 1];
            if (!ws.Cells.CellList.HasRow(row - 1)) return 0;
            if (row - 1 >= ws.Cells.CellList.Count) return 0;
            int Result = 0;
            ws.Cells.CellList[row - 1].Find(col - 1, ref Result);
            return Result + 1;
        }

        #endregion
        #endregion

        #region Indexed Color
        ///<inheritdoc />
        public override Color GetColorPalette(int index)
        {
            return FWorkbook.Globals.GetRgbColorPalette(index - 1);
        }

        ///<inheritdoc />
        public override void SetColorPalette(int index, Color value)
        {
            CheckConnected();
            FWorkbook.Globals.SetColorPalette(index - 1, value);
        }

        ///<inheritdoc />
        public override bool[] GetUsedPaletteColors
        {
            get
            {
                return FWorkbook.Globals.CellXF.GetUsedColors(ColorPaletteCount + 1, FWorkbook.Globals.Fonts, FWorkbook.Globals.Borders, FWorkbook.Globals.Patterns);
            }
        }

        ///<inheritdoc />
        public override int NearestColorIndex(Color value)
        {
            return NearestColorIndex(value, null);
        }

        ///<inheritdoc />
        public override int NearestColorIndex(Color value, bool[] UsedColors)
        {
            int Result = 1;
            double MinDist = -1;
            TLabColor refColor = value;
            for (int i = 1; i <= ColorPaletteCount; i++)
            {
                TLabColor p = FWorkbook.Globals.GetLabColorPalette(i - 1);

                double Distance = refColor.DistanceSquared(p);

                if (MinDist == -1 || Distance < MinDist)
                {
                    MinDist = Distance;
                    Result = i;
                    if (Distance == 0)
                    {
                        if (UsedColors != null) UsedColors[Result] = true;
                        return Result; //exact match...
                    }
                }
            }

            if (UsedColors == null) return Result;

            //Find the nearest color between the ones that are not in use.
            UsedColors[0] = true; //not really used
            UsedColors[1] = true; //pure black
            UsedColors[2] = true; //pure white

            int Result2 = -1;
            MinDist = -1;
            for (int i = 1; i <= ColorPaletteCount; i++)
            {
                if (i >= UsedColors.Length || UsedColors[i]) continue;
                TLabColor p = FWorkbook.Globals.GetLabColorPalette(i - 1);

                double Distance = refColor.DistanceSquared(p);
                if (MinDist == -1 || Distance < MinDist)
                {
                    MinDist = Distance;
                    Result2 = i;
                    if (Distance == 0)
                    {
                        if (UsedColors != null) UsedColors[Result2] = true;
                        return Result2; //exact match...
                    }
                }
            }

            if (Result2 < 0 || Result2 >= UsedColors.Length)
            {
                if (UsedColors != null) UsedColors[Result] = true;
                return Result;  //Not available colors to modify
            }
            SetColorPalette(Result2, value);
            UsedColors[Result2] = true;
            return Result2;
        }


        ///<inheritdoc />
        public override bool PaletteContainsColor(TExcelColor value)
        {
            return FWorkbook.Globals.PaletteContainsColor(value.ToColor(this));
        }

        ///<inheritdoc />
        public override void OptimizeColorPalette()
        {
            FWorkbook.Globals.CellXF.OptimizeColorPalette(ColorPaletteCount, FWorkbook.Globals, this);
        }


        #endregion

        #region Theme Color
        ///<inheritdoc />
        public override TDrawingColor GetColorTheme(TThemeColor themeColor)
        {
            return FWorkbook.Globals.Theme.Elements.ColorScheme[themeColor];
        }

        ///<inheritdoc />
        public override void SetColorTheme(TThemeColor themeColor, TDrawingColor value)
        {
            CheckWkConnected();
            FWorkbook.Globals.Theme.Elements.ColorScheme[themeColor] = value;
            FWorkbook.Globals.ThemeRecord.InvalidateData();
        }

        ///<inheritdoc />
        public override TThemeColor NearestColorTheme(Color value, out double tint)
        {
            tint = 0;
            CheckWkConnected();
            TThemeColor Result = TThemeColor.Foreground1;
            double MinDist = -1;
            double MinTint = 2;
            THSLColor hsl = value;
            double ColorHue = hsl.Hue;
            double ColorBrightness = hsl.Lum;
            double ColorSaturation = hsl.Sat;

            foreach (TThemeColor theme in TCompactFramework.EnumGetValues(typeof(TThemeColor)))
            {
                if (theme == TThemeColor.None) continue;
                Color p = GetColorTheme(theme).ToColor(this);
                THSLColor phsl = p;
                double hue = phsl.Hue;
                double sat = phsl.Sat;
                double ColorTint = 0;

                double Distance = THSLColor.DistanceSquared(hue, sat, ColorHue, ColorSaturation);

                if (MinDist == -1 || Distance <= MinDist) ColorTint = THSLColor.GetTint(ColorBrightness, phsl.Lum);

                if (MinDist == -1 || Distance < MinDist || (Distance == MinDist && Math.Abs(ColorTint) < MinTint))
                {
                    MinDist = Distance;
                    MinTint = Math.Abs(ColorTint);
                    Result = theme;
                    tint = ColorTint;
                    if (Distance == 0 && tint == 0) //if we don't check for tint = 0 black will be the same as white.
                    {
                        return Result; //exact match...
                    }
                }
            }

            return Result;
        }

        ///<inheritdoc />
        public override TThemeFont GetThemeFont(TFontScheme fontScheme)
        {
            switch (fontScheme)
            {
                case TFontScheme.None:
                    return null;

                case TFontScheme.Minor:
                    return FWorkbook.Globals.Theme.Elements.FontScheme.MinorFont;

                case TFontScheme.Major:
                    return FWorkbook.Globals.Theme.Elements.FontScheme.MajorFont;

            }

            return null;
        }

        ///<inheritdoc />
        public override void SetThemeFont(TFontScheme fontScheme, TThemeFont font)
        {
            switch (fontScheme)
            {
                case TFontScheme.None:
                    break;

                case TFontScheme.Minor:
                    FWorkbook.Globals.ThemeRecord.InvalidateData();
                    FWorkbook.Globals.Theme.Elements.FontScheme.MinorFont = font;
                    break;

                case TFontScheme.Major:
                    FWorkbook.Globals.ThemeRecord.InvalidateData();
                    FWorkbook.Globals.Theme.Elements.FontScheme.MajorFont = font;
                    break;

            }
        }

#if(FRAMEWORK30)
        ///<inheritdoc />
        public override TTheme GetTheme()
        {
            return FWorkbook.Globals.Theme.Clone();
        }

        ///<inheritdoc />
        public override void SetTheme(TTheme aTheme)
        {
            FWorkbook.Globals.ThemeRecord.Theme = aTheme;
            FWorkbook.Globals.ThemeRecord.InvalidateData();
        }
#endif
        #endregion

        #region Named Ranges
        ///<inheritdoc />
        public override int NamedRangeCount
        {
            get
            {
                CheckWkConnected();
                return FWorkbook.Globals.Names.Count;
            }
        }

        ///<inheritdoc />
        public override TXlsNamedRange GetNamedRange(int index)
        {
            CheckWkConnected();
            CheckRange(index, 1, NamedRangeCount, FlxParam.NamedRangeIndex);
            TNameRecord r = FWorkbook.Globals.Names[index - 1];

            int nrSheet = r.RangeSheet + 1;

            /* Docs are wrong. R.RangeSheet has the sheet index, not the externsheetindex
            int nrSheet=0;
            if (r.RangeSheet>= FWorkbook.Globals.References.ExternRefsCount)
            {
                nrSheet=-1;
            }
            else
            if (r.RangeSheet>=0)
            {
                nrSheet = FWorkbook.Globals.References.GetSheet(r.RangeSheet)+1;
            }
            */

            string RangeFormula = null;
            if (!IgnoreFormulaText)
            {
                RangeFormula = TFormulaConvertInternalToText.AsString(r.FormulaData, 0, 0, ActiveSheetObject.Cells.CellList);
            }

            return new TXlsNamedRange(r.Name, nrSheet, r.RefersToSheet(FWorkbook.Globals.References) + 1,
                r.R1 + 1, r.C1 + 1, r.R2 + 1, r.C2 + 1, r.OptionFlags, r.FormulaData, RangeFormula, r.Comment);
        }

        ///<inheritdoc />
        public override TXlsNamedRange GetNamedRange(string Name, int refersToSheetIndex)
        {
            for (int i = 1; i <= NamedRangeCount; i++)
            {
                TNameRecord r = FWorkbook.Globals.Names[i - 1];
                if (String.Equals(Name, r.Name, StringComparison.CurrentCultureIgnoreCase) && (refersToSheetIndex <= 0 || r.RefersToSheet(FWorkbook.Globals.References) + 1 == refersToSheetIndex))
                    return GetNamedRange(i);
            }
            return null;

        }
        ///<inheritdoc />
        public override TXlsNamedRange GetNamedRange(string Name, int refersToSheetIndex, int localSheetIndex)
        {
            for (int i = 1; i <= NamedRangeCount; i++)
            {
                TNameRecord r = FWorkbook.Globals.Names[i - 1];
                int nrSheet = r.RangeSheet + 1;
                if (String.Equals(Name, r.Name, StringComparison.CurrentCultureIgnoreCase) && (refersToSheetIndex <= 0 || r.RefersToSheet(FWorkbook.Globals.References) + 1 == refersToSheetIndex)
                    && (nrSheet == localSheetIndex))
                    return GetNamedRange(i);
            }
            return null;
        }

        ///<inheritdoc />
        public override int FindNamedRange(string Name, int localSheetIndex)
        {
            for (int i = 1; i <= NamedRangeCount; i++)
            {
                TNameRecord r = FWorkbook.Globals.Names[i - 1];
                int nrSheet = r.RangeSheet + 1;
                if (String.Equals(Name, r.Name, StringComparison.CurrentCultureIgnoreCase) && (nrSheet == localSheetIndex))
                    return i;
            }
            return -1;
        }

        internal override TParsedTokenList GetNamedRangeData(int nameIndex, out string externalName, out bool isAddin, out bool Error)
        {
            externalName = null;
            Error = false;
            isAddin = false;

            //CheckRange(nameIndex, 1, NamedRangeCount, FlxParam.NamedRangeIndex);
            if (nameIndex < 1 || nameIndex > NamedRangeCount) { Error = true; return null; }  //the above method can get too slow.

            TNameRecord r = FWorkbook.Globals.Names[nameIndex - 1];

            isAddin = r.IsAddin;
            if (isAddin)
            {
                externalName = r.Name;
                return null;
            }

            return r.Data;
        }

        internal override TParsedTokenList GetNamedRangeData(int externSheetIndex, int externNameIndex, out string externalBook, out string externalName, out int sheetIndexInOtherFile, out bool isAddin, out bool Error)
        {
            externalName = null;
            externalBook = null;
            isAddin = false;
            Error = false;
            sheetIndexInOtherFile = -1;

            if (!FWorkbook.Globals.References.IsLocalSheet(externSheetIndex))
            {
                isAddin = FWorkbook.Globals.References.IsAddinSheet(externSheetIndex);
                externalName = FWorkbook.Globals.References.GetExternName(externSheetIndex, externNameIndex, out externalBook, out sheetIndexInOtherFile);
                return null;
            }

            return GetNamedRangeData(externNameIndex, out externalName, out isAddin, out Error);
        }

        internal override INameRecordList GetNameRecordList()
        {
            return FWorkbook.Globals.Names;
        }

        internal override object EvaluateNamedRange(int nameIndex, int sheetIndex, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
        {
            string ExtName;
            bool IsAddin;
            bool Error;

            TParsedTokenList RangeData = GetNamedRangeData(nameIndex, out ExtName, out IsAddin, out Error);
            if (Error || RangeData == null || RangeData.IsEmpty) return TFlxFormulaErrorValue.ErrNA;
            TWorkbookInfo wi = new TWorkbookInfo(this, sheetIndex, 0, 0, 0, 0, 0, 0, false);
            return RangeData.EvaluateAll(wi, f, CalcState, CalcStack);
        }



        ///<inheritdoc />
        public override int SetNamedRange(TXlsNamedRange rangeData)
        {
            CheckConnected();
            TXlsNamedRange NewRangeData = (TXlsNamedRange)rangeData.Dec();
            NewRangeData.SheetIndex--;
            NewRangeData.NameSheetIndex--;
            NeedsRecalc = true;
            return FWorkbook.Globals.Names.AddName(NewRangeData,
                FWorkbook.Globals, ActiveSheetObject.Cells.CellList) + 1;
        }

        ///<inheritdoc />
        public override void SetNamedRange(int index, TXlsNamedRange rangeData)
        {
            CheckConnected();
            TXlsNamedRange NewRangeData = (TXlsNamedRange)rangeData.Dec();
            NewRangeData.SheetIndex--;
            NewRangeData.NameSheetIndex--;
            FWorkbook.Globals.Names.ReplaceName(index - 1, NewRangeData,
                FWorkbook.Globals, ActiveSheetObject.Cells.CellList);
            NeedsRecalc = true;
        }

        ///<inheritdoc />
        public override void DeleteNamedRange(int index)
        {
            CheckWkConnected();
            FWorkbook.Globals.Names.DeleteName(index - 1, FWorkbook);
            NeedsRecalc = true;
        }

        ///<inheritdoc />
        public override bool[] GetUsedNamedRanges()
        {
            CheckWkConnected();
            TDeletedRanges DeletedRanges = FWorkbook.FindUnreferencedRanges(-1, 0);
            bool[] Result = new bool[NamedRangeCount];

            for (int i = 0; i < Result.Length; i++)
            {
                Result[i] = DeletedRanges.Referenced(i);
            }
            return Result;
        }

        internal override int AddEmptyName(string name, int sheet)
        {
            TXlsNamedRange rangeData = new TXlsNamedRange(name, sheet - 1, 0, String.Empty);
            return FWorkbook.Globals.Names.AddName(rangeData,
                FWorkbook.Globals, null);
            
        }

        #endregion

        #region Copy and Paste (Clipboard)
        ///<inheritdoc />
        public override void CopyToClipboardFormat(TXlsCellRange range, StringBuilder textString, Stream xlsStream)
        {
            CheckConnected();
            CheckRowAndCol(range.Top, range.Left);
            CheckRowAndCol(range.Bottom, range.Right);

            Recalc(false);

            if (textString != null)
            {
                using (StringWriter sw = new StringWriter(textString))
                {
                    TextDelim.Write(sw, this, '\t', range, true);
                }
            }
            if (xlsStream != null)
            {
                using (MemoryStream MemFile = new MemoryStream(OtherStreams))
                {
                    using (TOle2File DataStream = new TOle2File(MemFile))
                    {
                        DataStream.PrepareForWrite(xlsStream, XlsConsts.WorkbookString, new string[0]);
                        FWorkbook.SaveRangeToStream(DataStream, new TSaveData(this, FWorkbook.Globals.Borders, FWorkbook.Globals, XlsBiffVersion, ErrorActions), FActiveSheet - 1, range.Dec());
                    }
                }
            }
        }

        ///<inheritdoc />
        public override void PasteFromTextClipboardFormat(int row, int col, TFlxInsertMode insertMode, string data)
        {
            CheckConnected();
            CheckRowAndCol(row, col);
            if (data == null) return;
            NeedsRecalc = true;

            XlsFile tmpWorkbook = new XlsFile();
            tmpWorkbook.NewFile();
            using (StringReader sr = new StringReader(data))
            {
                TextDelim.Read(sr, tmpWorkbook, '\t', 1, 1, null, null); //no dec here.
            }
            int cc = tmpWorkbook.ColCount;
            if (tmpWorkbook.RowCount == 0 || cc == 0) return;
            InsertAndCopyRange(new TXlsCellRange(1, 1, tmpWorkbook.RowCount, cc), row, col, 1, insertMode, TRangeCopyMode.All, tmpWorkbook, 1);
        }

        ///<inheritdoc />
        public override void PasteFromXlsClipboardFormat(int row, int col, TFlxInsertMode insertMode, Stream data)
        {
            CheckConnected();
            CheckRowAndCol(row, col);
            XlsFile tmpWorkbook = new XlsFile();
            tmpWorkbook.Open(data);

            NeedsRecalc = true;
            if (tmpWorkbook.SheetCount <= 0) return;
            tmpWorkbook.ActiveSheet = 1;
            if (tmpWorkbook.SheetType != TSheetType.Worksheet) return; //Biff8 only pastes one sheet
            TDimensionsRecord d = tmpWorkbook.FWorkbook.WorkSheets(0).OriginalDimensions;

            InsertAndCopyRange(new TXlsCellRange((int)d.FirstRow() + 1, (int)d.FirstCol() + 1, (int)d.LastRow(), (int)d.LastCol()), row, col, 1, insertMode, TRangeCopyMode.All, tmpWorkbook, 1);
        }


        #endregion

        #region Print
        ///<inheritdoc />
        public override string PageHeader
        {
            get
            {
                CheckConnected();
                return ActiveSheetObject.PageSetup.HeaderAndFooter.DefaultHeader;
            }
            set
            {
                CheckConnected();
                ActiveSheetObject.PageSetup.HeaderAndFooter.SetAllHeaders(value);
            }
        }

        ///<inheritdoc />
        public override string PageFooter
        {
            get
            {
                CheckConnected();
                return ActiveSheetObject.PageSetup.HeaderAndFooter.DefaultFooter;
            }
            set
            {
                CheckConnected();
                ActiveSheetObject.PageSetup.HeaderAndFooter.SetAllFooters(value);
            }
        }

        ///<inheritdoc />
        public override THeaderAndFooter GetPageHeaderAndFooter()
        {
            CheckConnected();
            return ActiveSheetObject.PageSetup.HeaderAndFooter;
        }

        ///<inheritdoc />
        public override void SetPageHeaderAndFooter(THeaderAndFooter headerAndFooter)
        {
            ActiveSheetObject.PageSetup.HeaderAndFooter = headerAndFooter;
        }

        private static void FindOnePart(string text, int p1, int p2, int p3, ref string part)
        {
            if (p1 >= 0)
            {
                int En = text.Length;
                if (p2 > p1 && p2 + 1 < En) En = p2 + 1;
                if (p3 > p1 && p3 + 1 < En) En = p3 + 1;
                part = text.Substring(p1 + 2, En - (p1 + 2));
            }
            else part = String.Empty;
            /* We can't do this here. If we do, and for example we have a header: &10&D (font 10, date), it will be
             * replaced with &1020/01/2004 which would mean font=1020. Also, if f.e. SheetName="Hoja&D", the &1 would be later replaced with a date.
            part=part.Replace(@"&A", SheetName);
            part=part.Replace(@"&D", DateTime.Now.Date.ToShortDateString());
            part=part.Replace(@"&T", DateTime.Now.ToShortTimeString());
            part=part.Replace(@"&P", pageNumber.ToString());
            part=part.Replace(@"&N", pageCount.ToString());
            part=part.Replace(@"&F", Path.GetFileName(FActiveFileName));
            part=part.Replace(@"&Z", Path.GetFullPath(FActiveFileName));
            */
        }

        ///<inheritdoc />
        public override void FillPageHeaderOrFooter(string fullText, ref string leftText, ref string centerText, ref string rightText)
        {
            string LS = @"&L";
            string RS = @"&R";
            string CS = @"&C";

            string s = fullText;
            if (fullText.Length > 2) s = fullText.Substring(0, 2);
            string aText = fullText;
            if (s != LS && s != RS && s != CS)
                aText = CS + fullText;
            int Pl = aText.IndexOf(LS);
            int Pc = aText.IndexOf(CS);
            int Pr = aText.IndexOf(RS);
            FindOnePart(aText, Pl, Pc, Pr, ref leftText);
            FindOnePart(aText, Pc, Pl, Pr, ref centerText);
            FindOnePart(aText, Pr, Pl, Pc, ref rightText);
        }

        ///<inheritdoc />
        public override string GetPageHeaderOrFooterAsHtml(string section, string imageTag, int pageNumber, int pageCount, THtmlVersion htmlVersion, Encoding encoding, IHtmlFontEvent onFont)
        {
            CheckConnected();
            if (section == null) return string.Empty;
            return TPageHeaderFooterRecord.AsHtml(this, section, imageTag, pageNumber, pageCount, htmlVersion, encoding, onFont);
        }


        ///<inheritdoc />
        public override void GetHeaderOrFooterImage(THeaderAndFooterKind headerAndFooterKind, THeaderAndFooterPos section, ref TXlsImgType imageType, Stream outStream)
        {
            CheckConnected();
            ActiveSheetObject.HeaderImages.GetDrawingFromStream(headerAndFooterKind, section, outStream, ref imageType);
        }

        ///<inheritdoc />
        public override THeaderOrFooterImageProperties GetHeaderOrFooterImageProperties(THeaderAndFooterKind headerAndFooterKind, THeaderAndFooterPos section)
        {
            CheckConnected();
            return ActiveSheetObject.HeaderImages.GetHeaderOrFooterImageProperties(headerAndFooterKind, section);
        }


        ///<inheritdoc />
        public override void SetHeaderOrFooterImage(THeaderAndFooterKind headerAndFooterKind, THeaderAndFooterPos section, byte[] data, TXlsImgType imageType, THeaderOrFooterImageProperties properties)
        {
            CheckConnected();

            if (data != null)
                ActiveSheetObject.HeaderImages.AssignHeaderOrFooterDrawing(headerAndFooterKind, section, data, imageType, properties);
            else
                ActiveSheetObject.HeaderImages.DeleteHeaderOrFooterImage(headerAndFooterKind, section);
        }


        ///<inheritdoc />
        public override bool PrintGridLines
        {
            get
            {
                CheckConnected();
                return ActiveSheetObject.SheetGlobals.PrintGridLines;
            }
            set
            {
                CheckConnected();
                ActiveSheetObject.SheetGlobals.PrintGridLines = value;
            }
        }

        ///<inheritdoc />
        public override bool PrintHeadings
        {
            get
            {
                CheckConnected();
                return ActiveSheetObject.SheetGlobals.PrintHeaders;
            }
            set
            {
                CheckConnected();
                ActiveSheetObject.SheetGlobals.PrintHeaders = value;
            }
        }

        ///<inheritdoc />
        public override bool PrintHCentered
        {
            get
            {
                CheckConnected();
                return ActiveSheetObject.PageSetup.HCenter;
            }
            set
            {
                CheckConnected();
                ActiveSheetObject.PageSetup.HCenter = value;
            }
        }

        ///<inheritdoc />
        public override bool PrintVCentered
        {
            get
            {
                CheckConnected();
                return ActiveSheetObject.PageSetup.VCenter;
            }
            set
            {
                CheckConnected();
                ActiveSheetObject.PageSetup.VCenter = value;
            }
        }

        ///<inheritdoc />
        public override TXlsMargins GetPrintMargins()
        {
            CheckConnected();
            return ActiveSheetObject.Margins;
        }

        ///<inheritdoc />
        public override void SetPrintMargins(TXlsMargins value)
        {
            CheckConnected();
            ActiveSheetObject.Margins = value;
        }

        ///<inheritdoc />
        public override bool PrintToFit
        {
            get
            {
                CheckConnected();
                return ActiveSheetObject.SheetGlobals.WsBool.FitToPage;
            }
            set
            {
                CheckConnected();
                ActiveSheetObject.SheetGlobals.WsBool.FitToPage = value;
            }
        }

        ///<inheritdoc />
        public override int PrintCopies
        {
            get
            {
                CheckConnected();
                return ActiveSheetObject.PageSetup.Setup.Copies;
            }
            set
            {
                CheckConnected();
                ActiveSheetObject.PageSetup.Setup.Copies = value;
            }
        }

        ///<inheritdoc />
        public override int PrintXResolution
        {
            get
            {
                CheckConnected();
                return ActiveSheetObject.PageSetup.Setup.HPrintRes;
            }
            set
            {
                CheckConnected();
                ActiveSheetObject.PageSetup.Setup.HPrintRes = value;
            }
        }

        ///<inheritdoc />
        public override int PrintYResolution
        {
            get
            {
                CheckConnected();
                return ActiveSheetObject.PageSetup.Setup.VPrintRes;
            }
            set
            {
                CheckConnected();
                ActiveSheetObject.PageSetup.Setup.VPrintRes = value;
            }
        }

        ///<inheritdoc />
        public override TPrintOptions PrintOptions
        {
            get
            {
                CheckConnected();
                return (TPrintOptions)ActiveSheetObject.PrintOptions;
            }
            set
            {
                CheckConnected();
                ActiveSheetObject.PrintOptions = (int)value;
            }
        }

        ///<inheritdoc />
        public override int PrintScale
        {
            get
            {
                CheckConnected();
                int aScale = ActiveSheetObject.PageSetup.Setup.Scale;
                if (aScale <= 0) return 100;
                return aScale;
            }
            set
            {
                CheckConnected();
                ActiveSheetObject.PageSetup.Setup.Scale = value;
            }
        }

        ///<inheritdoc />
        public override int? PrintFirstPageNumber
        {
            get
            {
                CheckConnected();
                return ActiveSheetObject.PageSetup.Setup.PageStart;
            }
            set
            {
                CheckConnected();
                ActiveSheetObject.PageSetup.Setup.PageStart = value;
            }
        }

        ///<inheritdoc />
        public override int PrintNumberOfHorizontalPages
        {
            get
            {
                CheckConnected();
                return ActiveSheetObject.PageSetup.Setup.FitWidth;
            }
            set
            {
                CheckConnected();
                ActiveSheetObject.PageSetup.Setup.FitWidth = value;
            }
        }

        ///<inheritdoc />
        public override int PrintNumberOfVerticalPages
        {
            get
            {
                CheckConnected();
                return ActiveSheetObject.PageSetup.Setup.FitHeight;
            }
            set
            {
                CheckConnected();
                ActiveSheetObject.PageSetup.Setup.FitHeight = value;
            }
        }

        ///<inheritdoc />
        public override TPaperSize PrintPaperSize
        {
            get
            {
                CheckConnected();
                return (TPaperSize)ActiveSheetObject.PageSetup.Setup.PaperSize;
            }
            set
            {
                CheckConnected();
                ActiveSheetObject.PageSetup.Setup.PaperSize = (int) value;
            }
        }

        ///<inheritdoc />
        public override TPaperDimensions PrintPaperDimensions
        {
            get
            {
                const long DM_PAPERSIZE = 0x00000002L;
                const long DM_PAPERLENGTH = 0x00000004L;
                const long DM_PAPERWIDTH = 0x00000008L;

                TPaperDimensions Result = new TPaperDimensions(PrintPaperSize);
                if (Result.Width > 0) return Result;

                TPrinterDriverSettings Pd = GetPrinterDriverSettings();
                if (Pd != null)
                {
                    byte[] PData = Pd.GetData();
                    if (PData != null && PData.Length > 66 + 8 + 12 && PData[0] == 0 && PData[1] == 0)
                    {
                        long Flags = BitOps.GetCardinal(PData, 66 + 8);

                        if ((Flags & DM_PAPERSIZE) != 0)
                        {
                            int PaperSize = BitConverter.ToInt16(PData, 66 + 8 + 6);
                            if (Enum.IsDefined(typeof(TPaperSize), PaperSize))
                                Result = new TPaperDimensions((TPaperSize)PaperSize);
                            else
                            {
                                Result = new TPaperDimensions(TPaperSize.A4);
                                Result.PaperName = String.Empty;
                            }
                        }
                        else
                        {
                            Result = new TPaperDimensions(TPaperSize.A4);
                            Result.PaperName = String.Empty;
                        }

                        if ((Flags & DM_PAPERWIDTH) != 0)
                        {
                            Result.PaperName = FlxMessages.GetString(FlxMessage.CustomPageSize);
                            Result.Width = TPaperDimensions.mm(BitConverter.ToInt16(PData, 66 + 8 + 10) / 10F);
                        }
                        if ((Flags & DM_PAPERLENGTH) != 0)
                        {
                            Result.PaperName = FlxMessages.GetString(FlxMessage.CustomPageSize);
                            Result.Height = TPaperDimensions.mm(BitConverter.ToInt16(PData, 66 + 8 + 8) / 10F);
                        }

                        if (Result.Width <= 0 || Result.Height <= 0)
                        {
                            Result = new TPaperDimensions(TPaperSize.A4);
                            Result.PaperName = String.Empty;
                        }
                        return Result;
                    }
                }
                Result = new TPaperDimensions(TPaperSize.A4);
                Result.PaperName = String.Empty;
                return Result;
            }
        }


        ///<inheritdoc />
        public override TPrinterDriverSettings GetPrinterDriverSettings()
        {
            CheckConnected();
            return ActiveSheetObject.PrinterDriverSettings;
        }

        ///<inheritdoc />
        public override void SetPrinterDriverSettings(TPrinterDriverSettings settings)
        {
            CheckConnected();
            ActiveSheetObject.PrinterDriverSettings = settings;
        }



        #endregion

        #region Images
        ///<inheritdoc />
        public override int ImageCount
        {
            get
            {
                CheckConnected();
                return ActiveSheetObject.Drawing.DrawingCount;
            }
        }

        ///<inheritdoc />
        public override void SetImage(int imageIndex, byte[] data, TXlsImgType imageType, bool usesObjectIndex, string objectPath)
        {
            CheckConnected();
            if (usesObjectIndex)
            {
                CheckRangeObjPath(objectPath, imageIndex, 1, ObjectCount, FlxParam.ImageIndex);
            }
            else
            {
                CheckRange(imageIndex, 1, ObjectCount, FlxParam.ImageIndex);
            }

            if (data != null)
                ActiveSheetObject.Drawing.AssignDrawing(imageIndex - 1, data, imageType, usesObjectIndex, objectPath);
        }

        ///<inheritdoc />
        public override void SetImageProperties(int imageIndex, TImageProperties imageProperties)
        {
            CheckConnected();
            CheckRange(imageIndex, 1, ImageCount, FlxParam.ImageIndex);
            if (imageProperties != null)
            {
                ActiveSheetObject.Drawing.SetImageProperties(imageIndex - 1, imageProperties.Dec(), ActiveSheetObject);
            }
        }


        ///<inheritdoc />
        public override string GetImageName(int imageIndex)
        {
            CheckConnected();
            CheckRange(imageIndex, 1, ImageCount, FlxParam.ImageIndex);
            return ActiveSheetObject.Drawing.DrawingName(imageIndex - 1);
        }


        ///<inheritdoc />
        public override void GetImage(int imageIndex, string objectPath, ref TXlsImgType imageType, Stream outStream, bool usesObjectIndex)
        {
            CheckConnected();
            if (usesObjectIndex)
            {
                CheckRangeObjPath(objectPath, imageIndex, 1, ObjectCount, FlxParam.ImageIndex);
            }
            else
            {
                CheckRange(imageIndex, 1, ImageCount, FlxParam.ImageIndex);
            } 
            
            ActiveSheetObject.Drawing.GetDrawingFromStream(imageIndex - 1, objectPath, outStream, ref imageType, usesObjectIndex);
        }

        ///<inheritdoc />
        public override TImageProperties GetImageProperties(int imageIndex)
        {
            CheckConnected();
            CheckRange(imageIndex, 1, ImageCount, FlxParam.ImageIndex);
            return ActiveSheetObject.Drawing.GetImageProperties(imageIndex - 1).Inc();
        }

        /// <summary>
        /// Only for internal testing. 
        /// </summary>
        /// <param name="imageIndex"></param>
        /// <returns>References to the blip on the workbook.</returns>
        internal long GetImageReferences(int imageIndex)
        {
            return ActiveSheetObject.Drawing.ReferencesCount(imageIndex - 1);
        }

        ///<inheritdoc />
        public override void AddImage(byte[] data, TXlsImgType imageType, TImageProperties imageProperties)
        {
            CheckConnected();
            ActiveSheetObject.Drawing.AddImage(this, data, imageType, imageProperties.Dec(), false, 
                imageProperties.ShapeName, ActiveSheetObject, ActiveSheet, null, false);
        }


        ///<inheritdoc />
        public override void DeleteImage(int imageIndex)
        {
            CheckConnected();
            CheckRange(imageIndex, 1, ImageCount, FlxParam.ImageIndex);
            ActiveSheetObject.Drawing.DeleteImage(imageIndex - 1);
        }


        ///<inheritdoc />
        public override void ClearImage(int imageIndex)
        {
            CheckConnected();
            CheckRange(imageIndex, 1, ImageCount, FlxParam.ImageIndex);
            ActiveSheetObject.Drawing.ClearImage(imageIndex - 1);
        }


        #region Objects
        ///<inheritdoc />
        public override int ObjectCount
        {
            get
            {
                CheckConnected();
                return ActiveSheetObject.Drawing.ObjectCount;
            }
        }

        ///<inheritdoc />
        public override int ImageIndexToObjectIndex(int imageIndex)
        {
            CheckConnected();
            CheckRange(imageIndex, 1, ImageCount, FlxParam.ImageIndex);
            return ActiveSheetObject.Drawing.ImageIndexToObjectIndex(imageIndex - 1) + 1;
        }

        ///<inheritdoc />
        public override int ObjectIndexToImageIndex(int objectIndex)
        {
            CheckConnected();
            CheckRange(objectIndex, 1, ObjectCount, FlxParam.ObjectIndex);
            return ActiveSheetObject.Drawing.ObjectIndexToImageIndex(objectIndex - 1) + 1;
        }


        ///<inheritdoc />
        public override string GetObjectName(int objectIndex)
        {
            CheckConnected();
            CheckRange(objectIndex, 1, ObjectCount, FlxParam.ObjectIndex);
            return ActiveSheetObject.Drawing.ObjectName(objectIndex - 1);
        }

        ///<inheritdoc />
        public override long GetObjectShapeId(int objectIndex)
        {
            CheckConnected();
            CheckRange(objectIndex, 1, ObjectCount, FlxParam.ObjectIndex);
            return ActiveSheetObject.Drawing.ShapeId(objectIndex - 1);
        }
       
        ///<inheritdoc />
        public override string FindObjectPath(string objectName)
        {
            CheckConnected();
            return ActiveSheetObject.Drawing.FindPath(objectName);
        }

        ///<inheritdoc />
        public override int FindObjectByShapeId(long ShapeId)
        {
            CheckConnected();
            return ActiveSheetObject.Drawing.FindShapeIdIndex(ShapeId) + 1;
        }

        ///<inheritdoc />
        public override TClientAnchor GetObjectAnchor(int objectIndex)
        {
            CheckConnected();
            CheckRange(objectIndex, 1, ObjectCount, FlxParam.ObjectIndex);
            return ActiveSheetObject.Drawing.GetObjectAnchor(objectIndex - 1).Inc();
        }

        ///<inheritdoc />
        public override void SetObjectAnchor(int objectIndex, string objectPath, TClientAnchor objectAnchor)
        {
            CheckConnected();
            CheckRangeObjPath(objectPath, objectIndex, 1, ObjectCount, FlxParam.ObjectIndex);
            ActiveSheetObject.Drawing.SetObjectAnchor(objectIndex - 1, objectPath, objectAnchor.Dec(), ActiveSheetObject);
        }

        private void IncAnchor(TShapeProperties Props)
        {
            if (Props.Anchor != null)
            {
                Props.Anchor.Row1++;
                Props.Anchor.Col1++;
                Props.Anchor.Row2++;
                Props.Anchor.Col2++;
            }

            for (int i = 1; i <= Props.ChildrenCount; i++)
            {
                IncAnchor(Props.Children(i));
            }
        }
        
        ///<inheritdoc />
        public override TShapeProperties GetObjectProperties(int objectIndex, bool getShapeOptions)
        {
            CheckConnected();
            CheckRange(objectIndex, 1, ObjectCount, FlxParam.ObjectIndex);
            TShapeProperties Result = ActiveSheetObject.Drawing.GetObjectProperties(objectIndex - 1, getShapeOptions);

            IncAnchor(Result);
            return Result;
        }

        ///<inheritdoc />
        public override TShapeProperties GetObjectPropertiesByShapeId(long shapeId, bool getShapeOptions)
        {
            CheckConnected();
            TShapeProperties Result = ActiveSheetObject.Drawing.GetObjectPropertiesByShapeId(shapeId, getShapeOptions);
            if (Result == null) return null;
            IncAnchor(Result);
            return Result;
        }

        ///<inheritdoc />
        public override void SetObjectText(int objectIndex, string objectPath, TRichString text)
        {
            CheckConnected();
            CheckRangeObjPath(objectPath, objectIndex, 1, ObjectCount, FlxParam.ObjectIndex);
            ActiveSheetObject.Drawing.SetObjectText(objectIndex - 1, objectPath, text, this, null);
        }

        ///<inheritdoc />
        public override void SetObjectName(int objectIndex, string objectPath, string name)
        {
            CheckConnected();
            CheckRangeObjPath(objectPath, objectIndex, 1, ObjectCount, FlxParam.ObjectIndex);
            ActiveSheetObject.Drawing.SetObjectName(objectIndex - 1, objectPath, name);
        }

        ///<inheritdoc />
        public override void SetObjectProperty(int objectIndex, string objectPath, TShapeOption shapeProperty, long value)
        {
            CheckConnected();
            CheckRangeObjPath(objectPath, objectIndex, 1, ObjectCount, FlxParam.ObjectIndex);
            ActiveSheetObject.Drawing.SetObjectProperty(objectIndex - 1, objectPath, shapeProperty, value);
        }

        ///<inheritdoc />
        public override void SetObjectProperty(int objectIndex, string objectPath, TShapeOption shapeProperty, double value)
        {
            CheckConnected();
            CheckRangeObjPath(objectPath, objectIndex, 1, ObjectCount, FlxParam.ObjectIndex);
            
            int i = (int)value;
            int f = (int)(value % 1.0) * 65536;
            long v = (i << 16) | f;
            ActiveSheetObject.Drawing.SetObjectProperty(objectIndex - 1, objectPath, shapeProperty, v);
        }

        ///<inheritdoc />
        public override void SetObjectProperty(int objectIndex, string objectPath, TShapeOption shapeProperty, int positionInSet, bool value)
        {
            CheckConnected();
            CheckRangeObjPath(objectPath, objectIndex, 1, ObjectCount, FlxParam.ObjectIndex);
            ActiveSheetObject.Drawing.SetObjectProperty(objectIndex - 1, objectPath, shapeProperty, positionInSet, value);
        }

        ///<inheritdoc />
        public override void SetObjectProperty(int objectIndex, string objectPath, TShapeOption shapeProperty, string text)
        {
            CheckConnected();
            CheckRangeObjPath(objectPath, objectIndex, 1, ObjectCount, FlxParam.ObjectIndex);
            ActiveSheetObject.Drawing.SetObjectProperty(objectIndex - 1, objectPath, shapeProperty, text);
        }

        ///<inheritdoc />
        public override void SetObjectProperty(int objectIndex, string objectPath, TShapeOption shapeProperty, THyperLink value)
        {
            CheckConnected();
            CheckRangeObjPath(objectPath, objectIndex, 1, ObjectCount, FlxParam.ObjectIndex);
            ActiveSheetObject.Drawing.SetObjectProperty(objectIndex - 1, objectPath, shapeProperty, value);
        }

        /*		///<inheritdoc />
                public override void AddAutoShape(TShapeType shapeType, TClientAnchor clientAnchor, TRichString text)
                {
                    CheckConnected();
                    CheckRange((int)shapeType, 1, 0x0FFF, FlxParam.AutoShapeIndex);
                    if (!ActiveSheetIsWorksheet) return;
                    ActiveSheetObject.AddAutoShape(shapeType, clientAnchor.Dec(), text);		
                }
        */


        ///<inheritdoc />
        public override void DeleteObject(int objectIndex, string objectPath)
        {
            CheckConnected();
            CheckRangeObjPath(objectPath, objectIndex, 1, ObjectCount, FlxParam.ObjectIndex);
            ActiveSheetObject.Drawing.DeleteObject(objectIndex - 1, objectPath);
        }

        ///<inheritdoc />
        public override void GetObjectsInRange(TXlsCellRange range, TExcelObjectList objectsInRange)
        {
            CheckConnected();
            ActiveSheetObject.Drawing.GetObjectsInRange(false, range.Top - 1, range.Bottom - 1, range.Left - 1, range.Right - 1, objectsInRange, null);
        }

        #region Linked cells
        ///<inheritdoc />
        public override TCellAddress GetObjectLinkedCell(int objectIndex, string objectPath)
        {
            CheckConnected();
            CheckRangeObjPath(objectPath, objectIndex, 1, ObjectCount, FlxParam.ObjectIndex);
            return ActiveSheetObject.Drawing.GetObjectLink(objectIndex - 1, objectPath, this); //tcelladdress is 1 based.
        }

        ///<inheritdoc />
        public override void SetObjectLinkedCell(int objectIndex, string objectPath, TCellAddress linkedCell)
        {
            CheckConnected();
            CheckRangeObjPath(objectPath, objectIndex, 1, ObjectCount, FlxParam.ObjectIndex);
            ActiveSheetObject.Drawing.SetObjectLink(objectIndex - 1, objectPath, linkedCell, null, this, false);
        }

        private object GetLinkedValue(int objectIndex, string objectPath, out bool IsLinked)
        {
            IsLinked = false;
            CheckConnected();
            CheckRangeObjPath(objectPath, objectIndex, 1, ObjectCount, FlxParam.ObjectIndex);
            TCellAddress LinkedCell = GetObjectLinkedCell(objectIndex, objectPath);
            if (LinkedCell != null)
            {
                IsLinked = true;
                int Sheet = ActiveSheet;
                if (!string.IsNullOrEmpty(LinkedCell.Sheet)) Sheet = GetSheetIndex(LinkedCell.Sheet, false);
                if (Sheet > 0)
                {
                    int XF = 1;
                    return GetCellValue(Sheet, LinkedCell.Row, LinkedCell.Col, ref XF); 
                }
            }
            return null;
        }

        private void SetLinkedValue(int objectIndex, string objectPath, bool selected, object value)
        {
            TCellAddress LinkedCell = GetObjectLinkedCell(objectIndex, objectPath);
            if (LinkedCell != null)
            {
                int Sheet = ActiveSheet;
                if (!string.IsNullOrEmpty(LinkedCell.Sheet)) Sheet = GetSheetIndex(LinkedCell.Sheet, false);
                if (Sheet <= 0) return;

                if (selected) SetCellValue(Sheet, LinkedCell.Row, LinkedCell.Col, value, -1);
                else
                {
                    //It has changed, so it needs to be deslected.
                    SetCellValue(Sheet, LinkedCell.Row, LinkedCell.Col, null, -1);
                }
            }
        }

        ///<inheritdoc />
        public override TCellAddressRange GetObjectInputRange(int objectIndex, string objectPath)
        {
            CheckConnected();
            CheckRangeObjPath(objectPath, objectIndex, 1, ObjectCount, FlxParam.ObjectIndex);
            return ActiveSheetObject.Drawing.GetObjectInputRange(objectIndex - 1, objectPath, this); //tcelladdress is 1 based.
        }

        ///<inheritdoc />
        public override void SetObjectInputRange(int objectIndex, string objectPath, TCellAddressRange inputRange)
        {
            CheckConnected();
            CheckRangeObjPath(objectPath, objectIndex, 1, ObjectCount, FlxParam.ObjectIndex);
            ActiveSheetObject.Drawing.SetFormulaRange(objectIndex - 1, objectPath, inputRange, null, this, false);
        }

        ///<inheritdoc />
        public override string GetObjectMacro(int objectIndex, string objectPath)
        {
            CheckConnected();
            CheckRangeObjPath(objectPath, objectIndex, 1, ObjectCount, FlxParam.ObjectIndex);
            return ActiveSheetObject.Drawing.GetObjectMacro(objectIndex - 1, objectPath, ActiveSheetObject.Cells.CellList); 
        }

        ///<inheritdoc />
        public override void SetObjectMacro(int objectIndex, string objectPath, string macro)
        {
            CheckConnected();
            TParsedTokenList Tokens = null;

            if (macro != null)
            {
                int ExternSheet = FWorkbook.Globals.References.AddSheet(SheetCount, 0xFFFE);

                int k = FindNamedRange(macro, 0);
                if (k <= 0)
                {
                    TXlsNamedRange nr = new TXlsNamedRange(macro, ExternSheet, 0x00, string.Empty);
                    k = SetNamedRange(nr);
                }

                Tokens = new TParsedTokenList(new TBaseParsedToken[] { new TNameXToken(ptg.NameX, ExternSheet, k) });
            }

            ActiveSheetObject.Drawing.SetButtonMacro(objectIndex - 1, objectPath, Tokens, this);
        }

        #endregion

        #region Checkboxes
        ///<inheritdoc />
        public override TCheckboxState GetCheckboxState(int objectIndex, string objectPath)
        {
            bool IsLinked;
            object v = GetLinkedValue(objectIndex, objectPath, out IsLinked);
            if (IsLinked)
            {
                if (v == null) return TCheckboxState.Unchecked;

                if (v is TFlxFormulaErrorValue)
                {
                    if ((TFlxFormulaErrorValue)v == TFlxFormulaErrorValue.ErrNA) return TCheckboxState.Indeterminate;
                    //if it isn't n/a, it's value doesn't matter. Checkbox stays as is.
                }
                else
                {
                    bool b;
                    if (TBaseParsedToken.ExtToBool(v, out b))
                    {
                        if (b) return TCheckboxState.Checked; else return TCheckboxState.Unchecked;
                    }
                }
            }

            return ActiveSheetObject.Drawing.GetCheckbox(objectIndex - 1, objectPath);
        }

        ///<inheritdoc />
        public override void SetCheckboxState(int objectIndex, string objectPath, TCheckboxState value)
        {
            CheckConnected();
            CheckRangeObjPath(objectPath, objectIndex, 1, ObjectCount, FlxParam.ObjectIndex);
            ActiveSheetObject.Drawing.SetCheckbox(objectIndex - 1, objectPath, value);

            object cv = null;
            switch (value)
            {
                case TCheckboxState.Checked:
                    cv = true;
                    break;

                case TCheckboxState.Unchecked:
                    cv = false;
                    break;

                case TCheckboxState.Indeterminate:
                    cv = TFlxFormulaErrorValue.ErrNA;
                    break;
            }
            SetLinkedValue(objectIndex, objectPath, true, cv);
        }

        ///<inheritdoc />
        [Obsolete("Use GetObjectLinkedCell instead.")]        
        public override TCellAddress GetCheckboxLinkedCell(int objectIndex, string objectPath)
        {
            CheckConnected();
            CheckRangeObjPath(objectPath, objectIndex, 1, ObjectCount, FlxParam.ObjectIndex);
            return ActiveSheetObject.Drawing.GetObjectLink(objectIndex - 1, objectPath, this); //tcelladdress is 1 based.
        }

        ///<inheritdoc />
        [Obsolete("Use SetObjectLinkedCell instead.")]
        public override void SetCheckboxLinkedCell(int objectIndex, string objectPath, TCellAddress linkedCell)
        {
            CheckConnected();
            CheckRangeObjPath(objectPath, objectIndex, 1, ObjectCount, FlxParam.ObjectIndex);
            ActiveSheetObject.Drawing.SetObjectLink(objectIndex - 1, objectPath, linkedCell, null, this, false);
        }

        ///<inheritdoc />
        public override int AddCheckbox(TClientAnchor anchor, TRichString text, TCheckboxState value, TCellAddress linkedCell, string name)
        {
            if (anchor == null) FlxMessages.ThrowException(FlxErr.ErrNullParameter, "anchor");

            CheckConnected();
            int Result = ActiveSheetObject.Drawing.AddCheckbox(anchor.Dec(), text, value, this, ActiveSheetObject, null, name) + 1;
            if (linkedCell != null) SetObjectLinkedCell(Result, null, linkedCell);
            return Result;
        }

        #endregion

        #region Radio buttons
        ///<inheritdoc />
        public override bool GetRadioButtonState(int objectIndex, string objectPath)
        {
            bool IsLinked;
            object v = GetLinkedValue(objectIndex, objectPath, out IsLinked);
            if (IsLinked)
            {
                if (v == null) return false;
                if (v is TFlxFormulaErrorValue)
                {
                    if ((TFlxFormulaErrorValue)v == TFlxFormulaErrorValue.ErrNA) return false;
                    //if it isn't n/a, its value doesn't matter. rb stays as is.
                }
                else
                {
                    double pd;
                    if (TBaseParsedToken.ExtToDouble(v, out pd)) //something like a string doesn't matter
                    {
                        if (pd <= 0 || pd >= int.MaxValue) return false; // in this case it doesn't matter, all cbs are unselected. 
                        int p = (int)pd;
                        return p == ActiveSheetObject.Drawing.GetRbPosition(objectIndex - 1, objectPath, this);
                    }
                }
            }
            return ActiveSheetObject.Drawing.GetRadioButton(objectIndex - 1, objectPath);
        }

        ///<inheritdoc />
        public override void SetRadioButtonState(int objectIndex, string objectPath, bool selected)
        {
            CheckConnected();
            CheckRangeObjPath(objectPath, objectIndex, 1, ObjectCount, FlxParam.ObjectIndex);
            int RbPosition;
            bool Changed;
            ActiveSheetObject.Drawing.SetRadioButton(objectIndex - 1, objectPath, selected, out RbPosition, this, out Changed);
            if (!Changed) return;

            SetLinkedValue(objectIndex, objectPath, selected, RbPosition);
        }

        ///<inheritdoc />
        public override int AddRadioButton(TClientAnchor anchor, TRichString text, string name)
        {
            if (anchor == null) FlxMessages.ThrowException(FlxErr.ErrNullParameter, "anchor");

            CheckConnected();
            int Result = ActiveSheetObject.Drawing.AddRadioButton(anchor.Dec(), text, TCheckboxState.Unchecked, this, ActiveSheetObject, null, name) + 1;
            return Result;
        }

        ///<inheritdoc />
        public override int AddGroupBox(TClientAnchor anchor, TRichString text, string name)
        {
            if (anchor == null) FlxMessages.ThrowException(FlxErr.ErrNullParameter, "anchor");

            CheckConnected();
            int Result = ActiveSheetObject.Drawing.AddGroupBox(anchor.Dec(), text, this, ActiveSheetObject, null, name) + 1;
            return Result;
        }
        #endregion

        #region Other objects
        ///<inheritdoc />
        public override int GetObjectSelection(int objectIndex, string objectPath)
        {
            bool IsLinked;
            object v = GetLinkedValue(objectIndex, objectPath, out IsLinked);
            if (IsLinked)
            {
                if (v == null) return 0;
                if (v is TFlxFormulaErrorValue)
                {
                    if ((TFlxFormulaErrorValue)v == TFlxFormulaErrorValue.ErrNA) return 0;
                    //if it isn't n/a, its value doesn't matter. rb stays as is.
                }
                else
                {
                    double pd;
                    if (TBaseParsedToken.ExtToDouble(v, out pd)) //something like a string doesn't matter
                    {
                        if (pd <= 0 || pd >= int.MaxValue) return 0; // in this case it doesn't matter, all cbs are unselected. 
                        int p = (int)pd;
                        return p;
                    }
                }
            }
            
            return ActiveSheetObject.Drawing.GetObjectSelection(objectIndex - 1, objectPath);
        }

        ///<inheritdoc />
        public override void SetObjectSelection(int objectIndex, string objectPath, int selectedItem)
        {
            CheckConnected();
            CheckRangeObjPath(objectPath, objectIndex, 1, ObjectCount, FlxParam.ObjectIndex);
            ActiveSheetObject.Drawing.SetObjectSelection(objectIndex - 1, objectPath, selectedItem);

            SetLinkedValue(objectIndex, objectPath, true, selectedItem);
        }

        ///<inheritdoc />
        public override TSpinProperties GetObjectSpinProperties(int objectIndex, string objectPath)
        {
            CheckConnected();
            CheckRangeObjPath(objectPath, objectIndex, 1, ObjectCount, FlxParam.ObjectIndex);
            return ActiveSheetObject.Drawing.GetObjectSpinProperties(objectIndex - 1, objectPath);
        }

        ///<inheritdoc />
        public override void SetObjectSpinProperties(int objectIndex, string objectPath, TSpinProperties spinProps)
        {
            CheckConnected();
            CheckRangeObjPath(objectPath, objectIndex, 1, ObjectCount, FlxParam.ObjectIndex);
            ActiveSheetObject.Drawing.SetObjectSpinProperties(objectIndex - 1, objectPath, spinProps);
        }

        ///<inheritdoc />
        public override int GetObjectSpinValue(int objectIndex, string objectPath)
        {
            bool IsLinked;
            object v = GetLinkedValue(objectIndex, objectPath, out IsLinked);
            if (IsLinked)
            {
                if (v == null) return 0;
                if (v is TFlxFormulaErrorValue)
                {
                    if ((TFlxFormulaErrorValue)v == TFlxFormulaErrorValue.ErrNA) return 0;
                    //if it isn't n/a, its value doesn't matter. rb stays as is.
                }
                else
                {
                    double pd;
                    if (TBaseParsedToken.ExtToDouble(v, out pd)) //something like a string doesn't matter
                    {
                        if (pd <= 0 || pd >= int.MaxValue) return 0; // in this case it doesn't matter, all cbs are unselected. 
                        int p = (int)pd;
                        return p;
                    }
                }
            }

            return ActiveSheetObject.Drawing.GetObjectSpinValue(objectIndex - 1, objectPath);
        }

        ///<inheritdoc />
        public override void SetObjectSpinValue(int objectIndex, string objectPath, int value)
        {
            CheckConnected();
            CheckRangeObjPath(objectPath, objectIndex, 1, ObjectCount, FlxParam.ObjectIndex);
            ActiveSheetObject.Drawing.SetObjectSpinValue(objectIndex - 1, objectPath, value);

            SetLinkedValue(objectIndex, objectPath, true, value);
        }

        ///<inheritdoc />
        public override int AddComboBox(TClientAnchor anchor, string name, TCellAddress linkedCell, TCellAddressRange inputRange, int selectedItem)
        {
            if (anchor == null) FlxMessages.ThrowException(FlxErr.ErrNullParameter, "anchor");

            CheckConnected();
            int Result = ActiveSheetObject.Drawing.AddComboBox(anchor.Dec(), selectedItem, this, ActiveSheetObject, null, name) + 1;
            if (linkedCell != null) SetObjectLinkedCell(Result, null, linkedCell);
            if (inputRange != null) SetObjectInputRange(Result, null, inputRange);
            return Result;
        }

        ///<inheritdoc />
        public override int AddListBox(TClientAnchor anchor, string name, TCellAddress linkedCell,
            TCellAddressRange inputRange, TListBoxSelectionType selectionType, int selectedItem)
        {
            if (anchor == null) FlxMessages.ThrowException(FlxErr.ErrNullParameter, "anchor");

            CheckConnected();
            int Result = ActiveSheetObject.Drawing.AddListBox(anchor.Dec(), selectedItem, this, ActiveSheetObject, null, name, selectionType) + 1;
            if (linkedCell != null) SetObjectLinkedCell(Result, null, linkedCell);
            if (inputRange != null) SetObjectInputRange(Result, null, inputRange);
            return Result;
        }

        ///<inheritdoc />
        public override int AddButton(TClientAnchor anchor, TRichString text, string name, string macro)
        {
            if (anchor == null) FlxMessages.ThrowException(FlxErr.ErrNullParameter, "anchor");

            CheckConnected();
            TObjectProperties Props = new TObjectProperties(anchor, name);
            Props.FShapeLine = new TShapeLine(true, null);
            Props.Print = false;
            Props.FText = text;
            Props.FTextProperties = new TObjectTextProperties(true, THFlxAlignment.center, TVFlxAlignment.center, TTextRotation.Normal);
            int Result = ActiveSheetObject.Drawing.AddButton(anchor.Dec(), this, ActiveSheetObject, Props) + 1;

            SetObjectMacro(Result, null, macro);
            return Result;
        }

        ///<inheritdoc />
        public override int AddLabel(TClientAnchor anchor, TRichString text, string name)
        {
            if (anchor == null) FlxMessages.ThrowException(FlxErr.ErrNullParameter, "anchor");

            CheckConnected();
            int Result = ActiveSheetObject.Drawing.AddLabel(anchor.Dec(), text, this, ActiveSheetObject, null, name) + 1;
            return Result;
        }

        ///<inheritdoc />
        public override int AddSpinner(TClientAnchor anchor, string name, TCellAddress linkedCell, TSpinProperties spinProps)
        {
            if (anchor == null) FlxMessages.ThrowException(FlxErr.ErrNullParameter, "anchor");

            CheckConnected();
            int Result = ActiveSheetObject.Drawing.AddSpinner(anchor.Dec(), this, ActiveSheetObject, null, name) + 1;
            if (spinProps != null) SetObjectSpinProperties(Result, null, spinProps);
            if (linkedCell != null) SetObjectLinkedCell(Result, null, linkedCell);
            return Result;
        }

        ///<inheritdoc />
        public override int AddScrollBar(TClientAnchor anchor, string name, TCellAddress linkedCell, TSpinProperties spinProps)
        {
            if (anchor == null) FlxMessages.ThrowException(FlxErr.ErrNullParameter, "anchor");

            CheckConnected();
            int Result = ActiveSheetObject.Drawing.AddScrollBar(anchor.Dec(), this, ActiveSheetObject, null, name) + 1;
            if (spinProps != null) SetObjectSpinProperties(Result, null, spinProps);
            if (linkedCell != null) SetObjectLinkedCell(Result, null, linkedCell);
            return Result;
        }
        #endregion

#if (!COMPACTFRAMEWORK && !MONOTOUCH && !SILVERLIGHT)
        ///<inheritdoc />
        public override Image RenderObject(int objectIndex, real dpi, TShapeProperties shapeProperties, SmoothingMode aSmoothingMode, InterpolationMode aInterpolationMode, bool antiAliased, bool returnImage, Color BackgroundColor, out PointF origin, out RectangleF imageDimensions, out Size imageSizePixels)
        {
            CheckConnected();
            CheckRange(objectIndex, 1, ObjectCount, FlxParam.ObjectIndex);

            return FlexCel.Render.TDrawObjects.RenderObject(this, dpi, aSmoothingMode, antiAliased, aInterpolationMode,
                objectIndex, shapeProperties, returnImage, BackgroundColor, out origin, out imageDimensions, out imageSizePixels);
        }

        ///<inheritdoc />
        public override Image RenderCells(int row1, int col1, int row2, int col2, bool drawBackground, real dpi, SmoothingMode aSmoothingMode, InterpolationMode aInterpolationMode, bool antiAliased)
        {
            CheckConnected();
            CheckRowAndCol(row1, col1);
            CheckRowAndCol(row2, col2);
            return FlexCel.Render.FlexCelRender.RenderCells(this, row1, col1, row2, col2, drawBackground, dpi, aSmoothingMode, aInterpolationMode, antiAliased);
        }
   
        ///<inheritdoc />
        public override RectangleF CellRangeDimensions(int row1, int col1, int row2, int col2, bool includeMargins)
        {
            return FlexCel.Render.FlexCelRender.CalcCellRangeSize(this, row1, col1, row2, col2, includeMargins);
        }
#endif

        #endregion
        #region Object position
        ///<inheritdoc />
        public override void SendToBack(int objectIndex)
        {
            CheckConnected();
            CheckRange(objectIndex, 1, ObjectCount, FlxParam.ObjectIndex);
            ActiveSheetObject.Drawing.SendToBack(objectIndex - 1);
        }

        ///<inheritdoc />
        public override void BringToFront(int objectIndex)
        {
            CheckConnected();
            CheckRange(objectIndex, 1, ObjectCount, FlxParam.ObjectIndex);
            ActiveSheetObject.Drawing.BringToFront(objectIndex - 1);
        }

        ///<inheritdoc />
        public override void SendForward(int objectIndex)
        {
            CheckConnected();
            CheckRange(objectIndex, 1, ObjectCount, FlxParam.ObjectIndex);
            ActiveSheetObject.Drawing.SendForward(objectIndex - 1);
        }

        ///<inheritdoc />
        public override void SendBack(int objectIndex)
        {
            CheckConnected();
            CheckRange(objectIndex, 1, ObjectCount, FlxParam.ObjectIndex);
            ActiveSheetObject.Drawing.SendBack(objectIndex - 1);
        }
        #endregion
        #endregion

        #region Comments
        ///<inheritdoc />
        public override int CommentRowCount()
        {
            CheckConnected();
            return ActiveSheetObject.Notes.Count;
        }

        ///<inheritdoc />
        public override int CommentCountRow(int row)
        {
            CheckConnected();
            CheckRowAndCol(row, 1);
            if (row - 1 >= ActiveSheetObject.Notes.Count) return 0;
            return ActiveSheetObject.Notes[row - 1].Count;
        }

        ///<inheritdoc />
        public override TRichString GetCommentRow(int row, int commentIndex)
        {
            CheckConnected();
            CheckRowAndCol(row, 1);
            CheckRange(commentIndex, 1, CommentCountRow(row), FlxParam.CommentIndex);
            return ActiveSheetObject.Notes[row - 1][commentIndex - 1].GetText();
        }

        ///<inheritdoc />
        public override int GetCommentRowCol(int row, int commentIndex)
        {
            CheckConnected();
            CheckRowAndCol(row, 1);
            CheckRange(commentIndex, 1, CommentCountRow(row), FlxParam.CommentIndex);
            return ActiveSheetObject.Notes[row - 1][commentIndex - 1].Col + 1;
        }

        ///<inheritdoc />
        public override TRichString GetComment(int row, int col)
        {
            CheckConnected();
            CheckRowAndCol(row, col);
            int index = -1;
            if (row > ActiveSheetObject.Notes.Count) return new TRichString();
            if (ActiveSheetObject.Notes[row - 1].Find(col - 1, ref index))
                return ActiveSheetObject.Notes[row - 1][index].GetText();
            else return new TRichString();
        }

        ///<inheritdoc />
        public override void SetCommentRow(int row, int commentIndex, TRichString value, TImageProperties commentProperties)
        {
            CheckConnected();
            CheckRowAndCol(row, 1);
            CheckRange(commentIndex, 1, CommentCountRow(row), FlxParam.CommentIndex);
            if (value.Length == 0)
            {
                ActiveSheetObject.Notes[row - 1].Delete(commentIndex - 1);
                return;
            }

            if (commentProperties != null)
            {
                int col = ActiveSheetObject.Notes[row - 1][commentIndex].Col + 1;
                ActiveSheetObject.Notes[row - 1].Delete(commentIndex - 1);
                ActiveSheetObject.Notes.AddNewComment(row - 1, col - 1, value, String.Empty, ActiveSheetObject.Drawing, commentProperties.Dec(), this, ActiveSheetObject, false);
            }
            else
                ActiveSheetObject.Notes[row - 1][commentIndex - 1].SetText(value);

        }

        private void SetComment(int row, int col, TRichString value, string author, TImageProperties commentProperties, bool deleteEmpty)
        {
            CheckConnected();
            CheckRowAndCol(row, col);
            int commentIndex = -1;
            bool Found = (row - 1 < ActiveSheetObject.Notes.Count)
                && ActiveSheetObject.Notes[row - 1].Find(col - 1, ref commentIndex);
            commentIndex++; //Just to be consistent.
            if ((value.Length == 0) && deleteEmpty)
            {
                if (Found) ActiveSheetObject.Notes[row - 1].Delete(commentIndex - 1);
                return;
            }

            if (Found && !deleteEmpty) value = ActiveSheetObject.Notes[row - 1][commentIndex - 1].GetText();

            if (commentProperties != null)
            {
                if (Found) ActiveSheetObject.Notes[row - 1].Delete(commentIndex - 1);
                ActiveSheetObject.Notes.AddNewComment(row - 1, col - 1, value, author, ActiveSheetObject.Drawing, commentProperties.Dec(), this, ActiveSheetObject, false);
            }
            else
            {
                if (Found) ActiveSheetObject.Notes[row - 1][commentIndex - 1].SetText(value);
                else
                {
                    TCommentProperties stdProperties = TCommentProperties.GetDefaultProps(row, col, this);
                    ActiveSheetObject.Notes.AddNewComment(row - 1, col - 1, value, author, ActiveSheetObject.Drawing, stdProperties.Dec(), this, ActiveSheetObject, false);
                }
            }
        }


        ///<inheritdoc />
        public override void SetComment(int row, int col, TRichString value, string author, TImageProperties commentProperties)
        {
            SetComment(row, col, value, author, commentProperties, true);
        }

        ///<inheritdoc />
        public override TCommentProperties GetCommentPropertiesRow(int row, int commentIndex)
        {
            CheckConnected();
            CheckRowAndCol(row, 1);
            CheckRange(commentIndex, 1, CommentCountRow(row), FlxParam.CommentIndex);
            TCommentProperties Result = ActiveSheetObject.GetCommentProperties(this, row - 1, commentIndex - 1);
            if (Result != null) Result = (TCommentProperties)Result.Inc();
            return Result;
        }

        ///<inheritdoc />
        public override TCommentProperties GetCommentProperties(int row, int col)
        {
            CheckConnected();
            CheckRowAndCol(row, col);
            int commentIndex = -1;
            bool Found = (row - 1 < ActiveSheetObject.Notes.Count)
                && ActiveSheetObject.Notes[row - 1].Find(col - 1, ref commentIndex);
            if (Found) return GetCommentPropertiesRow(row, commentIndex + 1);
            else return null;
        }

        ///<inheritdoc />
        public override void SetCommentPropertiesRow(int row, int commentIndex, TImageProperties commentProperties)
        {
            SetCommentRow(row, commentIndex, GetCommentRow(row, commentIndex), commentProperties);
        }

        ///<inheritdoc />
        public override void SetCommentProperties(int row, int col, TImageProperties commentProperties)
        {
            SetComment(row, col, new TRichString(), string.Empty, commentProperties, false);
        }
        #endregion

        #region Cell operations.

        #region Fast Swap
        private XlsFile swXLSGlobal;
        private ExcelFile swGlobal2;
        private int swSheet;

        private ExcelFile sourceW
        {
            get
            {
                if (swXLSGlobal != null)
                {
                    swXLSGlobal.FastActiveSheet(swSheet);
                }
                else swGlobal2.ActiveSheet = swSheet;
                return swGlobal2;
            }
        }
        private int dwSheet;
        private XlsFile destW
        {
            get
            {
                FastActiveSheet(dwSheet);
                return this;
            }
        }
        #endregion

        private int ReadFormat(int[] FormatCache, int cxf)
        {
            int destXF = FormatCache[cxf + 1];
            if (destXF == -2)
            {
                TFlxFormat fm = sourceW.GetFormat(cxf);
                if (fm.ParentStyle != null)
                {
                    TFlxFormat stylefmt = destW.GetStyle(fm.ParentStyle);
                    if (stylefmt == null)
                    {
                        destW.SetStyle(fm.ParentStyle, sourceW.GetStyle(fm.ParentStyle));
                    }
                }
                fm.LinkedStyle.AutomaticChoose = sourceW != destW; // if copying to the same file, we will keep the format the same.
                destXF = destW.AddFormat(fm);
                FormatCache[cxf + 1] = destXF;
            }
            return destXF;
        }

        private void CopyObjects(TXlsCellRange sourceRange, int destRow, int destCol, TFlxInsertMode insertMode, TRangeCopyMode copyMode, ExcelFile sourceWorkbook, int sourceSheet)
        {

            swXLSGlobal = (sourceWorkbook as XlsFile);
            swGlobal2 = sourceWorkbook;
            swSheet = sourceSheet;
            dwSheet = FActiveSheet;
            int SaveSourceSheet = swGlobal2.ActiveSheet;
            int SaveActiveSheet = FActiveSheet;  //This is because workbook and this might be equal.

            bool IsFullRange = swXLSGlobal != null && destRow == 1 && destCol == 1 && sourceRange.Left <= 1 && sourceRange.Right >= FlxConsts.Max_Columns + 1 && sourceRange.Top <= 1 && sourceRange.Bottom >= FlxConsts.Max_Rows + 1;
            try
            {
                try
                {
                    if (IsFullRange)
                    {
                        FWorkbook.Sheets[SaveActiveSheet - 1].Columns.DefColWidth = swXLSGlobal.FWorkbook.Sheets[sourceSheet - 1].Columns.DefColWidth;
                        FWorkbook.Sheets[SaveActiveSheet - 1].DefRowHeight = swXLSGlobal.FWorkbook.Sheets[sourceSheet - 1].DefRowHeight;
                    }

                    //Cells and rows.
                    int[] FormatCache = new int[sourceW.FormatCount + 5];
                    for (int i = 0; i < FormatCache.Length; i++) { FormatCache[i] = -2; }

                    int MaxRow = sourceW.RowCount;
                    if (sourceRange.Bottom < MaxRow) MaxRow = sourceRange.Bottom;

                    //Columns. Should be set before rows.
                    if (insertMode == TFlxInsertMode.ShiftColRight || (sourceRange.Top <= 1 && sourceRange.Bottom >= FlxConsts.Max_Rows + 1))
                        for (int c = sourceRange.Left; c <= sourceRange.Right; c++)
                            if (!sourceW.IsNotFormattedCol(c))
                            {
                                int co, cxf, cw;
                                cw = sourceW.GetColWidth(c);
                                cxf = sourceW.GetColFormat(c);
                                co = sourceW.GetColOptions(c);
                                int keeptogether = sourceW.GetKeepColsTogether(c);

                                int destXF = ReadFormat(FormatCache, cxf);

                                int cd = destCol + c - sourceRange.Left;
                                destW.SetColWidth(cd, cw);
                                destW.SetColFormat(cd, destXF, false);
                                if (cw != sourceW.DefaultColWidth || cw != destW.DefaultColWidth) co |= 0x02; //the column has no standard width. 
                                destW.SetColOptions(cd, co);

                                if (cd + 1 <= FlxConsts.Max_Columns) destW.KeepColsTogether(cd, cd + 1, keeptogether, true);
                            }

                    for (int r = sourceRange.Top; r <= MaxRow; r++)
                    {
                        int rd = destRow + r - sourceRange.Top;
                        int MaxCol = sourceW.ColCountInRow(r);

                        //Rows. Should go before cells.
                        if (insertMode == TFlxInsertMode.ShiftRowDown || (sourceRange.Left <= 1 && sourceRange.Right >= FlxConsts.Max_Columns + 1))
                            if (!sourceW.IsEmptyRow(r))
                            {
                                int ro, rxf, rh;
                                rh = sourceW.GetRowHeight(r);
                                rxf = sourceW.GetRowFormat(r);
                                ro = sourceW.GetRowOptions(r);
                                int keeptogether = sourceW.GetKeepRowsTogether(r);

                                int destXF = ReadFormat(FormatCache, rxf);

                                destW.SetRowHeight(rd, rh);
                                destW.SetRowFormat(rd, destXF);
                                destW.SetRowOptions(rd, ro);
                                if (rd + 1 <= FlxConsts.Max_Rows) destW.KeepRowsTogether(rd, rd + 1, keeptogether, true);
                            }

                        for (int c0 = 1; c0 <= MaxCol; c0++)
                        {
                            int c = sourceW.ColFromIndex(r, c0);
                            if (c < sourceRange.Left) continue;
                            if (c > sourceRange.Right) break;

                            CopyCell(sourceWorkbook, swSheet, dwSheet, r, c, rd, destCol + c - sourceRange.Left, copyMode);

                            //Merged Cells.
                            if (!IsFullRange)
                            {
                                TXlsCellRange MergedRange = sourceW.CellMergedBounds(r, c);
                                if (MergedRange.Top < sourceRange.Top) MergedRange.Top = sourceRange.Top;
                                if (MergedRange.Left < sourceRange.Left) MergedRange.Left = sourceRange.Left;
                                if (MergedRange.Right > sourceRange.Right) MergedRange.Right = sourceRange.Right;
                                if (MergedRange.Bottom > sourceRange.Bottom) MergedRange.Bottom = sourceRange.Bottom;

                                if (MergedRange.ColCount > 1 || MergedRange.RowCount > 1)
                                {
                                    destW.MergeCells(MergedRange.Top + destRow - sourceRange.Top, MergedRange.Left + destCol - sourceRange.Left, MergedRange.Bottom + destRow - sourceRange.Top, MergedRange.Right + destCol - sourceRange.Left);
                                }
                            }
                        }
                    }


                    //Comments
                    int MaxCommentRow = sourceW.CommentRowCount();
                    if (sourceRange.Bottom < MaxCommentRow) MaxCommentRow = sourceRange.Bottom;
                    for (int r = sourceRange.Top; r <= MaxCommentRow; r++)
                        for (int i = 1; i <= sourceW.CommentCountRow(r); i++)
                        {
                            TRichString comment = sourceW.GetCommentRow(r, i);
                            int c = sourceW.GetCommentRowCol(r, i);
                            TImageProperties ip = sourceW.GetCommentPropertiesRow(r, i);
                            if (c < sourceRange.Left || c > sourceRange.Right) continue;
                            if (ip != null)
                            {
                                ip.Anchor.Row1 += destRow - sourceRange.Top;
                                ip.Anchor.Row2 += destRow - sourceRange.Top;
                                ip.Anchor.Col1 += destCol - sourceRange.Left;
                                ip.Anchor.Col2 += destCol - sourceRange.Left;
                                destW.SetComment(destRow + r - sourceRange.Top, destCol + c - sourceRange.Left, comment, string.Empty, ip, false);
                            }
                        }

                    if (swXLSGlobal != null && FWorkbook.IsWorkSheet(SaveActiveSheet - 1) && swXLSGlobal.FWorkbook.IsWorkSheet(sourceSheet - 1))
                    {
                        //Drawings
                        TSheetInfo SheetInfo = new TSheetInfo(sourceSheet - 1, sourceSheet - 1, dwSheet - 1, swXLSGlobal.FWorkbook.Globals,
                            FWorkbook.Globals, swXLSGlobal.FWorkbook.Sheets[sourceSheet - 1], FWorkbook.Sheets[dwSheet - 1], SemiAbsoluteReferences);
                        FWorkbook.WorkSheets(SaveActiveSheet - 1).Drawing.CopyObjectsFrom(copyMode == TRangeCopyMode.AllIncludingDontMoveAndSizeObjects, destRow - sourceRange.Top, destCol - sourceRange.Left, swXLSGlobal.FWorkbook.WorkSheets(sourceSheet - 1).Drawing, sourceRange.Dec(), SheetInfo);
                        //HyperLinks
                        FWorkbook.WorkSheets(SaveActiveSheet - 1).HLinks.CopyObjectsFrom(swXLSGlobal.FWorkbook.WorkSheets(sourceSheet - 1).HLinks, sourceRange.Dec(), destRow - sourceRange.Top, destCol - sourceRange.Left, SheetInfo);

                        //AutoFilters
                        TXlsCellRange AutoFilter = GetAutoFilterRange();
                        if (AutoFilter == null) //no AutoFilters in the sheet.
                        {
                            AutoFilter = swXLSGlobal.GetAutoFilterRange();
                            if (AutoFilter != null) SetAutoFilter(AutoFilter.Top + destRow - sourceRange.Top, AutoFilter.Left + destCol - sourceRange.Left, AutoFilter.Right + destCol - sourceRange.Left);
                        }

                        //Page breaks

                        if (IsFullRange)
                        {
                            //CFs, MergedCells and data validations, when copying a sheet.
                            FWorkbook.WorkSheets(SaveActiveSheet - 1).MergedCells.CopyFrom(swXLSGlobal.FWorkbook.WorkSheets(sourceSheet - 1).MergedCells, SheetInfo);
                            FWorkbook.WorkSheets(SaveActiveSheet - 1).ConditionalFormats.CopyFrom(swXLSGlobal.FWorkbook.WorkSheets(sourceSheet - 1).ConditionalFormats, SheetInfo);
                            FWorkbook.WorkSheets(SaveActiveSheet - 1).DataValidation.CopyFrom(
                                swXLSGlobal.FWorkbook.WorkSheets(sourceSheet - 1).DataValidation, FWorkbook.WorkSheets(SaveActiveSheet - 1).Drawing, SheetInfo);
                            FWorkbook.WorkSheets(SaveActiveSheet - 1).DataValidation.ClearObjId();
                        }
                        else
                        {
                            //Currently we will only copy this if the sheet is the same.
                            if (swXLSGlobal == this && SaveActiveSheet == swSheet)
                            {
                                int RowCount = IsByRows(insertMode)? 1: 0;
                                int ColCount = RowCount == 1 ? 0 : 1;
                                //Conditional Formats
                                FWorkbook.WorkSheets(SaveActiveSheet - 1).ConditionalFormats.InsertAndCopyRange(sourceRange.Dec(), destRow - 1, destCol - 1, RowCount, ColCount, copyMode, insertMode, SheetInfo);
                                //Data Validation
                                FWorkbook.WorkSheets(SaveActiveSheet - 1).DataValidation.InsertAndCopyRange(sourceRange.Dec(), destRow - 1, destCol - 1, RowCount, ColCount, copyMode, insertMode, SheetInfo);
                            }
                        }
                    }
                }

                finally
                {
                    if (swXLSGlobal != null) swXLSGlobal.FastActiveSheet(SaveSourceSheet);
                    else swGlobal2.ActiveSheet = SaveSourceSheet;
                }
            }
            finally
            {
                FastActiveSheet(SaveActiveSheet);
            }
        }

        private bool IsByRows(TFlxInsertMode insertMode)
        {
            switch (insertMode)
            {
                case TFlxInsertMode.NoneRight:
                case TFlxInsertMode.ShiftColRight:
                case TFlxInsertMode.ShiftRangeRight:
                    return false;

            }
            return true;
        }

        private void CopyHeaderImages(TWorkbook sourceWorkbook, int sourceSheet)
        {
            //Header and Footer images. The text for header and footers is copied in CopySheetMisc.
            TSheet sourceTSheet = sourceWorkbook.Sheets[sourceSheet];

            foreach (THeaderAndFooterKind hkind in TCompactFramework.EnumGetValues(typeof(THeaderAndFooterKind)))
            {
                foreach (THeaderAndFooterPos hpos in TCompactFramework.EnumGetValues(typeof(THeaderAndFooterPos)))
                {
                    TXlsImgType ImgType = TXlsImgType.Unknown;
                    byte[] img = null;
                    using (MemoryStream ms = new MemoryStream())
                    {
                        sourceTSheet.HeaderImages.GetDrawingFromStream(hkind, hpos, ms, ref ImgType);
                        if (ms.Length > 0) img = ms.ToArray();
                    }
                    if (img != null)
                    {
                        THeaderOrFooterImageProperties ImgProperties = sourceTSheet.HeaderImages.GetHeaderOrFooterImageProperties(hkind, hpos);
                        FWorkbook.Sheets[FActiveSheet - 1].HeaderImages.AssignHeaderOrFooterDrawing(hkind, hpos, img, ImgType, ImgProperties);
                    }
                }
            }
        }


        ///<inheritdoc />
        public override void InsertAndCopyRange(TXlsCellRange sourceRange, int destRow, int destCol, int destCount, TFlxInsertMode insertMode, TRangeCopyMode copyMode, ExcelFile sourceWorkbook, int sourceSheet, TExcelObjectList ObjectsInRange)
        {
            CheckConnected();
            TXlsCellRange newCells = sourceRange.Dec();
            CheckRowAndCol(sourceRange.Top, sourceRange.Right);
            CheckRowAndCol(sourceRange.Bottom, sourceRange.Left);
            CheckRowAndCol(destRow, destCol);
            NeedsRecalc = true;

            if (newCells.Top > newCells.Bottom) Swap(ref newCells.Top, ref newCells.Bottom);
            if (newCells.Left > newCells.Right) Swap(ref newCells.Left, ref newCells.Right);
            TRangeCopyMode realCopyMode = copyMode;
            if (sourceWorkbook != null)
            {
                CheckRange(sourceSheet, 1, sourceWorkbook.SheetCount, FlxParam.SourceSheet);
                if (sourceWorkbook == this && sourceSheet == FActiveSheet)
                {
                    sourceWorkbook = null;
                }
                else
                {
                    //Only insert. We have to copy on another place.
                    realCopyMode = TRangeCopyMode.None;
                }
            }

            switch (insertMode)
            {
                case TFlxInsertMode.ShiftRangeDown:
                    FWorkbook.InsertAndCopyRange(FActiveSheet - 1, newCells, destRow - 1, destCol - 1, destCount, 0, realCopyMode, insertMode, SemiAbsoluteReferences, ObjectsInRange);
                    break;
                case TFlxInsertMode.ShiftRangeRight:
                    FWorkbook.InsertAndCopyRange(FActiveSheet - 1, newCells, destRow - 1, destCol - 1, 0, destCount, realCopyMode, insertMode, SemiAbsoluteReferences, ObjectsInRange);
                    break;
                case TFlxInsertMode.ShiftRowDown:
                    newCells.Left = 0;
                    newCells.Right = FlxConsts.Max_Columns;
                    FWorkbook.InsertAndCopyRange(FActiveSheet - 1, newCells, destRow - 1, destCol - 1, destCount, 0, realCopyMode, insertMode, SemiAbsoluteReferences, ObjectsInRange);
                    break;
                case TFlxInsertMode.ShiftColRight:
                    newCells.Top = 0;
                    newCells.Bottom = FlxConsts.Max_Rows;
                    FWorkbook.InsertAndCopyRange(FActiveSheet - 1, newCells, destRow - 1, destCol - 1, 0, destCount, realCopyMode, insertMode, SemiAbsoluteReferences, ObjectsInRange);
                    break;
                case TFlxInsertMode.NoneDown:
                case TFlxInsertMode.NoneRight:
                    int rCount = 1;
                    int cCount = 1;
                    if (insertMode == TFlxInsertMode.NoneDown) rCount = destCount; else cCount = destCount;
                    TXlsCellRange newSourceCells = new TXlsCellRange(destRow - 1, destCol - 1, destRow - 1 + newCells.RowCount * rCount - 1, destCol - 1 + newCells.ColCount * cCount - 1);
                    if (copyMode == TRangeCopyMode.Formats)
                    {
                        FWorkbook.Sheets[FActiveSheet - 1].ClearFormats(newSourceCells);
                    }
                    else
                    {
                        FWorkbook.Sheets[FActiveSheet - 1].ClearRange(newSourceCells);
                    }

                    if (sourceWorkbook == null) //local copy, but we will do it as if it was remote.
                    {
                        sourceWorkbook = this;
                        sourceSheet = FActiveSheet;
                    }
                    break;
            }

            //We are copying from another workbook.
            if (sourceWorkbook != null && copyMode != TRangeCopyMode.None)
            {
                if (insertMode == TFlxInsertMode.ShiftRangeDown || insertMode == TFlxInsertMode.ShiftRowDown || insertMode == TFlxInsertMode.NoneDown)
                {
                    int sourceRowLen = newCells.RowCount;
                    for (int k = 0; k < destCount * (sourceRowLen); k += sourceRowLen)
                        CopyObjects(newCells.Inc(), destRow + k, destCol, insertMode, copyMode, sourceWorkbook, sourceSheet);
                }
                else
                {
                    int sourceColLen = newCells.ColCount;
                    for (int k = 0; k < destCount * (sourceColLen); k += sourceColLen)
                        CopyObjects(newCells.Inc(), destRow, destCol + k, insertMode, copyMode, sourceWorkbook, sourceSheet);
                }
            }

            //Autofilter needs to be regen after this:
            //TXlsCellRange AutoFilter = GetAutoFilterRange();
            //if (AutoFilter != null) SetAutoFilter(AutoFilter);

        }


        ///<inheritdoc />
        public override void DeleteRange(TXlsCellRange cellRange, TFlxInsertMode insertMode)
        {
            DeleteRange(FActiveSheet, FActiveSheet, cellRange, insertMode);
        }

        ///<inheritdoc />
        public override void DeleteRange(int sheet1, int sheet2, TXlsCellRange cellRange, TFlxInsertMode insertMode)
        {
            CheckConnected();
            TXlsCellRange newCells = cellRange.Dec();
            CheckRowAndCol(cellRange.Top, cellRange.Right);
            CheckRowAndCol(cellRange.Bottom, cellRange.Left);
            NeedsRecalc = true;

            if (newCells.Top > newCells.Bottom) Swap(ref newCells.Top, ref newCells.Bottom);
            if (newCells.Left > newCells.Right) Swap(ref newCells.Left, ref newCells.Right);

            for (int sheet = sheet1; sheet <= sheet2; sheet++)
            {
                switch (insertMode)
                {
                    case TFlxInsertMode.ShiftRangeDown:
                    case TFlxInsertMode.ShiftRangeRight:
                    case TFlxInsertMode.NoneDown:
                    case TFlxInsertMode.NoneRight:
                        FWorkbook.DeleteRange(sheet - 1, newCells, insertMode);
                        break;
                    case TFlxInsertMode.ShiftRowDown:
                        newCells.Left = 0;
                        newCells.Right = FlxConsts.Max_Columns;
                        FWorkbook.DeleteRange(sheet - 1, newCells, insertMode);
                        break;
                    case TFlxInsertMode.ShiftColRight:
                        newCells.Top = 0;
                        newCells.Bottom = FlxConsts.Max_Rows;
                        FWorkbook.DeleteRange(sheet - 1, newCells, insertMode);
                        break;
                }
            }
        }

        private static void SaveHorizontalRanges(ref TXlsCellRange[] NewRange, TXlsCellRange CellRange, int newRow, int newCol, ref int drow2, ref int dcol2)
        {
            //On those cases source range will not move when inserting cells.
            if (newCol > CellRange.Left || newRow > CellRange.Bottom || newRow + CellRange.RowCount - 1 < CellRange.Top)
            {
                NewRange[0] = CellRange;
                return;
            }

            if (newRow < CellRange.Top)
            {
                NewRange[0] = new TXlsCellRange(CellRange.Top,
                    CellRange.Right + 1, newRow + CellRange.RowCount - 1, CellRange.Right + CellRange.ColCount);
                NewRange[1] = new TXlsCellRange(newRow + CellRange.RowCount,
                    CellRange.Left, CellRange.Bottom, CellRange.Right);
                drow2 = NewRange[0].RowCount;

                return;
            }
            if (newRow == CellRange.Top)
            {
                NewRange[0] = CellRange.Offset(newRow, CellRange.Right + 1);
                return;
            }

            NewRange[0] = new TXlsCellRange(CellRange.Top,
                CellRange.Left, newRow - 1, CellRange.Right);
            NewRange[1] = new TXlsCellRange(newRow,
                CellRange.Right + 1, CellRange.Bottom, CellRange.Right + CellRange.ColCount);
            drow2 = NewRange[0].RowCount;

            return;

        }

        private static void SaveVerticalRanges(ref TXlsCellRange[] NewRange, TXlsCellRange CellRange, int newRow, int newCol, ref int drow2, ref int dcol2)
        {
            //On those cases source range will not move when inserting cells.
            if (newRow > CellRange.Top || newCol > CellRange.Right || newCol + CellRange.ColCount - 1 < CellRange.Left)
            {
                NewRange[0] = CellRange;
                return;
            }

            if (newCol < CellRange.Left)
            {
                NewRange[0] = new TXlsCellRange(CellRange.Bottom + 1,
                    CellRange.Left, CellRange.Bottom + CellRange.RowCount, newCol + CellRange.ColCount - 1);
                NewRange[1] = new TXlsCellRange(CellRange.Top, newCol + CellRange.ColCount,
                    CellRange.Bottom, CellRange.Right);
                dcol2 = NewRange[0].ColCount;

                return;
            }
            if (newCol == CellRange.Left)
            {
                NewRange[0] = CellRange.Offset(CellRange.Bottom + 1, newCol);
                return;
            }

            NewRange[0] = new TXlsCellRange(CellRange.Top,
                CellRange.Left, CellRange.Bottom, newCol - 1);
            NewRange[1] = new TXlsCellRange(CellRange.Bottom + 1, newCol,
                CellRange.Bottom + CellRange.RowCount, CellRange.Right);
            dcol2 = NewRange[0].ColCount;

            return;

        }

        ///<inheritdoc />
        public override void MoveRange(TXlsCellRange cellRange, int newRow, int newCol, TFlxInsertMode insertMode)
        {
            CheckConnected();
            TXlsCellRange newCells = cellRange.Dec();
            CheckRowAndCol(cellRange.Top, cellRange.Right);
            CheckRowAndCol(cellRange.Bottom, cellRange.Left);
            CheckRowAndCol(newRow, newCol);

            if (cellRange.Top == newRow && cellRange.Left == newCol) return; //nothing to move, please move along.

            if (newCells.Top > newCells.Bottom) Swap(ref newCells.Top, ref newCells.Bottom);
            if (newCells.Left > newCells.Right) Swap(ref newCells.Left, ref newCells.Right);

            if (newRow + newCells.RowCount - 1 > FlxConsts.Max_Rows + 1) XlsMessages.ThrowException(XlsErr.ErrMoveRangeOutsideBounds);
            if (newCol + newCells.ColCount - 1 > FlxConsts.Max_Columns + 1) XlsMessages.ThrowException(XlsErr.ErrMoveRangeOutsideBounds);

            NeedsRecalc = true;

            switch (insertMode)
            {
                case TFlxInsertMode.ShiftRowDown:
                    newCells.Left = 0;
                    newCells.Right = FlxConsts.Max_Columns;
                    break;
                case TFlxInsertMode.ShiftColRight:
                    newCells.Top = 0;
                    newCells.Bottom = FlxConsts.Max_Rows;
                    break;
            }

            TXlsCellRange[] NewRange = new TXlsCellRange[2];
            int drow2 = 0; int dcol2 = 0;

            if (insertMode != TFlxInsertMode.NoneDown && insertMode != TFlxInsertMode.NoneRight)
            {
                if (newCells.HasRow(newRow - 1) && newCells.HasCol(newCol - 1)) XlsMessages.ThrowException(XlsErr.ErrMoveRangesCanNotIntersect);
                if ((insertMode == TFlxInsertMode.ShiftColRight || insertMode == TFlxInsertMode.ShiftRangeRight))
                {
                    SaveHorizontalRanges(ref NewRange, newCells, newRow - 1, newCol - 1, ref drow2, ref dcol2);
                }

                if ((insertMode == TFlxInsertMode.ShiftRowDown || insertMode == TFlxInsertMode.ShiftRangeDown))
                {
                    SaveVerticalRanges(ref NewRange, newCells, newRow - 1, newCol - 1, ref drow2, ref dcol2);
                }

                InsertAndCopyRange(cellRange, newRow, newCol, 1, insertMode, TRangeCopyMode.None);
            }
            else
            {
                NewRange[0] = newCells;
            }

            if (NewRange[0] != null) FWorkbook.MoveRange(FActiveSheet - 1, NewRange[0], newRow - 1, newCol - 1);
            if (NewRange[1] != null) FWorkbook.MoveRange(FActiveSheet - 1, NewRange[1], newRow - 1 + drow2, newCol - 1 + dcol2);

            if (insertMode != TFlxInsertMode.NoneDown && insertMode != TFlxInsertMode.NoneRight)
            {
                if (NewRange[0] != null) FWorkbook.DeleteRange(FActiveSheet - 1, NewRange[0], insertMode);
                if (NewRange[1] != null) FWorkbook.DeleteRange(FActiveSheet - 1, NewRange[1], insertMode);
            }

            RestoreObjectSizes();

        }


        #endregion

        #region Data Validation
        ///<inheritdoc />
        public override void ClearDataValidation()
        {
            CheckConnected();
            ActiveSheetObject.DataValidation.Clear();
        }

        ///<inheritdoc />
        public override void ClearDataValidation(TXlsCellRange range)
        {
            CheckConnected();
            ActiveSheetObject.DataValidation.ClearRange(range.Dec(), true);
        }

        ///<inheritdoc />
        public override void AddDataValidation(TXlsCellRange range, TDataValidationInfo validationInfo)
        {
            CheckConnected();
            ActiveSheetObject.DataValidation.AddRange(range.Dec(), validationInfo, ActiveSheetObject.Cells.CellList, true, false);
        }


        ///<inheritdoc />
        public override TDataValidationInfo GetDataValidation(int row, int col)
        {
            CheckConnected();
            CheckRowAndCol(row, col);
            return ActiveSheetObject.DataValidation.GetDataValidation(row - 1, col - 1, ActiveSheetObject.Cells.CellList, false);
        }


        #region Indexed Data Validation
        ///<inheritdoc />
        public override int DataValidationCount
        {
            get
            {
                CheckConnected();
                return ActiveSheetObject.DataValidation.Count;
            }
        }

        ///<inheritdoc />
        public override TDataValidationInfo GetDataValidationInfo(int index)
        {
            CheckConnected();
            CheckRange(index, 1, DataValidationCount, FlxParam.DataValidationIndex);
            return ActiveSheetObject.DataValidation.GetDataValidation(index - 1, ActiveSheetObject.Cells.CellList, false);
        }

        ///<inheritdoc />
        public override TXlsCellRange[] GetDataValidationRanges(int index)
        {
            CheckConnected();
            CheckRange(index, 1, DataValidationCount, FlxParam.DataValidationIndex);
            return ActiveSheetObject.DataValidation.GetDataValidationRange(index - 1, true);
        }

        #endregion

        #endregion

        #region HyperLinks

        ///<inheritdoc />
        public override int HyperLinkCount
        {
            get
            {
                CheckConnected();
                return ActiveSheetObject.HLinks.Count;
            }
        }

        ///<inheritdoc />
        public override THyperLink GetHyperLink(int hyperLinkIndex)
        {
            CheckConnected();
            CheckRange(hyperLinkIndex, 1, HyperLinkCount, FlxParam.HyperLinkIndex);
            return ActiveSheetObject.HLinks[hyperLinkIndex - 1].GetProperties();
        }

        ///<inheritdoc />
        public override void SetHyperLink(int hyperLinkIndex, THyperLink value)
        {
            CheckConnected();
            CheckRange(hyperLinkIndex, 1, HyperLinkCount, FlxParam.HyperLinkIndex);
            ActiveSheetObject.HLinks[hyperLinkIndex - 1].SetProperties(value);
        }

        ///<inheritdoc />
        public override TXlsCellRange GetHyperLinkCellRange(int hyperLinkIndex)
        {
            CheckConnected();
            CheckRange(hyperLinkIndex, 1, HyperLinkCount, FlxParam.HyperLinkIndex);
            return ActiveSheetObject.HLinks[hyperLinkIndex - 1].GetCellRange().Inc();
        }

        private TXlsCellRange ExpandAndDecRange(TXlsCellRange cellRange)
        {
            TXlsCellRange cellRangeA = ActiveSheetObject.CellMergedBounds(cellRange.Top - 1, cellRange.Left - 1);
            TXlsCellRange cellRangeB = cellRangeA;
            if (cellRange.Top != cellRange.Bottom || cellRange.Left != cellRange.Right) cellRangeB = ActiveSheetObject.CellMergedBounds(cellRange.Bottom - 1, cellRange.Right - 1);

            cellRangeA.Left = Math.Min(cellRangeA.Left, cellRangeB.Left);
            cellRangeA.Right = Math.Max(cellRangeA.Right, cellRangeB.Right);
            cellRangeA.Top = Math.Min(cellRangeA.Top, cellRangeB.Top);
            cellRangeA.Bottom = Math.Max(cellRangeA.Bottom, cellRangeB.Bottom);

            return cellRangeA;
        }

        ///<inheritdoc />
        public override void SetHyperLinkCellRange(int hyperLinkIndex, TXlsCellRange cellRange)
        {
            CheckConnected();
            CheckRange(hyperLinkIndex, 1, HyperLinkCount, FlxParam.HyperLinkIndex);

            ActiveSheetObject.HLinks[hyperLinkIndex - 1].SetCellRange(ExpandAndDecRange(cellRange));
        }

        ///<inheritdoc />
        public override void AddHyperLink(TXlsCellRange cellRange, THyperLink value)
        {
            CheckConnected();
            ActiveSheetObject.HLinks.Add(THLinkRecord.CreateNew(ExpandAndDecRange(cellRange), value));
        }

        ///<inheritdoc />
        public override void DeleteHyperLink(int hyperLinkIndex)
        {
            CheckConnected();
            CheckRange(hyperLinkIndex, 1, HyperLinkCount, FlxParam.HyperLinkIndex);
            ActiveSheetObject.HLinks.Delete(hyperLinkIndex - 1);
        }
        #endregion

        #region Group and Outline
        ///<inheritdoc />
        public override int GetRowOutlineLevel(int row)
        {
            CheckConnected();
            CheckRowAndCol(row, 1);
            return ActiveSheetObject.GetRowOutlineLevel(row - 1);
        }

        ///<inheritdoc />
        public override void SetRowOutlineLevel(int firstRow, int lastRow, int level)
        {
            CheckConnected();
            CheckRowAndCol(firstRow, 1);
            CheckRowAndCol(lastRow, 1);
            CheckRange(level, 0, 7, FlxParam.OutlineLevel);

            for (int i = firstRow - 1; i <= lastRow - 1; i++)
                ActiveSheetObject.SetRowOutlineLevel(i, level);
        }

        ///<inheritdoc />
        public override int GetColOutlineLevel(int col)
        {
            CheckConnected();
            CheckRowAndCol(col, 1);
            return ActiveSheetObject.GetColOutlineLevel(col - 1);
        }

        ///<inheritdoc />
        public override void SetColOutlineLevel(int firstCol, int lastCol, int level)
        {
            CheckConnected();
            CheckRowAndCol(1, firstCol);
            CheckRowAndCol(1, lastCol);
            CheckRange(level, 0, 7, FlxParam.OutlineLevel);

            for (int i = firstCol - 1; i <= lastCol - 1; i++)
                ActiveSheetObject.SetColOutlineLevel(i, level);
        }

        ///<inheritdoc />
        public override bool OutlineSummaryRowsBelowDetail
        {
            get
            {
                CheckConnected();
                return ActiveSheetObject.SheetGlobals.WsBool.RowSumsBelow;
            }
            set
            {
                CheckConnected();
                ActiveSheetObject.SheetGlobals.WsBool.RowSumsBelow = value;
            }
        }

        ///<inheritdoc />
        public override bool OutlineSummaryColsRightToDetail
        {
            get
            {
                CheckConnected();
                return ActiveSheetObject.SheetGlobals.WsBool.ColSumsRight;
            }
            set
            {
                CheckConnected();
                ActiveSheetObject.SheetGlobals.WsBool.ColSumsRight = value;
            }
        }

        ///<inheritdoc />
        public override bool OutlineAutomaticStyles
        {
            get
            {
                CheckConnected();
                return ActiveSheetObject.SheetGlobals.WsBool.ApplyStyles;
            }
            set
            {
                CheckConnected();
                ActiveSheetObject.SheetGlobals.WsBool.ApplyStyles = value;
            }
        }

        ///<inheritdoc />
        public override void CollapseOutlineRows(int level, TCollapseChildrenMode collapseChildren, int firstRow, int lastRow)
        {
            CheckConnected();
            if (level < 1) level = 1;
            if (level > 8) level = 8;

            CheckRowAndCol(firstRow, 1);
            CheckRowAndCol(lastRow, 1);

            if (lastRow < firstRow) return;
            if (lastRow > RowCount) lastRow = RowCount;

            bool PlusGoesDown = OutlineSummaryRowsBelowDetail;
            for (int row = firstRow; row <= lastRow; row++)
            {
                bool IsNode = false;
                if (PlusGoesDown)
                {
                    if (row - 2 > 0)
                    {
                        IsNode = ActiveSheetObject.GetRowOutlineLevel(row - 1) < ActiveSheetObject.GetRowOutlineLevel(row - 2);
                    }
                }
                else
                {
                    if (row - 1 > 0)
                    {
                        IsNode = ActiveSheetObject.GetRowOutlineLevel(row) > ActiveSheetObject.GetRowOutlineLevel(row - 1);
                    }
                }

                ActiveSheetObject.CollapseRows(row - 1, level, collapseChildren, IsNode);
            }
        }

        ///<inheritdoc />
        public override void CollapseOutlineCols(int level, TCollapseChildrenMode collapseChildren, int firstCol, int lastCol)
        {
            CheckConnected();
            if (level < 1) level = 1;
            if (level > 8) level = 8;

            CheckRowAndCol(1, firstCol);
            CheckRowAndCol(1, lastCol);

            if (lastCol < firstCol) return;
            int cc = ColCount;
            if (lastCol > cc) lastCol = cc;

            bool PlusGoesRight = OutlineSummaryColsRightToDetail;
            for (int col = firstCol; col <= lastCol; col++)
            {
                bool IsNode = false;
                if (PlusGoesRight)
                {
                    if (col - 2 > 0)
                    {
                        IsNode = ActiveSheetObject.GetColOutlineLevel(col - 1) < ActiveSheetObject.GetColOutlineLevel(col - 2);
                    }
                }
                else
                {
                    if (col - 1 > 0)
                    {
                        IsNode = ActiveSheetObject.GetColOutlineLevel(col) > ActiveSheetObject.GetColOutlineLevel(col - 1);
                    }
                }
                ActiveSheetObject.CollapseCols(col - 1, level, collapseChildren, IsNode);
            }
        }

        ///<inheritdoc />
        public override bool IsOutlineNodeRow(int row)
        {
            CheckConnected();
            CheckRowAndCol(row, 1);

            bool PlusGoesDown = OutlineSummaryRowsBelowDetail;
            bool Result = false;
            if (PlusGoesDown)
            {
                if (row - 2 > 0)
                {
                    Result = ActiveSheetObject.GetRowOutlineLevel(row - 1) < ActiveSheetObject.GetRowOutlineLevel(row - 2);
                }
            }
            else
            {
                if (row - 1 > 0)
                {
                    Result = ActiveSheetObject.GetRowOutlineLevel(row) > ActiveSheetObject.GetRowOutlineLevel(row - 1);
                }
            }

            return Result;
        }

        ///<inheritdoc />
        public override bool IsOutlineNodeCol(int col)
        {
            CheckConnected();
            CheckRowAndCol(1, col);

            bool PlusGoesRight = OutlineSummaryColsRightToDetail;
            bool Result = false;
            if (PlusGoesRight)
            {
                if (col - 2 > 0)
                {
                    Result = ActiveSheetObject.GetColOutlineLevel(col - 1) < ActiveSheetObject.GetColOutlineLevel(col - 2);
                }
            }
            else
            {
                if (col - 1 > 0)
                {
                    Result = ActiveSheetObject.GetColOutlineLevel(col) > ActiveSheetObject.GetColOutlineLevel(col - 1);
                }
            }

            return Result;
        }

        ///<inheritdoc />
        public override bool IsOutlineNodeCollapsedRow(int row)
        {
            if (!IsOutlineNodeRow(row)) return false;
            bool PlusGoesDown = OutlineSummaryRowsBelowDetail;
            int RowLevel = GetRowOutlineLevel(row);
            int PlusDir = PlusGoesDown ? -1 : +1;

            int r = row + PlusDir;
            int aCount = RowCount;
            while (r > 0 && r <= aCount)
            {
                int rl = GetRowOutlineLevel(r);
                if (rl <= RowLevel) return true;

                if (!ActiveSheetObject.GetRowHidden(r - 1)) return false;

                r += PlusDir;
            }

            return true;
        }

        ///<inheritdoc />
        public override bool IsOutlineNodeCollapsedCol(int col)
        {
            if (!IsOutlineNodeCol(col)) return false;
            bool PlusGoesRight = OutlineSummaryColsRightToDetail;
            int ColLevel = GetColOutlineLevel(col);
            int PlusDir = PlusGoesRight ? -1 : +1;

            int c = col + PlusDir;
            int aCount = ColCount;
            while (c > 0 && c <= aCount)
            {
                int cl = GetColOutlineLevel(c);
                if (cl <= ColLevel) return true;

                if (!ActiveSheetObject.GetColHidden(c - 1)) return false;

                c += PlusDir;
            }

            return true;
        }


        private void ExpandGroup(bool ColGroup, bool PlusGoesDown, int RowLevel, ref int r, bool Process)
        {
            int CMask = ColGroup ? 0x1000 : 0x10;
            int PlusDir = PlusGoesDown ? -1 : 1;

            if (Process)
            {
                if (ColGroup)
                {
                    ActiveSheetObject.SetColOptions(r - 1, ActiveSheetObject.GetColOptions(r - 1) & ~CMask);
                    ActiveSheetObject.SetColHidden(r - 1, false);
                }
                else
                {
                    ActiveSheetObject.SetRowOptions(r - 1, ActiveSheetObject.GetRowOptions(r - 1) & ~CMask);
                    ActiveSheetObject.SetRowHidden(r - 1, false);
                }
            }

            r += PlusDir;
            int aCount = ColGroup ? ColCount : RowCount;
            while (r > 0 && r <= aCount)
            {
                int rl = ColGroup ? GetColOutlineLevel(r) : GetRowOutlineLevel(r);
                if (rl <= RowLevel) return;
                if (rl == RowLevel + 1)
                {
                    if (Process)
                    {
                        if (ColGroup) ActiveSheetObject.SetColHidden(r - 1, false); else ActiveSheetObject.SetRowHidden(r - 1, false);
                    }
                    r += PlusDir;
                }
                else
                {
                    int PrevRowOptions = ColGroup ? ActiveSheetObject.GetColOptions(r - 1 - PlusDir) : ActiveSheetObject.GetRowOptions(r - 1 - PlusDir);
                    ExpandGroup(ColGroup, PlusGoesDown, rl - 1, ref r, Process && (PrevRowOptions & CMask) == 0);
                }
            }
        }



        ///<inheritdoc />
        public override void CollapseOutlineNodeRow(int row, bool collapse)
        {
            if (!IsOutlineNodeRow(row)) return;
            bool PlusGoesDown = OutlineSummaryRowsBelowDetail;
            int RowLevel = GetRowOutlineLevel(row);

            if (!collapse)
            {
                int r = row;
                ExpandGroup(false, PlusGoesDown, RowLevel, ref r, true);
                return;
            }

            if (PlusGoesDown)
            {
                for (int r = row; r > 0; r--)
                {
                    bool IsNode = IsOutlineNodeRow(r);
                    int rl = GetRowOutlineLevel(r);
                    if (r < row && rl <= RowLevel) return;

                    ActiveSheetObject.CollapseRows(r - 1, RowLevel + 1, TCollapseChildrenMode.DontModify, IsNode);
                }
            }
            else
            {
                for (int r = row; r <= RowCount; r++)
                {
                    bool IsNode = IsOutlineNodeRow(r);
                    int rl = GetRowOutlineLevel(r);
                    if (r > row && rl <= RowLevel) return;

                    ActiveSheetObject.CollapseRows(r - 1, RowLevel + 1, TCollapseChildrenMode.DontModify, IsNode);
                }
            }
        }

        ///<inheritdoc />
        public override void CollapseOutlineNodeCol(int col, bool collapse)
        {
            if (!IsOutlineNodeCol(col)) return;
            bool PlusGoesRight = OutlineSummaryColsRightToDetail;
            int ColLevel = GetColOutlineLevel(col);

            if (!collapse)
            {
                int c = col;
                ExpandGroup(true, PlusGoesRight, ColLevel, ref c, true);
                return;
            }

            if (PlusGoesRight)
            {
                for (int c = col; c > 0; c--)
                {
                    bool IsNode = IsOutlineNodeCol(c);
                    int cl = GetColOutlineLevel(c);
                    if (c < col && cl <= ColLevel) return;
                    ActiveSheetObject.CollapseCols(c - 1, ColLevel + 1, TCollapseChildrenMode.DontModify, IsNode);
                }
            }
            else
            {
                int aColCount = ColCount;
                for (int c = col; c <= aColCount; c++)
                {
                    bool IsNode = IsOutlineNodeCol(c);
                    int cl = GetColOutlineLevel(c);
                    if (c > col && cl <= ColLevel) return;
                    ActiveSheetObject.CollapseCols(c - 1, ColLevel + 1, TCollapseChildrenMode.DontModify, IsNode);
                }
            }
        }


        #endregion

        #region Protection
        internal override void SetModifyPassword(string modifyPassword, bool recommendReadOnly, string reservingUser)
        {
            CheckConnected();
            if (modifyPassword == null || modifyPassword.Length == 0)
            {
                FWorkbook.Globals.FileEncryption.WriteProt = null;
                FWorkbook.Globals.FileEncryption.FileSharing = null;
            }
            else
            {
                FWorkbook.Globals.FileEncryption.WriteProt = new TWriteProtRecord();
                FWorkbook.Globals.FileEncryption.FileSharing = new TFileSharingRecord(recommendReadOnly, modifyPassword, reservingUser, false);
            }
        }

        internal override bool HasModifyPassword
        {
            get
            {
                CheckConnected();
                return (FWorkbook.Globals.FileEncryption.WriteProt != null) &&
                    (FWorkbook.Globals.FileEncryption.FileSharing != null);
            }
        }

        internal override bool RecommendReadOnly
        {
            get
            {
                CheckConnected();
                return FWorkbook.Globals.FileEncryption.FileSharing != null && FWorkbook.Globals.FileEncryption.FileSharing.RecommendReadOnly;
            }
            set
            {
                CheckConnected();
                if (FWorkbook.Globals.FileEncryption.FileSharing == null)
                    FWorkbook.Globals.FileEncryption.FileSharing = new TFileSharingRecord(value, string.Empty, string.Empty, false);
                else
                    FWorkbook.Globals.FileEncryption.FileSharing.RecommendReadOnly = value;
            }
        }

        #region Workbook
        internal override void SetWorkbookProtection(string workbookPassword, TWorkbookProtectionOptions workbookProtectionOptions)
        {
            CheckConnected();
            if (FWorkbook.Globals.WorkbookProtection.Password == null)
                FWorkbook.Globals.WorkbookProtection.Password = new TPasswordRecord();
            FWorkbook.Globals.WorkbookProtection.Password.SetPassword(workbookPassword);

            this.WorkbookProtectionOptions = workbookProtectionOptions;
        }

        internal override bool HasWorkbookPassword
        {
            get
            {
                CheckConnected();
                TWorkbookProtection wp = FWorkbook.Globals.WorkbookProtection;
                return wp.Password != null && BitOps.GetWord(wp.Password.Data, 0) != 0 &&
                    ((wp.Protect != null && wp.Protect.Protected) || (wp.WindowProtect != null && wp.WindowProtect.Protected));

            }
        }

        internal override TWorkbookProtectionOptions WorkbookProtectionOptions
        {
            get
            {
                CheckConnected();
                TWorkbookProtectionOptions Result = new TWorkbookProtectionOptions();
                Result.Structure = FWorkbook.Globals.WorkbookProtection.Protect != null && FWorkbook.Globals.WorkbookProtection.Protect.Protected;
                Result.Window = FWorkbook.Globals.WorkbookProtection.WindowProtect != null && FWorkbook.Globals.WorkbookProtection.WindowProtect.Protected;
                return Result;
            }
            set
            {
                CheckConnected();

                if (FWorkbook.Globals.WorkbookProtection.Protect == null) FWorkbook.Globals.WorkbookProtection.Protect = new TProtectRecord();
                FWorkbook.Globals.WorkbookProtection.Protect.Protected = value == null ? false : value.Structure;

                if (FWorkbook.Globals.WorkbookProtection.WindowProtect == null) FWorkbook.Globals.WorkbookProtection.WindowProtect = new TWindowProtectRecord();
                FWorkbook.Globals.WorkbookProtection.WindowProtect.Protected = value == null ? false : value.Window;
            }
        }
        #endregion

        #region Shared Workbook
        internal override void SetSharedWorkbookProtection(string sharedWorkbookPassword, TSharedWorkbookProtectionOptions sharedWorkbookProtectionOptions)
        {
            CheckConnected();
            if (FWorkbook.Globals.WorkbookProtection.Prot4RevPass == null)
                FWorkbook.Globals.WorkbookProtection.Prot4RevPass = new TProt4RevPassRecord();
            FWorkbook.Globals.WorkbookProtection.Prot4RevPass.SetPassword(sharedWorkbookPassword);

            this.SharedWorkbookProtectionOptions = sharedWorkbookProtectionOptions;
        }

        internal override bool HasSharedWorkbookPassword
        {
            get
            {
                CheckConnected();
                TWorkbookProtection wp = FWorkbook.Globals.WorkbookProtection;
                return wp.Prot4RevPass != null && BitOps.GetWord(wp.Prot4RevPass.Data, 0) != 0 &&
                    (wp.Prot4Rev != null && wp.Prot4Rev.Protected);
            }

        }

        internal override TSharedWorkbookProtectionOptions SharedWorkbookProtectionOptions
        {
            get
            {
                CheckConnected();
                TSharedWorkbookProtectionOptions Result = new TSharedWorkbookProtectionOptions();
                Result.SharingWithTrackChanges = FWorkbook.Globals.WorkbookProtection.Prot4Rev != null && FWorkbook.Globals.WorkbookProtection.Prot4Rev.Protected;
                return Result;
            }
            set
            {
                CheckConnected();

                if (FWorkbook.Globals.WorkbookProtection.Prot4Rev == null) FWorkbook.Globals.WorkbookProtection.Prot4Rev = new TProt4RevRecord();
                FWorkbook.Globals.WorkbookProtection.Prot4Rev.Protected = value == null ? false : value.SharingWithTrackChanges; ;
            }
        }
        #endregion

        #region Sheet

        internal override void SetSheetProtection(string sheetPassword, TSheetProtectionOptions sheetProtectionOptions)
        {
            CheckConnected();
            TSheetProtection Sp = FWorkbook.Sheets[FActiveSheet - 1].SheetProtection;

            if (Sp.Password == null)
                Sp.Password = new TPasswordRecord();
            Sp.Password.SetPassword(sheetPassword);

            this.SheetProtectionOptions = sheetProtectionOptions;

        }

        internal override bool HasSheetPassword
        {
            get
            {
                CheckConnected();
                return FWorkbook.Sheets[FActiveSheet - 1].SheetProtection.Password != null &&
                    BitOps.GetWord(FWorkbook.Sheets[FActiveSheet - 1].SheetProtection.Password.Data, 0) != 0;
            }
        }

        internal override TSheetProtectionOptions SheetProtectionOptions
        {
            get
            {
                CheckConnected();
                return FWorkbook.Sheets[FActiveSheet - 1].GetSheetProtectionOptions();
            }
            set
            {
                CheckConnected();
                FWorkbook.Sheets[FActiveSheet - 1].SetSheetProtectionOptions(value);
            }
        }

        #endregion

        internal override string WriteAccess
        {
            get
            {
                CheckConnected();
                TWriteAccessRecord Result = FWorkbook.Globals.FileEncryption.WriteAccess;
                if (Result == null) return String.Empty;
                return Result.UserName;
            }
            set
            {
                CheckConnected();
                TWriteAccessRecord Result = FWorkbook.Globals.FileEncryption.WriteAccess;
                if (Result == null) return;
                Result.UserName = value;
            }
        }

        #endregion

        #region Cell selection
        ///<inheritdoc />
        public override void SelectCells(TXlsCellRange[] cellRange)
        {
            CheckConnected();
            if (cellRange == null || cellRange.Length == 0) cellRange = new TXlsCellRange[] { new TXlsCellRange(1, 1, 1, 1) };

            TXlsCellRange[] newRange = new TXlsCellRange[cellRange.Length];
            for (int i = 0; i < cellRange.Length; i++)
            {
                if (cellRange[i] == null) FlxMessages.ThrowException(FlxErr.ErrInvalidRange, FlxConvert.ToString(null));
                CheckRowAndCol(cellRange[i].Top, cellRange[i].Left);
                CheckRowAndCol(cellRange[i].Bottom, cellRange[i].Right);
                newRange[i] = cellRange[i].Offset(cellRange[i].Top - 1, cellRange[i].Left - 1);

            }

            ActiveSheetObject.Window.Selection.Select(ActiveSheetObject.GetActivePaneForSelection(), newRange, -1, -1, 0);

        }

        ///<inheritdoc />
        public override TXlsCellRange[] GetSelectedCells()
        {
            CheckConnected();
            TXlsCellRange[] cellRange = ActiveSheetObject.Window.Selection.GetSelection(ActiveSheetObject.GetActivePaneForSelection());
            if (cellRange == null) return new TXlsCellRange[] { new TXlsCellRange(1, 1, 1, 1) };
            TXlsCellRange[] newRange = new TXlsCellRange[cellRange.Length];
            for (int i = 0; i < cellRange.Length; i++)
            {
                if (cellRange[i] == null) FlxMessages.ThrowException(FlxErr.ErrInvalidRange, FlxConvert.ToString(null));
                newRange[i] = cellRange[i].Offset(cellRange[i].Top + 1, cellRange[i].Left + 1);
            }

            return newRange;
        }

        ///<inheritdoc />
        public override void ScrollWindow(TPanePosition panePosition, int row, int col)
        {
            CheckConnected();
            CheckRowAndCol(row, col);
            FWorkbook.Sheets[FActiveSheet - 1].ScrollWindow(panePosition, row - 1, col - 1);
        }

        ///<inheritdoc />
        public override TCellAddress GetWindowScroll(TPanePosition panePosition)
        {
            CheckConnected();
            TCellAddress ca = FWorkbook.Sheets[FActiveSheet - 1].GetWindowScroll(panePosition);
            ca.Row++;
            ca.Col++;
            return ca;
        }




        #endregion

        #region Freeze Panes
        ///<inheritdoc />
        public override void FreezePanes(TCellAddress cell)
        {
            CheckConnected();
            if (cell != null)
            {
                CheckRowAndCol(cell.Row, cell.Col);
                FWorkbook.Sheets[FActiveSheet - 1].FreezePanes(cell.Row - 1, cell.Col - 1);
            }
            else
            {
                FWorkbook.Sheets[FActiveSheet - 1].FreezePanes(0, 0);
            }
        }

        ///<inheritdoc />
        public override TCellAddress GetFrozenPanes()
        {
            CheckConnected();
            TCellAddress Result = FWorkbook.Sheets[FActiveSheet - 1].GetFrozenPanes();
            Result.Col++;
            Result.Row++;
            return Result;
        }

        ///<inheritdoc />
        public override void SplitWindow(int xOffset, int yOffset)
        {
            CheckConnected();
            FWorkbook.Sheets[FActiveSheet - 1].SplitWindow(xOffset, yOffset);
        }

        ///<inheritdoc />
        public override TPoint GetSplitWindow()
        {
            CheckConnected();
            return FWorkbook.Sheets[FActiveSheet - 1].GetSplitWindow();
        }


        #endregion

        #region Document Properties
        internal override object GetStandardProperty(TPropertyId PropertyId)
        {
            bool IsExtended = (int)PropertyId >= 0xFFFF;
            TOle2Properties Props = IsExtended ? Ole2ExtProperties : Ole2StdProperties;
            if (IsExtended) PropertyId -= 0xFFFF;

            if (Props == null)
            {
                Props = new TOle2Properties();
                if (IsExtended)
                {
                    Ole2ExtProperties = Props;
                }
                else
                {
                    Ole2StdProperties = Props;
                }

                using (MemoryStream MemFile = new MemoryStream(OtherStreams))
                {
                    using (TOle2File DataStream = new TOle2File(MemFile))
                    {
                        if (IsExtended)
                        {
                            DataStream.SelectStream(XlsConsts.DocumentPropertiesStringExtended);
                        }
                        else
                        {
                            DataStream.SelectStream(XlsConsts.DocumentPropertiesString);
                        }
                        Props.Load(DataStream);
                    }
                }
            }
            return Props.GetValue((UInt32)PropertyId);

        }
        #endregion

        #region Charts
        ///<inheritdoc />
        public override int ChartCount
        {
            get
            {
                CheckConnected();
                TSheet Sheet = FWorkbook.Sheets[FActiveSheet - 1];
                return Sheet.ChartCount;
            }
        }

        ///<inheritdoc />
        public override ExcelChart GetChart(int objectIndex, string objectPath)
        {
            CheckConnected();
            TFlxChart Chart = null;
            if (objectPath == null) Chart = ActiveSheetObject as TFlxChart;
            if (Chart != null) return new XlsChart(this, Chart);

            CheckRangeObjPath(objectPath, objectIndex, 1, ObjectCount, FlxParam.ObjectIndex);
            Chart = ActiveSheetObject.Drawing.GetChart(objectIndex - 1, objectPath);
            if (Chart == null) return null;
            return new XlsChart(this, Chart);
        }



        #endregion

        #region Sort, Search and Replace
        ///<inheritdoc />
        public override TCellAddress Find(object value, TXlsCellRange Range, TCellAddress Start, bool ByRows, bool CaseInsensitive, bool SearchInFormulas, bool WholeCellContents)
        {
            TSearch SearchObject = new TSearch(OptionsDates1904, value, CaseInsensitive, SearchInFormulas, WholeCellContents);
            TSearchAndReplace.Search(this, Range, Start, ByRows, SearchObject);
            return SearchObject.Cell;
        }

        ///<inheritdoc />
        public override int Replace(object oldValue, object newValue, TXlsCellRange Range, bool CaseInsensitive, bool SearchInFormulas, bool WholeCellContents)
        {
            TReplace SearchObject = new TReplace(OptionsDates1904, oldValue, newValue, CaseInsensitive, SearchInFormulas, WholeCellContents);
            TSearchAndReplace.Search(this, Range, null, false, SearchObject);
            return SearchObject.ReplaceCount;
        }

        ///<inheritdoc />
        public override void Sort(TXlsCellRange Range, bool ByRows, int[] Keys, TSortOrder[] SortOrder, IComparer Comparer)
        {
            TSortRange.Sort(this, Range, ByRows, Keys, SortOrder, Comparer);
        }


        #endregion

        #region Options
        ///<inheritdoc />
        public override bool OptionsDates1904
        {
            get
            {
                return FWorkbook.Globals.Dates1904;
            }
            set
            {
                CheckConnected();
                FWorkbook.Globals.Dates1904 = value;
            }
        }

        ///<inheritdoc />
        public override bool OptionsR1C1
        {
            get
            {
                CheckConnected();
                return FWorkbook.R1C1;
            }
            set
            {
                CheckConnected();
                FWorkbook.R1C1 = value;
            }
        }

        ///<inheritdoc />
        public override bool OptionsSaveExternalLinkValues
        {
            get
            {
                CheckConnected();
                return FWorkbook.Globals.SaveExternalLinkValues;
            }
            set
            {
                CheckConnected();
                FWorkbook.Globals.SaveExternalLinkValues = value;
            }
        }


        ///<inheritdoc />
        public override bool OptionsPrecisionAsDisplayed
        {
            get
            {
                CheckConnected();
                return FWorkbook.Globals.PrecisionAsDisplayed;
            }
            set
            {
                CheckConnected();
                FWorkbook.Globals.PrecisionAsDisplayed = value;
            }
        }

        ///<inheritdoc />
        public override int OptionsMultithreadRecalc
        {
            get
            {
                CheckConnected();
                return FWorkbook.Globals.MultithreadRecalc;
            }
            set
            {
                CheckConnected();
                FWorkbook.Globals.MultithreadRecalc = value;
            }
        }

        ///<inheritdoc />
        public override bool OptionsForceFullRecalc
        {
            get
            {
                CheckConnected();
                return FWorkbook.Globals.ForceFullRecalc;
            }
            set
            {
                CheckConnected();
                FWorkbook.Globals.ForceFullRecalc = value;
            }
        }

        ///<inheritdoc />
        public override bool OptionsAutoCompressPictures
        {
            get
            {
                CheckConnected();
                return FWorkbook.Globals.AutoCompressPictures;
            }
            set
            {
                CheckConnected();
                FWorkbook.Globals.AutoCompressPictures = value;
            }
        }

        ///<inheritdoc />
        public override bool OptionsCheckCompatibility
        {
            get
            {
                CheckConnected();
                return FWorkbook.Globals.CheckCompatibility;
            }
            set
            {
                CheckConnected();
                FWorkbook.Globals.CheckCompatibility = value;
            }
        }

        ///<inheritdoc />
        public override bool OptionsBackup
        {
            get
            {
                CheckConnected();
                return FWorkbook.Globals.Backup;
            }
            set
            {
                CheckConnected();
                FWorkbook.Globals.Backup = value;
            }
        }

        ///<inheritdoc />
        public override TSheetCalcMode OptionsRecalcMode
        {
            get
            {
                CheckConnected();
                return FWorkbook.Globals.CalcOptions.CalcMode;
            }
            set
            {
                CheckConnected();
                FWorkbook.Globals.CalcOptions.CalcMode = value;
            }
        }

        #endregion

        #region AutoFilter
        ///<inheritdoc />
        public override void SetAutoFilter(int row, int col1, int col2)
        {
            CheckConnected();
            CheckRowAndCol(row, col1);
            CheckRowAndCol(row, col2);
            ActiveSheetObject.SetAutoFilter(FActiveSheet - 1, row - 1, col1 - 1, col2 - 1);
        }

        ///<inheritdoc />
        public override void RemoveAutoFilter()
        {
            CheckConnected();
            ActiveSheetObject.RemoveAutoFilter();
        }

        ///<inheritdoc />
        public override bool HasAutoFilter()
        {
            CheckConnected();
            return ActiveSheetObject.HasAutoFilter();
        }

        ///<inheritdoc />
        public override bool HasAutoFilter(int row, int col)
        {
            CheckConnected();
            CheckRowAndCol(row, col);
            return ActiveSheetObject.HasAutoFilter(FActiveSheet - 1, row - 1, col - 1);
        }

        ///<inheritdoc />
        public override TXlsCellRange GetAutoFilterRange()
        {
            CheckConnected();
            TXlsCellRange Result = ActiveSheetObject.GetAutoFilterRange(FActiveSheet - 1);
            if (Result == null) return null;
            return Result.Inc();
        }

        #endregion

        #region Recalculation
        ///<inheritdoc />
        public override TRecalcMode RecalcMode
        {
            get
            {
                return FRecalcMode;
            }
            set
            {
                if (FWorkbook != null && FWorkbook.Sheets.Count > 0)
                    FlxMessages.ThrowException(FlxErr.ErrCantChangeRecalcMode);

                FRecalcMode = value;
            }
        }

        ///<inheritdoc />
        public override void Recalc(bool forced)
        {
            CheckConnected();
            if (Workspace != null) Workspace.Recalc(forced);
            else
            {
                FRecalculating = true;
                try
                {
                    if (RecalcForced && (RecalcMode != TRecalcMode.Smart || NeedsRecalc))
                        FWorkbook.ForceAutoRecalc();

                    if (forced || PendingRecalc())
                    {
                        FWorkbook.CleanFlags();
                        FWorkbook.Recalc(this, null);
                    }
                    NeedsRecalc = false;
                }
                finally
                {
                    FRecalculating = false;
                }
            }
        }

        internal override void InternalRecalc(bool forced, TUnsupportedFormulaList Ufl)
        {
            if (forced || PendingRecalc())
            {
                FWorkbook.Recalc(this, Ufl);
            }
            NeedsRecalc = false;
        }


        ///<inheritdoc />
        public override TUnsupportedFormulaList RecalcAndVerify()
        {
            CheckConnected();
            if (Workspace != null) return Workspace.RecalcAndVerify();
            else
            {
                try
                {
                    FRecalculating = true;
                    TUnsupportedFormulaList Ufl = new TUnsupportedFormulaList();
                    FWorkbook.CleanFlags();
                    FWorkbook.Recalc(this, Ufl);
                    return Ufl;

                }
                finally
                {
                    FRecalculating = false;
                }
            }
        }


        ///<inheritdoc />
        public override bool RecalcForced
        {
            get
            {
                return FRecalcForced;
            }
            set
            {
                FRecalcForced = value;
            }
        }

        ///<inheritdoc />
        public override object RecalcCell(int sheet, int row, int col, bool forced)
        {
            CheckConnected();
            CheckRowAndCol(row, col);
            CheckSheet(sheet);

            if (forced || PendingRecalc())
            {
                FWorkbook.CleanFlags();
                GetCellValueAndRecalc(sheet, row, col, new TCalcState(), new TCalcStack());
            }

            int XF = -1;
            TFormula f = GetCellValue(sheet, row, col, ref XF) as TFormula;
            if (f == null) return null;
            return f.Result;
        }

        ///<inheritdoc />
        public override object RecalcExpression(string expression, bool forced)
        {
            CheckConnected();
            if (forced || PendingRecalc())
            {
                FWorkbook.CleanFlags();
            }

            return ActiveSheetObject.Cells.CellList.CalcExpression(this, FActiveSheet, expression);
        }

        internal override bool NeedsRecalc
        {
            get
            {
                return FNeedsRecalc;
            }
            set
            {
                FNeedsRecalc = value;
                if (RecalcMode == TRecalcMode.OnEveryChange && value)
                    Recalc();
            }
        }

        internal override bool Recalculating
        {
            get
            {
                return FRecalculating;
            }
        }

        internal override bool ReorderCalcChain(int SheetBase1, IFormulaRecord i1, IFormulaRecord i2)
        {
            TSheet Ws = FWorkbook.Sheets[SheetBase1 - 1];
            return Ws.Cells.CellList.FormulaCache.ReorderCalcChain(i1, i2);
        }

        internal override ExcelFile GetSupportingFile(string fileName)
        {
            if (Workspace == null) return null;
            ExcelFile Result = Workspace[fileName];
            if (Result != null) return Result;
            Result = Workspace[Path.GetFileName(fileName)];
            if (Result != null) return Result;

            return Workspace.GetLinkedFile(fileName);
        }

        internal override void SetRecalculating(bool value)
        {
            FRecalculating = value;
        }

        internal override void CleanFlags()
        {
            FWorkbook.CleanFlags();
        }


        #endregion

        #region What-If Tables
        ///<inheritdoc />

        public override TCellAddress[] GetWhatIfTableList()
        {
            CheckConnected();
            return ActiveSheetObject.Cells.CellList.GetTables();
        }

        ///<inheritdoc />
        public override TXlsCellRange GetWhatIfTable(int sheet, int row, int col, out TCellAddress rowInputCell, out TCellAddress colInputCell)
        {
            rowInputCell = null;
            colInputCell = null;

            CheckConnected();
            CheckRowAndCol(row, col);
            CheckSheet(sheet);

            TSheet Ws = FWorkbook.Sheets[sheet - 1];
            return Ws.Cells.CellList.GetTable(sheet - 1, row - 1, col - 1, out rowInputCell, out colInputCell);
        }

        ///<inheritdoc />
        public override void SetWhatIfTable(TXlsCellRange range, TCellAddress rowInputCell, TCellAddress colInputCell)
        {
            CheckConnected();
            FActiveSheetObject.Cells.CellList.AddWhatIfTable(range.Dec(), rowInputCell, colInputCell, -1);
        }
        #endregion

        #region External Links
        
        ///<inheritdoc />
        public override int LinkCount
        {
            get 
            {
                CheckConnected();
                int d = FWorkbook.Globals.References.LocalSupBook < 0 ? 0 : 1;
                return FWorkbook.Globals.References.Supbooks.Count - d;
            }
        }

        ///<inheritdoc />
        public override string GetLink(int index)
        {
            int i = index - 1;
            CheckRange(index, 1, LinkCount, FlxParam.Index);
            int lsb = FWorkbook.Globals.References.LocalSupBook;
            if (i >= lsb) i++;
            return FWorkbook.Globals.References.Supbooks[i].BookName();
        }

        ///<inheritdoc />
        public override void SetLink(int index, string value)
        {
            int i = index - 1;
            CheckRange(index, 1, LinkCount, FlxParam.Index);
            int lsb = FWorkbook.Globals.References.LocalSupBook;
            if (i >= lsb) i++;
            FWorkbook.Globals.References.Supbooks[i].SetBookName(value);
            
        }
        #endregion

        #region Misc
        ///<inheritdoc />
        public override void ConvertFormulasToValues(bool onlyExternal)
        {
            CheckConnected();
            Recalc(false);
            FWorkbook.ConvertFormulasToValues(FActiveSheet - 1, onlyExternal);
        }

        ///<inheritdoc />
        public override void ConvertExternalNamesToRefErrors()
        {
            CheckConnected();
            Recalc(false);
            FWorkbook.ConvertExternalNamesToRefErrors();
        }

        ///<inheritdoc />
        public override real HeightCorrection
        {
            get
            {
                return FHeightCorrection;
            }
            set
            {
                if (value > 0 && value < 2) FHeightCorrection = value;
            }
        }

        ///<inheritdoc />
        public override TFileFormats FileFormatWhenOpened
        {
            get { return FFileFormatWhenOpened; }
        }

        ///<inheritdoc />
        public override double Linespacing
        {
            get
            {
                return FLinespacing;
            }
            set
            {
                FLinespacing = value;
            }
        }

        ///<inheritdoc />
        public override bool IgnoreFormulaText
        {
            get
            {
                return FIgnoreFormulaText;
            }
            set
            {
                FIgnoreFormulaText = value;
            }
        }


        ///<inheritdoc />
        public override real WidthCorrection
        {
            get
            {
                return FWidthCorrection;
            }
            set
            {
                if (value > 0 && value < 2) FWidthCorrection = value;
            }
        }




        #endregion

        #region Custom Formulas
        ///<inheritdoc />
        public override void AddUserDefinedFunction(TUserDefinedFunctionScope scope, TUserDefinedFunctionLocation location, TUserDefinedFunction userFunction)
        {
            switch (scope)
            {
                case TUserDefinedFunctionScope.Global:
                    lock (GlobalFormulaFunctionAccess)
                    {
                        GlobalFormulaFunctions.Add(new TUserDefinedFunctionContainer(location, userFunction));
                    }
                    break;
                case TUserDefinedFunctionScope.Local:
                    LocalFormulaFunctions.Add(new TUserDefinedFunctionContainer(location, userFunction));
                    break;
                default:
                    break;
            }
        }

        ///<inheritdoc />
        public override void ClearUserDefinedFunctions(TUserDefinedFunctionScope scope)
        {
            switch (scope)
            {
                case TUserDefinedFunctionScope.Global:
                    lock (GlobalFormulaFunctionAccess)
                    {
                        GlobalFormulaFunctions.Clear();
                    }
                    break;
                case TUserDefinedFunctionScope.Local:
                    LocalFormulaFunctions.Clear();
                    break;
                default:
                    break;
            }
        }

        ///<inheritdoc />
        public override object EvaluateUserDefinedFunction(string functionName, TUdfEventArgs arguments, object[] parameters)
        {
            TUserDefinedFunctionContainer FunctionContainer = GetUserDefinedFunction(functionName);
            if (FunctionContainer == null) return TFlxFormulaErrorValue.ErrName;

            return FunctionContainer.Function.Evaluate(arguments, parameters);
        }

        ///<inheritdoc />
        public override bool IsDefinedFunction(string functionName, out TUserDefinedFunctionLocation location)
        {
            TUserDefinedFunctionContainer FunctionContainer = GetUserDefinedFunction(functionName);

            location = TUserDefinedFunctionLocation.External;
            if (FunctionContainer != null) location = FunctionContainer.Location;
            return FunctionContainer != null;
        }

        internal override TUserDefinedFunctionContainer GetUserDefinedFunction(string functionName)
        {
            TUserDefinedFunctionContainer FunctionContainer = BuiltInFormulaFunctions.GetValue(functionName);
            if (FunctionContainer == null)
            {
                FunctionContainer = LocalFormulaFunctions.GetValue(functionName);
                if (FunctionContainer == null)
                {
                    lock (GlobalFormulaFunctionAccess)
                    {
                        FunctionContainer = GlobalFormulaFunctions.GetValue(functionName);
                    }
                }
            }
            return FunctionContainer;
        }

        internal override TUserDefinedFunctionContainer GetUserDefinedFunctionFromDisplayName(string DisplayFunctionName)
        {
            TUserDefinedFunctionContainer FunctionContainer = BuiltInFormulaFunctions.GetValueFromDisplayName(DisplayFunctionName);
            if (FunctionContainer == null)
            {
                FunctionContainer = LocalFormulaFunctions.GetValueFromDisplayName(DisplayFunctionName);
                if (FunctionContainer == null)
                {
                    lock (GlobalFormulaFunctionAccess)
                    {
                        FunctionContainer = GlobalFormulaFunctions.GetValueFromDisplayName(DisplayFunctionName);
                    }
                }
            }
            return FunctionContainer;
        }

        internal override void EnsureAddInExternalName(string functionName, out int externSheet, out int externName)
        {
            FWorkbook.Globals.References.AddAddinExternalName(functionName, out externSheet, out externName);
        }

        internal override void EnsureAddInInternalName(string functionName, bool AddErrorDataToFormula, out int nameIndex)
        {
            nameIndex = FWorkbook.Globals.Names.AddAddin(functionName, AddErrorDataToFormula);
        }

        internal override int EnsureExternName(int ExternSheet, string Name)
        {
            return FWorkbook.Globals.References.EnsureExternName(ExternSheet, Name);
        }

        internal override void AddUnsupported(TUnsupportedFormulaErrorType ErrorType, string FuncName)
        {
            if (FUnsupportedFormulaList != null)
                FUnsupportedFormulaList.Add(
                    new TUnsupportedFormula(ErrorType,
                    FUnsupportedFormulaList.CellAddress, FuncName, ActiveFileName));
        }

        internal override void SetUnsupportedFormulaCellAddress(TCellAddress aCellAddress)
        {
            if (FUnsupportedFormulaList != null) FUnsupportedFormulaList.CellAddress = aCellAddress;
        }

        internal override void SetUnsupportedFormulaList(TUnsupportedFormulaList Ufl)
        {
            FUnsupportedFormulaList = Ufl;
        }

        #region ExternSheet
        private static void GetExternalFileLink(string ExternSheet, out string FileName, out string SheetNames)
        {
            int wbpos = ExternSheet.LastIndexOf(TBaseFormulaParser.ft(TFormulaToken.fmWorkbookClose));
            if (wbpos > 0) //This is a normal ref like c:\test[text.xls]sheet1!a1 .When there is no sheet, like in names, we don't use [], which makes it impossible to know if this is an external ref or a sheeet ref.
            {
                FileName = ExternSheet.Substring(0, wbpos);
                FileName = FileName.Replace(TBaseFormulaParser.fts(TFormulaToken.fmWorkbookOpen), "");
                SheetNames = ExternSheet.Substring(wbpos + 1);
                return;
            }

            //No external name detected
            FileName = null;
            SheetNames = ExternSheet;
        }

        private void GetExternSheetIndexes(string ExternSheet, out int Sheet1, out int Sheet2, bool ReadingXlsx)
        {
            string[] Sheets = ExternSheet.Split(TFormulaMessages.TokenChar(TFormulaToken.fmRangeSep));
            if (Sheets.Length > 2 || Sheets.Length < 1)
                FlxMessages.ThrowException(FlxErr.ErrInvalidRef, ExternSheet);

            if (!FWorkbook.FindSheet(Sheets[0], out Sheet1))
            {
                if (!ReadingXlsx)
                {
                    FlxMessages.ThrowException(FlxErr.ErrInvalidSheet, Sheets[0]);
                }
                Sheet1 = 0xFFFF; //invalid ref
            }

            Sheet2 = Sheet1;
            if (Sheets.Length > 1)
                if (!FWorkbook.FindSheet(Sheets[1], out Sheet2))
                    FlxMessages.ThrowException(FlxErr.ErrInvalidSheet, Sheets[1]);
        }

        private static void SplitSheetNames(string SheetNames, out string Sheet1, out string Sheet2)
        {
            string[] Sheets = SheetNames.Split(TFormulaMessages.TokenChar(TFormulaToken.fmRangeSep));
            if (Sheets.Length > 2 || Sheets.Length < 1)
                FlxMessages.ThrowException(FlxErr.ErrInvalidRef, SheetNames);

            Sheet1 = Sheets[0];
            if (Sheets.Length > 1) Sheet2 = Sheets[1]; else Sheet2 = Sheet1;
        }

        internal override int GetExternSheet(string ExternSheet, bool ReadingXlsx)
        {
            bool IsLocal;
            int Sheet1;
            return GetExternSheet(ExternSheet, true, ReadingXlsx, out IsLocal, out Sheet1);
        }

        internal override int GetExternSheet(string ExternSheet, bool IsCellReference, bool ReadingXlsx, out bool IsLocal, out int Sheet1)
        {
            IsLocal = false;
            Sheet1 = -1;
            int Sheet2;
            string FileName; string SheetNames;
            GetExternalFileLink(ExternSheet, out FileName, out SheetNames);

            if (FileName == null || FileName.Trim().Length == 0) //reference is to the same file
            {
                IsLocal = true;

                GetExternSheetIndexes(ExternSheet, out Sheet1, out Sheet2, ReadingXlsx);
                if (!IsCellReference && Sheet1 != Sheet2) FlxMessages.ThrowException(FlxErr.ErrInvalidRef, SheetNames);

                return FWorkbook.Globals.References.AddSheet(FWorkbook.Globals.SheetCount, Sheet1, Sheet2);
            }
            else
            {
                string SheetName1; string SheetName2;
                SplitSheetNames(SheetNames, out SheetName1, out SheetName2);
                if (!IsCellReference && SheetName1 != SheetName2) FlxMessages.ThrowException(FlxErr.ErrInvalidRef, SheetNames);

                if (IsCellReference && (SheetName1 == null || SheetName2 == null || SheetName1.Length == 0 || SheetName2.Length == 0))
                    FlxMessages.ThrowException(FlxErr.ErrInvalidSheet, String.Empty);  //External names might not point to any sheet, external 3d refs must point to existing sheets.

                if (ReadingXlsx)
                {
                    int SupBookIndex = Convert.ToInt32(FileName, CultureInfo.InvariantCulture);
                    IsLocal = SupBookIndex == 0;
                    return FWorkbook.Globals.References.AddSheetFromXlsxFile(SupBookIndex, SheetName1, SheetName2);
                }
                return FWorkbook.Globals.References.AddSheetFromExternalFile(FileName, SheetName1, SheetName2); 
            }
        }
        #endregion

        #region Macros
        ///<inheritdoc />
        public override void RemoveMacros()
        {
            CheckConnected();
            FWorkbook.Globals.HasMacro = false;
#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
            MacroData = null;
#endif
            using (MemoryStream MemFile = new MemoryStream(OtherStreams))
            {
                using (MemoryStream NewMemFile = new MemoryStream())
                {
                    using (TOle2File DataStream = new TOle2File(MemFile))
                    {
                        DataStream.PrepareForWrite(NewMemFile, XlsConsts.WorkbookString, XlsConsts.VBAStreams);
                    }
                    OtherStreams = NewMemFile.ToArray();  //do this after the using.
                }
            }
        }

        ///<inheritdoc />
        public override bool HasMacros()
        {
            CheckConnected();
            return FWorkbook.Globals.HasMacro;
        }

        internal override byte[] GetMacroData()
        {
            CheckConnected();
#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
            if (MacroData != null) return MacroData;
            using (MemoryStream MemFile = new MemoryStream(OtherStreams))
            {
                using (MemoryStream NewMemFile = new MemoryStream())
                {
                    using (TOle2File DataStream = new TOle2File(MemFile))
                    {
                        DataStream.GetStorages(XlsConsts.VBAMainStreamFullPath, NewMemFile);
                    }
                    return NewMemFile.ToArray();  //do this after the using.
                }
            }
#else
            return null;
#endif

        }
        internal override bool HasMacroXlsm()
        {
#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
            return MacroData != null;

#else
            return false;
#endif
        }

        internal override void SetMacrodata(byte[] aMacroData)
        {
            if (aMacroData == null) return;
#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
            MacroData = aMacroData;
#endif
            FWorkbook.Globals.HasMacro = true;            
        }
        #endregion

        #endregion
        #endregion


    }
}
