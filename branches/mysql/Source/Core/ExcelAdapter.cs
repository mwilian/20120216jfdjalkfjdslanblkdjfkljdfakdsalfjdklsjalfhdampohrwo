using System;
using System.IO;
using System.Globalization;
using System.Text;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;

#if (MONOTOUCH)
	using real = System.Single;
	using System.Drawing;
    using Color = MonoTouch.UIKit.UIColor;
    using Image = MonoTouch.UIKit.UIImage;
#else
	#if (WPF)
	using RectangleF = System.Windows.Rect;
	using PointF = System.Windows.Point;
	using real = System.Double;
	using System.Windows.Media;
	using System.Windows.Controls;
	#else
	using real = System.Single;
	using System.Drawing;
	using System.Drawing.Imaging;
	using System.Drawing.Drawing2D;
	#endif
#endif


namespace FlexCel.Core
{
	/// <summary>
	/// Interface a FlexCel engine has to implement to be used with FlexCelReport.
	/// </summary>
	/// <remarks>This is an abstract class encapsulating the FlexCel API. Any implementation on the API must derive from this class.<br/>
	/// FlexCel provides an implementation of this interface on the class <see cref="FlexCel.XlsAdapter.XlsFile"/>.
	/// </remarks>
    [ClassInterface(ClassInterfaceType.None)
	]
    public abstract class ExcelFile : IFlexCelFontList, IEmbeddedObjects, IFlexCelPalette, IRowColSize, IEnumerable<CellValue>
	{
		#region Globals
		bool FAllowOverwritingFiles;
		TProtection FProtection;
		TDocumentProperties FDocumentProperties;
		private TExcelFileErrorActions FErrorActions = TExcelFileErrorActions.ErrorOnXlsxMissingPart;
		private bool FSemiAbsoluteReferences;
		internal TWorkspace Workspace;
        private TFileFormats FDefaultFileFormat;
        private TXlsBiffVersion FXlsBiffVersion;
        private TReferenceStyle FFormulaReferenceStyle;
        private bool FVirtualMode;


        internal const bool FTesting = false; //Only for testing, should be true on normal work.

		#endregion

        #region Events
        /// <summary>
        /// If you assign this event FlexCel will not load the file into memory when opening a file, allowing you to 
        /// open very big files using little memory. This event will be called for every value read from the file, and then 
        /// the value will be discarded, instead of loaded into memory. Look for "Virtual Mode" in the Performance Guide for more information.
        /// </summary>
        public event VirtualCellReadEventHandler VirtualCellRead;

        /// <summary>
        /// When in virtual mode (<see cref="VirtualCellRead"/> is assigned) this event will be called
        /// after the sheet names have been read, but before starting to read the cells. You can use 
        /// this event to know how many sheets you are reading.
        /// </summary>
        public event VirtualCellStartReadingEventHandler VirtualCellStartReading;

        /// <summary>
        /// When in virtual mode (<see cref="VirtualCellRead"/> is assigned) this event will be called
        /// after the file has been processed. You can use it to do cleanup. 
        /// </summary>
        public event VirtualCellEndReadingEventHandler VirtualCellEndReading;

        #endregion

        #region Constructors
        /// <summary>
		/// Initializes ExcelFile fields.
		/// </summary>
		protected ExcelFile()
		{
			FProtection=new TProtection(this);
			FDocumentProperties = new TDocumentProperties(this);
            FDefaultFileFormat = TFileFormats.Automatic;
		}
		#endregion

        #region Virtual Mode
        /// <summary>
        /// Set this value to true to turn Virtual Mode on. "Virtual Mode" is explained in the Performance Guide.
        /// </summary>
        public bool VirtualMode
        {
            get { return FVirtualMode; }
            set { FVirtualMode = value; }
        }

        /// <summary>
        /// Replace this method if you want to override this event in a derived class.
        /// </summary>
        protected internal virtual void OnVirtualCellRead(ExcelFile xls, VirtualCellReadEventArgs e)
        {
            if (VirtualCellRead != null) VirtualCellRead(xls, e);
        }        
        
        /// <summary>
        /// Replace this method if you want to override this event in a derived class.
        /// </summary>
        protected internal virtual void OnVirtualCellStartReading(ExcelFile xls, VirtualCellStartReadingEventArgs e)
        {
            if (VirtualCellStartReading != null) VirtualCellStartReading(xls, e);
        }

        /// <summary>
        /// Replace this method if you want to override this event in a derived class.
        /// </summary>
        protected internal virtual void OnVirtualCellEndReading(ExcelFile xls, VirtualCellEndReadingEventArgs e)
        {
            if (VirtualCellEndReading != null) VirtualCellEndReading(xls, e);
        }
        #endregion

        #region File Managment
        /// <summary>
        /// Determines the default file format used by Excel when saving a file without specifying one, and when the file format can't be 
        /// determined from the extension of the file. If set to Automatic (The default) the file will be saved in the same format it was opened.
        /// That is, if you opened an xlsx file it will be saved as xlsx. If you opened an xls file (or created it with XlsFile.NewFile()) it will be saved as xls.
        /// When this property is automatic, text files will be saved as xls.
        /// </summary>
        public TFileFormats DefaultFileFormat { get { return FDefaultFileFormat; } set { FDefaultFileFormat = value; } }

        /// <summary>
        /// Defines the Excel mode used in this thread.
        /// Note that while on v2007 (the default) you still can make xls 97 spreadsheets, so the only reason to change this setting
        /// is if you have any compatibility issues (for example your code depends on a sheet having 65536 rows).
        /// <b>IMPORTANT: Do NOT change this value after reading a workbook</b>. Also, remember that the value is changed for all the reports in all threads.
        /// </summary>
        public static TExcelVersion ExcelVersion { get { return FlxConsts.ExcelVersion; } set { FlxConsts.ExcelVersion = value; } }


        /// <summary>
        /// This property lets you know if the version of FlexCel.dll you are using supports XLSX file format.
        /// Currently XLSX is only supported in .NET 3.0 or newer.
        /// </summary>
#if (FRAMEWORK30)
        public static readonly bool SupportsXlsx = true;
#else
        public static readonly bool SupportsXlsx = false;
#endif
            

        /// <summary>
        /// Xls files created by Excel 2007 have additional records that allow the generated file to store characteristics not available in Excel 2003 or older.
        /// (Like for example True color for cells instead of 54 colors). When opening an xls file created by Excel 2007 in Excel 2007, Excel will be able to read those values back.
        /// <br></br>By default FlexCel will read those extra records and when reading, and identify the file it creates as created by Excel 2007 when writing, so when you open it in Excel 2007 it will read those additional records.
        /// If for any reason you prefer FlexCel to behave as Excel 2003, saving the files as if they were created by Excel 2003 (So Excel 2007 will ignore the additional characteristics),
        /// and also stop FlexCel from reading those extra records, just change the value of this property.
        /// </summary>
        public TXlsBiffVersion XlsBiffVersion { get { return FXlsBiffVersion; } set { FXlsBiffVersion = value; } }

        /// <summary>
        /// Empty files created by different versions of Excel can have different characteristics. For example, the default font in an Excel 2003 
        /// file is Arial, while the default in 2007 is Calibri. This property returns the version of file that is loaded into FlexCel.
        /// When calling <see cref="NewFile(int, TExcelFileFormat)"/> or when opening a new file, FlexCel will update the value of this property.
        /// /// </summary>
        public abstract TExcelFileFormat ExcelFileFormat { get; }

        /// <summary>
        /// Defines what FlexCel will do when it finds a reference to the last row or column in an Excel 97-2003 spreadsheet, and it is upgrading to Excel 2007.
        /// If false (the default) row 65536 will be updated to row 1048576, and column 256 to column 16384.
        /// If true, references will stay the same. <b>Note: </b> This is a static global property, so it affects all threads running.
        /// </summary>
        public static bool KeepMaxRowsAndColumsWhenUpdating { get { return FlxConsts.KeepMaxRowsAndColumsWhenUpdating; } set { FlxConsts.KeepMaxRowsAndColumsWhenUpdating = value; } }



		/// <summary>
		/// The file we are working on. When we save the file with another name, it changes.
		/// When we open a stream, it is set to "".
		/// This value is also used to get the text of Headers and Footers (when using the filename macro).
		/// When using the filename macro on headers/footers, make sure you set this value to what you want.
		/// </summary>
		public abstract string ActiveFileName{get;set;}

		/// <summary>
		/// Creates a new empty file, with 3 empty sheets.
		/// </summary>
		/// <remarks>
		/// The file created will have empty properties, (author, description, etc).
		/// If you want to create a more personalized file when you call NewFile 
		/// (for example with a given Author, or the sheets names on your language), there are 2 options:
		///    <list type="number">
		///    <item>Don't use NewFile, but open an existing file you can modify.</item>
		///    <item>Or, you can replace the file "EmptyWorkbook.xls" (on xlsadapter folder) with your own and recompile.
		///    </item></list> 
		/// </remarks>
        public void NewFile()
        {
            NewFile(3);
        }

        /// <inheritdoc cref = "NewFile(int, TExcelFileFormat)" />
        public void NewFile(int aSheetCount)
        {
            NewFile(aSheetCount, TExcelFileFormat.v2003);
        }

        /// <summary>
        /// Creates a new empty file, with the specified number of sheets.
        /// </summary>
        /// <param name="aSheetCount">Number of sheets for the new file.</param>
        /// <param name="fileFormat">Different Excel versions save different empty files. By default, FlexCel will create a new file that looks
        /// like a file created by Excel 2003, but you can change the version of the new file created with this parameter.</param>
        /// <remarks>
        /// The file created will have empty properties, (author, description, etc).
        /// If you want to create a more personalized file when you call NewFile 
        /// (for example with a given Author, or the sheets names on your language), there are 2 options:
        ///    <list type="number">
        ///    <item>Don't use NewFile, but open an existing file you can modify.</item>
        ///    <item>Or, you can replace the files "EmptyWorkbook.xls", "EmptyWorkbook2007.xls" and "EmptyWorkbook2010.xls" (on xlsadapter folder) with your own and recompile.
        ///    </item></list> 
        /// </remarks>
        public abstract void NewFile(int aSheetCount, TExcelFileFormat fileFormat);


		/// <summary>
		/// Loads a new Spreadsheet form disk.
		/// </summary>
		/// <param name="fileName">File to open.</param>
		public void Open(string fileName)
		{
			Open(fileName, TFileFormats.Automatic, '\t', 1, 1, null);
		}

		/// <summary>
		/// Loads a new Spreadsheet form a stream.
		/// </summary>
		/// <param name="aStream">Stream to Load, must be a seekable stream. Verify it is on the correct position.</param>
		public void Open(Stream aStream)
		{
			Open(aStream, TFileFormats.Automatic, '\t', 1, 1, null);
		}

        ///<inheritdoc cref = "Open(string, TFileFormats, char, int, int, ColumnImportType[], string[], Encoding, bool)" />
        public void Open(string fileName, TFileFormats fileFormat, char delimiter, int firstRow, int firstCol, ColumnImportType[] columnFormats)
		{
			Open(fileName, fileFormat, delimiter, firstRow, firstCol, columnFormats, Encoding.Default, true);
        }

        ///<inheritdoc cref = "Open(string, TFileFormats, char, int, int, ColumnImportType[], string[], Encoding, bool)" />
        public void Open(string fileName, TFileFormats fileFormat, char delimiter, int firstRow, int firstCol, ColumnImportType[] columnFormats, Encoding fileEncoding, bool detectEncodingFromByteOrderMarks)
        {
            Open(fileName, fileFormat, delimiter, firstRow, firstCol, columnFormats, null, fileEncoding, detectEncodingFromByteOrderMarks);
        }

        ///<inheritdoc cref = "Open(Stream, TFileFormats, char, int, int, ColumnImportType[], string[], Encoding, bool)" />
        public void Open(Stream aStream, TFileFormats fileFormat, char delimiter, int firstRow, int firstCol, ColumnImportType[] columnFormats)
        {
            Open(aStream, fileFormat, delimiter, firstRow, firstCol, columnFormats, Encoding.Default, true);
        }

        ///<inheritdoc cref = "Open(Stream, TFileFormats, char, int, int, ColumnImportType[], string[], Encoding, bool)" />
        public void Open(Stream aStream, TFileFormats fileFormat, char delimiter, int firstRow, int firstCol, ColumnImportType[] columnFormats, bool detectEncodingFromByteOrderMarks)
        {
            Open(aStream, fileFormat, delimiter, firstRow, firstCol, columnFormats, Encoding.Default, detectEncodingFromByteOrderMarks);
        }

        ///<inheritdoc cref = "Open(Stream, TFileFormats, char, int, int, ColumnImportType[], string[], Encoding, bool)" />
        public void Open(Stream aStream, TFileFormats fileFormat, char delimiter, int firstRow, int firstCol, ColumnImportType[] columnFormats, Encoding fileEncoding, bool detectEncodingFromByteOrderMarks)
        {
            Open(aStream, fileFormat, delimiter, firstRow, firstCol, columnFormats, null, fileEncoding, detectEncodingFromByteOrderMarks);
        }

        /// <summary>
        /// Loads a new Spreadsheet form disk, on one of the specified formats.
        /// </summary>
        /// <param name="fileName">File to open</param>
        /// <param name="fileFormat">List with possible file formats to try</param>
        /// <param name="delimiter">Delimiter used to separate columns, if the format is <see cref="TFileFormats.Text"/></param>
        /// <param name="firstRow">First row where we will copy the cells on the new sheet, for <see cref="TFileFormats.Text"/> </param>
        /// <param name="firstCol">First column where we will copy the cells on the new sheet, for <see cref="TFileFormats.Text"/></param>
        /// <param name="columnFormats">An array of <see cref="ColumnImportType"/> elements, telling how each column should be imported.</param>
        ///<param name="fileEncoding">Encoding used by the file we are reading, when opening a Text-delimited file (csv or txt). 
        /// This parameter has no effect on xls files. If ommited, it is assumed to be Encoding.Default</param>
        ///<param name="detectEncodingFromByteOrderMarks">This parameter only applies when reading Text files. It is the same on the constructor of a StreamReader, and it says if BOM must be used at the beginning of the file. It defaults to true.</param>
        /// <param name="dateFormats">A list of formats allowed for dates and times, when opening text files. Windows is a little liberal in what it thinks can be a date, and it can convert things 
        /// like "1.2" into dates. By setting this property, you can ensure the dates are only in the formats you expect. If you leave it null, we will trust "DateTime.TryParse" to guess the correct values. This value has no meaning in normal xls files, only text files.</param>
        public abstract void Open(string fileName, TFileFormats fileFormat, char delimiter, int firstRow, int firstCol, ColumnImportType[] columnFormats, string[] dateFormats, Encoding fileEncoding, bool detectEncodingFromByteOrderMarks);

        /// <summary>
        /// Loads a new Spreadsheet form a stream, on one of the specified formats.
        /// </summary>
        /// <param name="aStream">Stream to open, must be a seekable stream. Verify it is on the correct position.</param>
        /// <param name="fileFormat">List with possible file formats to try</param>
        /// <param name="delimiter">Delimiter used to separate columns, if the format is <see cref="TFileFormats.Text"/></param>
        /// <param name="firstRow">First row where we will copy the cells on the new sheet, for <see cref="TFileFormats.Text"/> </param>
        /// <param name="firstCol">First column where we will copy the cells on the new sheet, for <see cref="TFileFormats.Text"/></param>
        /// <param name="columnFormats">An array of <see cref="ColumnImportType"/> elements, telling how each column should be imported.<br/> See the example
        ///<param name="fileEncoding">Encoding used by the file we are reading, when opening a Text-delimited file (csv or txt). 
        /// This parameter has no effect on xls files. If ommited, it is assumed to be Encoding.Default.</param>
        ///<param name="detectEncodingFromByteOrderMarks">This parameter only applies when reading Text files. It is the same on the constructor of a StreamReader, and it says if BOM must be used at the beginning of the file. It defaults to true.</param>
        /// for more information on how to use it.
        /// </param>
        /// <param name="dateFormats">A list of formats allowed for dates and times, when opening text files. Windows is a little liberal in what it thinks can be a date, and it can convert things 
        /// like "1.2" into dates. By setting this property, you can ensure the dates are only in the formats you expect. If you leave it null, we will trust "DateTime.TryParse" to guess the correct values. This value has no meaning in normal xls files, only text files.</param>
        /// <example>
        /// Imagine you have a file with 20 columns, and column 2 has numbers you want to be imported as text (like phone numbers), and you don't want to import column 10.
        /// <br/>You can use the following code to do it:
        /// <code>
        ///   ColumnImportType[] ColTypes = new ColumnImportType[10]; //You just need to define 10 items, all other columns after 10 will be imported with default formatting.
        ///   ColTypes[1] = ColumnImportType.Text; //Import whatever is in column 2 as text.
        ///   ColTypes[9] = ColumnImportType.Skip; //don't import column 10.
        ///
        ///   xls.Open("csv.csv", TFileFormats.Text, ',', 1,1,ColTypes);
        /// </code>
        /// </example>
		public abstract void Open(Stream aStream, TFileFormats fileFormat, char delimiter, int firstRow, int firstCol, ColumnImportType[] columnFormats, string[] dateFormats, Encoding fileEncoding, bool detectEncodingFromByteOrderMarks);

        /// <summary>
        /// Imports a text file (character-delimited columns) into the current sheet. Note that this method won't clear any existing data.
        /// </summary>
        /// <param name="fileName">File with the text to import.</param>
        /// <param name="delimiter">Character used to separate columns.</param>
        /// <param name="firstRow">Row in the Active sheet where we will start importing the text file.</param>
        /// <param name="firstCol">Column in the Active sheet where we will start importing the text file.</param>
        /// <param name="columnFormats">An array of <see cref="ColumnImportType"/> elements, telling how each column should be imported.<br/> See the example
        /// in <see cref="Open(Stream, TFileFormats, char, int, int, ColumnImportType[])"/> for more information on how to use it.
        /// </param>
        ///<param name="fileEncoding">Encoding used by the file we are reading. If ommited, it is assumed to be Encoding.Default.</param>
        ///<param name="detectEncodingFromByteOrderMarks">It is the same on the constructor of a StreamReader, and it says if BOM must be used at the beginning of the file. It defaults to true.</param>
        public void Import(string fileName, int firstRow, int firstCol, char delimiter, ColumnImportType[] columnFormats, Encoding fileEncoding, bool detectEncodingFromByteOrderMarks)
        {
            Import(fileName, firstRow, firstCol, delimiter, columnFormats, null, fileEncoding, detectEncodingFromByteOrderMarks);
        }

        /// <summary>
        /// Imports a text file (character-delimited columns) into the current sheet. Note that this method won't clear any existing data.
        /// </summary>
        /// <param name="fileName">File with the text to import.</param>
        /// <param name="delimiter">Character used to separate columns.</param>
        /// <param name="firstRow">Row in the Active sheet where we will start importing the text file.</param>
        /// <param name="firstCol">Column in the Active sheet where we will start importing the text file.</param>
        /// <param name="columnFormats">An array of <see cref="ColumnImportType"/> elements, telling how each column should be imported.<br/> See the example
        /// in <see cref="Open(Stream, TFileFormats, char, int, int, ColumnImportType[])"/> for more information on how to use it.
        /// </param>
        ///<param name="fileEncoding">Encoding used by the file we are reading. If ommited, it is assumed to be Encoding.Default.</param>
        ///<param name="detectEncodingFromByteOrderMarks">It is the same on the constructor of a StreamReader, and it says if BOM must be used at the beginning of the file. It defaults to true.</param>
        /// <param name="dateFormats">A list of formats allowed for dates and times, when opening text files. Windows is a little liberal in what it thinks can be a date, and it can convert things 
        /// like "1.2" into dates. By setting this property, you can ensure the dates are only in the formats you expect. If you leave it null, we will trust "DateTime.TryParse" to guess the correct values. This value has no meaning in normal xls files, only text files.</param>
        public void Import(string fileName, int firstRow, int firstCol, char delimiter, ColumnImportType[] columnFormats, string[] dateFormats, Encoding fileEncoding, bool detectEncodingFromByteOrderMarks)
        {
            using (FileStream f = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) //FileShare.ReadWrite is the way we have to open a file even if it is being used by excel.
            {
                Import(f, firstRow, firstCol, delimiter, columnFormats, dateFormats, fileEncoding, detectEncodingFromByteOrderMarks);
            }
        }

        /// <summary>
        /// Imports a text file (character-delimited columns) into the current sheet. Note that this method won't clear any existing data.
        /// </summary>
        /// <param name="aStream">Stream with the text to import.</param>
        /// <param name="delimiter">Character used to separate columns.</param>
        /// <param name="firstRow">Row in the Active sheet where we will start importing the text file.</param>
        /// <param name="firstCol">Column in the Active sheet where we will start importing the text file.</param>
        /// <param name="columnFormats">An array of <see cref="ColumnImportType"/> elements, telling how each column should be imported.<br/> See the example
        /// in <see cref="Open(Stream, TFileFormats, char, int, int, ColumnImportType[])"/> for more information on how to use it.
        /// </param>
        ///<param name="fileEncoding">Encoding used by the file we are reading. If ommited, it is assumed to be Encoding.Default.</param>
        ///<param name="detectEncodingFromByteOrderMarks">It is the same on the constructor of a StreamReader, and it says if BOM must be used at the beginning of the file. It defaults to true.</param>
        public void Import(Stream aStream, int firstRow, int firstCol, char delimiter, ColumnImportType[] columnFormats, Encoding fileEncoding, bool detectEncodingFromByteOrderMarks)
        {
            Import(aStream, firstRow, firstCol, delimiter, columnFormats, null, fileEncoding, detectEncodingFromByteOrderMarks);
        }

        /// <summary>
        /// Imports a text file (character-delimited columns) into the current sheet. Note that this method won't clear any existing data.
        /// </summary>
        /// <param name="aStream">Stream with the text to import.</param>
        /// <param name="delimiter">Character used to separate columns.</param>
        /// <param name="firstRow">Row in the Active sheet where we will start importing the text file.</param>
        /// <param name="firstCol">Column in the Active sheet where we will start importing the text file.</param>
        /// <param name="columnFormats">An array of <see cref="ColumnImportType"/> elements, telling how each column should be imported.<br/> See the example
        /// in <see cref="Open(Stream, TFileFormats, char, int, int, ColumnImportType[])"/> for more information on how to use it.
        /// </param>
        ///<param name="fileEncoding">Encoding used by the file we are reading. If ommited, it is assumed to be Encoding.Default.</param>
        ///<param name="detectEncodingFromByteOrderMarks">It is the same on the constructor of a StreamReader, and it says if BOM must be used at the beginning of the file. It defaults to true.</param>
        /// <param name="dateFormats">A list of formats allowed for dates and times, when opening text files. Windows is a little liberal in what it thinks can be a date, and it can convert things 
        /// like "1.2" into dates. By setting this property, you can ensure the dates are only in the formats you expect. If you leave it null, we will trust "DateTime.TryParse" to guess the correct values. This value has no meaning in normal xls files, only text files.</param>
        public void Import(Stream aStream, int firstRow, int firstCol, char delimiter, ColumnImportType[] columnFormats, string[] dateFormats, Encoding fileEncoding, bool detectEncodingFromByteOrderMarks)
        {
            using (StreamReader sr = new StreamReader(new TUndisposableStream(aStream), fileEncoding, detectEncodingFromByteOrderMarks))
            {
                Import(sr, firstRow, firstCol, delimiter, columnFormats, dateFormats);
            }
        }

        /// <summary>
        /// Imports a text file (character-delimited columns) into the current sheet. Note that this method won't clear any existing data.
        /// </summary>
        /// <param name="aTextReader">TextReader with the text to import.</param>
        /// <param name="delimiter">Character used to separate columns.</param>
        /// <param name="firstRow">Row in the Active sheet where we will start importing the text file.</param>
        /// <param name="firstCol">Column in the Active sheet where we will start importing the text file.</param>
        /// <param name="columnFormats">An array of <see cref="ColumnImportType"/> elements, telling how each column should be imported.<br/> See the example
        /// in <see cref="Open(Stream, TFileFormats, char, int, int, ColumnImportType[])"/> for more information on how to use it.
        /// </param>
        /// <param name="dateFormats">A list of formats allowed for dates and times, when opening text files. Windows is a little liberal in what it thinks can be a date, and it can convert things 
        /// like "1.2" into dates. By setting this property, you can ensure the dates are only in the formats you expect. If you leave it null, we will trust "DateTime.TryParse" to guess the correct values. This value has no meaning in normal xls files, only text files.</param>
        public abstract void Import(TextReader aTextReader, int firstRow, int firstCol, char delimiter, ColumnImportType[] columnFormats, string[] dateFormats);

        /// <summary>
        /// Imports a text file (fixed length columns) into the current sheet. Note that this method won't clear any existing data.
        /// </summary>
        /// <param name="fileName">File with the text to import.</param>
        /// <param name="firstRow">Row in the Active sheet where we will start importing the text file.</param>
        /// <param name="firstCol">Column in the Active sheet where we will start importing the text file.</param>
        /// <param name="columnWidths">An array with the column widths for every column you want to import.</param>
        /// <param name="columnFormats">An array of <see cref="ColumnImportType"/> elements, telling how each column should be imported.<br/> See the example
        /// in <see cref="Open(Stream, TFileFormats, char, int, int, ColumnImportType[])"/> for more information on how to use it.
        /// </param>
        ///<param name="fileEncoding">Encoding used by the file we are reading. If ommited, it is assumed to be Encoding.Default.</param>
        ///<param name="detectEncodingFromByteOrderMarks">It is the same on the constructor of a StreamReader, and it says if BOM must be used at the beginning of the file. It defaults to true.</param>
        public void Import(string fileName, int firstRow, int firstCol, int[] columnWidths, ColumnImportType[] columnFormats, Encoding fileEncoding, bool detectEncodingFromByteOrderMarks)
        {
            Import(fileName, firstRow, firstCol, columnWidths, columnFormats, null, fileEncoding, detectEncodingFromByteOrderMarks);
        }

        /// <summary>
        /// Imports a text file (fixed length columns) into the current sheet. Note that this method won't clear any existing data.
        /// </summary>
        /// <param name="fileName">File with the text to import.</param>
        /// <param name="firstRow">Row in the Active sheet where we will start importing the text file.</param>
        /// <param name="firstCol">Column in the Active sheet where we will start importing the text file.</param>
        /// <param name="columnWidths">An array with the column widths for every column you want to import.</param>
        /// <param name="columnFormats">An array of <see cref="ColumnImportType"/> elements, telling how each column should be imported.<br/> See the example
        /// in <see cref="Open(Stream, TFileFormats, char, int, int, ColumnImportType[])"/> for more information on how to use it.
        /// </param>
        ///<param name="fileEncoding">Encoding used by the file we are reading. If ommited, it is assumed to be Encoding.Default.</param>
        ///<param name="detectEncodingFromByteOrderMarks">It is the same on the constructor of a StreamReader, and it says if BOM must be used at the beginning of the file. It defaults to true.</param>
        /// <param name="dateFormats">A list of formats allowed for dates and times, when opening text files. Windows is a little liberal in what it thinks can be a date, and it can convert things 
        /// like "1.2" into dates. By setting this property, you can ensure the dates are only in the formats you expect. If you leave it null, we will trust "DateTime.TryParse" to guess the correct values. This value has no meaning in normal xls files, only text files.</param>
        public void Import(string fileName, int firstRow, int firstCol, int[] columnWidths, ColumnImportType[] columnFormats, string[] dateFormats, Encoding fileEncoding, bool detectEncodingFromByteOrderMarks)
        {
            using (FileStream f = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) //FileShare.ReadWrite is the way we have to open a file even if it is being used by excel.
            {
                Import(f, firstRow, firstCol, columnWidths, columnFormats, dateFormats, fileEncoding, detectEncodingFromByteOrderMarks);
            }
        }

        /// <summary>
        /// Imports a text file (fixed length columns) into the current sheet. Note that this method won't clear any existing data.
        /// </summary>
        /// <param name="aStream">Stream with the text to import.</param>
        /// <param name="firstRow">Row in the Active sheet where we will start importing the text file.</param>
        /// <param name="firstCol">Column in the Active sheet where we will start importing the text file.</param>
        /// <param name="columnWidths">An array with the column widths for every column you want to import.</param>
        /// <param name="columnFormats">An array of <see cref="ColumnImportType"/> elements, telling how each column should be imported.<br/> See the example
        /// in <see cref="Open(Stream, TFileFormats, char, int, int, ColumnImportType[])"/> for more information on how to use it.
        /// </param>
        ///<param name="fileEncoding">Encoding used by the file we are reading. If ommited, it is assumed to be Encoding.Default.</param>
        ///<param name="detectEncodingFromByteOrderMarks">It is the same on the constructor of a StreamReader, and it says if BOM must be used at the beginning of the file. It defaults to true.</param>
        public void Import(Stream aStream, int firstRow, int firstCol, int[] columnWidths, ColumnImportType[] columnFormats, Encoding fileEncoding, bool detectEncodingFromByteOrderMarks)
        {
            Import(aStream, firstRow, firstCol, columnWidths, columnFormats, null, fileEncoding, detectEncodingFromByteOrderMarks);
        }

        /// <summary>
        /// Imports a text file (fixed length columns) into the current sheet. Note that this method won't clear any existing data.
        /// </summary>
        /// <param name="aStream">Stream with the text to import.</param>
        /// <param name="firstRow">Row in the Active sheet where we will start importing the text file.</param>
        /// <param name="firstCol">Column in the Active sheet where we will start importing the text file.</param>
        /// <param name="columnWidths">An array with the column widths for every column you want to import.</param>
        /// <param name="columnFormats">An array of <see cref="ColumnImportType"/> elements, telling how each column should be imported.<br/> See the example
        /// in <see cref="Open(Stream, TFileFormats, char, int, int, ColumnImportType[])"/> for more information on how to use it.
        /// </param>
        ///<param name="fileEncoding">Encoding used by the file we are reading. If ommited, it is assumed to be Encoding.Default.</param>
        ///<param name="detectEncodingFromByteOrderMarks">It is the same on the constructor of a StreamReader, and it says if BOM must be used at the beginning of the file. It defaults to true.</param>
        /// <param name="dateFormats">A list of formats allowed for dates and times, when opening text files. Windows is a little liberal in what it thinks can be a date, and it can convert things 
        /// like "1.2" into dates. By setting this property, you can ensure the dates are only in the formats you expect. If you leave it null, we will trust "DateTime.TryParse" to guess the correct values. This value has no meaning in normal xls files, only text files.</param>
        public void Import(Stream aStream, int firstRow, int firstCol, int[] columnWidths, ColumnImportType[] columnFormats, string[] dateFormats , Encoding fileEncoding, bool detectEncodingFromByteOrderMarks)
        {
            using (StreamReader sr = new StreamReader(new TUndisposableStream(aStream), fileEncoding, detectEncodingFromByteOrderMarks))
            {
                Import(sr, firstRow, firstCol, columnWidths, columnFormats, dateFormats);
            }
        }

        /// <summary>
        /// Imports a text file (fixed length columns) into the current sheet. Note that this method won't clear any existing data.
        /// </summary>
        /// <param name="aTextReader">StreamReader with the text to import.</param>
        /// <param name="firstRow">Row in the Active sheet where we will start importing the text file.</param>
        /// <param name="firstCol">Column in the Active sheet where we will start importing the text file.</param>
        /// <param name="columnWidths">An array with the column widths for every column you want to import.</param>
        /// <param name="columnFormats">An array of <see cref="ColumnImportType"/> elements, telling how each column should be imported.<br/> See the example
        /// in <see cref="Open(Stream, TFileFormats, char, int, int, ColumnImportType[])"/> for more information on how to use it.
        /// </param>
        public void Import(TextReader aTextReader, int firstRow, int firstCol, int[] columnWidths, ColumnImportType[] columnFormats)
        {
            Import(aTextReader, firstRow, firstCol, columnWidths, columnFormats, null);
        }

        /// <summary>
        /// Imports a text file (fixed length columns) into the current sheet. Note that this method won't clear any existing data.
        /// </summary>
        /// <param name="aTextReader">StreamReader with the text to import.</param>
        /// <param name="firstRow">Row in the Active sheet where we will start importing the text file.</param>
        /// <param name="firstCol">Column in the Active sheet where we will start importing the text file.</param>
        /// <param name="columnWidths">An array with the column widths for every column you want to import.</param>
        /// <param name="columnFormats">An array of <see cref="ColumnImportType"/> elements, telling how each column should be imported.<br/> See the example
        /// in <see cref="Open(Stream, TFileFormats, char, int, int, ColumnImportType[])"/> for more information on how to use it.
        /// </param>
        /// <param name="dateFormats">A list of formats allowed for dates and times, when opening text files. Windows is a little liberal in what it thinks can be a date, and it can convert things 
        /// like "1.2" into dates. By setting this property, you can ensure the dates are only in the formats you expect. If you leave it null, we will trust "DateTime.TryParse" to guess the correct values. This value has no meaning in normal xls files, only text files.</param>
        public abstract void Import(TextReader aTextReader, int firstRow, int firstCol, int[] columnWidths, ColumnImportType[] columnFormats, string[] dateFormats);

        
        /// <summary>
		/// Saves the file to disk, on native format.
		/// </summary>
		/// <param name="fileName">File to save. If <see cref="AllowOverwritingFiles"/> is false, then fileName MUST NOT exist.</param>
		public void Save(string fileName)
		{
			Save(fileName,  TFileFormats.Automatic, '\t');
		}

		/// <summary>
		/// Saves the file to a stream, on native format.
		/// </summary>
		/// <param name="aStream">Stream where to save the file. Must be a seekable stream.</param>
		public void Save(Stream aStream)
		{
			Save(aStream,  TFileFormats.Automatic, '\t');
		}

		/// <summary>
		/// Saves the file to a disk.
		/// </summary>
		/// <param name="fileName">File to save. If <see cref="AllowOverwritingFiles"/> is false, then fileName MUST NOT exist.</param>
        /// <param name="fileFormat">File format. If file format is text, a tab will be used as delimiter.  Automatic will try to guess it from the filename, if present.</param>
		public void Save(string fileName, TFileFormats fileFormat)
		{
			Save(fileName, fileFormat, '\t');
		}

		/// <summary>
		/// Saves the file to a stream.
		/// </summary>
		/// <param name="aStream">Stream where to save the file. Must be a seekable stream.</param>
        /// <param name="fileFormat">File format. If file format is text, a tab will be used as delimiter.  Automatic will try to guess it from the filename, if present.</param>
		public void Save(Stream aStream, TFileFormats fileFormat)
		{
			Save(aStream, fileFormat, '\t');
		}

		/// <summary>
		/// Saves the file to a disk.
		/// </summary>
		/// <param name="fileName">File to save. If <see cref="AllowOverwritingFiles"/> is false, then fileName MUST NOT exist.</param>
        /// <param name="fileFormat">File format. Automatic will try to guess it from the filename, if present.</param>
		/// <param name="delimiter">Delimiter to use if FileFormat is <see cref="TFileFormats.Text"/></param>
		public void Save(string fileName, TFileFormats fileFormat, char delimiter)
		{
			Save(fileName, fileFormat, delimiter, Encoding.Default);
		}

		/// <summary>
		/// Saves the file to a stream.
		/// </summary>
		/// <param name="aStream">Stream where to save the file. Must be a seekable stream.</param>
        /// <param name="fileFormat">File format. Automatic will try to guess it from the filename, if present.</param>
		/// <param name="delimiter">Delimiter to use if FileFormat is <see cref="TFileFormats.Text"/></param>
		public void Save(Stream aStream, TFileFormats fileFormat, char delimiter)
		{
			Save(aStream, fileFormat, delimiter, Encoding.Default);
		}

		/// <summary>
		/// Saves the file to a disk.
		/// </summary>
		/// <param name="fileName">File to save. If <see cref="AllowOverwritingFiles"/> is false, then fileName MUST NOT exist.</param>
        /// <param name="fileFormat">File format.  Automatic will try to guess it from the filename, if present.</param>
		/// <param name="delimiter">Delimiter to use if FileFormat is <see cref="TFileFormats.Text"/></param>
		/// <param name="fileEncoding">Encoding for the generated file, when writing a Text-delimited file (csv or txt). 
		/// This parameter has no effect on xls files. If ommited, Encoding.Default will be used. Note that to create a file with BOM (byte order marker) you need to specify an encoding here, the same as you do with a StreamWriter.</param>
		public abstract void Save(string fileName, TFileFormats fileFormat, char delimiter, Encoding fileEncoding);

		/// <summary>
		/// Saves the file to a stream.
		/// </summary>
		/// <param name="aStream">Stream where to save the file. Must be a seekable stream.</param>
		/// <param name="fileFormat">File format. Automatic will try to guess it from the filename, if present.</param>
		/// <param name="delimiter">Delimiter to use if FileFormat is <see cref="TFileFormats.Text"/></param>
		/// <param name="fileEncoding">Encoding for the generated file, when writing a Text-delimited file (csv or txt). 
		/// This parameter has no effect on xls files. If ommited, Encoding.Default will be used. Note that to create a file with BOM (byte order marker) you need to specify an encoding here, the same as you do with a StreamWriter.</param>
		public abstract void Save(Stream aStream, TFileFormats fileFormat, char delimiter, Encoding fileEncoding);

        /// <summary>
        /// Exports a range of cells from the active sheet into a text file (character delimited columns).
        /// </summary>
        /// <param name="fileName">File where we want to save the data.</param>
        /// <param name="range">Range of cells to export. If you want to export the full sheet, set it to null.</param>
        /// <param name="delimiter">Character used to delimit the fields in the exported file. You might normally use a comma (',') or a tab here. </param>
        /// <param name="exportHiddenRowsOrColumns">If true, hidden rows and columns will be exported. If false, they will be ignored.</param>
        /// <param name="fileEncoding">Encoding for the generated file. If you are unsure, you can use Encoding.Default here. Note that to create a file with BOM (byte order marker) you need to specify an encoding here, the same as you do with a StreamWriter.</param>
        public void Export(string fileName, TXlsCellRange range, char delimiter, bool exportHiddenRowsOrColumns, Encoding fileEncoding)
        {
            try
            {
                FileMode fm = FileMode.CreateNew;
                if (AllowOverwritingFiles) fm = FileMode.Create;
                FileAccess fa = FileAccess.Write;
                using (FileStream f = new FileStream(fileName, fm, fa))
                {
                    Export(f, range, delimiter, exportHiddenRowsOrColumns, fileEncoding);
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
        }

        /// <summary>
        /// Exports a range of cells from the active sheet into a text file (character delimited columns).
        /// </summary>
        /// <param name="aStream">Stream where we want to save the file.</param>
        /// <param name="range">Range of cells to export. If you want to export the full sheet, set it to null.</param>
        /// <param name="delimiter">Character used to delimit the fields in the exported file. You might normally use a comma (',') or a tab here. </param>
        /// <param name="exportHiddenRowsOrColumns">If true, hidden rows and columns will be exported. If false, they will be ignored.</param>
        /// <param name="fileEncoding">Encoding for the generated file. If you are unsure, you can use Encoding.Default here. Note that to create a file with BOM (byte order marker) you need to specify an encoding here, the same as you do with a StreamWriter.</param>
        public void Export(Stream aStream, TXlsCellRange range, char delimiter, bool exportHiddenRowsOrColumns, Encoding fileEncoding)
        {
            using (StreamWriter sw = new StreamWriter(new TUndisposableStream(aStream), fileEncoding))
            {
                Export(sw, range, delimiter, exportHiddenRowsOrColumns);
            }
        }

        /// <summary>
        /// Exports a range of cells from the active sheet into a text file (character delimited columns).
        /// </summary>
        /// <param name="aTextWriter">TextWriter where we want to save the file.</param>
        /// <param name="range">Range of cells to export. If you want to export the full sheet, set it to null.</param>
        /// <param name="delimiter">Character used to delimit the fields in the exported file. You might normally use a comma (',') or a tab here. </param>
        /// <param name="exportHiddenRowsOrColumns">If true, hidden rows and columns will be exported. If false, they will be ignored.</param>
        public abstract void Export(TextWriter aTextWriter, TXlsCellRange range, char delimiter, bool exportHiddenRowsOrColumns);

        /// <inheritdoc cref="Export(string, TXlsCellRange, int, int[], bool, Encoding, bool)" />
        public void Export(string fileName, TXlsCellRange range, int charactersForFirstColumn, int[] columnWidths, bool exportHiddenRowsOrColumns, Encoding fileEncoding)
        {
            Export(fileName, range, charactersForFirstColumn, columnWidths, exportHiddenRowsOrColumns, fileEncoding, false);
        }

        /// <summary>
        /// Exports a range of cells from the active sheet into a text file (fixed length columns).
        /// </summary>
        /// <param name="fileName">File where we want to save the data.</param>
        /// <param name="range">Range of cells to export. If you want to export the full sheet, set it to null.</param>
        /// <param name="charactersForFirstColumn">This value only has effect if columnWidths is null. It will specify how many characters
        /// to use for the first column, and all other columns will be determinded acording to their ratio with the first.<br>
        /// </br>For example, if the first column is 150 pixels wide and you specify "8" for this parameter, the first column will be padded to 8 characters when exporting.
        /// If the second column is 300 pixels wide, then it will be padded to 16 characters and so on. As this might not be 100% exact and depend in pixel measurements,
        /// you might want to specify columnWidths parameter instead of using this one. <br></br>
        /// Note: Setting this parameter to a negative value will assume the text in the collumns is already padded, and won't attempt to do any padding.
        /// Use this value if your data is padded in the spreadsheet itself.
        /// </param>
        /// <param name="columnWidths">Array with the number of charaters that will be assigned to every column when exporting. Supplying this array
        /// allows you to specify exactly how many characters you want for every field, and that might be really necessary to interop with other applications.
        /// But you can also leave this parameter null and specify "charactersForFirstColumn" to let FlexCel calculate how many characters to apply for every field.</param>
        /// <param name="exportHiddenRowsOrColumns">If true, hidden rows and columns will be exported. If false, they will be ignored.</param>
        /// <param name="fileEncoding">Encoding for the generated file. If you are unsure, you can use Encoding.Default here. Note that to create a file with BOM (byte order marker) you need to specify an encoding here, the same as you do with a StreamWriter.</param>
        /// <param name="exportTextOutsideCells">If true and the cell text spans over more than one empty cell to the right, that text will be exported. When false (the default) only text that fits in the cell will be exported.
        /// When this value is true the printout will look better, but it will not be possible to reimport the data as the columns are lost.</param>
        public void Export(string fileName, TXlsCellRange range, int charactersForFirstColumn, int[] columnWidths, bool exportHiddenRowsOrColumns, Encoding fileEncoding, bool exportTextOutsideCells)
        {
            try
            {
                FileMode fm = FileMode.CreateNew;
                if (AllowOverwritingFiles) fm = FileMode.Create;
                FileAccess fa = FileAccess.Write;
                using (FileStream f = new FileStream(fileName, fm, fa))
                {
                    Export(f, range, charactersForFirstColumn, columnWidths, exportHiddenRowsOrColumns, fileEncoding, exportTextOutsideCells);
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
        }

        /// <inheritdoc cref="Export(Stream, TXlsCellRange, int, int[], bool, Encoding, bool)" />
        public void Export(Stream aStream, TXlsCellRange range, int charactersForFirstColumn, int[] columnWidths, bool exportHiddenRowsOrColumns, Encoding fileEncoding)
        {
            Export(aStream, range, charactersForFirstColumn, columnWidths, exportHiddenRowsOrColumns, fileEncoding, false);
        }

        /// <summary>
        /// Exports a range of cells from the active sheet into a text file (fixed length columns).
        /// </summary>
        /// <param name="aStream">Stream where we want to save the file.</param>
        /// <param name="range">Range of cells to export. If you want to export the full sheet, set it to null.</param>
        /// <param name="charactersForFirstColumn">This value only has effect if columnWidths is null. It will specify how many characters
        /// to use for the first column, and all other columns will be determinded acording to their ratio with the first.<br>
        /// </br>For example, if the first column is 150 pixels wide and you specify "8" for this parameter, the first column will be padded to 8 characters when exporting.
        /// If the second column is 300 pixels wide, then it will be padded to 16 characters and so on. As this might not be 100% exact and depend in pixel measurements,
        /// you might want to specify columnWidths parameter instead of using this one. <br></br>
        /// Note: Setting this parameter to a negative value will assume the text in the collumns is already padded, and won't attempt to do any padding.
        /// Use this value if your data is padded in the spreadsheet itself.
        /// </param>
        /// <param name="columnWidths">Array with the number of charaters that will be assigned to every column when exporting. Supplying this array
        /// allows you to specify exactly how many characters you want for every field, and that might be really necessary to interop with other applications.
        /// But you can also leave this parameter null and specify "charactersForFirstColumn" to let FlexCel calculate how many characters to apply for every field.</param>
        /// <param name="exportHiddenRowsOrColumns">If true, hidden rows and columns will be exported. If false, they will be ignored.</param>
        /// <param name="fileEncoding">Encoding for the generated file. If you are unsure, you can use Encoding.Default here. Note that to create a file with BOM (byte order marker) you need to specify an encoding here, the same as you do with a StreamWriter.</param>
        /// <param name="exportTextOutsideCells">If true and the cell text spans over more than one empty cell to the right, that text will be exported. When false (the default) only text that fits in the cell will be exported.
        /// When this value is true the printout will look better, but it will not be possible to reimport the data as the columns are lost.</param>
        public void Export(Stream aStream, TXlsCellRange range, int charactersForFirstColumn, int[] columnWidths, bool exportHiddenRowsOrColumns, Encoding fileEncoding, bool exportTextOutsideCells)
        {
            using (StreamWriter sw = new StreamWriter(new TUndisposableStream(aStream), fileEncoding))
            {
                Export(sw, range, charactersForFirstColumn, columnWidths, exportHiddenRowsOrColumns, exportTextOutsideCells);
            }
        }

        /// <inheritdoc cref="Export(TextWriter, TXlsCellRange, int, int[], bool, bool)" />
        public void Export(TextWriter aTextWriter, TXlsCellRange range, int charactersForFirstColumn, int[] columnWidths, bool exportHiddenRowsOrColumns)
        {
            Export(aTextWriter, range, charactersForFirstColumn, columnWidths, exportHiddenRowsOrColumns, false);
        }

        /// <summary>
        /// Exports a range of cells from the active sheet into a text file (fixed length columns).
        /// </summary>
        /// <param name="aTextWriter">TextWriter where we want to save the file.</param>
        /// <param name="range">Range of cells to export. If you want to export the full sheet, set it to null.</param>
        /// <param name="charactersForFirstColumn">This value only has effect if columnWidths is null. It will specify how many characters
        /// to use for the first column, and all other columns will be determinded acording to their ratio with the first.<br>
        /// </br>For example, if the first column is 150 pixels wide and you specify "8" for this parameter, the first column will be padded to 8 characters when exporting.
        /// If the second column is 300 pixels wide, then it will be padded to 16 characters and so on. As this might not be 100% exact and depend in pixel measurements,
        /// you might want to specify columnWidths parameter instead of using this one. <br></br>
        /// Note: Setting this parameter to a negative value will assume the text in the collumns is already padded, and won't attempt to do any padding.
        /// Use this value if your data is padded in the spreadsheet itself.
        /// </param>
        /// <param name="columnWidths">Array with the number of charaters that will be assigned to every column when exporting. Supplying this array
        /// allows you to specify exactly how many characters you want for every field, and that might be really necessary to interop with other applications.
        /// But you can also leave this parameter null and specify "charactersForFirstColumn" to let FlexCel calculate how many characters to apply for every field.</param>
        /// <param name="exportHiddenRowsOrColumns">If true, hidden rows and columns will be exported. If false, they will be ignored.</param>
        /// <param name="exportTextOutsideCells">If true and the cell text spans over more than one empty cell to the right, that text will be exported. When false (the default) only text that fits in the cell will be exported.
        /// When this value is true the printout will look better, but it will not be possible to reimport the data as the columns are lost.</param>
        public abstract void Export(TextWriter aTextWriter, TXlsCellRange range, int charactersForFirstColumn, int[] columnWidths, bool exportHiddenRowsOrColumns, bool exportTextOutsideCells);

        /// <summary>
        /// This method will save the file in a format that will remain the same if the file is not modified. Normal xls files contain TimeStamp
        /// fields that might be modified when the file is downloaded or just copied.<br></br>
        /// While you will not be able to load the file saved, you might use this method to create a hash of a file and compare it to others to know if something changed. 
        /// <br></br><br></br>
        /// This overload will not save cell selections or the active sheet, and it is equivalent to calling <see cref="SaveForHashing(Stream, TExcludedRecords)"/> with
        /// the excludedRecords parameter set to TExcludedRecords.All. Use <see cref="SaveForHashing(Stream, TExcludedRecords)"/> for more control on which records to exclude.
        /// </summary>
        /// <param name="aStream">Stream where the file will be saved. You will probably want to hash this stream to store the corresponding hash.</param>
        /// <remarks>This method will not save the file in any readable format, and <b>the file format might change between FlexCel versions.</b>
        /// The only thing it guarantees is that the hashes for 2 identical xls files will be the same, for the same FlexCel version. Once you upgrade version, hashes might have to be rebuilt.
        /// <br></br><br></br>Also note that this method is useful to detect changes when the file is not edited in Excel. If you open the file in Excel and save it again,
        /// Excel will change a lot of reserevd bits, and the files will be too different for this method to have the same hashes. This is only to detect changes
        /// when copying or downloding an xls file. If you want to compare just cell contents, you might compare the files saved as CSV.
        /// </remarks>
        /// <example>
        /// The following method will calculate the hash for an existing file:
        /// <code>
        /// private byte[] GetHash(string FileName)
        /// {
        ///    XlsFile xls = new XlsFile(FileName);
        ///    using (System.Security.Cryptography.SHA1CryptoServiceProvider hasher = new System.Security.Cryptography.SHA1CryptoServiceProvider())
        ///    {
        ///         using (MemoryStream ms = new MemoryStream())
        ///         {
        ///             xls.SaveForHashing(ms);
        ///             ms.Position = 0;
        ///             return hasher.ComputeHash(ms);
        ///         }
        ///     }
        /// }
        /// </code>
        /// </example>
        public void SaveForHashing(Stream aStream)
        {
            SaveForHashing(aStream, TExcludedRecords.All);
        }

        /// <summary>
        /// This method will save the file in a format that will remain the same if the file is not modified. Normal xls files contain TimeStamp
        /// fields that might be modified when the file is downloaded or just copied.<br></br>
        /// While you will not be able to load the file saved, you might use this method to create a hash of a file and compare it to others to know if something changed. 
        /// </summary>
        /// <param name="aStream">Stream where the file will be saved. You will probably want to hash this stream to store the corresponding hash.</param>
        /// <param name="excludedRecords">A list with all the records you don't wish to include in the saved file (like for example cell selection). You will normally will want to 
        /// specify <b>TExcludedRecords.All</b> here, but you can OR different members of the TExcludedRecords enumerations for more control on what is saved.</param>
        /// <remarks>This method will not save the file in any readable format, and <b>the file format might change between FlexCel versions.</b>
        /// The only thing it guarantees is that the hashes for 2 identical xls files will be the same, for the same FlexCel version. Once you upgrade version, hashes might have to be rebuilt.
        /// <br></br><br></br>Also note that this method is useful to detect changes when the file is not edited in Excel. If you open the file in Excel and save it again,
        /// Excel will change a lot of reserevd bits, and the files will be too different for this method to have the same hashes. This is only to detect changes
        /// when copying or downloding an xls file. If you want to compare just cell contents, you might compare the files saved as CSV.
        /// </remarks>
        /// <example>
        /// The following method will calculate the hash for an existing file:
        /// <code>
        /// private byte[] GetHash(string FileName)
        /// {
        ///    XlsFile xls = new XlsFile(FileName);
        ///    using (System.Security.Cryptography.SHA1CryptoServiceProvider hasher = new System.Security.Cryptography.SHA1CryptoServiceProvider())
        ///    {
        ///         using (MemoryStream ms = new MemoryStream())
        ///         {
        ///             xls.SaveForHashing(ms);
        ///             ms.Position = 0;
        ///             return hasher.ComputeHash(ms);
        ///         }
        ///     }
        /// }
        /// </code>
        /// </example>
        public abstract void SaveForHashing(Stream aStream, TExcludedRecords excludedRecords);

		/// <summary>
		/// Determines if a call to "Save()" will automatically overwrite an existing file or not.
		/// </summary>
		public bool AllowOverwritingFiles {get {return FAllowOverwritingFiles;} set {FAllowOverwritingFiles=value;}}

		/// <summary>
		/// Determines if the file is a template (xlt format instead of xls). Both file formats are nearly identical, but there is an extra record
		/// needed so the file is a proper xlt template.
		/// </summary>
		public abstract bool IsXltTemplate{get;set;}


		#endregion

		#region Sheet Operations

		/// <summary>
		/// The Sheet where we are working on, 1-based(First sheet is 1, not 0). 
		/// Always set this property before working on a file.
		/// You can read or write this value.
		/// </summary>
		public abstract int ActiveSheet{get;set;}

		/// <summary>
		/// The sheet where we are working on, referred by name instead of by index.
		/// To change the active sheet name, use <see cref="SheetName"/>
		/// </summary>
		public abstract string ActiveSheetByName{get;set;}
       
		/// <summary>
		/// Finds a sheet name on the workbook and returns its index.
		/// If it doesn't find the sheet it will raise an exception. 
		/// See <see cref="GetSheetIndex(System.String,System.Boolean)"/> for a method that will not throw an exception.
		/// To change the active sheet by the name, use <see cref="ActiveSheetByName"/>. See also <see cref="GetSheetName"/>
		/// </summary>
		/// <param name="sheetName">Sheet you want to find.</param>
		/// <returns>The sheet index if the sheet exists, or throws an Exception otherwise.</returns>
		public int GetSheetIndex(string sheetName)
		{
			return GetSheetIndex(sheetName, true);
		}

		/// <summary>
		/// Finds a sheet name on the workbook and returns its index.
		/// Depending on the "throwException" parameter, this method will raise
		/// an exception or return -1 if the sheet does not exist.
		/// To change the active sheet by the name, use <see cref="ActiveSheetByName"/>. See also <see cref="GetSheetName"/>
		/// </summary>
		/// <param name="sheetName">Sheet you want to find.</param>
		/// <param name="throwException">When sheetName does not exist, having
		/// throwException = true will raise an Exception. throwException = false will 
		/// return -1 if the sheet is not found.</param>
		public abstract int GetSheetIndex(string sheetName, bool throwException);

		/// <summary>
		/// Returns the sheet name for a given index.
		/// To change the active sheet by the name, use <see cref="ActiveSheetByName"/>. See also <see cref="GetSheetIndex(string)"/>
		/// </summary>
		public abstract string GetSheetName(int sheetIndex);

		/// <summary>
		/// Internal use. Returns the sheets (like sheet1:sheet2) for a given externsheet.
		/// </summary>
		internal abstract void GetSheetsFromExternSheet(int externSheet, out int Sheet1, out int Sheet2, out bool ExternalSheets, out string ExternBookName);

		/// <summary>
		/// The number of sheets on the file.
		/// </summary>
		public abstract int SheetCount{get;}

		/// <summary>
		/// Reads and changes the name of the active sheet. To switch to another sheet by its name, use <see cref="ActiveSheetByName"/>
		/// </summary>
		/// <remarks>
		/// This name must be unique in the file. It can be up to 31 characters long, and it can't contain the
		/// characters &quot;/&quot;, &quot;\&quot;, &quot;?&quot;, &quot;[&quot;, &quot;]&quot;, &quot;*&quot;
		/// or &quot;:&quot;<para></para>
		/// </remarks>
		public abstract string SheetName{get;set;}

		/// <summary>
		/// Returns or sets the codename of a sheet, that is an unique identifier assigned to the sheet when it is created. 
		/// Codenames are useful because they never change once the file is created, and they are what macros reference.
        /// <b>Very important! Don't change a codename once it has ben created if you have macros or other objects that might reference them.</b>
		/// </summary>
        public abstract string SheetCodeName { get; }

		/// <summary>
		/// Sets the visibility of the active sheet.
		/// </summary>
		public abstract TXlsSheetVisible SheetVisible{get;set;}

		/// <summary>
		/// Reads/Writes the zoom of the current sheet.
		/// </summary>
		public abstract int SheetZoom{get;set;}

        /// <summary>
        /// This is the first sheet that will be visible in the bar of sheet tabs at the bottom. Normally you will want this to be 1.
        /// Note that every time you change <see cref="ActiveSheet"/> this value gets reset, because it makes no sense to preserve it.
        /// <br></br>If you want to change it, change it before saving. The same way, to read it, read it just after opening the file.
        /// <br></br>Please also note that if the first sheet you select is hidden, FlexCel will ignore this value and select a visible sheet. (otherwise Excel would crash)
        /// </summary>
        public abstract int FirstSheetVisible { get; set; }

		/// <summary>
		/// Reads/Writes the color of the current sheet tab. <see cref="TExcelColor.Automatic"/> to specify no color.
		/// </summary>
		public abstract TExcelColor SheetTabColor{get;set;}

        /// <summary>
        /// Inserts an empty sheet at the end of the file. This is equivalent to calling InsertAndCopySheets(0, SheetCount, 1).
        /// If you need to insert more than one sheet, or insert it at the middle of existing sheets, use <see cref="InsertAndCopySheets(int, int, int)"/> instead.
        /// </summary>
        public void AddSheet()
        {
            InsertAndCopySheets(0, SheetCount, 1);
        }

		/// <summary>
		/// Inserts and copies sheet "CopyFrom", SheetCount times before InsertBefore.
		/// To insert empty sheets, set CopyFrom = 0 (You migth also call <see cref="AddSheet"/>).
		/// </summary>
		/// <param name="copyFrom">The sheet index of the sheet we are copying from (1 based). Set it to 0 to insert Empty sheets.</param>
		/// <param name="insertBefore">The sheet before which we will insert (1 based). This might be SheetCount+1, to insert at the end of the Workbook.</param>
		/// <param name="aSheetCount">The number of sheets to insert.</param>
		public void InsertAndCopySheets (int copyFrom, int insertBefore, int aSheetCount)
		{
			InsertAndCopySheets(copyFrom, insertBefore, aSheetCount, null);
		}
        
		/// <summary>
		/// Inserts and copies sheet "CopyFrom", SheetCount times before InsertBefore.
		/// If sourceWorkbook is not null, sheets will be copied from another file.
		/// To insert empty sheets, set CopyFrom=0. 
		/// </summary>
		/// <remarks>When copying sheets from another file, the algorithm used is not so efficient as
		/// when copying from the same file. Whenever possible, don't copy from different files.</remarks>
		/// <param name="copyFrom">The sheet index of the sheet we are copying from (1 based). Set it to 0 to insert Empty sheets.</param>
        /// <param name="insertBefore">The sheet before which we will insert (1 based). This might be SheetCount+1, to insert at the end of the Workbook.</param>
		/// <param name="aSheetCount">The number of sheets to insert.</param>
		/// <param name="sourceWorkbook">Workbook from where the sheet will be copied from. Null to copy from the same file.</param>
		public abstract void InsertAndCopySheets (int copyFrom, int insertBefore, int aSheetCount, ExcelFile sourceWorkbook);

		/// <summary>
		/// Inserts and copies all the sheets in the "CopyFrom" array, before InsertBefore into another workbook.
		/// </summary>
		/// <remarks>When copying sheets from another file, the algorithm used is not so efficient as
		/// when copying from the same file. 
		/// <p></p>
		/// Use this version of the method only when copying from 2 different files. If you have for example a formula "=Sheet2!A1" in Sheet1,
		/// and another formula "=Sheet1!A1" in Sheet2, you need to use this method to copy them, as Copying them one by one
		/// (with the standard <see cref="FlexCel.Core.ExcelFile.InsertAndCopySheets(System.Int32,System.Int32,System.Int32,FlexCel.Core.ExcelFile)"/>
		/// method) will return an error of formula references not found.
		/// <p></p>
		/// You can also use <see cref="ConvertFormulasToValues"/> method to avoid formulas before copying between workbooks.
		/// </remarks>
		/// <param name="copyFrom">The sheet index of all the sheets we are copying from (1 based).</param>
        /// <param name="insertBefore">The sheet before which we will insert (1 based). This might be SheetCount+1, to insert at the end of the Workbook.</param>
		/// <param name="sourceWorkbook">Workbook from where the sheet will be copied from. On this overloaded version it cannot be null.</param>
		public abstract void InsertAndCopySheets (int[] copyFrom, int insertBefore, ExcelFile sourceWorkbook);

		/// <summary>
		/// Clears all data on the active sheet, but does not delete it. <seealso cref="DeleteSheet"/>
		/// </summary>
		public abstract void ClearSheet();

		/// <summary>
		/// Deletes the active sheet and aSheetCount-1 sheets more to the right. 
		/// It will change all formula references to those sheets to invalid. 
        /// Note that to add a sheet, you need to use <see cref="InsertAndCopySheets(int, int, int)"/>
        /// <seealso cref="ClearSheet"/>
		/// </summary>
		/// <param name="aSheetCount">The number of sheets to delete</param>
		public abstract void DeleteSheet(int aSheetCount);

		/// <summary>
		/// True if the gray grid lines are shown on the Active sheet. You can also set this option with <see cref="SheetOptions"/>
		/// </summary>
		public abstract bool ShowGridLines{get;set;}

		/// <summary>
		/// When true, the formula text will be displayed instead of the formula value. You can also set this option with <see cref="SheetOptions"/>
		/// </summary>
		public abstract bool ShowFormulaText{get;set;}

		/// <summary>
		/// Color of the grid separator lines.
		/// </summary>
		public abstract TExcelColor GridLinesColor{get;set;}

		/// <summary>
		/// When true number 0 will be shown as empty. You can also set this option with <see cref="SheetOptions"/>
		/// </summary>
		public abstract bool HideZeroValues{get;set;}

		/// <summary>
		/// Use this property to know it the <see cref="ActiveSheet"/> is a worksheet, a chart sheet or other.
		/// </summary>
		public abstract TSheetType SheetType{get;}

		/// <summary>
		/// This property groups a lot of properties of the sheet, like for example if it is showing formula texts or the results.
		/// Most of this properties can be changed directly from XlsFile, but this method allows you to change them all together,
		/// or to easily copy the options from one file to another.  Look also at <see cref="SheetWindowOptions"/> for options that affect all sheets.
		/// </summary>
		public abstract TSheetOptions SheetOptions{get;set;}

		/// <summary>
		/// This property groups a lot of properties of all the sheets in the workbook, like for example if the sheet tab bar at the bottom is visible.
		///  Look also at <see cref="SheetOptions"/> for options that affect only the active sheet.
		/// </summary>
		public abstract TSheetWindowOptions SheetWindowOptions{get;set;}

		#endregion

		#region Page Breaks
		/// <summary>
		/// True if the sheet has a Manual Horizontal page break on the row.
		/// </summary>
		/// <param name="row">Row to check.</param>
		/// <returns></returns>
		public abstract bool HasHPageBreak(int row);
		/// <summary>
		/// True if the sheet has a Manual Vertical page break on the column.
		/// </summary>
		/// <param name="col">Column to check</param>
		/// <returns></returns>
		public abstract bool HasVPageBreak(int col);

        /// <inheritdoc cref="InsertVPageBreak(int, bool)" />
        public void InsertHPageBreak(int row)
        {
            InsertHPageBreak(row, false);
        }

        /// <summary>
        /// Inserts an Horizontal Page Break at the specified row. If there is one already, it will do nothing.
        /// If the number of pagebreaks is bigger than the maximum Excel can admit, it will add it anyway, but you might get an 
        /// exception when saving the file as xls. Exporting as images or PDF will use those additional page breaks.
        /// To control what to do when there are too many page breaks, see <see cref="ErrorActions"/>
        /// </summary>
        /// <param name="row">Row where to insert the page break. All row numbers are 1-based, and the breaks occur after the row.</param>
        /// <param name="aGoesAfter">This is used by FlexCelReport to add page breaks that behave as if they affected the next column, not the column to the left.</param>
        internal abstract void InsertHPageBreak(int row, bool aGoesAfter);

        /// <inheritdoc cref="InsertVPageBreak(int, bool)" />
        public void InsertVPageBreak(int col)
        {
            InsertVPageBreak(col, false);
        }

        /// <summary>
        /// Inserts a Vertical Page Break at the specified column. If there is one already, it will do nothing.
        /// If the number of pagebreaks is bigger than the maximum Excel can admit, it will add it anyway, but you might get an 
        /// exception when saving the file as xls. Exporting as images or PDF will use those additional page breaks.
        /// To control what to do when there are too many page breaks, see <see cref="ErrorActions"/>
        /// </summary>
        /// <param name="col">Column where to insert the page break All column numbers are 1-based, and the breaks occur after the column.</param>
        /// <param name="aGoesAfter">This is used by FlexCelReport to add page breaks that behave as if they affected the next column, not the column to the left.</param>
        internal abstract void InsertVPageBreak(int col, bool aGoesAfter);
        
		/// <summary>
		/// Deletes all manual page breaks at row. If there is no manual page break on row, this method will do nothing.
		/// </summary>
		/// <param name="row">Row where to delete the Page break.</param>
		public abstract void DeleteHPageBreak(int row);
		/// <summary>
		/// Deletes all manual page breaks at col. If there is no manual page break on col, this method will do nothing.
		/// </summary>
		/// <param name="col">Column where to delete the Page break</param>
		public abstract void DeleteVPageBreak(int col);

		/// <summary>
		/// Deletes all manual page breaks on the active sheet.
		/// </summary>
		public abstract void ClearPageBreaks();

		#endregion

		#region Intelligent Page Breaks

        /// <summary>
        /// Tells FlexCel that it must try to keep together the rows between row1 and row2 (inclusive) when printing.  
        /// This method does nothing to the resulting Excel file since this is not an Excel feature. To actually
        /// do something, you need to call <see cref="AutoPageBreaks()"/> after calling this method.
        /// </summary>
        /// <example>
        /// If you call:
        /// <code>
        ///   XlsFile.KeepRowsTogether(1, 10, 1, true);
        ///   XlsFile.KeepRowsTogether(3, 4, 2, true);
        /// </code>
        /// This will define two groups. FlexCel will try to keep rows 1 to 10 in one page. But if this is impossible, it will try to keep at least rows 3 and 4 together.<br/>
        /// Values of the "level" parameter mean the strongness of the link. Rows with higher level values have a stronger link, and if it is impossible to keep all rows together, FlexCel will try to keep the higher level rows.<br/>
        /// <br/>In this case, Rows 3 and 4 have a higher level, so if rows 1 to 10 cannot be kept together in one page, FlexCel will try to keep at least rows 3 and 4.
        /// Normally you will want to use higher level values completely inside groups with lower levels.
        /// If we called:
        /// <code>
        ///   XlsFile.KeepRowsTogether(1, 10, 1, true);
        ///   XlsFile.KeepRowsTogether(3, 4, 1, true);
        /// </code>
        /// The second call would do nothing, since the level is the same, and rows 1 to 10 are already linked.
        /// </example>
        /// <param name="row1">First row of the group you want to keep together.</param>
        /// <param name="row2">Last row of the group you want to keep together.</param>
        /// <param name="level">Set this parameter to 0 to remove the condition to keep rows together. Any bigger than zero value will mean that
        /// the rows must be kept together. You can use more than one level to tell FlexCel to try to keep different groups together. 
        /// If all rows cannot be kept together in one page, FlexCel will try to keep as much rows with higher levels as possible. See the example for more information.</param>
        /// <param name="replaceLowerLevels">If true, all existing level values in the row range will be replaced. If false, the new level values will be written 
        /// only if they are bigger than the existing ones. You can use the false setting to set many values in any order. <br/>
        /// For example, if you first call KeepRowsTogether(2, 3, 5, false) and then KeepRowsTogether(1, 10, 1, false), rows 2 and 3 will keep the
        /// level in 5. If you did so with this parameter true, the second call would replace the levels of rows 2 and 3 to level 1, making all row levels between 1 and 10 equal to 1.</param>
        public abstract void KeepRowsTogether(int row1, int row2, int level, bool replaceLowerLevels);

        /// <summary>
        /// Tells FlexCel that it must try to keep together the columns between col1 and col2 (inclusive) when printing.  
        /// This method does nothing to the resulting Excel file since this is not an Excel feature. To actually
        /// do something, you need to call <see cref="AutoPageBreaks()"/> after calling this method.
        /// </summary>
        /// <example>
        /// If you call:
        /// <code>
        ///   XlsFile.KeepColsTogether(1, 10, 1, true);
        ///   XlsFile.KeepColsTogether(3, 4, 2, true);
        /// </code>
        /// This will define two groups. FlexCel will try to keep columns 1 to 10 in one page. But if this is impossible, it will try to keep at least columns 3 and 4 together.<br/>
        /// Values of the "level" parameter mean the strongness of the link. Columns with higher level values have a stronger link, and if it is impossible to keep all columns together, FlexCel will try to keep the higher level columns.<br/>
        /// <br/>In this case, Columns 3 and 4 have a higher level, so if columns 1 to 10 cannot be kept together in one page, FlexCel will try to  keep at least columns 3 and 4.
        /// Normally you will want to use higher level values completely inside groups with lower levels.
        /// If we called:
        /// <code>
        ///   XlsFile.KeepColsTogether(1, 10, 1, true);
        ///   XlsFile.KeepColsTogether(3, 4, 1, true);
        /// </code>
        /// The second call would do nothing, since the level is the same, and columns 1 to 10 are already linked.
        /// </example>
        /// <param name="col1">First column of the group you want to keep together.</param>
        /// <param name="col2">Last column of the group you want to keep together.</param>
        /// <param name="level">Set this parameter to 0 to remove the condition to keep columns together. Any bigger than zero value will mean that
        /// the rows must be kept together. You can use more than one level to tell FlexCel to try to keep different groups together,
        /// If all columns cannot be kept together in one page, FlexCel will try to keep as much columns with higher levels as possible. See the example for more information.</param>
        /// <param name="replaceLowerLevels">If true, all existing level values in the column range will be replaced. If false, the new level values will be written 
        /// only if they are bigger than the existing ones. You can use the false setting to set many values in any order. <br/>
        /// For example, if you first call KeepColsTogether(2, 3, 5,false) and then KeepColsTogether(1, 10, 1, false), columns 2 and 3 will keep the
        /// level in 5. If you did so with this parameter true, the second call would replace the levels of columns 2 and 3 to level 1, making all column levels between 1 and 10 equal to 1.</param>
        public abstract void KeepColsTogether(int col1, int col2, int level, bool replaceLowerLevels);

        /// <summary>
        /// Clears all the "KeepTogether" links in the current page.
        /// </summary>
        public void ClearKeepRowsAndColsTogether()
        {
            if (RowCount > 0)
            {
                KeepRowsTogether(1, RowCount, 0, true);
                KeepColsTogether(1, ColCount, 0, true);
            }
        }

        /// <summary>
        /// Returns the value of level for a row as set in <see cref="KeepRowsTogether"/>. Note that the last value of a "keep together" range is 0.
        /// For example, if you set KeepRowsTogether(1, 3, 8, true); GetKeepRowsTogether will return 8 for rows 1 and 2, and 0 for row 3.
        /// </summary>
        /// <param name="row">Row index. (1 based)</param>
        /// <returns>The Keep together level of the row.</returns>
        public abstract int GetKeepRowsTogether(int row);

        /// <summary>
        /// Returns the value of level for a column as set in <see cref="KeepColsTogether"/>. Note that the last value of a "keep together" range is 0.
        /// For example, if you set KeepColsTogether(1, 3, 8, true); GetKeepColsTogether will return 8 for columns 1 and 2, and 0 for column 3.
        /// </summary>
        /// <param name="col">Column index. (1 based)</param>
        /// <returns>The Keep together level of the column.</returns>
        public abstract int GetKeepColsTogether(int col);

		/// <summary>
		/// Returns true if there is any row marked as keeptogether in the sheet.
		/// This method traverses every row to find out, so it acn be somehow slow and you should not call it too often.
		/// </summary>
		/// <returns></returns>
		public abstract bool HasKeepRowsTogether();

		/// <summary>
		/// Returns true if there is any column marked as keeptogether in the sheet.
		/// This method traverses every column to find out, so it acn be somehow slow and you should not call it too often.
		/// </summary>
		public abstract bool HasKeepColsTogether();



        /// <summary>
        /// This method will create manual page breaks in the sheet to try to keep together the rows and columns marked with 
        /// <see cref="KeepRowsTogether"/> and <see cref="KeepColsTogether"/>.
        /// It might be desirable to clear all manual page breaks (with <see cref="ClearPageBreaks"/>) before calling this method, so it has more freedom
        /// to place the new ones. If you call this method twice without removing the old page breaks, it will add the page breaks to the existing ones.
        /// </summary>
        /// <remarks>This methood is the same as calling AutoPageBreaks(20, 95)</remarks>
        public void AutoPageBreaks()
        {
            AutoPageBreaks(20, 95);
        }


        /// <summary>
        /// This method will create manual page breaks in the sheet to try to keep together the rows and columns marked with 
        /// <see cref="KeepRowsTogether"/> and <see cref="KeepColsTogether"/>.
        /// It might be desirable to clear all manual page breaks (with <see cref="ClearPageBreaks"/>) before calling this method, so it has more freedom
        /// to place the new ones. If you call this method twice without removing the old page breaks, it will add the page breaks to the existing ones.
        /// </summary>
        /// <param name="PercentOfUsedSheet">Percentage of the sheet that must be used in any page when fitting the rows and columns.
        /// A value of zero means that no part of the sheet must be used, so FlexCel might add a page break after a single row in a page, leaving it almost completely blank.<br/>
        /// A value of 50% means that half of the page must be used. This means that FlexCel will add a page break only if there is 50% of the current page already used.<br/>
        /// A value of 100% will do nothing, since the sheet must be completely used, and so FlexCel can never add a page break.<br/></param>
        /// <param name="PageScale">This parameter must be between 50 and 100, and means how much smaller page will be considered in order to calculate the page breaks.
        /// <br/> A value of 100 means that the size used in the calculation will be the real size of the page, and while this will always work fine when exporting to pdf 
        /// or exporting to images, when printing from Excel might result in a page break that is placed a little after where it should go and an empty page for certain printers.
        /// (Page size in Excel is different for different printers) Normally a value around 95 is the recommended value for this parameter.<br/>
        /// If you need to do a finer grain adjustment, you can use <see cref="AutoPageBreaks(int, RectangleF)"/>.
        /// </param>
        public abstract void AutoPageBreaks(int PercentOfUsedSheet, int PageScale);

        /// <summary>
        /// This method will create manual page breaks in the sheet to try to keep together the rows and columns marked with 
        /// <see cref="KeepRowsTogether"/> and <see cref="KeepColsTogether"/>.
        /// It might be desirable to clear all manual page breaks (with <see cref="ClearPageBreaks"/>) before calling this method, so it has more freedom
        /// to place the new ones. If you call this method twice without removing the old page breaks, it will add the page breaks to the existing ones.
        /// </summary>
        /// <param name="PercentOfUsedSheet">Percentage of the sheet that must be used in any page when fitting the rows and columns.
        /// A value of zero means that no part of the sheet must be used, so FlexCel might add a page break after a single row in a page, leaving it almost completely blank.<br/>
        /// A value of 50% means that half of the page must be used. This means that FlexCel will add a page break only if there is 50% of the current page already used.<br/>
        /// A value of 100% will do nothing, since the sheet must be completely used, and so FlexCel can never add a page break.<br/></param>
        /// <param name="PageBounds">You can customize a custom page size here. If width or height of this parameter is 0, the paper size specified in the file
        /// will be used. There is normaly no need to set this parameter, unless you want to fine tune the results.</param>
        public abstract void AutoPageBreaks(int PercentOfUsedSheet, RectangleF PageBounds);

		#endregion

		#region Cell Value
		/// <summary>
		/// Reads a Cell Value and Format.
		/// </summary>
		/// <param name="row">Row, 1 based.</param>
		/// <param name="col">Column, 1 based.</param>
		/// <param name="XF">XF format.</param>
		/// <returns>Object with the value. It can be null, a double, a string, a boolean, a
		///  <see cref="FlexCel.Core.TFormula"/>, a <see cref="FlexCel.Core.TFlxFormulaErrorValue"/> or
		///  a <see cref="FlexCel.Core.TRichString"/>. Dates are returned as doubles. See the Reading Files demo to know how to use each type of the objects returned.</returns>
		public abstract object GetCellValue(int row, int col, ref int XF);

		/// <summary>
		/// Reads a Cell Value and Format from a sheet that is not the active sheet.
		/// </summary>
		/// <param name="sheet">Sheet where is the cell you want to get the value.</param>
		/// <param name="row">Row, 1 based.</param>
		/// <param name="col">Column, 1 based.</param>
		/// <param name="XF">XF format.</param>
		/// <returns>Object with the value. It can be null, a double, a string, a boolean, a
		///  <see cref="FlexCel.Core.TFormula"/>, a <see cref="FlexCel.Core.TFlxFormulaErrorValue"/> or
		///  a <see cref="FlexCel.Core.TRichString"/>. Dates are returned as doubles. See the Reading Files demo to know how to use each type of the objects returned.</returns>
		public abstract object GetCellValue(int sheet, int row, int col, ref int XF);

		/// <summary>
		/// Reads a Cell Value and Format, using a column index for faster access. Normal GetCellValue(row, col)
		/// has to search for the column on a sorted list. If you are looping from 1 to <see cref="ColCountInRow(int)"/>
		/// this method is faster.
		/// </summary>
        /// <param name="sheet">Sheet where the cell is, 1 based.</param>
		/// <param name="row">Row, 1 based.</param>
		/// <param name="colIndex">Column index, 1 based.</param>
		/// <param name="XF">XF format.</param>
		/// <returns>Object with the value. It can be null, a double, a string, a boolean, a
		///  <see cref="FlexCel.Core.TFormula"/>, a <see cref="FlexCel.Core.TFlxFormulaErrorValue"/> or
		///  a <see cref="FlexCel.Core.TRichString"/>. Dates are returned as doubles. See the Reading Files demo to know how to use each type of the objects returned.</returns>
		public abstract object GetCellValueIndexed(int sheet, int row, int colIndex, ref int XF);

        /// <inheritdoc cref="GetCellValueIndexed(int, int, int, ref int)" />
        public object GetCellValueIndexed(int row, int colIndex, ref int XF)
        {
            return GetCellValueIndexed(ActiveSheet, row, colIndex, ref XF);
        }

        /// <summary>
        /// This is used in virtual mode to know how many sheets have already been loaded.
        /// </summary>
        /// <returns></returns>
        internal abstract int PartialSheetCount();


		/// <summary>
		/// This is used internally to get the value of another part of the workbook.
		/// No checks are made, and we try to recalculate the value before sending it.
		/// </summary>
        internal abstract object GetCellValueAndRecalc(int sheet, int row, int col, TCalcState CalcState, TCalcStack CalcStack);

		/// <summary>
		/// Sets the value and format on a cell.
		/// </summary>
		/// <remarks>This method will enter the datatype of the object you pass to it. For example, if you set value="1"
		/// the string "1" will be entered on the cell. To convert a string to the best representation (on this case a number), use <see cref="SetCellFromString(int, int, TRichString, int)"/>
		/// To enter a HTML formatted string, use <see cref="SetCellFromHtml(int, int, string, int)"/> 
		/// </remarks>
		/// <param name="row">Row, 1 based.</param>
		/// <param name="col">Column, 1 based.</param>
		/// <param name="value">Value to set.</param>
		/// <param name="XF">Format to Set. You normally get this number with <see cref="AddFormat"/> function. Use -1 to keep format unchanged.</param>
		public abstract void SetCellValue(int row, int col, object value, int XF);

		/*/// <summary>
		/// Sets a value on a range of cells. If value is a single object, it will be repeated on all the range. If it is 
		/// an array, each element will be entered into the corresponding position. If it is a DataSet, it will fill the
		/// range with values form the dataset.
		/// While this method is normally a shortcut for calling SetCellValue in a for loop, this method is the
		/// only way to enter an ARRAY FORMULA that spans over more than one cell.
		/// </summary>
		/// <param name="row1">First row where to enter the data. (1 based)</param>
		/// <param name="col1">First column where to enter the data. (1 based)</param>
		/// <param name="row2">Last row where to enter the data. If this value is &lt; 1 and value is
		/// a dataset or an array, then all the data on the array or dataset will be entered.</param>
		/// <param name="col2">Last column where to enter the data. If this value is &lt; 1 and value is
		/// a dataset or an array, then all the data on the array or dataset will be entered.</param>
		/// <param name="values">Value to enter. It might be a single object, an array or a DataSet. It might also be an ARRAY FORMULA.</param>
		/// <remarks>
		/// This method does NOT improve performance over multiple calls to SetCellValue. Use it if the data *already*
		/// is on an array or a DataSet, but do not fill an array with the data to later call this method. If the
		/// data is not in an array or dataset, just loop trough it and call SetCellValue(row, col, value) for each value.
		/// </remarks>
		/// <example>
		/// To enter an matrix with the values {1,2},{3,4} on cells a1:b2 and an Array formula with the transposed results
		/// on cells c1:d2, use the following code:
		/// <code>
		/// int[,] Values = {1,2;3,4};
		/// xls.SetCellValue(1, 1, 2, 2, Values);
		/// xls.SetCellValue(1, 3, 2, 5, new TFormula("=Transpose(A1:B2)"));
		/// </code>
		/// </example>
		public abstract void SetCellValue(int row1, int col1, int row2, int col2, object values);*/

		/// <summary>
		/// Reads a Cell Value.
		/// </summary>
		/// <remarks>This method will return the real value stored on the cell. For example, if you have "1.3" formatted as "1.30",
		/// GetCellValue will return the number 1.3. To get a string with the formatted value, see <see cref="GetStringFromCell(int, int)"/></remarks>
		/// <param name="row">Row, 1 based.</param>
		/// <param name="col">Column, 1 based.</param>
		/// <returns>Object with the value. It can be null, a double, a string, a boolean, a
		///  <see cref="FlexCel.Core.TFormula"/>, a <see cref="FlexCel.Core.TFlxFormulaErrorValue"/> or
		///  <see cref="FlexCel.Core.TRichString"/>. Dates are returned as doubles. See the Reading Files demo to know how to use each type of the objects returned.</returns>
		public object GetCellValue(int row, int col)
		{
			int tempXF=-1; 
			return GetCellValue(row, col, ref tempXF);
		}

		/// <summary>
		/// Sets the value on a cell.
		/// </summary>
		/// <remarks>This method will enter the datatype of the object you pass to it. For example, if you set value="1"
		/// the string "1" will be entered on the cell. To convert a string to the best representation (on this case a number), use <see cref="SetCellFromString(int, int, TRichString, int)"/>.
		/// To enter a HTML formatted string, use <see cref="SetCellFromHtml(int, int, string, int)"/> 
		/// </remarks>
		/// <param name="row">Row, 1 based.</param>
		/// <param name="col">Column, 1 based.</param>
		/// <param name="value">Value to set.</param>
		public void SetCellValue(int row, int col, object value)
		{
			SetCellValue(row, col, value, -1);
		}

        /// <summary>
        /// Sets the value on a cell.
        /// </summary>
        /// <remarks>This method will enter the datatype of the object you pass to it. For example, if you set value="1"
        /// the string "1" will be entered on the cell. To convert a string to the best representation (on this case a number), use <see cref="SetCellFromString(int, int, TRichString, int)"/>.
        /// To enter a HTML formatted string, use <see cref="SetCellFromHtml(int, int, string, int)"/> 
        /// </remarks>
        /// <param name="sheet">Sheet number, 1 based</param>
        /// <param name="row">Row, 1 based.</param>
        /// <param name="col">Column, 1 based.</param>
        /// <param name="value">Value to set.</param>
        /// <param name="XF">Format to Set. You normally get this number with <see cref="AddFormat"/> function. Use -1 to keep format unchanged.</param>
        public abstract void SetCellValue(int sheet, int row, int col, object value, int XF);

        /// <inheritdoc cref="ConvertString(TRichString, ref int, string[])" />
        public object ConvertString(TRichString value, ref int XF)
        {
            return ConvertString(value, ref XF, null);
        }

        /// <summary>
        /// Converts a string into the best datatype (a boolean, a number, etc)
        /// </summary>
        /// <remarks>See <see cref="SetCellFromString(int, int, TRichString, int)"/> for more information.</remarks>
        /// <param name="value">RichString to convert.</param>
        /// <param name="XF">XF of the cell. It might be modified, for example, if you are entering a date.</param>
        /// <returns>value converted to the best datatype.</returns>
        /// <param name="dateFormats">A list of formats allowed for dates and times. Windows is a little liberal in what it thinks can be a date, and it can convert things 
        /// like "1.2" into dates. By setting this property, you can ensure the dates are only in the formats you expect. If you leave it null, we will trust "DateTime.TryParse" to guess the correct values.</param>
        public abstract object ConvertString(TRichString value, ref int XF, string[] dateFormats);

        /// <inheritdoc cref="SetCellFromString(int, int, TRichString, int, string[])" />
        public void SetCellFromString(int row, int col, string value)
        {
            SetCellFromString(row, col, new TRichString(value));
        }

        /// <inheritdoc cref="SetCellFromString(int, int, TRichString, int, string[])" />
		public void SetCellFromString(int row, int col, TRichString value)
		{
			SetCellFromString(row, col, value, -1, null);
		}

        /// <inheritdoc cref="SetCellFromString(int, int, TRichString, int, string[])" />
        public void SetCellFromString(int row, int col, string value, string[] dateFormats)
        {
            SetCellFromString(row, col, new TRichString(value), dateFormats);
        }

        /// <inheritdoc cref="SetCellFromString(int, int, TRichString, int, string[])" />
        public void SetCellFromString(int row, int col, TRichString value, string[] dateFormats)
        {
            SetCellFromString(row, col, value, -1, dateFormats);
        }

        /// <inheritdoc cref="SetCellFromString(int, int, TRichString, int, string[])" />
        public void SetCellFromString(int row, int col, string value, int XF)
        {
            SetCellFromString(row, col, value, XF, null);
        }

        /// <inheritdoc cref="SetCellFromString(int, int, TRichString, int, string[])" />
        public void SetCellFromString(int row, int col, TRichString value, int XF)
        {
            SetCellFromString(row, col, value, XF, null);
        }

        /// <inheritdoc cref="SetCellFromString(int, int, TRichString, int, string[])" />
        public void SetCellFromString(int row, int col, string value, int XF, string[] dateFormats)
        {
            SetCellFromString(row, col, new TRichString(value), XF, dateFormats);
        }

        /// <summary>
        /// Converts a string to the best datatype, and the enters it into a cell.
        /// </summary>
        /// <remarks>
        /// When using CellValue to set a cell value, you have to know the datatype you want to enter. 
        /// That is, if you have a string s="1/1/2002" and call SetCellValue(1,1,s); the cell A1 will 
        /// end up with a string "1/1/2002" and not with a date. The same if you have a string holding a number.
        /// <para></para>
        /// <para>SetCellFromString tries to solve this problem. When you call SetCellFromString(1,1,s) it will see:</para>
        /// <list type="number">
        ///  <item> If s contains a valid number: If it does, it will enter the number into the cell, and not the string s</item>
        ///  <item> If s contains a boolean:      If s equals the words "TRUE" or "FALSE" (or whatever you define on the constants TxtTrue and TxtFalse) it will enter the boolean into the cell</item>
        ///  <item> If s contains a date:         If s is a valid date (according to your windows settings, or with a list of allowed date/time formats) it will enter a number into the cell and format the cell as a date. (see UsingFlexCelAPI.pdf)</item>
        ///  <item> In any other case, it will enter the string s into the cell.</item>
        ///  </list>
        /// </remarks>
        /// <param name="row">Cell Row (1 based)</param>
        /// <param name="col">Cell Column (1 based)</param>
        /// <param name="value">Value to enter into the cell.</param>
        /// <param name="XF">New XF of the cell. It can be modified, i.e. if you enter a date, the XF will be converted to a Date XF.</param>
        /// <param name="dateFormats">A list of formats allowed for dates and times. Windows is a little liberal in what it thinks can be a date, and it can convert things 
        /// like "1.2" into dates. By setting this property, you can ensure the dates are only in the formats you expect. If you leave it null, we will trust "DateTime.TryParse" to guess the correct values.</param>
        public abstract void SetCellFromString(int row, int col, TRichString value, int XF, string[] dateFormats);


        /// <inheritdoc cref="GetStringFromCell(int, int, ref int, ref Color)" />
        public TRichString GetStringFromCell(int row, int col)
		{
			int XF=-1;
			Color aColor=ColorUtil.Empty;
			return GetStringFromCell(row, col, ref XF, ref aColor);
		}

        /// <summary>
        /// This method will return a rich string that is formatted similar to the way Excel shows it.
        /// For example, if you have "1.0" on a cell, and the cell is formatted as exponential, this method will return "1.0e+1"
        /// It might also change the color depending on the value and format. (for example, red for negative numbers) 
        /// </summary>
        /// <param name="row">Cell Row (1 based). </param>
        /// <param name="col">Cell Column (1 based)</param>
        /// <param name="XF">The resulting XF for the cell.</param>
        /// <param name="aColor">Resulting color of the string. If for example you define red for negative numbers, and the result is red, this will
        /// be returned on aColor. If there is not color info on the format, it will remain unchanged.</param>
        /// <returns>A rich string with the cell value.</returns>
        public abstract TRichString GetStringFromCell(int row, int col, ref int XF, ref Color aColor);

        /// <inheritdoc cref="SetCellFromHtml(int, int, string, int)" />
        public void SetCellFromHtml(int row, int col, string htmlText)
		{
			SetCellFromHtml(row, col, htmlText, -1);
		}

        /// <summary>
        /// Enters an HTML formatted string into a cell, and tries to match the Excel formats with the Html formatting tags.
        /// Note that the rich text inside Excel is more limited than xls (you are limited to only changing font attributes),
        /// so many tags from the HTML tags might be ignored. Whenever a tag is not understood or cannot be mapped into Excel,
        /// it will just be omitted. For a list of supported tags, see the <b>Remarks</b> section.
        /// </summary>
        /// <remarks>
        /// Any tag (between brackets) will be stripped from the entered text. If you want to
        /// enter a &lt; sign, you need to enter "&amp;lt;". For a list of all "&amp;" possible symbols
        /// consult any HTML reference. &amp;[HexadecimalUnicodeCharacer] are supported too. <br/>
        /// All tags are case insensitive. <br/>
        /// Spaces will be eliminated from the string as in normal HTML. To enter a "hard" space you can enter "&amp;nbsp;"
        /// (non breaking space). Also, if text is inside "pre" tags, spaces will be preserved.
        /// <br/><br/>
        /// 
        /// The supported tags are: 
        /// <list type="bullet">
        /// <item><b>b</b> or <b>strong</b> - Bold.</item>
        /// <item><b>i</b> or <b>em</b>- Italics.</item>
        /// <item><b>u</b> - Underline.</item>
        /// <item><b>s</b> or <b>strike</b>- Strikeout.</item>
        /// <item><b>sub</b> - Subscript.</item>
        /// <item><b>sup</b> - Superscript.</item>
        /// <item><b>tt</b> - Use monospace font.</item>
        /// <item><b>pre</b> - Preserve spaces and returns. On normal mode, if you have 2 spaces on a string, only one will be kept.</item>
        /// <item><b>font</b> - Change the used font. This tag behaves as normal HTML, and you can specify 
        /// <i>color</i>, <i>face</i>, <i>point-size</i> or <i>size</i> as attributes. Size might be between 1 and 7, 
        /// and the size in points are 8, 9, 12, 14, 18, 24, 34 for each size. You can also specify relative sizes (for example -1)</item>
        /// <item><b>h1</b>..<b>h6</b> - Header fonts.</item>
        /// <item><b>small</b> - Use a smaller font. This is equivalent to &lt;font size = '-1'&gt;</item>
        /// <item><b>big</b> - Use a bigger font. This is equivalent to &lt;font size = '+1'&gt;</item>
        /// </list>
        /// </remarks>
        /// <param name="row">Cell Row (1 based)</param>
        /// <param name="col">Cell Column (1 based)</param>
        /// <param name="htmlText">Text with an html formatted string.</param>
        /// <param name="XF">Format for the cell. It can be -1 to keep the existing format.</param>
		public abstract void SetCellFromHtml(int row, int col, string htmlText, int XF);

        /// <summary>
		/// Returns the contents of formatted cell as HTML. Conditional formats are not applied, you need to call <see cref="ConditionallyModifyFormat"/> to the cell style for that. 
		/// If <see cref="ShowFormulaText"/> is true, it will return the formula text instead of the value.
		///  For a list of html tags that might be returned, see the <b>Remarks</b> section.
		/// </summary>
		/// <remarks>
		/// Text will be returned as standard HTML, on the version specified. This means that 
		/// for example a &lt; sign will be returned as "&amp;lt;". For a list of all "&amp;" possible symbols
		/// consult any HTML reference. &amp;[HexadecimalUnicodeCharacer] might be returned too. <br/>
		/// <br/>
		/// Multiple spaces will be returned as "&amp;nbsp;"
		/// (non breaking space) and line breaks will be returned as &lt;br&gt; tags. 
		/// <br/><br/>
		/// 
		/// The tags that might be returned by this method are: 
		/// If htmlStyle is Simple:
		/// <list type="bullet">
		/// <item><b>b</b> - Bold.</item>
		/// <item><b>i</b> - Italics.</item>
		/// <item><b>u</b> - Underline.</item>
		/// <item><b>s</b> - Strikeout.</item>
		/// <item><b>sub</b> - Subscript.</item>
		/// <item><b>sup</b> - Superscript.</item>
		/// <item><b>font</b> - Change the used font. This tag behaves as normal HTML, and you can specify 
		/// <i>color</i>, <i>face</i>, <i>point-size</i> or <i>size</i> as attributes. Size might be between 1 and 7, 
		/// and the size in points are 8, 9, 12, 14, 18, 24, 34 for each size. You can also specify relative sizes (for example -1)</item>
		/// </list>
		/// <br/>
		/// If htmlStlye is Css:
		/// The equivalent CSS style commands to the ones listed under htmlStyle = Simple.
		/// </remarks>
		/// <param name="row">Cell Row (1 based)</param>
		/// <param name="col">Cell Column (1 based)</param>
		/// <param name="htmlVersion">Version of the html returned. In XHTML, single tags have a "/" at the end, while in 4.0 they don't.</param>
		/// <param name="htmlStyle">Specifies if to use simple tags or for the returned HTML.</param>
		/// <param name="encoding">Encoding for the returned string. Use UTF-8 if in doubt.</param>
		/// <returns>An Html formatted string with the cell contents.</returns>
		public abstract string GetHtmlFromCell(int row, int col, THtmlVersion htmlVersion, THtmlStyle htmlStyle, Encoding encoding);


		/// <summary>
		/// Copies <b>one</b> cell from one workbook to another. If the cell has a formula, it will be offset so it matches the new destination.
		/// <b>Note:</b> You will normally not need this method. To copy a range of cells from a workbook to another use 
		/// <see cref="InsertAndCopyRange(TXlsCellRange, int, int, int, TFlxInsertMode, TRangeCopyMode, ExcelFile, int )"/> instead.
		/// To copy a full sheet from one file to another, use <see cref="InsertAndCopySheets(int, int, int, ExcelFile)"/>.
		/// </summary>
		/// <param name="sourceWorkbook">File from where we want to copy the cell.</param>
		/// <param name="sourceSheet">Sheet in sourceWorkbook where the data is.</param>
		/// <param name="destSheet">Sheet in this file where we want to copy the data.</param>
		/// <param name="sourceRow">Row on the source file of the cell (1 based)</param>
		/// <param name="sourceCol">Column on the source file of the cell (1 based)</param>
		/// <param name="destRow">Row on the destination file of the cell (1 based)</param>
		/// <param name="destCol">Column on the destination file of the cell (1 based)</param>
		/// <param name="copyMode">How the cell will be copied.</param>
		public abstract void CopyCell(ExcelFile sourceWorkbook, int sourceSheet, int destSheet, int sourceRow, int sourceCol, int destRow, int destCol, TRangeCopyMode copyMode);

		#endregion

		#region Cell Format
		#region XF
		/// <summary>
		/// Returns the format definition for a given format index. <b>Note that this method will only return
        /// Cell formats. If you want to read a Style format, use <see cref="GetStyle(int)"/></b>
		/// </summary>
		/// <param name="XF">Format index 0-Based</param>
		/// <returns>Format definition</returns>
		public abstract TFlxFormat GetFormat(int XF);
		
		/// <summary>
		/// Number of custom formats defined in all the sheet. When calling GetFormat(XF), 0&lt;=XF&lt;FormatCount.
		/// </summary>
		public abstract int FormatCount{get;}

		/// <summary>
		/// Returns Excel standard format for an empty cell. (NORMAL format)
		/// </summary>
		public TFlxFormat GetDefaultFormat{get {return GetFormat(DefaultFormatId);}}

		/// <summary>
		/// Returns Excel standard format for the normal style. "Normal" style applies to the headers "A", "B" ... at the top of the columns
		/// and "1", "2"... at the left of the rows. This method is the same as calling 
        /// <i>xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Normal, 0))</i>
		/// You normally will want to use <see cref="GetDefaultFormat"/> instead of this method.
		/// </summary>
        public TFlxFormat GetDefaultFormatNormalStyle { get { return GetStyle(GetBuiltInStyleName(TBuiltInStyle.Normal, 0)); } }

		/// <summary>
		/// Returns XF identifier for the style that applies to all empty cells. Note that this is different
		/// from the "Normal" style as defined inside Excel. This is a "Cell" format that can be applied to cells,
		/// while "Normal" is a "Style" format that is applied to this cell format.
		/// </summary>
		public int DefaultFormatId{get {return FlxConsts.DefaultFormatId;}}

		/// <summary>
		/// Adds a new format to the Excel format list. If it already exists, it doesn't add a new one, so you can use this method for searching too.
		/// </summary>
		/// <param name="format">Format to add to the list.</param>
		/// <returns>Position on the list for the format.</returns>
		public abstract int AddFormat(TFlxFormat format);

		/// <summary>
		/// Sets the font definition for a given format index. Normally it is of not use, (you should use AddFont or AddFormat instead) but could be used
		/// to change the default format. (using SetFormat(0, fmt); ). This emthod will change style XFs and CellXfs, depending if aFormat 
        /// is a StyleXF or a CellXF.
		/// </summary>
		/// <param name="formatIndex">Format index. 0-based</param>
		/// <param name="aFormat">Format definition</param>
		public abstract void SetFormat(int formatIndex, TFlxFormat aFormat);

		#endregion

		#region Font
		/// <summary>
		/// Returns the font definition for a given font index.
		/// </summary>
		/// <param name="fontIndex">Font index. 0-based</param>
		/// <returns>Font definition</returns>
		public abstract TFlxFont GetFont(int fontIndex);

		/// <summary>
		/// Sets the font definition for a given font index. Normally it is of not use, (you should use AddFont or AddFormat instead) but could be used
		/// to change the default font format. (using SetFont(0, font); )
		/// </summary>
		/// <param name="fontIndex">Font index. 0-based</param>
		/// <param name="aFont">Font definition</param>
		public abstract void SetFont(int fontIndex, TFlxFont aFont);

		/// <summary>
		/// Number of fonts defined in all the sheet. When calling GetFont(fontIndex), 0&lt;=fontIndex&lt;FormatCount.
		/// </summary>
		public abstract int FontCount{get;}

		/// <summary>
		/// Returns Excel standard font for an empty cell.
		/// </summary>
		public abstract TFlxFont GetDefaultFont{get;}

		/// <summary>
		/// Returns Excel font for the "normal" style. This style is used to draw the row and column headings.
		/// </summary>
		public abstract TFlxFont GetDefaultFontNormalStyle{get;}

		/// <summary>
		/// Adds a new font to the excel font list.  If it already exists, it doesn't add a new one, so you can use this method for searching too.
		/// </summary>
		/// <param name="font">Font to add to the list.</param>
		/// <returns>The position on the list for the added font.</returns>
		public abstract int AddFont(TFlxFont font);
		#endregion

		/// <summary>
		/// Sets the Cell format (XF) on a given cell.
		/// You can create new formats using the  <see cref="AddFormat"/> function.
		/// <seealso cref="GetCellFormat(System.Int32,System.Int32)"/><seealso cref="AddFormat"/>
		/// </summary>
		/// <param name="row">Row index of the cell (1 based)</param>
		/// <param name="col">Column index of the cell (1 based)</param>
		/// <param name="XF">XF Format index. See Using FlexCel API.pdf.></param>
		public abstract void SetCellFormat(int row, int col, int XF);

		/// <summary>
		/// Sets the Cell format (XF) on a range of cells.
		/// You can create new formats using the  <see cref="AddFormat"/> function.
		/// <seealso cref="GetCellFormat(System.Int32,System.Int32)"/><seealso cref="AddFormat"/>
		/// </summary>
		/// <param name="row1">Row index of the top cell on the range (1 based)</param>
		/// <param name="col1">Column index of the left cell on the range (1 based)</param>
		/// <param name="row2">Row index of the bottom cell on the range (1 based)</param>
		/// <param name="col2">Column index of the right cell on the range (1 based)</param>
		/// <param name="XF">XF Format index. See Using FlexCel API.pdf.></param>
		public abstract void SetCellFormat(int row1, int col1, int row2, int col2, int XF);

        /// <inheritdoc cref="SetCellFormat(int, int, int, int, TFlxFormat, TFlxApplyFormat, bool)" />
        public void SetCellFormat(int row1, int col1, int row2, int col2, TFlxFormat newFormat, TFlxApplyFormat applyNewFormat)
		{
			SetCellFormat(row1, col1, row2, col2, newFormat, applyNewFormat, false);
		}

		/// <summary>
		/// Changes part of the Cell format on a range of cells. WARNING! This method is slower than the other SetCellFormat versions, use it only
		/// if you do not care about maximum performance or if you just can't use the other SetCellFormat versions.
		/// This particular version of SetCellFormat has to read the format on each cell, modify it and write it back.
		/// While still very fast, it is not as fast as just setting the format on a cell.
		/// </summary>
		/// <remarks>
		/// You can use this method for example to add a border on the top of a row of cells, keeping the existing font and pattern styles on the range.
		/// </remarks>
		/// <param name="row1">Row index of the top cell on the range (1 based)</param>
		/// <param name="col1">Column index of the left cell on the range (1 based)</param>
		/// <param name="row2">Row index of the bottom cell on the range (1 based)</param>
		/// <param name="col2">Column index of the right cell on the range (1 based)</param>
		/// <param name="newFormat">Format to apply to the cells.</param>
		/// <param name="applyNewFormat">Indicates which properties of newFormat will be applied to the cells.</param> 
		/// <param name="exteriorBorders">When true, the format for the border will be applied only to the outer cells in the range. This can
		/// be useful for example to draw a box around a range of cells, but not drawing borders inside the range. 
        /// Other parameters, like the cell background, will still be applied to the full range.</param>
		public abstract void SetCellFormat(int row1, int col1, int row2, int col2, TFlxFormat newFormat, TFlxApplyFormat applyNewFormat, bool exteriorBorders);

        /// <inheritdoc cref="GetCellFormat(int, int, int)" />
        public abstract int GetCellFormat(int row, int col);

        /// <summary>
        /// Cell Format for a given cell
        /// This method gets the Format number (XF) of a cell. 
        /// You can create new formats using the  <see cref="AddFormat"/> function.
        /// <seealso cref="AddFormat"/><seealso cref="SetCellFormat(int, int, int)"/><seealso cref="GetFormat"/><seealso cref="GetCellVisibleFormat(System.Int32,System.Int32)"/>
        /// </summary>
        /// <remarks>This method DOES NOT return format for an empty cell, even if it
        /// has a column or a row format. For the visible format of the cell, see <see cref="GetCellVisibleFormat(System.Int32,System.Int32)"/></remarks>
        /// <example>
        /// To copy the format on cell A1 to B2 you should write:
        /// <code>
        /// XlsFile.SetCellFormat(2,2,XlsFile.GetCellFormat(1,1));
        /// </code></example>
        /// <param name="row">Row Index (1 based)</param>
        /// <param name="col">Column Index (1 based)</param>
        /// <returns>XF for the cell.</returns>
        ///<param name="sheet">Sheet index (1 based).</param>
		public abstract int GetCellFormat(int sheet, int row, int col);

        /// <inheritdoc cref="GetCellVisibleFormat(int, int, int)" />
        public int GetCellVisibleFormat(int row, int col)
		{
			int XF = GetCellFormat(row, col);
			if (XF < 0)
			{
				XF = GetRowFormat(row);
				if (XF < 0) XF = GetColFormat(col);
			}
			if (XF < 0) XF = DefaultFormatId;
			return XF;
		}

		/// <summary>
		/// Cell Format for a given cell, including the format of the row and the column.
		/// <remarks>This might return format even if the cell is empty, if the column or the row have format.
		/// For the real format of the cell, see <see cref="GetCellFormat(int, int)"/>
		/// </remarks>
		/// <seealso cref="AddFormat"/><seealso cref="SetCellFormat(int, int, int)"/><seealso cref="GetFormat"/><seealso cref="GetCellFormat(System.Int32,System.Int32)"/><seealso cref="GetCellVisibleFormatDef(System.Int32,System.Int32)"/>
		/// </summary>
		/// <param name="sheet">Sheet index (1 based)</param>
		/// <param name="row">Row Index (1 based)</param>
		/// <param name="col">Column Index (1 based)</param>
		/// <returns>XF for the cell</returns>
		public int GetCellVisibleFormat(int sheet, int row, int col)
		{
			int XF = GetCellFormat(sheet, row, col);
			if (XF < 0)
			{
				XF=GetRowFormat(sheet, row);
				if (XF < 0) XF=GetColFormat(sheet, col);
			}
			if (XF < 0) XF = DefaultFormatId;
			return XF;
		}

        /// <inheritdoc cref="GetCellVisibleFormatDef(int, int, int)" />
        public TFlxFormat GetCellVisibleFormatDef(int row, int col)
		{
			return GetFormat(GetCellVisibleFormat(row, col));
		}

		/// <summary>
		/// Cell Format for a given cell, including the format of the row and the column.
		/// <remarks>This might return format even if the cell is empty, if the column or the row have format.
		/// For the real format of the cell, see <see cref="GetCellFormat(System.Int32,System.Int32)"/>
		/// This is a shortcut for GetCellVisibleFormat, returning the final Format struct.
		/// </remarks>
		/// <seealso cref="AddFormat"/><seealso cref="SetCellFormat(int, int, int)"/><seealso cref="GetFormat"/><seealso cref="GetCellFormat(System.Int32,System.Int32)"/><seealso cref="GetCellVisibleFormat(System.Int32,System.Int32)"/>
		/// </summary>
		/// <param name="sheet">Sheet index. (1 based)</param>
		/// <param name="row">Row Index (1 based)</param>
		/// <param name="col">Column Index (1 based)</param>
		/// <returns>Format for the cell</returns>
		public TFlxFormat GetCellVisibleFormatDef(int sheet, int row, int col)
		{
			return GetFormat(GetCellVisibleFormat(sheet, row, col));
		}

		/// <summary>
		/// Modifies the format of the specified cell if it has a conditional format active.
		/// Returns a modified format with the applied conditional format if there was any change, null otherwise.
		/// </summary>
		/// <param name="format">Original format of the cell.</param>
		/// <param name="row">Row of the cell (1 based)</param>
		/// <param name="col">Column of the cell (1 based)</param>
		/// <returns>If the format is modified by a conditional format, it returns the new format.
		/// If there are no changes returns null, to avoid creating new instances of TFlxFormat.</returns>
		public abstract TFlxFormat ConditionallyModifyFormat(TFlxFormat format, int row, int col);

		/*
		/// <summary>
		/// Sets a conditional format for a range of cells. You can use null on conditionalFormat parameter to clear all conditional formats on the range.
		/// </summary>
		/// <param name="firstRow">First row to format (1 based).</param>
		/// <param name="firstCol">First column to format (1 based).</param>
		/// <param name="lastRow">Last row to format (1 based).</param>
		/// <param name="lastCol">Last column to format (1 based).</param>
		/// <param name="conditionalFormat">List of Conditional formats to apply. Set it to null to clear the conditional formats on the range.</param>
		public abstract void SetConditionalFormat(int firstRow, int firstCol, int lastRow, int lastCol, TConditionalFormatRule[] conditionalFormat);

		/// <summary>
		/// Returns the number of conditional format blocks on the list. You can use this value to loop on them and retrieve the individual ones with <see cref="GetConditionalFormat"/>
		/// </summary>
		public abstract int ConditionalFormatCount{get;}

		/// <summary>
		/// One of the entries on the conditional format list of this file.
		/// </summary>
		/// <param name="index">Index to the conditional format list. (1 based)</param>
		/// <param name="range">The range of cells where the conditional format is applied.</param>
		/// <returns>The list of conditional format definitions for the index.</returns>
		public abstract TConditionalFormatRule[] GetConditionalFormat(int index, out TXlsCellRange range);
*/
		#endregion

		#region Styles
		/// <summary>
		/// Returns the number of named styles in the file.
		/// </summary>
		public abstract int StyleCount {get;}

		/// <summary>
		/// Gets the name of the style at position index. (1 based).
		/// </summary>
		/// <param name="index">Position in the list of styles (1 based).</param>
		/// <returns>The name of the style.</returns>
		public abstract string GetStyleName(int index);

		/// <summary>
		/// Returns the named style at position index for the workbook
		/// </summary>
		/// <param name="index">Position in the list of styles (1 based).</param>
		/// <returns>The named style definition.</returns>
		public abstract TFlxFormat GetStyle(int index);

		/// <summary>
		/// Returns a named style for the workbook. You can also use this method to check if a style exists or not.
        /// The returned style will have the "IsStyle" property set to true, so you can't apply it directly to a cell.
        /// If you want to apply the result of this method to a cell, use <see cref="GetStyle(string,bool)"/>.
		/// </summary>
		/// <param name="name">Name for the style. It might be an user defined name, or a built-in name. You can get
		/// a list of buitin names with <see cref="GetBuiltInStyleName(TBuiltInStyle,int)"/></param>
		/// <returns>The style definition, or null if the style doesn't exists.</returns>
        public TFlxFormat GetStyle(string name)
        {
            return GetStyle(name, false);
        }

        /// <summary>
        /// Returns a named style for the workbook. You can also use this method to check if a style exists or not.
        /// </summary>
        /// <param name="name">Name for the style. It might be an user defined name, or a built-in name. You can get
        /// a list of buitin names with <see cref="GetBuiltInStyleName(TBuiltInStyle,int)"/></param>
        /// <param name="convertToCellStyle">If true, the returned style will have the "IsStyle" property set to false, 
        /// so you can apply this TFlxFormat to a cell. If false IsStyle will be true and you can use the format definition in style definitions.
        /// Setting this parameter to true is exactly the same as setting it to false and setting "IsStyle" property in the result to false, and also setting
        /// the parent of the resulting cell format to the cell style.</param>
        /// <returns>The style definition, or null if the style doesn't exists.</returns>
        public abstract TFlxFormat GetStyle(string name, bool convertToCellStyle);

		/// <summary>
		/// Renames an existing style. Note that this might be an user-defined style, you can't rename built-in styles.
		/// </summary>
		/// <param name="oldName">Name of the existing style in the workbook.</param>
		/// <param name="newName">New name for the style. It must not exist.</param>
		public abstract void RenameStyle(string oldName, string newName);

		/// <summary>
		/// Modifies an existing style if name already exists, or creates a new style if it doesn't. 
		/// </summary>
		/// <param name="name">Name for the style. It might be an user defined name, or a built-in name. You can get
		/// a list of buit.in names with <see cref="GetBuiltInStyleName(TBuiltInStyle,int)"/></param>
		/// <param name="fmt">The new style definition.</param>
		public abstract void SetStyle(string name, TFlxFormat fmt);

		/// <summary>
		/// Returns a named style for the workbook.
		/// </summary>
		/// <param name="name">Name for the style. It must be an user defined name.</param>
		public abstract void DeleteStyle(string name);

		/// <summary>
		/// Returns the name for a built-in style.
		/// </summary>
		/// <param name="style">Style you want find out the name.</param>
		/// <param name="level">Used only if style is <see cref="TBuiltInStyle.RowLevel"/> or <see cref="TBuiltInStyle.ColLevel"/>. It specifies the level of the outline,
		/// and must be a number between 1 and 7. Keep it 0 for all other styles.</param>
		/// <returns>The name for the Built in style.</returns>
		public abstract string GetBuiltInStyleName(TBuiltInStyle style, int level);

        /// <summary>
        /// Tries to convert a string into an built-in style identifier. Will return true if styleName can be converted, false otherwise.
        /// </summary>
        /// <param name="styleName">Style that we want to convert to built-in style.</param>
        /// <param name="style">Returns the built-in style. This value is only valid if this method returns true.</param>
        /// <param name="level">Returns the level built-in style (1 based). This value is only valid if this method returns true, and only applies to outline styles. It will be 0 for non outline styles.</param>
        /// <returns>True is styleNameis a built-in style (and thus style and level are valid), false otherwise.</returns>
        public abstract bool TryGetBuiltInStyleType(string styleName, out TBuiltInStyle style, out int level);
		#endregion

		#region Merged Cells
		/// <summary>
		/// Merged Range where the cell is.
		/// <seealso cref="MergeCells"/><seealso cref="UnMergeCells"/><seealso cref="CellMergedListCount"/><seealso cref="CellMergedList"/>
		/// </summary>
		/// <example>
		/// If you have a merged cell in range A1: B2, calling CellMergedBounds on any of the cells: A1, B1, A2, B2 will return A1:B2.
		/// </example>
		/// <param name="row">Row Index (1 based)</param>
		/// <param name="col">Column Index (1 based)</param>
		/// <returns>The range where the cell is</returns>
		public abstract TXlsCellRange CellMergedBounds(int row, int col);

		/// <summary>
		/// Merges a number of cells into one.
		/// <seealso cref="CellMergedBounds"/><seealso cref="UnMergeCells"/><seealso cref="CellMergedListCount"/><seealso cref="CellMergedList"/>
		/// </summary>
		/// <param name="firstRow">First row of the merged cell.</param>
		/// <param name="firstCol">First column of the merged cell.</param>
		/// <param name="lastRow">Last row of the merged cell.</param>
		/// <param name="lastCol">Last column of the merged cell.</param>
		/// <remarks>This method might return a bigger range than the one you specify.
		/// For example, if you have a merged cell (A1:C1), calling MergeCells(1,2,2,2) will 
		/// merge the cell with the existing one, returning ONE merged cell (A1:C2), and not (B1:B2)
		/// <p></p>
		/// Other thing to consider, is that Excel honors individual linestyles inside a merged cell.
		/// That is, you can have only the first 2 columns of the merged cell with a line and the others without.
		/// Normally, after merging the cells, you will want to call <see cref="SetCellFormat(int, int, int)"/> on the range
		/// to make them all similar. FlexCel does not do it by default, to give the choice to you.</remarks>
		public abstract void MergeCells(int firstRow, int firstCol, int lastRow, int lastCol);

		/// <summary>
		/// Unmerges the range of cells. The coordinates have to be exact, if there is no
		/// merged cell with the exact coordinates, nothing will be done.
		/// <seealso cref="CellMergedBounds"/><seealso cref="MergeCells"/><seealso cref="CellMergedListCount"/><seealso cref="CellMergedList"/>
		/// </summary>
		/// <param name="firstRow">First row of the merged cell.</param>
		/// <param name="firstCol">First column of the merged cell.</param>
		/// <param name="lastRow">Last row of the merged cell.</param>
		/// <param name="lastCol">Last column of the merged cell.</param>
		public abstract void UnMergeCells(int firstRow, int firstCol, int lastRow, int lastCol);

		/// <summary>
		/// For using with <see cref="CellMergedList"/> on a loop: for (int i=1;i &lt;= CellMergedListCount;i++) DoSomething(CellMergedList(i))...
		/// </summary>
		public abstract int CellMergedListCount{get;} 
		
		/// <summary>
		/// The Merged cell at position index on the mergedcell list.
		/// <seealso cref="CellMergedListCount"/>
		/// </summary>
		/// <param name="index">index on the list (1 based)</param>
		/// <returns>The merged cell at position index.</returns>
		public abstract TXlsCellRange CellMergedList(int index);
		#endregion

		#region Rows and Cols
		/// <summary>
		/// Number of rows actually used on the sheet. 
		/// </summary>
		public abstract int RowCount{get;}

        /// <summary>
        /// Number of rows actually used on a given sheet. 
        /// </summary>
        public abstract int GetRowCount(int sheet);

		/// <summary>
		/// Number of columns actually used on the active sheet. <b>Note that this method is *slow*</b> as it needs to loop over all the rows to find out
        /// the biggest used column. <b>Never</b> use it in a loop like "for (int  col = 1; col &lt;= xls.ColCount; col++)". Instead try to use <see cref="ColCountInRow(int)"/>.
        /// If you *need* to use ColCount, cache its value first:
        /// <code>
        /// int RowCount = xls.RowCount;
        /// int ColCount = xls.ColCount;
        /// for (int  row = 1; row &lt;= RowCount; row++)
        /// {
        ///    for (int  col = 1; col &lt;= ColCount; col++)
        ///    {
        /// </code>
        /// Remember that loops in C# will evaluate the second parameter every time the loop is executed.
		/// </summary>
		public abstract int ColCount{get;}

        /// <summary>
        /// Number of columns actually used on a given sheet. 
        /// </summary>
        /// <param name="sheet">Sheet index where you want to find the columns (1 based)</param>
        public int GetColCount(int sheet)
        {
            return GetColCount(sheet, true);
        }

        /// <summary>
        /// Number of columns actually used on a given sheet. 
        /// </summary>
        /// <param name="sheet">Sheet index where you want to find the columns (1 based)</param>
        /// <param name="includeFormattedColumns">If true (the default) formatted columns (for example a column you selected and painted yellow)
        /// will be included in the count, even if it doesn't have data.
        /// </param>
        public abstract int GetColCount(int sheet, bool includeFormattedColumns);

		/// <summary>
		/// True if the specified row does not have any cells, nor any format on it.
		/// In short, this row has never been used.
		/// </summary>
		/// <param name="row">Row to test (1-based)</param>
		public abstract bool IsEmptyRow(int row);

		/// <summary>
		/// True if the specified column does not have any format applied on it.
		/// </summary>
		/// <param name="col">Column to test (1-based)</param>
		public abstract bool IsNotFormattedCol(int col);

        /// <inheritdoc cref="GetRowFormat(int, int)" />
        public abstract int GetRowFormat(int row);

        /// <summary>
        /// Gets the XF format for the specified row, -1 if the row doesn't have format.
        /// </summary>
        /// <param name="row">Row index (1-based)</param>
        /// <returns>XF format.</returns>
        ///<param name="sheet">Sheet index (1 based).</param>
		public abstract int GetRowFormat(int sheet, int row);


        /// <summary>
		/// Sets the XF format for the entire row.
		/// </summary>
		/// <param name="row">Row index (1-based)</param>
		/// <param name="XF">XF format.</param>
		public void SetRowFormat(int row, int XF)
		{
			SetRowFormat(row, XF, true);
		}

		/// <summary>
		/// Sets the XF format for the entire row.
		/// </summary>
		/// <param name="row">Row index (1-based)</param>
		/// <param name="XF">XF format.</param>
		/// <param name="resetRow">When true, all existing cells on the row will be reset to this format.
		/// This is the standard Excel behavior and the recommended option. If you don't care about existing cells, 
		/// you can speed up this method by setting it to false.</param>
		public abstract void SetRowFormat(int row, int XF, bool resetRow);

		/// <summary>
		/// Sets the format characteristics specified in ApplyFormat for the entire row.
		/// </summary>
		/// <param name="row">Row index (1-based)</param>
		/// <param name="newFormat">Format to apply.</param>
		/// <param name="applyNewFormat">Indicates which properties of newFormat will be applied to the cells.</param> 
		/// <param name="resetRow">When true, all existing cells on the row will be reset to this format.
		/// This is the standard Excel behavior and the recommended option. If you don't care about existing cells, 
		/// you can speed up this method by setting it to false.</param>
		public abstract void SetRowFormat(int row, TFlxFormat newFormat, TFlxApplyFormat applyNewFormat, bool resetRow);


        /// <inheritdoc cref="GetColFormat(int, int)" />
        public abstract int GetColFormat(int col);

        /// <summary>
        /// Gets the XF format for the specified column, -1 if the column doesn't have format.
        /// </summary>
        /// <param name="col">Column index (1-based)</param>
        /// <param name="sheet">Sheet index (1 based).</param>
        /// <returns>XF format.</returns>
		public abstract int GetColFormat(int sheet, int col);


        /// <summary>
		/// Sets the format for an entire column.
		/// </summary>
		/// <param name="col">Column to set.</param>
		/// <param name="XF">XF Format index.</param>
		/// <param name="resetColumn">When true, all existing cells on the column will be reset to this format.
		/// This is the standard Excel behavior and the recommended option. If you don't care about existing cells, 
		/// you can speed up this method by setting it to false.</param>
		public abstract void SetColFormat(int col, int XF, bool resetColumn);

		/// <summary>
		/// Sets the format for an entire column.
		/// </summary>
		/// <param name="col">Column to set.</param>
		/// <param name="XF">XF Format index.</param>
		public void SetColFormat(int col, int XF)
		{
			SetColFormat(col, XF, true);
		}

		/// <summary>
		/// Sets the format characteristics specified in ApplyFormat for the entire column.
		/// </summary>
		/// <param name="col">Column to set.</param>
		/// <param name="newFormat">Format to apply.</param>
		/// <param name="applyNewFormat">Indicates which properties of newFormat will be applied to the cells.</param> 
		/// <param name="resetColumn">When true, all existing cells on the column will be reset to this format.
		/// This is the standard Excel behavior and the recommended option. If you don't care about existing cells, 
		/// you can speed up this method by setting it to false.</param>
		public abstract void SetColFormat(int col, TFlxFormat newFormat, TFlxApplyFormat applyNewFormat, bool resetColumn);


		/// <summary>
		/// Returns all Row options at once (if the row is autosize, if it is hidden, etc). 
		/// </summary>
		/// <param name="row">Row Index (1 based)</param>
		/// <remarks>
		/// To get individual values, use the corresponding methods (i.e. <see cref="GetAutoRowHeight"/>)
		/// Use this method only to copy the options from one row to another. 
		/// </remarks>
		/// <example>
		/// To copy all the row options from row 1 to 2, use
		/// <code>SetRowOptions(2,GetRowOptions(1));</code>
		/// This is much faster than assigning each option alone.
		/// </example>
		/// <returns>Row options</returns>
		public abstract int GetRowOptions(int row);

		/// <summary>
		/// Sets all Row options at once (if the row is autosize, if it is hidden, etc). 
		/// </summary>
		/// <param name="row">Row Index (1 based)</param>
		/// <param name="options">A flag with all row options</param>
		/// <remarks>
		/// To set individual values, use the corresponding methods (i.e. <see cref="SetAutoRowHeight"/>)
		/// Use this method only to copy the options from one row to another. 
		/// </remarks>
		/// <example>
		/// To copy all the row options from row 1 to 2, use
		/// <code>SetRowOptions(2,GetRowOptions(1));</code>
		/// This is much faster than assigning each option alone.
		/// </example>
		public abstract void SetRowOptions(int row, int options);
        
		/// <summary>
		/// Returns all Column options at once (if the column is hidden, etc). 
		/// </summary>
		/// <param name="col">Column Index (1 based)</param>
		/// <remarks>
		/// To get individual values, use the corresponding methods (i.e. <see cref="GetColHidden"/>)
		/// Use this method only to copy the options from one column to another. 
		/// </remarks>
		/// <example>
		/// To copy all the column options from column 1 to 2, use
		/// <code>SetColOptions(2,GetColOptions(1));</code>
		/// This is much faster than assigning each option alone.
		/// </example>
		public abstract int GetColOptions(int col);

		/// <summary>
		/// Sets all Column options at once (if the column is hidden, etc). 
		/// </summary>
		/// <param name="col">Column Index (1 based)</param>
		/// <param name="options">A flag with all column options.</param>
		/// <remarks>
		/// To set individual values, use the corresponding methods (i.e. <see cref="SetColHidden"/>)
		/// Use this method only to copy the options from one column to another. 
		/// </remarks>
		/// <example>
		/// To copy all the column options from column 1 to 2, use
		/// <code>SetColOptions(2,GetColOptions(1));</code>
		/// This is much faster than assigning each option alone.
		/// </example>
		public abstract void SetColOptions(int col, int options);

		/// <summary>
		/// Returns the current Row height, in Excel internal units. (1/20th of a point)
		/// </summary>
		/// <remarks>Use <see cref="FlxConsts.RowMult"/> to convert the internal units to pixels.</remarks>
		/// <example><code>RowHeightInPixels=GetRowHeight(Row)/FlxConsts.RowMult;</code></example>
		/// <param name="row">Row Index (1 based)</param>
		/// <returns>Row height in internal excel units.</returns>
		public abstract int GetRowHeight(int row);

		/// <summary>
		/// Returns the current Row height, in Excel internal units. (1/20th of a point)
		/// </summary>
		/// <remarks>Use <see cref="FlxConsts.RowMult"/> to convert the internal units to pixels.</remarks>
		/// <example><code>RowHeightInPixels=GetRowHeight(Row)/FlxConsts.RowMult;</code></example>
		/// <param name="row">Row Index (1 based)</param>
		/// <param name="HiddenIsZero">If true, the height returned for a hidden row will be 0 and not its real height.</param>
		/// <returns>Row height in internal excel units.</returns>
		public int GetRowHeight(int row, bool HiddenIsZero)
		{
			if (HiddenIsZero && GetRowHidden(row)) return 0;
			else return GetRowHeight(row);
		}

		/// <summary>
		/// Returns the current Row height for a given sheet, in Excel internal units. (1/20th of a point)
		/// </summary>
		/// <remarks>Use <see cref="FlxConsts.RowMult"/> to convert the internal units to pixels.</remarks>
		/// <example><code>RowHeightInPixels=GetRowHeight(Row)/ExcelMetrics.RowMult(Workbook);</code></example>
		/// <param name="sheet">Sheet where to look for the height.</param>
		/// <param name="row">Row Index (1 based)</param>
		/// <param name="HiddenIsZero">If true, the height returned for a hidden row will be 0 and not its real height.</param>
		/// <returns>Row height in internal Excel units.(1/20th of a point)</returns>
		public abstract int GetRowHeight(int sheet, int row, bool HiddenIsZero);


		/// <summary>
		/// Sets the current Row height, in Excel internal units. (1/20th of a point)
		/// </summary>
		/// <remarks>Use <see cref="FlxConsts.RowMult"/> to convert the internal units to pixels.</remarks>
		/// <example><code>RowHeightInPixels=GetRowHeight(Row)/FlxConsts.RowMult;</code></example>
		/// <param name="row">Row Index (1 based)</param>
		/// <param name="height">Row height, in Excel internal units. (1/20th of a point). See <see cref="FlxConsts.RowMult"/></param>
		public abstract void SetRowHeight(int row, int height);
        
		/// <summary>
		/// Returns the current Column width, in Excel internal units. (Character width of "font 0" / 256)
		/// </summary>
		/// <remarks>Use <see cref="ExcelMetrics.ColMult"/> to convert the internal units to pixels.</remarks>
		/// <example><code>ColWidthInPixels=GetColWidth(Col)/ExcelMetrics.ColMult(Workbook);</code></example>
		/// <param name="col">Column Index (1 based)</param>
		/// <returns>Column width in internal excel units.(Character width of "font 0" / 256)</returns>
		public abstract int GetColWidth(int col);

		/// <summary>
		/// Returns the current Column width, in Excel internal units. (Character width of font 0 / 256)
		/// </summary>
		/// <remarks>Use <see cref="ExcelMetrics.ColMult"/> to convert the internal units to pixels.</remarks>
		/// <example><code>ColWidthInPixels=GetColWidth(Col)/ExcelMetrics.ColMult(Workbook);</code></example>
		/// <param name="col">Column Index (1 based)</param>
		/// <param name="HiddenIsZero">If true, the width returned for a hidden column will be 0 and not its real width.</param>
		/// <returns>Column width in internal excel units.(Character width of font 0 / 256)</returns>
		public abstract int GetColWidth(int col, bool HiddenIsZero);

		/// <summary>
		/// Returns the current Column width for a given sheet, in Excel internal units. (Character width of font 0 / 256)
		/// </summary>
		/// <remarks>Use <see cref="ExcelMetrics.ColMult"/> to convert the internal units to pixels.</remarks>
		/// <example><code>ColWidthInPixels=GetColWidth(Col)/ExcelMetrics.ColMult(Workbook);</code></example>
		/// <param name="sheet">Sheet where to look for the width.</param>
		/// <param name="col">Column Index (1 based)</param>
		/// <param name="HiddenIsZero">If true, the width returned for a hidden column will be 0 and not its real width.</param>
		/// <returns>Column width in internal excel units.(Character width of font 0 / 256)</returns>
		public abstract int GetColWidth(int sheet, int col, bool HiddenIsZero);

		/// <summary>
		/// Sets the current Column width, in Excel internal units. (Character width of font 0 / 256)
		/// </summary>
		/// <remarks>Use <see cref="FlxConsts.ColMult"/> to convert the internal units to pixels.</remarks>
		/// <example><code>ColWidthInPixels=GetColWidth(Col)/FlxConsts.ColMult;</code></example>
		/// <param name="col">Column Index (1 based)</param>
		/// <param name="width">Column width, in Excel internal units. (Character width of font 0 / 256). See <see cref="FlxConsts.ColMult"/></param>
		public abstract void SetColWidth(int col, int width);

        internal abstract void SetColWidthInternal(int col, int width);

		/// <summary>
		/// The default height for empty rows, in Excel internal units. (1/20th of a point). <b>IMPORTANT: </b> For this property
        /// to have any effect, you also need to set <see cref="DefaultRowHeightAutomatic"/> = false
		/// </summary>
		/// <remarks>Use <see cref="FlxConsts.RowMult"/> to convert the internal units to pixels.</remarks>
		/// <example>
        /// <code>
        /// RowHeightInPixels=GetRowHeight(Row)/FlxConsts.RowMult;
        /// </code></example>
		public abstract int DefaultRowHeight{get;set;}

        /// <summary>
        /// When this property is true, the row height for empty rows is calculated with the height of the "Normal" font and will
        /// change if you change the Normal style. When false, the value in <see cref="DefaultRowHeight"/> will be used.
        /// </summary>
        public abstract bool DefaultRowHeightAutomatic { get; set; }

		/// <summary>
		/// The default width for empty columns, in Excel internal units. (Character width of font 0 / 256)
		/// </summary>
		/// <remarks>Use <see cref="FlxConsts.ColMult"/> to convert the internal units to pixels.</remarks>
		/// <example><code>ColWidthInPixels=GetColWidth(Col)/FlxConsts.ColMult;</code></example>
		public abstract int DefaultColWidth{get;set;}

		/// <summary>
		/// Returns true if the row is hidden.
		/// </summary>
		/// <param name="row">Row index (1 based).</param>
		/// <returns>True if the row is hidden.</returns>
		public abstract bool GetRowHidden(int row);

		/// <summary>
		/// Returns true if the row is hidden. This method does not care about ActiveSheet.
		/// </summary>
		/// <param name="sheet">Sheet index (1 based) for the row.</param>
		/// <param name="row">Row index (1 based).</param>
		/// <returns>True if the row is hidden.</returns>
		public abstract bool GetRowHidden(int sheet, int row);

		/// <summary>
		/// Hides or shows an specific row.
		/// </summary>
		/// <param name="row">Row index (1 based)</param>
		/// <param name="hide">If true, row will be hidden, if false it will be visible.</param>
		public abstract void SetRowHidden(int row, bool hide);
        
		/// <summary>
		/// Returns true if the column is hidden.
		/// </summary>
		/// <param name="col">Column index (1 based)</param>
		/// <returns>True if the column is hidden.</returns>
		public abstract bool GetColHidden(int col);

		/// <summary>
		/// Hides or shows an specific column.
		/// </summary>
		/// <param name="col">Column index (1 based).</param>
		/// <param name="hide">If true, column will be hidden, if false it will be visible.</param>
		public abstract void SetColHidden(int col, bool hide);

		/// <summary>
		/// Returns if the row is adjusting its size to the cell (the default) or if it has a fixed height.
		/// </summary>
		/// <remarks>
		/// By default, Excel rows auto adapt their size to the font size. 
		/// If you set the row height manually, it will remain fixed to this size until you set 
		/// AutoFit (Menu->Format->Row->AutoFit) back. 
		/// </remarks>
		/// <param name="row">Row index (1-based)</param>
		/// <returns>True if AutoFit is on for the row, False if it has a fixed size</returns>
		public abstract bool GetAutoRowHeight(int row);

		/// <summary>
		/// Sets the current row to automatically autosize to the biggest cell or not.
		/// </summary>
		/// <remarks>
		/// By default, Excel rows auto adapt their size to the font size. 
		/// If you set the row height manually, it will remain fixed to this size until you set 
		/// AutoFit (Menu->Format->Row->AutoFit) back. 
		/// </remarks>
		/// <param name="row">Row index (1-based)</param>
		/// <param name="autoRowHeight">If true, row will have autofit.</param>
		public abstract void SetAutoRowHeight(int row, bool autoRowHeight);

		/// <summary>
		/// Autofits a range of rows so they adapt their height to show all the text inside. Note that due to GDI+ / GDI incompatibilities,
		/// the height calculated by FlexCel will not be exactly the same than the one calculated by Excel. So when you open this workbook
		/// in Excel, Excel will re calculate the row heights to what it believe is best. You can change this behaviour specifying keepHeightAutomatic = false.
		/// </summary>
		/// <remarks>THIS METHOD DOES NOT WORK ON COMPACT FRAMEWORK.</remarks>
		/// <param name="row1">First row to autofit.</param>
		/// <param name="row2">Last row to autofit.</param>
		/// <param name="autofitNotAutofittingRows">When you are autofitting a range of rows, some rows might not be 
		/// set to Autofit in Excel. When this parameter is true, those rows will be autofitted anyway.</param>
		/// <param name="keepHeightAutomatic">If true, rows will be still autoheight when you open the file in Excel, so Excel
		/// will recalculate the values, probably changing the page breaks. If you set it to false, rows will be fixed in size,
		/// and when you open it on Excel they will remain so.</param>
		/// <param name="adjustment">You will normally want to set this parameter to 1, which means that autofit will be made with standard measurements.
		/// If you set it to for example 1.1, then rows will be adjusted to 110% percent of what their calculated height was. 
		/// Use this parameter to fine-tune autofiting, if for example rows are too small when opening the file in Excel.</param>
        /// <param name="adjustmentFixed">You will normally set this parameter to 0, which means standard autofit. If you set it to a value, the row will be
        /// made larger by that ammount from the calculated autofit. Different from the "adjustment" parameter, this parameter adds a fixed size to the row
        /// and not a percentage. The final size of the row will be:  FinalSize = CalulatedAutoFit * adjustment + adjusmentFixed</param>
        /// <param name="minHeight">Minimum final height for the row to autofit. If the calculated value is less than minHeight, row size will be set to minHeight.
        /// <br/>A negative value on minHeight means the row size will be no smaller than the original height.</param>
        /// <param name="maxHeight">Maximum final height for the row to autofit. If the calculated value is more than maxHeigth, row size will be set to maxHeight.
        /// <br/>maxHeight = 0 means no maxHeight.
        /// <br/>A negative value on maxHeight means the row size will be no bigger than the original height.
        /// </param>
        public void AutofitRow(int row1, int row2, bool autofitNotAutofittingRows, bool keepHeightAutomatic, real adjustment, int adjustmentFixed, int minHeight, int maxHeight)
		{
			AutofitRow(row1, row2, autofitNotAutofittingRows, keepHeightAutomatic, adjustment, adjustmentFixed, minHeight, maxHeight, TAutofitMerged.OnLastCell);
		}

		/// <summary>
		/// Autofits a range of rows so they adapt their height to show all the text inside. Note that due to GDI+ / GDI incompatibilities,
		/// the height calculated by FlexCel will not be exactly the same than the one calculated by Excel. So when you open this workbook
		/// in Excel, Excel will re calculate the row heights to what it believe is best. You can change this behaviour specifying keepHeightAutomatic = false.
		/// </summary>
		/// <remarks>THIS METHOD DOES NOT WORK ON COMPACT FRAMEWORK.</remarks>
		/// <param name="row1">First row to autofit.</param>
		/// <param name="row2">Last row to autofit.</param>
		/// <param name="autofitNotAutofittingRows">When you are autofitting a range of rows, some rows might not be 
		/// set to Autofit in Excel. When this parameter is true, those rows will be autofitted anyway.</param>
		/// <param name="keepHeightAutomatic">If true, rows will be still autoheight when you open the file in Excel, so Excel
		/// will recalculate the values, probably changing the page breaks. If you set it to false, rows will be fixed in size,
		/// and when you open it on Excel they will remain so.</param>
		/// <param name="adjustment">You will normally want to set this parameter to 1, which means that autofit will be made with standard measurements.
		/// If you set it to for example 1.1, then rows will be adjusted to 110% percent of what their calculated height was. 
		/// Use this parameter to fine-tune autofiting, if for example rows are too small when opening the file in Excel.</param>
		public void AutofitRow(int row1, int row2, bool autofitNotAutofittingRows, bool keepHeightAutomatic, real adjustment)
		{
			AutofitRow(row1, row2, autofitNotAutofittingRows, keepHeightAutomatic, adjustment, 0, 0, 0);
		}

        /// <summary>
        /// Autofits a range of rows so they adapt their height to show all the text inside. Note that due to GDI+ / GDI incompatibilities,
        /// the height calculated by FlexCel will not be exactly the same than the one calculated by Excel. So when you open this workbook
        /// in Excel, Excel will re calculate the row heights to what it believe is best. You can change this behaviour specifying keepHeightAutomatic = false.
        /// </summary>
        /// <remarks>THIS METHOD DOES NOT WORK ON COMPACT FRAMEWORK.</remarks>
        /// <param name="row1">First row to autofit.</param>
        /// <param name="row2">Last row to autofit.</param>
        /// <param name="autofitNotAutofittingRows">When you are autofitting a range of rows, some rows might not be 
        /// set to Autofit in Excel. When this parameter is true, those rows will be autofitted anyway.</param>
        /// <param name="keepHeightAutomatic">If true, rows will be still autoheight when you open the file in Excel, so Excel
        /// will recalculate the values, probably changing the page breaks. If you set it to false, rows will be fixed in size,
        /// and when you open it on Excel they will remain so.</param>
        /// <param name="adjustment">You will normally want to set this parameter to 1, which means that autofit will be made with standard measurements.
        /// If you set it to for example 1.1, then rows will be adjusted to 110% percent of what their calculated height was. 
        /// Use this parameter to fine-tune autofiting, if for example rows are too small when opening the file in Excel.</param>
        /// <param name="adjustmentFixed">You will normally set this parameter to 0, which means standard autofit. If you set it to a value, the row will be
        /// made larger by that ammount from the calculated autofit. Different from the "adjustment" parameter, this parameter adds a fixed size to the row
        /// and not a percentage. The final size of the row will be:  FinalSize = CalulatedAutoFit * adjustment + adjusmentFixed</param>
        /// <param name="minHeight">Minimum final height for the row to autofit. If the calculated value is less than minHeight, row size will be set to minHeight.
        /// <br/>A negative value on minHeight means the row size will be no smaller than the original height.</param>
        /// <param name="maxHeight">Maximum final height for the row to autofit. If the calculated value is more than maxHeigth, row size will be set to maxHeight.
        /// <br/>maxHeight = 0 means no maxHeight.
        /// <br/>A negative value on maxHeight means the row size will be no bigger than the original height.
        /// </param>
        /// <param name="autofitMerged">Specifies which row in a merged cell using more than one row will be used to autofit the merged cell.
        /// If you don't specify this parameter, it will be the last row.</param>
        public abstract void AutofitRow(int row1, int row2, bool autofitNotAutofittingRows, bool keepHeightAutomatic, real adjustment, int adjustmentFixed, int minHeight, int maxHeight, TAutofitMerged autofitMerged);

		/// <summary>
		/// Autofits a row so it adapts its height to show all the text inside. It does not matter if the row is set to Autofit or not in Excel.
		/// Note that due to GDI+ / GDI incompatibilities,
		/// the height calculated by FlexCel will not be exactly the same than the one calculated by Excel. So when you open this workbook
		/// in Excel, Excel will re calculate the row heights to what it believe is best. You can change this behaviour specifying keepHeightAutomatic = false.
		/// </summary>
		/// <remarks>THIS METHOD DOES NOT WORK ON COMPACT FRAMEWORK.</remarks>
		/// <param name="row">Row to Autofit.</param>
		/// <param name="keepHeightAutomatic">If true, rows will be still autoheight when you open the file in Excel, so Excel
		/// will recalculate the values, probably changing the page breaks. If you set it to false, rows will be fixed in size,
		/// and when you open it on Excel they will remain so.</param>
		/// <param name="adjustment">You will normally want to set this parameter to 1, which means that autofit will be made with standard measurements.
		/// If you set it to for example 1.1, then rows will be adjusted to 110% percent of what their calculated height was. 
		/// Use this parameter to fine-tune autofiting, if for example rows are too small when opening the file in Excel.</param>
		public void AutofitRow(int row, bool keepHeightAutomatic, real adjustment)
		{
			AutofitRow(row, row, true, keepHeightAutomatic, adjustment);
		}

		/// <summary>
		/// Autofits a column so it adapts its width to show all the text inside.
		/// </summary>
		/// <remarks>THIS METHOD DOES NOT WORK ON COMPACT FRAMEWORK.</remarks>
		/// <param name="col">Column to Autofit.</param>
		/// <param name="ignoreStrings">When true, strings will not be considered for the autofit. Only numbers will.</param>
		/// <param name="adjustment">You will normally want to set this parameter to 1, which means that autofit will be made with standard measurements.
		/// If you set it to for example 1.1, then columns will be adjusted to 110% percent of what their calculated width was. 
		/// Use this parameter to fine-tune autofiting, if for example columns are too small when opening the file in Excel.</param>
		public void AutofitCol(int col, bool ignoreStrings, real adjustment)
		{
			AutofitCol(col, col, ignoreStrings, adjustment);
		}

		/// <summary>
		/// Autofits a range of columns so they adapt their width to show all the text inside. 
		/// </summary>
		/// <remarks>THIS METHOD DOES NOT WORK ON COMPACT FRAMEWORK.</remarks>
		/// <param name="col1">First column to Autofit.</param>
		/// <param name="col2">Last column to Autofit.</param>
		/// <param name="ignoreStrings">When true, strings will not be considered for the autofit. Only numbers will.</param>
		/// <param name="adjustment">You will normally want to set this parameter to 1, which means that autofit will be made with standard measurements.
		/// If you set it to for example 1.1, then columns will be adjusted to 110% percent of what their calculated width was. 
		/// Use this parameter to fine-tune autofiting, if for example columns are too small when opening the file in Excel.</param>
		public void AutofitCol(int col1, int col2, bool ignoreStrings, real adjustment)
		{
			AutofitCol(col1, col2, ignoreStrings, adjustment, 0, 0, 0);
		}

        /// <summary>
        /// Autofits a range of columns so they adapt their width to show all the text inside. 
        /// </summary>
        /// <remarks>THIS METHOD DOES NOT WORK ON COMPACT FRAMEWORK.</remarks>
        /// <param name="col1">First column to Autofit.</param>
        /// <param name="col2">Last column to Autofit.</param>
        /// <param name="ignoreStrings">When true, strings will not be considered for the autofit. Only numbers will.</param>
        /// <param name="adjustment">You will normally want to set this parameter to 1, which means that autofit will be made with standard measurements.
        /// If you set it to for example 1.1, then columns will be adjusted to 110% percent of what their calculated width was. 
        /// Use this parameter to fine-tune autofiting, if for example columns are too small when opening the file in Excel.</param>
        /// <param name="adjustmentFixed">You will normally set this parameter to 0, which means standard autofit. If you set it to a value, the column will be
        /// made larger by that ammount from the calculated autofit. Different from the "adjustment" parameter, this parameter adds a fixed size to the column
        /// and not a percentage. The final size of the column will be:  FinalSize = CalulatedAutoFit * adjustment + adjusmentFixed</param>
        /// <param name="minWidth">Minimum final width for the column to autofit. If the calculated value is less than minWidth, column size will be set to minWidth.
        /// <br/>A negative value on minWidth means the column size will be no smaller than the original width.</param>
        /// <param name="maxWidth">Maximum final width for the column to autofit. If the calculated value is more than maxWidth, column size will be set to maxWidth.
        /// <br/>maxWidth = 0 means no maxWidth.
        /// <br/>A negative value on maxWidth means the column size will be no bigger than the original width.
        /// </param>
		public void AutofitCol(int col1, int col2, bool ignoreStrings, real adjustment, int adjustmentFixed, int minWidth, int maxWidth)
		{
			AutofitCol(col1, col2, ignoreStrings, adjustment, adjustmentFixed, minWidth, maxWidth, TAutofitMerged.OnLastCell);
		}


		/// <summary>
		/// Autofits a range of columns so they adapt their width to show all the text inside. 
		/// </summary>
		/// <remarks>THIS METHOD DOES NOT WORK ON COMPACT FRAMEWORK.</remarks>
		/// <param name="col1">First column to Autofit.</param>
		/// <param name="col2">Last column to Autofit.</param>
		/// <param name="ignoreStrings">When true, strings will not be considered for the autofit. Only numbers will.</param>
		/// <param name="adjustment">You will normally want to set this parameter to 1, which means that autofit will be made with standard measurements.
		/// If you set it to for example 1.1, then columns will be adjusted to 110% percent of what their calculated width was. 
		/// Use this parameter to fine-tune autofiting, if for example columns are too small when opening the file in Excel.</param>
		/// <param name="adjustmentFixed">You will normally set this parameter to 0, which means standard autofit. If you set it to a value, the column will be
		/// made larger by that ammount from the calculated autofit. Different from the "adjustment" parameter, this parameter adds a fixed size to the column
		/// and not a percentage. The final size of the column will be:  FinalSize = CalulatedAutoFit * adjustment + adjusmentFixed</param>
		/// <param name="minWidth">Minimum final width for the column to autofit. If the calculated value is less than minWidth, column size will be set to minWidth.
		/// <br/>A negative value on minWidth means the column size will be no smaller than the original width.</param>
		/// <param name="maxWidth">Maximum final width for the column to autofit. If the calculated value is more than maxWidth, column size will be set to maxWidth.
		/// <br/>maxWidth = 0 means no maxWidth.
		/// <br/>A negative value on maxWidth means the column size will be no bigger than the original width.
		/// </param>
		/// <param name="autofitMerged">Specifies which column in a merged cell using more than one column will be used to autofit the merged cell.
		/// If you don't specify this parameter, it will be the last column.</param>
		public abstract void AutofitCol(int col1, int col2, bool ignoreStrings, real adjustment, int adjustmentFixed, int minWidth, int maxWidth, TAutofitMerged autofitMerged);


		/// <summary>
		/// Autofits all rhe rows on all sheets on a workbook that are set to autofit so they adapt their height to show all the text inside. 
		/// Note that due to GDI+ / GDI incompatibilities,
		/// the heights calculated by FlexCel will not be exactly the same than the ones calculated by Excel. So when you open this workbook
		/// in Excel, Excel might re calculate the row heights to what it believe is best. You can change this behaviour specifying keepSizesAutomatic = false.
		/// <remarks>THIS METHOD DOES NOT WORK ON COMPACT FRAMEWORK.</remarks>
		/// </summary>
		/// <param name="autofitNotAutofittingRows">When you are autofitting a range of rows, some rows might not be 
		/// set to Autofit in Excel. When this parameter is true, those rows will be autofitted anyway.</param>
		/// <param name="keepSizesAutomatic">
		/// When true, no modifications will be done to the "autofit" status of the rows. When false, all rows will be marked as "no autofit", so when you open this file
		/// in Excel it will not be resized by Excel, and the printing/export to pdf from Excel will be the same as FlexCel.
		/// </param>
		/// <param name="adjustment">You will normally want to set this parameter to 1, which means that autofit will be made with standard measurements.
		/// If you set it to for example 1.1, then rows will be adjusted to 110% percent of what their calculated height was. 
		/// Use this parameter to fine-tune autofiting, if for example rows are too small when opening the file in Excel.</param>
		public void AutofitRowsOnWorkbook(bool autofitNotAutofittingRows, bool keepSizesAutomatic, real adjustment)
		{
			AutofitRowsOnWorkbook(autofitNotAutofittingRows, keepSizesAutomatic, adjustment, 0, 0, 0);
		}

        /// <summary>
        /// Autofits all rhe rows on all sheets on a workbook that are set to autofit so they adapt their height to show all the text inside. 
        /// Note that due to GDI+ / GDI incompatibilities,
        /// the heights calculated by FlexCel will not be exactly the same than the ones calculated by Excel. So when you open this workbook
        /// in Excel, Excel might re calculate the row heights to what it believe is best. You can change this behaviour specifying keepSizesAutomatic = false.
        /// <remarks>THIS METHOD DOES NOT WORK ON COMPACT FRAMEWORK.</remarks>
        /// </summary>
        /// <param name="autofitNotAutofittingRows">When you are autofitting a range of rows, some rows might not be 
        /// set to Autofit in Excel. When this parameter is true, those rows will be autofitted anyway.</param>
        /// <param name="keepSizesAutomatic">
        /// When true, no modifications will be done to the "autofit" status of the rows. When false, all rows will be marked as "no autofit", so when you open this file
        /// in Excel it will not be resized by Excel, and the printing/export to pdf from Excel will be the same as FlexCel.
        /// </param>
        /// <param name="adjustment">You will normally want to set this parameter to 1, which means that autofit will be made with standard measurements.
        /// If you set it to for example 1.1, then rows will be adjusted to 110% percent of what their calculated height was. 
        /// Use this parameter to fine-tune autofiting, if for example rows are too small when opening the file in Excel.</param>
        /// <param name="adjustmentFixed">You will normally set this parameter to 0, which means standard autofit. If you set it to a value, the row will be
        /// made larger by that ammount from the calculated autofit. Different from the "adjustment" parameter, this parameter adds a fixed size to the row
        /// and not a percentage. The final size of the row will be:  FinalSize = CalulatedAutoFit * adjustment + adjusmentFixed</param>
        /// <param name="minHeight">Minimum final height for the row to autofit. If the calculated value is less than minHeight, row size will be set to minHeight.
        /// <br/>A negative value on minHeight means the row size will be no smaller than the original height.</param>
        /// <param name="maxHeight">Maximum final height for the row to autofit. If the calculated value is more than maxHeigth, row size will be set to maxHeight.
        /// <br/>maxHeight = 0 means no maxHeight.
        /// <br/>A negative value on maxHeight means the row size will be no bigger than the original height.
        /// </param>
		public void AutofitRowsOnWorkbook(bool autofitNotAutofittingRows, bool keepSizesAutomatic, real adjustment, int adjustmentFixed, int minHeight, int maxHeight)
		{
			AutofitRowsOnWorkbook(autofitNotAutofittingRows, keepSizesAutomatic, adjustment, adjustmentFixed, minHeight, maxHeight, TAutofitMerged.OnLastCell);
		}

		/// <summary>
		/// Autofits all rhe rows on all sheets on a workbook that are set to autofit so they adapt their height to show all the text inside. 
		/// Note that due to GDI+ / GDI incompatibilities,
		/// the heights calculated by FlexCel will not be exactly the same than the ones calculated by Excel. So when you open this workbook
		/// in Excel, Excel might re calculate the row heights to what it believe is best. You can change this behaviour specifying keepSizesAutomatic = false.
		/// <remarks>THIS METHOD DOES NOT WORK ON COMPACT FRAMEWORK.</remarks>
		/// </summary>
		/// <param name="autofitNotAutofittingRows">When you are autofitting a range of rows, some rows might not be 
		/// set to Autofit in Excel. When this parameter is true, those rows will be autofitted anyway.</param>
		/// <param name="keepSizesAutomatic">
		/// When true, no modifications will be done to the "autofit" status of the rows. When false, all rows will be marked as "no autofit", so when you open this file
		/// in Excel it will not be resized by Excel, and the printing/export to pdf from Excel will be the same as FlexCel.
		/// </param>
		/// <param name="adjustment">You will normally want to set this parameter to 1, which means that autofit will be made with standard measurements.
		/// If you set it to for example 1.1, then rows will be adjusted to 110% percent of what their calculated height was. 
		/// Use this parameter to fine-tune autofiting, if for example rows are too small when opening the file in Excel.</param>
		/// <param name="adjustmentFixed">You will normally set this parameter to 0, which means standard autofit. If you set it to a value, the row will be
		/// made larger by that ammount from the calculated autofit. Different from the "adjustment" parameter, this parameter adds a fixed size to the row
		/// and not a percentage. The final size of the row will be:  FinalSize = CalulatedAutoFit * adjustment + adjusmentFixed</param>
		/// <param name="minHeight">Minimum final height for the row to autofit. If the calculated value is less than minHeight, row size will be set to minHeight.
		/// <br/>A negative value on minHeight means the row size will be no smaller than the original height.</param>
		/// <param name="maxHeight">Maximum final height for the row to autofit. If the calculated value is more than maxHeigth, row size will be set to maxHeight.
		/// <br/>maxHeight = 0 means no maxHeight.
		/// <br/>A negative value on maxHeight means the row size will be no bigger than the original height.
		/// </param>
        /// <param name="autofitMerged">Specifies which row in a merged cell using more than one row will be used to autofit the merged cell.
        /// If you don't specify this parameter, it will be the last row.</param>
		public abstract void AutofitRowsOnWorkbook(bool autofitNotAutofittingRows, bool keepSizesAutomatic, real adjustment, int adjustmentFixed, int minHeight, int maxHeight, TAutofitMerged autofitMerged);

		/// <summary>
		/// Marks a row as candidate for future autofit. Note that this method will NOT change anything on the file. It just "marks" the row 
		/// so you can use it later with <see cref="AutofitMarkedRowsAndCols(bool, bool, real)"/>. To change the actual autofit status on the xls file, use <see cref="SetAutoRowHeight"/>
		/// <b>NOTE</b>: This method will not mark empty rows.
		/// </summary>
		/// <remarks>
		/// You can use this method for "delay-marking" rows that you will want to autofit later, but that you cannot autofit yet since
		/// they are not filled with data. There is normally no need to use this method, but it is used on report generation to "mark" &lt;#row height(autofit)&gt;
		/// tags so those rows and columns can be autofitted once the data on the sheet has been filled.
		/// </remarks>
		/// <param name="row">Row index (1 based)</param>
		/// <param name="autofit">Set this to true to mark the row for autofitting, false for removing the row from autofitting list.</param>
		/// <param name="adjustment">You will normally want to set this parameter to 1, which means that autofit will be made with standard measurements.
		/// If you set it to for example 1.1, then rows will be adjusted to 110% percent of what their calculated height was. 
		/// Use this parameter to fine-tune autofiting, if for example rows are too small when opening the file in Excel.</param>
		public void MarkRowForAutofit(int row, bool autofit, real adjustment)
		{
			MarkRowForAutofit(row, autofit, adjustment, 0, 0, 0, false);
		}

        /// <summary>
        /// Marks a row as candidate for future autofit. Note that this method will NOT change anything on the file. It just "marks" the row 
        /// so you can use it later with <see cref="AutofitMarkedRowsAndCols(bool, bool, real)"/>. To change the actual autofit status on the xls file, use <see cref="SetAutoRowHeight"/>
        /// <b>NOTE</b>: This method will not mark empty rows.
        /// </summary>
        /// <remarks>
        /// You can use this method for "delay-marking" rows that you will want to autofit later, but that you cannot autofit yet since
        /// they are not filled with data. There is normally no need to use this method, but it is used on report generation to "mark" &lt;#row height(autofit)&gt;
        /// tags so those rows and columns can be autofitted once the data on the sheet has been filled.
        /// </remarks>
        /// <param name="row">Row index (1 based)</param>
        /// <param name="autofit">Set this to true to mark the row for autofitting, false for removing the row from autofitting list.</param>
        /// <param name="adjustment">You will normally want to set this parameter to 1, which means that autofit will be made with standard measurements.
        /// If you set it to for example 1.1, then rows will be adjusted to 110% percent of what their calculated height was. 
        /// Use this parameter to fine-tune autofiting, if for example rows are too small when opening the file in Excel.</param>
        /// <param name="adjustmentFixed">You will normally set this parameter to 0, which means standard autofit. If you set it to a value, the row will be
        /// made larger by that ammount from the calculated autofit. Different from the "adjustment" parameter, this parameter adds a fixed size to the row
        /// and not a percentage. The final size of the row will be:  FinalSize = CalulatedAutoFit * adjustment + adjusmentFixed</param>
        /// <param name="minHeight">Minimum final height for the row to autofit. If the calculated value is less than minHeight, row size will be set to minHeight.
        /// <br/>A negative value on minHeight means the row size will be no smaller than the original height.</param>
        /// <param name="maxHeight">Maximum final height for the row to autofit. If the calculated value is more than maxHeigth, row size will be set to maxHeight.
        /// <br/>maxHeight = 0 means no maxHeight.
        /// <br/>A negative value on maxHeight means the row size will be no bigger than the original height.
        /// </param>
        /// <param name="isMerged">If true, only the cell will be autofitted, not the whole row.</param>
        public abstract void MarkRowForAutofit(int row, bool autofit, real adjustment, int adjustmentFixed, int minHeight, int maxHeight, bool isMerged);

		/// <summary>
		/// Marks a column as candidate for future autofit. Note that this method will NOT change anything on the file. It just "marks" the column 
		/// so you can use it later with <see cref="AutofitMarkedRowsAndCols(bool, bool, real)"/>.
		/// </summary>
		/// <remarks>
		/// You can use this method for "delay-marking" columns that you will want to autofit later, but that you cannot autofit yet since
		/// they are not filled with data.There is normally no need to use this method, but it is used on report generation to "mark" &lt;#col width(autofit)&gt;
		/// tags so those rows and columns can be autofitted once the data on the sheet has been filled.
		/// </remarks>
		/// <param name="col">Column index (1 based)</param>
		/// <param name="autofit">Set this to true to mark the column for autofitting, false for removing the column from autofitting list.</param>
		/// <param name="adjustment">You will normally want to set this parameter to 1, which means that autofit will be made with standard measurements.
		/// If you set it to for example 1.1, then columns will be adjusted to 110% percent of what their calculated width was. 
		/// Use this parameter to fine-tune autofiting, if for example columns are too small when opening the file in Excel.</param>
		public void MarkColForAutofit(int col, bool autofit, real adjustment)
		{
			MarkColForAutofit(col, autofit, adjustment, 0, 0, 0, false);
		}

        /// <summary>
        /// Marks a column as candidate for future autofit. Note that this method will NOT change anything on the file. It just "marks" the column 
        /// so you can use it later with <see cref="AutofitMarkedRowsAndCols(bool, bool, real)"/>.
        /// </summary>
        /// <remarks>
        /// You can use this method for "delay-marking" columns that you will want to autofit later, but that you cannot autofit yet since
        /// they are not filled with data.There is normally no need to use this method, but it is used on report generation to "mark" &lt;#col width(autofit)&gt;
        /// tags so those rows and columns can be autofitted once the data on the sheet has been filled.
        /// </remarks>
        /// <param name="col">Column index (1 based)</param>
        /// <param name="autofit">Set this to true to mark the column for autofitting, false for removing the column from autofitting list.</param>
        /// <param name="adjustment">You will normally want to set this parameter to 1, which means that autofit will be made with standard measurements.
        /// If you set it to for example 1.1, then columns will be adjusted to 110% percent of what their calculated width was. 
        /// Use this parameter to fine-tune autofiting, if for example columns are too small when opening the file in Excel.</param>
        /// <param name="adjustmentFixed">You will normally set this parameter to 0, which means standard autofit. If you set it to a value, the row will be
        /// made larger by that ammount from the calculated autofit. Different from the "adjustment" parameter, this parameter adds a fixed size to the column
        /// and not a percentage. The final size of the column will be:  FinalSize = CalulatedAutoFit * adjustment + adjusmentFixed</param>
        /// <param name="minWidth">Minimum final width for the column to autofit. If the calculated value is less than minWidth, column size will be set to minWidth.
        /// <br/>A negative value on minWidth means the column size will be no smaller than the original width.</param>
        /// <param name="maxWidth">Maximum final width for the column to autofit. If the calculated value is more than maxWidth, column size will be set to maxWidth.
        /// <br/>maxWidth = 0 means no maxWidth.
        /// <br/>A negative value on maxWidth means the column size will be no bigger than the original width.
        /// </param>
		/// <param name="isMerged">If true, only the cell will be autofitted, not the whole column.</param>
		public abstract void MarkColForAutofit(int col, bool autofit, real adjustment, int adjustmentFixed, int minWidth, int maxWidth, bool isMerged);

		/// <summary>
		/// Autofits all the rows and columns on a sheet that have been previously marked with the <see cref="MarkRowForAutofit(int, bool, real)"/>  and <see cref="MarkColForAutofit(int, bool, real)"/> methods.
		/// </summary>
		/// <param name="keepSizesAutomatic">
		/// When true, no modifications will be done to the "autofit" status of the rows. When false, all rows will be marked as "no autofit", so when you open this file
		/// in Excel it will not be resized by Excel, and the sizes when printing/export to pdf from Excel will be the same as FlexCel, even when some cells might appear "cut" when printing on Excel.
		/// </param>
		/// <param name="ignoreStringsOnColumnFit">When true, cells containing strings will not be autofitted.</param>
		/// <param name="adjustment">You will normally want to set this parameter to 1, which means that autofit will be made with standard measurements.
		/// If you set it to for example 1.1, then columns and rows will be adjusted to 110% percent of what their calculated width and heigth was. 
		/// Use this parameter to fine-tune autofiting, if for example columns are too small when opening the file in Excel.</param>
		public void AutofitMarkedRowsAndCols(bool keepSizesAutomatic, bool ignoreStringsOnColumnFit, real adjustment)
		{
			AutofitMarkedRowsAndCols(keepSizesAutomatic, ignoreStringsOnColumnFit, adjustment, 0, 0, 0, 0, 0);
		}

		/// <summary>
		/// Autofits all the rows and columns on a sheet that have been previously marked with the <see cref="MarkRowForAutofit(int, bool, real)"/>  and <see cref="MarkColForAutofit(int, bool, real)"/> methods.
		/// </summary>
		/// <param name="keepSizesAutomatic">
		/// When true, no modifications will be done to the "autofit" status of the rows. When false, all rows will be marked as "no autofit", so when you open this file
		/// in Excel it will not be resized by Excel, and the sizes when printing/export to pdf from Excel will be the same as FlexCel, even when some cells might appear "cut" when printing on Excel.
		/// </param>
		/// <param name="ignoreStringsOnColumnFit">When true, cells containing strings will not be autofitted.</param>
		/// <param name="adjustment">You will normally want to set this parameter to 1, which means that autofit will be made with standard measurements.
		/// If you set it to for example 1.1, then columns and rows will be adjusted to 110% percent of what their calculated width and heigth was. 
		/// Use this parameter to fine-tune autofiting, if for example columns are too small when opening the file in Excel.</param>
		/// <param name="adjustmentFixed">You will normally set this parameter to 0, which means standard autofit. If you set it to a value, the row will be
		/// made larger by that ammount from the calculated autofit. Different from the "adjustment" parameter, this parameter adds a fixed size to the row
		/// and not a percentage. The final size of the row will be:  FinalSize = CalulatedAutoFit * adjustment + adjusmentFixed</param>
		/// <param name="minHeight">Minimum final height for the row to autofit. If the calculated value is less than minHeight, row size will be set to minHeight.
		/// <br/>A negative value on minHeight means the row size will be no smaller than the original height.</param>
		/// <param name="maxHeight">Maximum final height for the row to autofit. If the calculated value is more than maxHeigth, row size will be set to maxHeight.
		/// <br/>maxHeight = 0 means no maxHeight.
		/// <br/>A negative value on maxHeight means the row size will be no bigger than the original height.
		/// </param>
		/// <param name="minWidth">Minimum final width for the column to autofit. If the calculated value is less than minWidth, column size will be set to minWidth.
		/// <br/>A negative value on minWidth means the column size will be no smaller than the original width.</param>
		/// <param name="maxWidth">Maximum final width for the column to autofit. If the calculated value is more than maxWidth, column size will be set to maxWidth.
		/// <br/>maxWidth = 0 means no maxWidth.
		/// <br/>A negative value on maxWidth means the column size will be no bigger than the original width.
		/// </param>
		public void AutofitMarkedRowsAndCols(bool keepSizesAutomatic, bool ignoreStringsOnColumnFit, real adjustment, int adjustmentFixed, int minHeight, int maxHeight, int minWidth, int maxWidth)
		{
			AutofitMarkedRowsAndCols(keepSizesAutomatic, ignoreStringsOnColumnFit, adjustment, adjustmentFixed, minHeight, maxHeight, minWidth, maxWidth, TAutofitMerged.OnLastCell);
		}

        /// <summary>
        /// Autofits all the rows and columns on a sheet that have been previously marked with the <see cref="MarkRowForAutofit(int, bool, real)"/>  and <see cref="MarkColForAutofit(int, bool, real)"/> methods.
        /// </summary>
        /// <param name="keepSizesAutomatic">
        /// When true, no modifications will be done to the "autofit" status of the rows. When false, all rows will be marked as "no autofit", so when you open this file
        /// in Excel it will not be resized by Excel, and the sizes when printing/export to pdf from Excel will be the same as FlexCel, even when some cells might appear "cut" when printing on Excel.
        /// </param>
        /// <param name="ignoreStringsOnColumnFit">When true, cells containing strings will not be autofitted.</param>
        /// <param name="adjustment">You will normally want to set this parameter to 1, which means that autofit will be made with standard measurements.
        /// If you set it to for example 1.1, then columns and rows will be adjusted to 110% percent of what their calculated width and heigth was. 
        /// Use this parameter to fine-tune autofiting, if for example columns are too small when opening the file in Excel.</param>
        /// <param name="adjustmentFixed">You will normally set this parameter to 0, which means standard autofit. If you set it to a value, the row will be
        /// made larger by that ammount from the calculated autofit. Different from the "adjustment" parameter, this parameter adds a fixed size to the row
        /// and not a percentage. The final size of the row will be:  FinalSize = CalulatedAutoFit * adjustment + adjusmentFixed</param>
        /// <param name="minHeight">Minimum final height for the row to autofit. If the calculated value is less than minHeight, row size will be set to minHeight.
        /// <br/>A negative value on minHeight means the row size will be no smaller than the original height.</param>
        /// <param name="maxHeight">Maximum final height for the row to autofit. If the calculated value is more than maxHeigth, row size will be set to maxHeight.
        /// <br/>maxHeight = 0 means no maxHeight.
        /// <br/>A negative value on maxHeight means the row size will be no bigger than the original height.
        /// </param>
        /// <param name="minWidth">Minimum final width for the column to autofit. If the calculated value is less than minWidth, column size will be set to minWidth.
        /// <br/>A negative value on minWidth means the column size will be no smaller than the original width.</param>
        /// <param name="maxWidth">Maximum final width for the column to autofit. If the calculated value is more than maxWidth, column size will be set to maxWidth.
        /// <br/>maxWidth = 0 means no maxWidth.
        /// <br/>A negative value on maxWidth means the column size will be no bigger than the original width.
        /// </param>
        /// <param name="autofitMerged">Specifies which row in a merged cell using more than one row, or which column in a merged cell with more than one column will be used to autofit the merged cell.
        /// If you don't specify this parameter, it will be the last row or column in the merged range.</param>
        public abstract void AutofitMarkedRowsAndCols(bool keepSizesAutomatic, bool ignoreStringsOnColumnFit, real adjustment, int adjustmentFixed, int minHeight, int maxHeight, int minWidth, int maxWidth, TAutofitMerged autofitMerged);


		#region Cols by index.  
		/// <summary>
		/// This method returns the existing columns on ONE ROW.
		/// You can use this together with <see cref="ColFromIndex(int,int)"/> and <see cref="ColToIndex(int,int)"/> to iterate faster on a block.
		/// </summary>
		/// <param name="row">Row index. (1-based)</param>
		/// <returns>The number of existing columns on one row.</returns>
		/// <example>
		/// Instead of writing:
		///   <code>
		///   for (int r=1; r&lt;=File.RowCount;r++)
		///     for (int c=1; clt;=File.ColCount;c++)
		///       DoSomething(r,c);
		///   </code> 
		///  You can use:
		///   <code>
		///   for (int r=1; r&lt;=File.RowCount;r++)
		///     for (int c=1; c&lt;=File.ColCountInRow(r);c++)
		///       DoSomething(r,ColFromIndex(c));
		///   </code>   
		///</example>
		public abstract int ColCountInRow(int row);

		/// <summary>
		/// This method returns the existing columns on ONE ROW, for a given sheet.
		/// You can use this together with <see cref="ColFromIndex(int,int)"/> and <see cref="ColToIndex(int,int)"/> to iterate faster on a block.
		/// </summary>
		/// <param name="sheet">Sheet where we are working. It might be different from ActiveSheet.</param>
		/// <param name="row">Row index. (1-based)</param>
		/// <returns>The number of existing columns on one row.</returns>
		/// <example>
		/// Instead of writing:
		///   <code>
		///   for (int r=1; r&lt;=File.RowCount;r++)
		///     for (int c=1; clt;=File.ColCount;c++)
		///       DoSomething(r,c);
		///   </code> 
		///  You can use:
		///   <code>
		///   for (int r=1; r&lt;=File.RowCount;r++)
		///     for (int c=1; c&lt;=File.ColCountInRow(r);c++)
		///       DoSomething(r,ColFromIndex(c));
		///   </code>   
		///</example>
		public abstract int ColCountInRow(int sheet, int row);

		/// <summary>
		/// This is the column (1 based) for a given ColIndex. See <see cref="ColCountInRow(int)"/> for an example.
		/// </summary>
		/// <param name="row">Row (1 based)</param>
		/// <param name="colIndex">The index on the column list for the row. (1 based)</param>
		/// <returns>The column (1 based) for the corresponding item.</returns>
		public abstract int ColFromIndex(int row, int colIndex);

		/// <summary>
		/// This is the column (1 based) for a given ColIndex and sheet. See <see cref="ColCountInRow(int)"/> for an example.
		/// </summary>
		/// <param name="sheet">Sheet where we are working. It might be different from ActiveSheet.</param>
		/// <param name="row">Row (1 based)</param>
		/// <param name="colIndex">The index on the column list for the row. (1 based)</param>
		/// <returns>The column (1 based) for the corresponding item.</returns>
		public abstract int ColFromIndex(int sheet, int row, int colIndex);

		/// <summary>
		/// This is the inverse of <see cref="ColFromIndex(int,int)"/>. It will return the index on the 
		/// internal column array from the row for an existing column. If the column doesn't exist, it will return the 
		/// index of the "LAST existing column less than col", plus 1.
		/// </summary>
		/// <example>
		/// To loop on all the existing cells on a row you can use:
		/// int LastCIndex = xls.ColToIndex(row, LastColumn + 1);
		/// if (xls.ColFromIndex(LastCIndex) > LastColumn) LastCIndex --;  // LastColumn does not exist.
		/// for (int cIndex = xls.ColToIndex(FirstColumn); cIndex &lt;= LastCIndex; cIndex++)
		/// {
		///     xls.GetCellValueIndexed(row, cIndex, ref XF);
		/// }
		/// </example>
		/// <param name="row">Row (1 based)</param>
		/// <param name="col">Column (1 based)</param>
		/// <returns>The index on the column list for the row. (1 based)</returns>
		public abstract int ColToIndex(int row, int col);

		/// <summary>
		/// This is the inverse of <see cref="ColFromIndex(int,int)"/>. It will return the index on the 
		/// internal column array from the row for an existing column. If the column doesn't exist, it will return the 
		/// index of the "LAST existing column less than col", plus 1.
		/// </summary>
		/// <example>
		/// To loop on all the existing cells on a row you can use:
		/// int LastCIndex = xls.ColToIndex(row, LastColumn + 1);
		/// if (xls.ColFromIndex(LastCIndex) > LastColumn) LastCIndex --;  // LastColumn does not exist.
		/// for (int cIndex = xls.ColToIndex(FirstColumn); cIndex &lt;= LastCIndex; cIndex++)
		/// {
		///     xls.GetCellValueIndexed(row, cIndex, ref XF);
		/// }
		/// </example>
		/// <param name="sheet">Sheet where we are working. It might be different from ActiveSheet.</param>
		/// <param name="row">Row (1 based)</param>
		/// <param name="col">Column (1 based)</param>
		/// <returns>The index on the column list for the row. (1 based)</returns>
		public abstract int ColToIndex(int sheet, int row, int col);
		#endregion
		#endregion

		#region Indexed Color
		/// <summary>
		/// Changes a color on the Excel color palette.
		/// </summary>
		/// <param name="index">Index of the entry to change. Must be 1&lt;=indexlt;=<see cref="ColorPaletteCount"/></param>
		/// <param name="value">Color to set.</param>
		public abstract void SetColorPalette(int index, Color value);


		/// <summary>
		/// Returns a color from the color palette. This method will throw an exception if its "index" parameter
		/// is bigger than <see cref="ColorPaletteCount"/>, (for example, for an automatic color).
		/// To get the real color, use <see cref="GetColorPalette(int, Color)"/>
		/// </summary>
		/// <param name="index">Index of the entry to return. Must be 1&lt;=index&lt;=<see cref="ColorPaletteCount"/></param>
		/// <returns>Color at position index.</returns>
		public abstract Color GetColorPalette(int index);

		/// <summary>
		/// Returns a color from the color palette. If the index is not into the range 1&lt;=index&lt;=<see cref="ColorPaletteCount"/>
		/// this method will return the automaticColor.
		/// </summary>
		/// <remarks>ColorIndexes returned by FlexCel might be &lt;=0 or &gt;<see cref="ColorPaletteCount"/> if the color is
		/// set to Automatic.<p>Automatic color is white for backgrounds, black for foregrounds and gray for gridlines.</p></remarks>
		/// <param name="index"></param>
		/// <param name="automaticColor"></param>
		/// <returns></returns>
		public Color GetColorPalette(int index, Color automaticColor)
		{
			if (index<=0 || index> ColorPaletteCount) return automaticColor;
			return GetColorPalette(index);
		}
        
		/// <summary>
		/// The number of entries on an Excel color palette. This is  always 56.
		/// </summary>
		public int ColorPaletteCount
		{
			get
			{
				return 56;
			}
		}

		/// <summary>
		/// Returns a list of the used colors on the palette. You can use it as an entry to <see cref="NearestColorIndex(Color)"/>
		/// to modify the palette.
		/// </summary>
		public abstract bool[] GetUsedPaletteColors{get;}

		/// <summary>
		/// Returns the most similar entry on the excel palette for a given color.
		/// </summary>
		/// <param name="value">Color we want to use.</param>
		/// <returns>Most similar color on the Excel palette.</returns>
		public abstract int NearestColorIndex(Color value);

		/// <summary>
        /// <b>IMPORTANT:</b> Since FlexCel 5.1, using <see cref="OptimizeColorPalette"/> before saving should normally be used instead of this method
        /// to get an optimized palette. Just enter the true colors in FlexCel, and call <see cref="OptimizeColorPalette"/> before saving.<br></br>
		/// Returns the most similar entry on the excel palette for a given color.
		/// If UsedColors is not null, it will try to modify the Excel color palette
		/// to get a better match on the color, modifying among the not used colors. 
		/// Note that modifying the standard palette
		/// might result on a file that is not easy to edit on Excel later, since it does not have the standard Excel colors.
		/// </summary>
		/// <param name="value">Color we want to use.</param>
		/// <param name="UsedColors">If null, this behaves like the standard NearestColorIndex. 
		/// To get a list of used colors for the first call, use <see cref="GetUsedPaletteColors"/>.
		/// After the first call, keep using the same UsedColors structure and do not call GetUsedPaletteColors again, to avoid overwriting colors
		/// that are not yet inserted into the xls file with new ones. You can call GetUsedPaletteColors only after you added the format with <see cref="AddFormat"/>
		/// </param>
		/// <returns>Most similar color on the Excel palette.</returns>
		public abstract int NearestColorIndex(Color value, bool[] UsedColors);

        /// <summary>
        /// Returns true if the internal color palette contains the exact specified color. Note that Excel 2007 doesn't use the color palette, so this
        /// method is not needed there.
        /// </summary>
        /// <param name="value">Color to check.</param>
        /// <returns>True if color is defined.</returns>
        public abstract bool PaletteContainsColor(TExcelColor value);

        /// <summary>
        /// Changes the colors in the color palette so they can represent better the colors in use. This method will change the colors not used in the palette
        /// by colors used in the sheet. If there are more unique colors in the sheet than the 56 available in the palette, only the first colors will be changed.
        /// <br></br>
        /// When FlexCel saves an xls file, it saves the color information twice: The real color for Excel 2007 and newer, and the indexed color
        /// for older Excel versions. This method optimizes the palette of indexed colors so they look better in Excel 2003 or older. It doesn't effect
        /// Excel 2007 or newer at all.
        /// </summary>
        public abstract void OptimizeColorPalette();
		#endregion

        #region Theme Color
#if (FRAMEWORK30)
        /// <summary>
        /// Changes a color on the Excel theme. Only has effect in Excel 2007, and you need .NET 3.5 or newer to use this method.
        /// <br></br>If you want to change the full theme, use <see cref="GetTheme"/> and <see cref="SetTheme"/>
        /// </summary>
        /// <param name="themeColor">Color of the theme to change.</param>
        /// <param name="value">Color to set.</param>
#else
        /// <summary>
        /// This method doesn't work in .NET 2.0
        /// </summary>
        /// <param name="themeColor">Color of the theme to change.</param>
        /// <param name="value">Color to set.</param>
#endif
        public abstract void SetColorTheme(TThemeColor themeColor, TDrawingColor value);


#if (FRAMEWORK30)
        /// <summary>
        /// Returns a color from the active theme palette. 
        /// Only has effect in Excel 2007, and you need .NET 3.5 or newer to get a color different from the standard "office" theme.
        /// <br></br> To get the full theme, look at <see cref="GetTheme"/>
        /// </summary>
        /// <param name="themeColor">Color of the theme to get.</param>
        /// <returns>Color for the given theme.</returns>
#else
        /// <summary>
        /// This method doesn't work in .NET 2.0
        /// </summary>
        /// <param name="themeColor">Color of the theme to get.</param>
        /// <returns>Color for the given theme.</returns>
#endif         
        public abstract TDrawingColor GetColorTheme(TThemeColor themeColor);

        /// <summary>
        /// Returns the most similar entry on the theme palette for a given color.
        /// </summary>
        /// <param name="value">Color we want to use.</param>
        /// <param name="tint">Returns the tint to apply to the theme color.</param>
        /// <returns>Most similar color on the theme palette.</returns>
        public abstract TThemeColor NearestColorTheme(Color value, out double tint);

        /// <summary>
        /// Gets the major of minor font scheme in the theme.
        /// </summary>
        /// <param name="fontScheme">Font Scheme we want to get (either minor or major). Using "none" here will return null.</param>
        /// <returns>Font definition.</returns>
        public abstract TThemeFont GetThemeFont(TFontScheme fontScheme);

        /// <summary>
        /// Sets either the minor or the major font for the theme.
        /// </summary>
        /// <param name="fontScheme">Font Scheme we want to set (either minor or major). Using "none" here will do nothing.</param>
        /// <param name="font">Font definition.</param>
        public abstract void SetThemeFont(TFontScheme fontScheme, TThemeFont font);


#if (FRAMEWORK30)
        /// <summary>
        /// This is an advanced method, that allows you to get the full theme in use. Normally you will just want to replace colors, and you can do this with
        /// <see cref="SetColorTheme"/> and <see cref="GetColorTheme"/> methods. Much of the functionality in a theme applies to PowerPoint, not Excel.<br/>
        /// This method is only available in .NET 3.5 or newer
        /// </summary>
        /// <returns></returns>
        public abstract TTheme GetTheme();

        /// <summary>
        /// This is an advanced method, that allows you to set the full theme in use. Normally you will just want to replace colors, and you can do this with
        /// <see cref="SetColorTheme"/> and <see cref="GetColorTheme"/> methods.<br/>
        /// This method is only available in .NET 3.5 or newer
        /// </summary>
        /// <param name="aTheme">Theme to set. You would normally use the result from <see cref="GetTheme"/> here, or you might load a method from a ".tmx" file.
        /// There are many standard tmx files available in an Office instalation under the "Document Themes Version" folder</param>
        public abstract void SetTheme(TTheme aTheme);
#endif

        #endregion

        #region Named Ranges
        /// <summary>
		/// The count of all named ranges on the file.
		/// </summary>
		public abstract int NamedRangeCount{get;}

		/// <summary>
		/// Returns the Named Range Definition. If the range is not user defined (like "Print_Area")
		/// it will have a one-char name, and the value is on the enum <see cref="InternalNameRange"/>
		/// </summary>
		/// <remarks><seealso cref="InternalNameRange"/></remarks>
		/// <param name="index">Index of the named range.</param>
		/// <returns>Named Range</returns>
		public abstract TXlsNamedRange GetNamedRange(int index);

		/// <summary>
		/// Returns the Named Range Definition. If the range is not user defined (like "Print_Area")
		/// it will have a one-char name, and the value is on the enum <see cref="InternalNameRange"/>
		/// </summary>
		/// <remarks><seealso cref="InternalNameRange"/></remarks>
		/// <param name="Name">Name of the range we are looking for. Case insensitive.</param>
		/// <param name="refersToSheetIndex">Sheet where the range refers to. A range with the same name might be defined on more than one sheet.<br></br>
		/// To get the first range with the name, make refersToSheetIndex&lt;=0
        /// <br></br>Note that a name can be stored in one sheet and refer to a different sheet. Or it might not refer to any sheet at all. (for example, the name "=name1+1"
        /// doesn't refer to any sheet.) And you could have a name "=Sheet1!a1" defined in sheet2. This name will "refer to" sheet1, but be defined in sheet2.
        /// <br></br>While names are normally all defined globally (sheet 0), if you need to get the names that are stored in an specific sheet, use
        /// <see cref="GetNamedRange(string, int, int)"/> instead.</param>
		/// <returns>Named Range or null if the range does not exist.</returns>
		public abstract TXlsNamedRange GetNamedRange(string Name, int refersToSheetIndex);

		/// <summary>
		/// Returns the Named Range Definition. If the range is not user defined (like "Print_Area")
		/// it will have a one-char name, and the value is on the enum <see cref="InternalNameRange"/>
		/// </summary>
		/// <remarks><seealso cref="InternalNameRange"/></remarks>
		/// <param name="Name">Name of the range we are looking for. Case insensitive.</param>
		/// <param name="refersToSheetIndex">Sheet where the range refers to. Note that this is <b>not</b> where the name is stored. Normally names are all
        /// stored globally (in sheet 0), but they refer to an specific sheet. For example, if you define the name "=Sheet2!a1" globally, it will be stored in sheet 0, but will refer to sheet2.
        /// <br></br>A name like "=Sheet1:Sheet2!A1 doesn't refer to any sheet.</param>
		/// <param name="localSheetIndex">Sheet where the range is stored. A range might be stored local to a sheet, or global (Excel default).
		/// To get a global range, make localSheetIndex=0</param>
		/// <returns>Named Range or null if the range does not exist.</returns>
		public abstract TXlsNamedRange GetNamedRange(string Name, int refersToSheetIndex, int localSheetIndex);

		/// <summary>
		/// Returns the index (0 based) on the list of named ranges for a given name and local sheet. If the range is not found, this method will return -1
		/// You could use <see cref="GetNamedRange(int)"/> to get the name definition, or directly call <see cref="GetNamedRange(string,int,int)"/>
		/// to get a named range knowing its name and sheet position.
		/// </summary>
		/// <param name="Name">Name of the range we are looking for. Case insensitive.</param>
		/// <param name="localSheetIndex">Sheet where the range is stored. A range might be stored local to a sheet, or global (Excel default).
		/// To get a global range, make localSheetIndex=0</param>
		/// <returns>The index (0 based) in the list of named ranges, or -1 if the range is not found.</returns>
		public abstract int FindNamedRange(string Name, int localSheetIndex);

		/// <summary>
		/// Internal use. We could call GetNamedRange, but this one is faster.
		/// </summary>
		internal abstract TParsedTokenList GetNamedRangeData(int nameIndex, out string externalName, out bool isAddin, out bool Error);

		/// <summary>
		/// Internal use. We could call GetNamedRange, but this one is faster.
		/// </summary>
		internal abstract TParsedTokenList GetNamedRangeData(int externSheetIndex, int externNameIndex, out string externalBook, out string externalName, out int sheetIndexInOtherFile, out bool isAddin, out bool Error);

		/// <summary>
		/// Internal use for the formula evaluator in the report tags.
		/// </summary>
		/// <returns></returns>
		internal abstract INameRecordList GetNameRecordList();

		/// <summary>
		/// Internal use. Evaluates a named range and returns the result.
		/// </summary>
        internal abstract object EvaluateNamedRange(int externNameIndex, int SheetIndex, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack);

		/// <summary>
		/// Modifies or adds a Named Range. If the named range exists, it will be modified, else it will be added.
		/// If the range is not user defined (like "Print_Area")
		/// it will have a one-char name, and the value is on the enum <see cref="InternalNameRange"/>
		/// Look at the example for more information.
		/// </summary>
		/// <example>
		/// This will create a range for repeating the first 2 columns and rows on each printed page (on sheet 1):
		/// <code>
		///    Xls.SetNamedRange(new TXlsNamedRange(TXlsNamedRange.GetInternalName(InternalNameRange.Print_Titles),1, 0, "=1:2,A:B"));
		/// </code>
		/// Note that in this example in particular (Print_Titles), the range has to have full rows/columns, as this is what Excel expects.
		/// You should also use "A:B" notation instead of the full "A1:B65536" name, so it will work in Excel 2007 too.
		/// </example>
		/// <param name="rangeData">Data of the named range. You don't need to specify the RPN Array.</param>
        /// <returns>The name index of the inserted or modified range (1 based).</returns>
		public abstract int SetNamedRange(TXlsNamedRange rangeData);

		/// <summary>
		/// Modifies a Named Range in the specified position. You could normally use <see cref="SetNamedRange(TXlsNamedRange)"/> to do this,
		/// but if you want to modify the name of the named range, then you need to use this overloaded version. <see cref="SetNamedRange(TXlsNamedRange)"/>
		/// would add a new range instead of modifying the existing one if you change the name.
		/// Look at the example for more information on how to use it.
		/// </summary>
		/// <example>
		/// This will modify the name of the named range at position 3:
		/// <code>
		///    TXlsNamedRange R = Xls.GetNamedRange(3);
		///    R.Name = "MyNewName";
		///    Xls.SetNamedRange(3, R);
		/// </code>
		/// </example>
		/// <param name="index">Index of the named range we are trying to modify.</param>
		/// <param name="rangeData">Data of the named range. You don't need to specify the RPN Array.</param>
		public abstract void SetNamedRange(int index, TXlsNamedRange rangeData);
		

		/// <summary>
		/// Deletes the name at the specified position. <b>Important:</b> If the name you are trying to delete is referenced by any formula/chart/whatever in your file,
		/// the name will <b>not actually be deleted</b> but hidden. <br/> 
        /// You won't see the name in Excel or in the formula, but it will be there and you can see it from FlexCel.<br/>You can use <see cref="GetUsedNamedRanges"/> to learn if a range might be deleted.<br/>
		/// <i>Also, note that if you later delete the formulas that reference those ranges FlexCel will remove those hanging ranges when saving.</i>
		/// The only hidden ranges that will be present in the final file will be those that have active formulas referencing them.
		/// <br/><br/><b>Important:</b>If the name wasn't deleted, <see cref="NamedRangeCount"/> will not change. This means that you can't have code like this:
		/// <code>
		///   while (xls.NamedRangeCount > 0) //WRONG! This loop might never end.
		///   {
		///     xls.DeleteNamedRange(1);  //Might not be deleted, and NamedRangeCount will never be 0.
		///   }
		/// </code>
		/// <br/>
		/// The correct code in this case would be:
		/// <code>
		///   for (int i = xls.NamedRangeCount; i > 0; i--) xls.DeleteNamedRange(i);
		/// </code>
		/// <br/>
		/// <code lang = "vbnet">
		///   For i = xls.NamedRangeCount To 1 Step -1 xls.DeleteNamedRange(i);
		/// </code>
		/// <br/>
		/// <code lang = "Delphi .NET" title = "Delphi .NET">
		///   for i := xls.NamedRangeCount downto 1 do xls.DeleteNamedRange(i);
		/// </code>
		/// </summary>
		/// <param name="index">Index of the name to delete (1 based).</param>
		public abstract void DeleteNamedRange(int index);

        /// <summary>
        /// Returns an array of booleans where each value indicates if the name at position "i-1" is used by any formula, chart, or object in the file. If the name is in use, it can't be deleted. Note that the index here is
        /// Zero-based, different from all other Name indexes in FlexCel, because arrays in C# are always 0-based. So UsedRange[0] corresponds to GetNamedRange(1) and so on.
        /// </summary>
        /// <returns>An array of booleans indicating whether each name is used or not.</returns>
        public abstract bool[] GetUsedNamedRanges();

        internal abstract int AddEmptyName(string name, int sheet);

		#endregion

		#region Copy and Paste
		/// <summary>
		/// Copies the active sheet to a clipboard stream, on native and text formats.
		/// </summary>
		/// <remarks>See the copy and paste demo.</remarks>
		/// <param name="textString">StringBuilder where the text will be copied. Leave it null to not copy to text.</param>
		/// <param name="xlsStream">Stream where the xls native info will be copied.</param>
		public void CopyToClipboard(StringBuilder textString, Stream xlsStream)
		{

			if (RowCount <= 0) return;
            int cc = ColCount;
            if (cc <= 0) return;
            CopyToClipboardFormat(new TXlsCellRange(1, 1, RowCount, cc), textString, xlsStream);
		}

		/// <summary>
		/// Copies a range on the active sheet to a clipboard stream, on native and text formats.
		/// </summary>
		/// <remarks>See the copy and paste demo.</remarks>
		/// <param name="range">Range with the cells to copy.</param>
		/// <param name="textString">StringBuilder where the text will be copied. Leave it null to not copy to text.</param>
		/// <param name="xlsStream">Stream where the xls native info will be copied.</param>
		public abstract void CopyToClipboardFormat(TXlsCellRange range, StringBuilder textString, Stream xlsStream);

		/// <summary>
		/// Pastes the clipboard contents beginning on cells row, col.
		/// </summary>
		/// <remarks>See the copy and paste demo.</remarks>
		/// <param name="row">First row where to paste.</param>
		/// <param name="col">First column where to paste.</param>
		/// <param name="data">A stream containing a Native xls format.</param>
		/// <param name="insertMode">How the pasted cells will be inserted on the file.</param>
		public abstract void PasteFromXlsClipboardFormat(int row, int col, TFlxInsertMode insertMode, Stream data);

		/// <summary>
		/// Pastes the clipboard contents beginning on cells row, col.
		/// </summary>
		/// <remarks>See the copy and paste demo.</remarks>
		/// <param name="row">First row where to paste.</param>
		/// <param name="col">First column where to paste.</param>
		/// <param name="data">A string containing a tab separated text format.</param>
		/// <param name="insertMode">How the pasted cells will be inserted on the file.</param>
		public abstract void PasteFromTextClipboardFormat(int row, int col, TFlxInsertMode insertMode, string data);
		#endregion

		#region Print
		/// <summary>
		/// Page header on the active sheet.
        /// <b>Note that this property sets the same header for the all the pages.</b> In Excel 2007 or newer you can set a different
        /// header for the first page, or odd/even pages. If you want to control these options, see <see cref="GetPageHeaderAndFooter()"/> and <see cref="SetPageHeaderAndFooter(THeaderAndFooter)"/>.
        /// <br/>
		/// A page header is a string that contains the text for the 3 parts of the header.<p></p>
		/// The Left section begins with &amp;L, the Center section with &amp;C and the Right with &amp;R<p></p>
		/// For example, the text"&amp;LThis goes at the left!&amp;CThis is centered!&amp;RThis is right aligned"
		/// will write text to all the sections.
		/// <p></p><p></p>
		/// This is the full list of macros you can include:
		///<p></p>
		/// <list type="bullet">
		/// <item>&amp;&amp; The "&amp;" character itself</item>
		/// <item>&amp;L Start of the left section</item>
		/// <item>&amp;C Start of the centered section</item>
		/// <item>&amp;R Start of the right section</item>
		/// <item>&amp;P Current page number</item>
		/// <item>&amp;N Page count</item>
		/// <item>&amp;D Current date</item>
		/// <item>&amp;T Current time</item>
		/// <item>&amp;A Sheet name</item>
		/// <item>&amp;F File name without path</item>
		/// <item>&amp;Z File path without file name (XP or Newer)</item>
		/// <item>&amp;G Picture (XP or Newer)</item>
		/// <item>&amp;U Underlining on/off</item>
		/// <item>&amp;E Double underlining on/off</item>
		/// <item>&amp;S Strikeout on/off</item>
		/// <item>&amp;X Superscript on/off</item>
		/// <item>&amp;Y Subscript on/off</item>
		/// <item>&amp;"&lt;fontname&gt;" Set new font &lt;fontname&gt;</item>
		/// <item>&amp;"&lt;fontname&gt;,&lt;fontstyle&gt;" Set new font with specified style &lt;fontstyle&gt;. The style &lt;fontstyle&gt; is in most cases
		/// one of "Regular", "Bold", "Italic", or "Bold Italic". But this setting is dependent on the
		/// used font, it may differ (localized style names, or "Standard", "Oblique", ...).</item>
		/// <item>&amp;&lt;fontheight&gt; Set font height in points (&lt;fontheight&gt; is a decimal value). If this command is followed
		/// by a plain number to be printed in the header, it will be separated from the font height
		/// with a space character.</item>
		/// </list>
		/// <p></p>
		/// Normally, the easiest way to find out which header string you need is to create an xls file on Excel, add a header, open the file
		/// with FlexCel and take a look at the generated header (You can use the ApiMate tool for that).
		/// </summary>
		public abstract string PageHeader{get;set;}

		/// <summary>
		/// Page footer on the active sheet. For a description on the format of the string, see <see cref="PageHeader"/>
		/// </summary>
		public abstract string PageFooter{get;set;}

        /// <summary>
        /// This method will return all the headers and footers in a sheet.
        /// </summary>
        /// <returns></returns>
        public abstract THeaderAndFooter GetPageHeaderAndFooter();

        /// <summary>
        /// This method will set all the headers and footers in a sheet. If you want a simple header or footer for all the pages, you might want to use <see cref="PageHeader"/> and <see cref="PageFooter"/>
        /// </summary>
        /// <param name="headerAndFooter">Structure with the headers and footers definition.</param>
        /// <returns></returns>
        public abstract void SetPageHeaderAndFooter(THeaderAndFooter headerAndFooter);

		/// <summary>
		/// Given a Page Header or footer string including macros (like [FileName] or [PageNo]), this method
		/// will return the strings that go into the left, right and middle sections.
		/// </summary>
		/// <param name="fullText">Header or footer text.</param>
		/// <param name="leftText">Text that should be left justified.</param>
		/// <param name="centerText">Text that should be centered.</param>
		/// <param name="rightText">Text that should be right justified.</param>
		public abstract void FillPageHeaderOrFooter(string fullText, ref string leftText, ref string centerText, ref string rightText);

		/// <summary>
		/// Converts a section of a page header or footer into an HTML string. 
		/// </summary>
		/// <param name="section">Text to convert. You will normally get this parameter calling <see cref="FillPageHeaderOrFooter"/></param>
		/// <param name="imageTag">Tag to call an image for the section. It must be in the form: "&lt;img src=...&gt;" 
        /// and you can get the image with <see cref="GetHeaderOrFooterImage(THeaderAndFooterKind, THeaderAndFooterPos, ref TXlsImgType)"/> . If null or empty, no img tag will be present in the resulting html, even if the section includes an image.</param>
		/// <param name="pageNumber">Page we are printing. This parameter will be used if you have text like "Page 1 of 3" in the header.<br/>
		/// If you are exporting to html, this value should be 1, since there are no page breaks in an html doc.</param>
		/// <param name="pageCount">Number of pages in the document. This parameter will be used if you have text like "Page 1 of 3" in the header.<br/>
		/// If you are exporting to html, this value should be 1, since there are no page breaks in an html doc</param>
		/// <param name="htmlVersion">Version of html we are targeting. In Html 4 &lt;br&gt; is valid and &lt;br/&gt; is not. In XHtml the inverse is true.</param>
		/// <param name="encoding">Code page used to encode the string. Normally this is UTF-8</param>
		/// <param name="onFont">Method that can customize the fonts used in the resulting string. It can be null if you don't want to do any modification to the fonts.</param>
		/// <returns>An html string with the section.</returns>
		public abstract string GetPageHeaderOrFooterAsHtml(string section, string imageTag, int pageNumber, int pageCount, THtmlVersion htmlVersion, Encoding encoding, IHtmlFontEvent onFont);

		/// <summary>
		/// This method returns the images associated to a given section of the header or footer.
		/// There can be only one image per section, and you refer it from the header string 
		/// (see <see cref="PageHeader"/> and <see cref="PageFooter"/>) by writing &amp;G.
		/// NOTE THAT YOU CAN ONLY USE HEADER AND FOOTER GRAPHICS ON EXCEL XP AND NEWER. Excel 2000 and 97
		/// will still open the file, but they will show no graphics.
		/// </summary>
        /// <param name="headerAndFooterKind">Type of page for which we want to retrieve the image. You will normally get this value from <see cref="THeaderAndFooter.GetHeaderAndFooterKind"/>.</param>
		/// <param name="section">Section of the header or footer for which we want to retrieve the image.</param>
		/// <param name="imageType"><b>Returns</b> the image type for the data returned. (If it is a bmp, jpg or other)</param>
		/// <param name="outStream">Stream where the image data will be copied.</param>
		public abstract void GetHeaderOrFooterImage(THeaderAndFooterKind headerAndFooterKind, THeaderAndFooterPos section, ref TXlsImgType imageType, Stream outStream);

		/// <summary>
		/// This method returns the images associated to a given section of the header or footer.
		/// There can be only one image per section, and you refer it from the header string 
		/// (see <see cref="PageHeader"/> and <see cref="PageFooter"/>) by writing &amp;G.
		/// NOTE THAT YOU CAN ONLY USE HEADER AND FOOTER GRAPHICS ON EXCEL XP AND NEWER. Excel 2000 and 97
		/// will still open the file, but they will show no graphics.
		/// </summary>
        /// <param name="headerAndFooterKind">Type of page for which we want to retrieve the image. You will normally get this value from <see cref="THeaderAndFooter.GetHeaderAndFooterKind"/>.</param>
        /// <param name="section">Section of the header or footer for which we want to retrieve the image.</param>
		/// <param name="imageType"><b>Returns</b> the image type for the data returned. (If it is a bmp, jpg or other)</param>
		/// <returns>Bytes for the image. Null if there is no image on this position.</returns>
		public byte[] GetHeaderOrFooterImage(THeaderAndFooterKind headerAndFooterKind, THeaderAndFooterPos section, ref TXlsImgType imageType)
		{
			using (MemoryStream ms = new MemoryStream())
			{
				GetHeaderOrFooterImage(headerAndFooterKind, section, ref imageType, ms);
				if (ms.Length==0) return null;
				return ms.ToArray();
			}
		}

		/// <summary>
		/// Returns the image position and size.
		/// </summary>
		/// <param name="section">Section of the header or footer for which we want to retrieve the image properties.</param>
        /// <param name="headerAndFooterKind">Type of page for which we want to retrieve the image. You will normally get this value from <see cref="THeaderAndFooter.GetHeaderAndFooterKind"/>.</param>
        /// <returns>Image properties. Null if there is no image on this section.</returns>
		public abstract THeaderOrFooterImageProperties GetHeaderOrFooterImageProperties(THeaderAndFooterKind headerAndFooterKind, THeaderAndFooterPos section);

		/// <summary>
		/// This method sets the image associated to a given section of the header or footer.
		/// There can be only one image per section, and you refer it from the header string 
		/// (see <see cref="PageHeader"/> and <see cref="PageFooter"/>) by writing &amp;G.
		/// NOTE THAT YOU CAN ONLY USE HEADER AND FOOTER GRAPHICS ON EXCEL XP AND NEWER. Excel 2000 and 97
		/// will still open the file, but they will show no graphics.
		/// ALSO, NOTE that only setting the image will not display it. You need to write &amp;G in 
		/// the corresponding <see cref="PageHeader"/> or <see cref="PageFooter"/>
		/// </summary>
        /// <param name="headerAndFooterKind">Type of page for which we want to set the image. You will normally get this value from <see cref="THeaderAndFooter.GetHeaderAndFooterKind"/>.</param>
        /// <param name="section">Section of the header or footer for which we want to set the image.</param>
		/// <param name="imageType">The image type for the data sent. (If it is a bmp, jpg or other)</param>
		/// <param name="data">Image data.</param>
		/// <param name="properties">Image size.</param>
        public abstract void SetHeaderOrFooterImage(THeaderAndFooterKind headerAndFooterKind, THeaderAndFooterPos section, byte[] data, TXlsImgType imageType, THeaderOrFooterImageProperties properties);
		
		/// <summary>
		/// This method sets the image associated to a given section of the header or footer.
		/// There can be only one image per section, and you refer it from the header string 
		/// (see <see cref="PageHeader"/> and <see cref="PageFooter"/>) by writing &amp;G.
		/// NOTE THAT YOU CAN ONLY USE HEADER AND FOOTER GRAPHICS ON EXCEL XP AND NEWER. Excel 2000 and 97
		/// will still open the file, but they will show no graphics.
		/// ALSO, NOTE that only setting the image will not display it. You need to write &amp;G in 
		/// the corresponding <see cref="PageHeader"/> or <see cref="PageFooter"/>
		/// This methods will try to automatically guess/convert the image type
		/// of the data to the better fit.
		/// </summary>
        /// <param name="headerAndFooterKind">Type of page for which we want to set the image. You will normally get this value from <see cref="THeaderAndFooter.GetHeaderAndFooterKind"/>.</param>
        /// <param name="section">Section of the header or footer for which we want to set the image.</param>
		/// <param name="data">Image data.</param>
		/// <param name="properties">Image Size.</param>
        public void SetHeaderOrFooterImage(THeaderAndFooterKind headerAndFooterKind, THeaderAndFooterPos section, byte[] data, THeaderOrFooterImageProperties properties)
		{
			if (data==null) 
			{
				SetHeaderOrFooterImage(headerAndFooterKind, section, data, TXlsImgType.Bmp, properties);
				return;
			}

			data = ImageUtils.StripOLEHeader(data);
			TXlsImgType imgType= ImageUtils.GetImageType(data);

            ImageUtils.CheckImgValid(ref data, ref imgType, false);
        
			SetHeaderOrFooterImage(headerAndFooterKind, section, data, imgType, properties);            
		}

		/// <summary>
		/// True if the gray grid lines are printed when printing the spreadsheet.
		/// </summary>
		public abstract bool PrintGridLines{get;set;}

		/// <summary>
		/// When true the row and column labels (A,B...etc for columns, 1,2... for rows) will be printed.
		/// </summary>
		public abstract bool PrintHeadings{get;set;}

		/// <summary>
		/// When true the sheet will print horizontally centered on the page.
		/// </summary>
		public abstract bool PrintHCentered{get;set;}

		/// <summary>
		/// When true the sheet will print vertically centered on the page.
		/// </summary>
		public abstract bool PrintVCentered{get;set;}

		/// <summary>
		/// Gets the Margins on the active sheet.
		/// </summary>
		public abstract TXlsMargins GetPrintMargins();

		/// <summary>
		/// Sets the Margins on the active sheet.
		/// </summary>
		//Note that this MUST be a method, not a property, to avoid someone calling "x.PrintMargins.Left=zz" which would do nothing, as SetPrintMargins will never be called.
		public abstract void SetPrintMargins(TXlsMargins value);

		/// <summary>
		/// If true, sheet will be configured to fit on <see cref="PrintNumberOfHorizontalPages"/> x <see cref="PrintNumberOfVerticalPages"/>.
		/// </summary>
		public abstract bool PrintToFit{get;set;}

		/// <summary>
		/// Number of copies to print.
		/// </summary>
		public abstract int PrintCopies{get;set;}

		/// <summary>
		/// Horizontal printer resolution on DPI.
		/// </summary>
		public abstract int PrintXResolution{get;set;}

		/// <summary>
		/// Vertical printer resolution on DPI.
		/// </summary>
		public abstract int PrintYResolution{get;set;}

		/// <summary>
		/// Misc options.
		/// </summary>
		public abstract TPrintOptions PrintOptions{get;set;}

		/// <summary>
		/// Percent to grow/shrink the sheet.
		/// </summary>
		public abstract int PrintScale{get;set;}

        /// <summary>
        /// Page number that will be assigned to the first sheet when printing. (So it will show in page headers/footers).
        /// You might set this value to null to keep the page automatic. Also, the value returned here will be null it
        /// this value is not set (Set to Automatic)
        /// </summary>
        public abstract int? PrintFirstPageNumber { get; set; }

		/// <summary>
        /// If set, the sheet will be printed on at most this number of horizontal pages. Use 0 to have unlimited horizontal pages while still limiting the vertical pages with <see cref="PrintNumberOfVerticalPages"/>. (see "Preparing for printing" in the Pdf Api Guide) <seealso cref="PrintNumberOfVerticalPages"/>
		/// </summary>
		public abstract int PrintNumberOfHorizontalPages{get;set;}

		/// <summary>
        /// If set, the sheet will be printed on at most this number of vertical pages. Use 0 to have unlimited vertical pages while still limiting the horizontal pages with <see cref="PrintNumberOfHorizontalPages"/>. (see "Preparing for printing" in the Pdf Api Guide)<seealso cref="PrintNumberOfHorizontalPages"/>
		/// </summary>
		public abstract int PrintNumberOfVerticalPages{get;set;}

		/// <summary>
		/// Pre-defined standard paper size. If you want to set up a printer specific paper size, see <see cref="SetPrinterDriverSettings"/>
		/// </summary>
		public abstract TPaperSize PrintPaperSize{get;set;}

		/// <summary>
		/// Returns the dimensions for the selected paper. See also <see cref="PrintPaperSize"/>.
		/// </summary>
		public abstract TPaperDimensions PrintPaperDimensions{get;}

		/// <summary>
		/// Returns printer driver settings. This method is not intended to be used alone,
		/// but together with <see cref="SetPrinterDriverSettings"/> to copy printer driver information from a file
		/// to another.
		/// </summary>
		/// <remarks>
		/// Excel stores printer settings in <b>two places</b>
		/// <list type="number">
		/// <item>
		/// Standard printer settings: You can set/read this with <see cref="PrintPaperSize"/>, 
		/// <see cref="PrintScale"/>, <see cref="PrintOptions"/> (Landscape inside printOptions only),
		/// <see cref="PrintXResolution"/>, <see cref="PrintYResolution"/> and <see cref="PrintCopies"/>
		/// </item>
		/// <item>
		/// Printer driver settings: You can access this with GetPrinterDriverSettings and SetPrinterDriverSettings.
		/// </item>
		/// </list>
		/// <b>NOTE THAT THOSE PLACES STORE DUPLICATED INFORMATION.</b>  For example, Excel stores the PageSize on both <i>Standard printer settings</i>
		/// and <i>Printer driver settings.</i><p></p> Always that a value is stored on both places, <i>(1) Standard printer settings</i> takes preference.
		/// <p></p>If you set PaperSize=A4 on standard settings and PaperSize=A5 on driver settings, A4 will be used.
		/// </remarks>
		/// <example>To copy the printer driver information from an empty template to the working file, use:
		/// <code>Xls1.SetPrinterDriverSettings(Xls.GetPrinterDriverSettings());</code>
		/// If you have defined a printer specific paper size, and you want to use it, you should call
		/// <code>Xls1.PrintPaperSize=Xls2.PrintPaperSize;</code>
		/// after copying the driver settings.
		/// </example>
		/// <returns>The printer driver settings.</returns>
		public abstract TPrinterDriverSettings GetPrinterDriverSettings();

		/// <summary>
		/// Sets printer driver information. This method is not intended to be used alone,
		/// but together with <see cref="GetPrinterDriverSettings"/> to copy printer driver information from a file
		/// to another.
		/// </summary>
		/// <remarks>
		/// Excel stores printer settings in <b>two places</b>
		/// <list type="number">
		/// <item>
		/// Standard printer settings: You can set/read this with <see cref="PrintPaperSize"/>, 
		/// <see cref="PrintScale"/>, <see cref="PrintOptions"/> (Landscape inside printOptions only),
		/// <see cref="PrintXResolution"/>, <see cref="PrintYResolution"/> and <see cref="PrintCopies"/>
		/// </item>
		/// <item>
		/// Printer driver settings: You can access this with GetPrinterDriverSettings and SetPrinterDriverSettings.
		/// </item>
		/// </list>
		/// <b>NOTE THAT THOSE PLACES STORE DUPLICATED INFORMATION.</b>  For example, Excel stores the PageSize on both <i>Standard printer settings</i>
		/// and <i>Printer driver settings.</i><p></p> Always that a value is stored on both places, <i>(1) Standard printer settings</i> takes preference.
		/// <p></p>If you set PaperSize=A4 on standard settings and PaperSize=A5 on driver settings, A4 will be used.
		/// </remarks>
		/// <example>To copy the printer driver information from an empty template to the working file, use:
		/// <code>Xls1.SetPrinterDriverSettings(Xls.GetPrinterDriverSettings());</code>
		/// If you have defined a printer specific paper size, and you want to use it, you should call
		/// <code>Xls1.PrintPaperSize=Xls2.PrintPaperSize;</code>
		/// after copying the driver settings.
		/// </example>
		/// <param name="settings">Printer driver information obtained with <see cref="GetPrinterDriverSettings"/>. Use null to remove the printer settings.</param>
		public abstract void SetPrinterDriverSettings(TPrinterDriverSettings settings);
		#endregion

		#region Images
		/// <summary>
		/// The number of images on the active sheet.
		/// </summary>
		public abstract int ImageCount{get;}

        /// <summary>
        /// Sets the image data and / or image properties of an existing image.
        /// </summary>
        /// <param name="imageIndex">Index of the image on the sheet array (1-based)</param>
        /// <param name="data">Image data.</param>
        /// <param name="imageType">Image type of the new data.</param>
        public void SetImage(int imageIndex, byte[] data, TXlsImgType imageType)
        {
            SetImage(imageIndex, data, imageType, false, null);
        }

		/// <summary>
		/// Sets the image data and / or image properties of an existing image.
		/// </summary>
		/// <param name="imageIndex">Index of the image on the sheet array (1-based)</param>
		/// <param name="data">Image data.</param>
		/// <param name="imageType">Image type of the new data.</param>
        /// <param name="usesObjectIndex">If false (the default) then imageIndex is an index to the list of images.
        /// When true imageIndex is an index to the list of all objects in the sheet. When you have the object id, you can avoid calling
        /// <see cref="ObjectIndexToImageIndex"/> which is a slow method, by setting this parameter to true.</param>
        /// <param name="objectPath">Path to the object, when the object is grouped with others. This parameter only
        /// has meaning if usesObjectIndex is true. <br></br>
        /// If it is "absolute"(it starts with "\\"), then the path includes the objectIndex, and the objectIndex is
        /// not used. An object path of "\\1\\2\\3" is exactly the same as using objectIndex = 1 and objectPath = "2\\3"</param>
        public abstract void SetImage(int imageIndex, byte[] data, TXlsImgType imageType, bool usesObjectIndex, string objectPath);

        /// <summary>
		/// Sets the image data for an existing image. It will try to automatically guess/convert the image type
		/// of the data to the better fit.
		/// </summary>
		/// <param name="imageIndex">Index of the image on the sheet array (1-based)</param>
		/// <param name="data">Image data.</param>
        public void SetImage(int imageIndex, byte[] data)
        {
            SetImage(imageIndex, data, false, null);
        }

		/// <summary>
		/// Sets the image data for an existing image. It will try to automatically guess/convert the image type
		/// of the data to the better fit.
		/// </summary>
		/// <param name="imageIndex">Index of the image on the sheet array (1-based)</param>
		/// <param name="data">Image data.</param>
        /// <param name="usesObjectIndex">If false (the default) then imageIndex is an index to the list of images.
        /// When true imageIndex is an index to the list of all objects in the sheet. When you have the object id, you can avoid calling
        /// <see cref="ObjectIndexToImageIndex"/> which is a slow method, by setting this parameter to true.</param>
        /// <param name="objectPath">Path to the object, when the object is grouped with others. This parameter only
        /// has meaning if usesObjectIndex is true.<br></br>
        /// If it is "absolute"(it starts with "\\"), then the path includes the objectIndex, and the objectIndex is
        /// not used. An object path of "\\1\\2\\3" is exactly the same as using objectIndex = 1 and objectPath = "2\\3"</param>
        public void SetImage(int imageIndex, byte[] data, bool usesObjectIndex, string objectPath)
		{
			if (data==null) 
			{
				SetImage(imageIndex, data, TXlsImgType.Bmp, usesObjectIndex, objectPath);
				return;
			}

			data = ImageUtils.StripOLEHeader(data);
			TXlsImgType imgType= ImageUtils.GetImageType(data);

            ImageUtils.CheckImgValid(ref data, ref imgType, false);
        
			SetImage(imageIndex, data, imgType, usesObjectIndex, objectPath);            
		}

        		/// <summary>
		/// WARNING!  Not CF compliant.
		/// </summary>
		/// <param name="imageIndex">Image Index. 1-Based.</param>
		/// <param name="Img">Image to replace.</param>
        /// <remarks>                    
		/// Saving a WMF or EMF Image is not currently supported by the .NET framework.
		/// If you pass a MetaFile to this method, it will be saved as png. 
		/// For inserting a REAL wmf into excel use <see cref="AddImage(Stream, TXlsImgType, TImageProperties)"/>
		/// </remarks>
		/// <platform><frameworks><compact>false</compact></frameworks></platform>
        public void SetImage(int imageIndex, Image Img)
        {
            SetImage(imageIndex, Img, false, null);
        }

		/// <summary>
		/// WARNING!  Not CF compliant.
		/// </summary>
		/// <param name="imageIndex">Image Index. 1-Based.</param>
		/// <param name="Img">Image to replace.</param>
        /// <param name="usesObjectIndex">If false (the default) then imageIndex is an index to the list of images.
        /// When true imageIndex is an index to the list of all objects in the sheet. When you have the object id, you can avoid calling
        /// <see cref="ObjectIndexToImageIndex"/> which is a slow method, by setting this parameter to true.</param>
        /// <param name="objectPath">Path to the object, when the object is grouped with others. This parameter only
        /// has meaning if usesObjectIndex is true.<br></br>
        /// If it is "absolute"(it starts with "\\"), then the path includes the objectIndex, and the objectIndex is
        /// not used. An object path of "\\1\\2\\3" is exactly the same as using objectIndex = 1 and objectPath = "2\\3"</param>
        /// <remarks>                    
		/// Saving a WMF or EMF Image is not currently supported by the .NET framework.
		/// If you pass a MetaFile to this method, it will be saved as png. 
		/// For inserting a REAL wmf into excel use <see cref="AddImage(Stream, TXlsImgType, TImageProperties)"/>
		/// </remarks>
		/// <platform><frameworks><compact>false</compact></frameworks></platform>
		public void SetImage(int imageIndex, Image Img, bool usesObjectIndex, string objectPath)
		{
#if (MONOTOUCH)
			using (MonoTouch.Foundation.NSData st = Img.AsPNG())
			{
				byte[] MemData = new byte[st.Length];
				using (Stream str = st.AsStream())
				{
					str.Write(MemData,0, MemData.Length);
				    SetImage(imageIndex, MemData, TXlsImgType.Png, usesObjectIndex, objectPath);
				}
			}			
#else
#if(!COMPACTFRAMEWORK && !SILVERLIGHT)
			TXlsImgType imageType= TXlsImgType.Unknown;
			using (MemoryStream MemStream=new MemoryStream())
			{
				if (Img.RawFormat.Equals(ImageFormat.Jpeg))
				{
					Img.Save(MemStream, ImageFormat.Jpeg);
					imageType= TXlsImgType.Jpeg;
				}
					/*                else
										if (Img.RawFormat.Equals(ImageFormat.Wmf))
									{
										Img.Save(MemStream, ImageFormat.Wmf);
										imageType= TXlsImgType.Wmf;
									}
									else
										if (Img.RawFormat.Equals(ImageFormat.Emf))
									{
										Img.Save(MemStream, ImageFormat.Emf);
										imageType= TXlsImgType.Emf;
									}
					*/                else
				{
					Img.Save(MemStream, ImageFormat.Png);
					imageType= TXlsImgType.Png;       
				}
				SetImage(imageIndex, MemStream.ToArray(), imageType, usesObjectIndex, objectPath);
			}
#else
            throw new MissingMemberException("SetImage(int imageIndex, Image Img)");
#endif
#endif
		}

		/// <summary>
		/// Sets the image properties of an existing image.
		/// </summary>
		/// <param name="imageIndex">Index of the image on the sheet array (1-based)</param>
		/// <param name="imageProperties">Image size, placement, etc. </param>
		public abstract void SetImageProperties(int imageIndex, TImageProperties imageProperties);

		/// <summary>
		/// Returns the image name at position imageIndex.
		/// </summary>
		/// <remarks>Normally image names are automatically assigned by Excel, and are on the form
		/// "picture1", "picture2", etc. But you can name an image on Excel by selecting and then typing its name
		/// on the name combo box. (the combobox at the upper left corner on Excel). After that, you can use its name
		/// to identify it here.</remarks>
		/// <param name="imageIndex">Index of the image (1 based)</param>
		/// <returns>Image name.</returns>
		public abstract string GetImageName(int imageIndex);

		/// <summary>
		/// Returns an image bytes and type.
		/// </summary>
		/// <param name="imageIndex">Index of the image. (1 based)</param>
		/// <param name="imageType"><b>Returns</b> the image type for the data returned. (If it is a bmp, jpg or other)</param>
		/// <returns>Image data. Use the returned imageType to find out the format of the stored image.</returns>
		public byte[] GetImage(int imageIndex, ref TXlsImgType imageType)
		{
			using (MemoryStream ms = new MemoryStream())
			{
				GetImage(imageIndex, ref imageType, ms);
				return ms.ToArray();
			}
		}

		/// <summary>
		/// Returns an image and its type.
		/// </summary>
		/// <param name="imageIndex">Index of the image. (1 based)</param>
		/// <param name="imageType"><b>Returns</b> the image type for the data returned. (If it is a bmp, jpg or other)</param>
		/// <param name="outStream">Stream where the image data will be copied.</param>
		public void GetImage(int imageIndex, ref TXlsImgType imageType, Stream outStream)
		{
			GetImage(imageIndex, String.Empty, ref imageType, outStream);
		}

        /// <summary>
        /// Returns an image and its type.
        /// </summary>
        /// <param name="imageIndex">Index of the image. (1 based)</param>
        /// <param name="objectPath">Object path to the image when it is a grouped image. For toplevel images you can use String.Empty. In other case, you need to use the value returned by <see cref="GetObjectProperties"/>
        /// <br></br>
        /// If it is "absolute"(it starts with "\\"), then the path includes the objectIndex, and the objectIndex is
        /// not used. An object path of "\\1\\2\\3" is exactly the same as using objectIndex = 1 and objectPath = "2\\3"</param>
        /// <param name="imageType"><b>Returns</b> the image type for the data returned. (If it is a bmp, jpg or other)</param>
        /// <param name="outStream">Stream where the image data will be copied.</param>
        public void GetImage(int imageIndex, string objectPath, ref TXlsImgType imageType, Stream outStream)
        {
            GetImage(imageIndex, objectPath, ref imageType, outStream, false);
        }

		/// <summary>
        /// Returns an image and its type.
        /// </summary>
		/// <param name="imageIndex">Index of the image. (1 based)</param>
		/// <param name="objectPath">Object path to the image when it is a grouped image. For toplevel images you can use String.Empty. In other case, you need to use the value returned by <see cref="GetObjectProperties"/>
        /// <br></br>
        /// If it is "absolute"(it starts with "\\"), then the path includes the objectIndex, and the objectIndex is
        /// not used. An object path of "\\1\\2\\3" is exactly the same as using objectIndex = 1 and objectPath = "2\\3"</param>
		/// <param name="imageType"><b>Returns</b> the image type for the data returned. (If it is a bmp, jpg or other)</param>
		/// <param name="outStream">Stream where the image data will be copied.</param>
        /// <param name="usesObjectIndex">If false (the default) then imageIndex is an index to the list of images.
        /// When true imageIndex is an index to the list of all objects in the sheet. When you have the object id, you can avoid calling
        /// <see cref="ObjectIndexToImageIndex"/> which is a slow method, by setting this parameter to true.</param>
        public abstract void GetImage(int imageIndex, string objectPath, ref TXlsImgType imageType, Stream outStream, bool usesObjectIndex);

		/// <summary>
		/// Returns image position and size.
		/// </summary>
		/// <param name="imageIndex">Index of the image (1 based)</param>
		/// <returns>Image position and size.</returns>
		public abstract TImageProperties GetImageProperties(int imageIndex);

		/// <summary>
		/// Adds an image to the active sheet.
		/// </summary>
		/// <param name="data">byte array with the image data.</param>
		/// <param name="imageType">Type of image you are inserting (bmp, jpg, etc).</param>
		/// <param name="imageProperties">Placement and other properties of the image.</param>
        public abstract void AddImage(byte[] data, TXlsImgType imageType, TImageProperties imageProperties);

		/// <summary>
		/// Adds an image to the active sheet.
		/// </summary>
		/// <param name="aStream">Stream containing the image data.</param>
		/// <param name="imageType">Type of image you are inserting (bmp, jpg, etc).</param>
		/// <param name="imageProperties">Placement and other properties of the image.</param>
		public void AddImage(Stream aStream, TXlsImgType imageType, TImageProperties imageProperties)
		{
			MemoryStream ms= aStream as MemoryStream;
			if (ms!=null) //Small optimization to avoid another copy
			{
				AddImage(ms.ToArray(), imageType, imageProperties);
			}
			else
			{
				byte[] buffer= new byte[aStream.Length];
				Sh.Read(aStream, buffer, 0, buffer.Length);
				AddImage(buffer, imageType, imageProperties);
			}
		}

		/// <summary>
		/// WARNING!  Not CF compliant. If you don't have the image already created, prefer 
		/// using <see cref="AddImage(Stream, TImageProperties)"/>, as it is faster.
		/// </summary>
		/// <param name="img">Image to insert.</param>
		/// <param name="imageProperties">Image size/position</param>
		/// <remarks>                    
		/// Saving a WMF or EMF Image is not currently supported by the .NET framework.
		/// If you pass a MetaFile to this method, it will be saved as png. 
		/// For inserting a REAL wmf into excel use <see cref="AddImage(Stream, TImageProperties)"/>
		/// </remarks>
		/// <platform><frameworks><compact>false</compact></frameworks></platform>
		public void AddImage(Image img, TImageProperties imageProperties)
		{
#if (MONOTOUCH)
			using (MonoTouch.Foundation.NSData st = img.AsPNG())
			{
				byte[] MemData = new byte[st.Length];
				using (Stream str = st.AsStream())
				{
					str.Write(MemData,0, MemData.Length);
				    AddImage(MemData, TXlsImgType.Png, imageProperties);
				}
			}			
#else

#if(!COMPACTFRAMEWORK && !SILVERLIGHT)
			TXlsImgType imageType= TXlsImgType.Unknown;
			using (MemoryStream MemStream=new MemoryStream())
			{
				if (img.RawFormat.Equals(ImageFormat.Jpeg))
				{
					img.Save(MemStream, ImageFormat.Jpeg);
					imageType= TXlsImgType.Jpeg;
				}
					// Saving a WMF and EMF is not currently supported by the framework!
					// They will be saved as png. For inserting a REAL wmf into excel use AddImage(Stream, TImageProperties)
					/*                else
										if (img.RawFormat.Equals(ImageFormat.Wmf))
									{
										img.Save(MemStream, ImageFormat.Wmf);
										imageType= TXlsImgType.Wmf;
									}
									else
										if (img.RawFormat.Equals(ImageFormat.Emf))
									{
										img.Save(MemStream, ImageFormat.Emf);
										imageType= TXlsImgType.Emf;
									}
					*/
				else
				{
					img.Save(MemStream, ImageFormat.Png);
					imageType= TXlsImgType.Png;       
				}
				AddImage(MemStream.ToArray(), imageType, imageProperties);
			}
#else
            throw new MissingMemberException("AddImage(Image img, TImageProperties imageProperties)");
#endif
#endif
		}

		/// <summary>
		/// WARNING!  Not CF compliant. If you don't have the image already created, prefer 
		/// using <see cref="AddImage(Stream, TImageProperties)"/>, as it is faster.
		/// </summary>
		/// <param name="img">Image to insert.</param>
		/// <param name="row">Row Index (1 based)</param>
		/// <param name="col">Column Index (1 based)</param>
		/// <remarks>                    
		/// Saving a WMF or EMF Image is not currently supported by the .NET framework.
		/// If you pass a MetaFile to this method, it will be saved as png. 
		/// For inserting a REAL wmf into excel use <see cref="AddImage(Stream, TImageProperties)"/>
		/// </remarks>
		/// <platform><frameworks><compact>false</compact></frameworks></platform>
		public void AddImage(int row, int col, Image img)
		{
#if (COMPACTFRAMEWORK || (!FRAMEWORK30 && !MONOTOUCH))
            int imgHeight = img.Height;
            int imgWidth = img.Width;
#else
            int imgHeight = img.Height();
            int imgWidth = img.Width();
#endif
            TImageProperties imageProperties=new TImageProperties(new TClientAnchor(TFlxAnchorType.MoveAndResize, row, 0, col, 0, imgHeight, imgWidth, this),String.Empty, String.Empty,new TCropArea(), 
				FlxConsts.NoTransparentColor, FlxConsts.DefaultBrightness, FlxConsts.DefaultContrast, FlxConsts.DefaultGamma, 
                true, true, false, false, false, true, true, null, null, 
                true, true, false, false, true);
			AddImage(img, imageProperties);
		}

		/// <summary>
		/// Adds an image to the active sheet. It will try to automatically guess/convert the image type
		/// of the data to the better fit.
		/// </summary>
		/// <param name="data">Image data.</param>
		/// <param name="imageProperties">Image Properties.</param>
		public void AddImage(byte[] data, TImageProperties imageProperties)
		{
#if(!COMPACTFRAMEWORK  || FRAMEWORK20)
			if (data==null) 
			{
				AddImage(data, TXlsImgType.Bmp, imageProperties);
				return;
			}

			data = ImageUtils.StripOLEHeader(data);
			TXlsImgType imgType = ImageUtils.GetImageType(data);
            ImageUtils.CheckImgValid(ref data, ref imgType, false);
        
			AddImage(data, imgType, imageProperties);
#else
            throw new MissingMemberException("AddImage(byte[] data, TImageProperties imageProperties)");
#endif

		}

		/// <summary>
		/// Adds an image to the active sheet.
		/// </summary>
		/// <param name="aStream">Stream containing the image data.</param>
		/// <param name="imageProperties">Placement and other properties of the image.</param>
		public void AddImage(Stream aStream, TImageProperties imageProperties)
		{
			MemoryStream ms= aStream as MemoryStream;
			if (ms!=null) //Small optimization to avoid another copy
			{
				AddImage(ms.ToArray(), imageProperties); //Do not use GetBuffer, as it pads the stream with 0's
			}
			else
			{
				byte[] buffer= new byte[aStream.Length-aStream.Position];
				Sh.Read(aStream, buffer, 0, buffer.Length);
				AddImage(buffer, imageProperties);
			}
		}

    
		/// <summary>
		/// Deletes the image at position imageIndex.
		/// </summary>
		/// <param name="imageIndex">Index of the image to delete. (1 based)</param>
		public abstract void DeleteImage(int imageIndex);

		/// <summary>
		/// Clears the image at position imageIndex, leaving an empty white box.
		/// </summary>
		/// <param name="imageIndex">Index of the image to clear. (1 based)</param>
		public abstract void ClearImage(int imageIndex);


		#region Objects
		/// <summary>
		/// Count of all graphical objects on the sheet. They can be charts, images, shapes, etc.
		/// </summary>
		public abstract int ObjectCount{get;}

		/// <summary>
		/// Returns the general index on the object list for an image. You can use then this index on SendToBack, for example.
		/// </summary>
		/// <param name="imageIndex">Image index on the image array.</param>
		/// <returns>Image index on the total objects array.</returns>
		public abstract int ImageIndexToObjectIndex(int imageIndex);

		/// <summary>
        /// Returns the index on the image collection of an object. <b>Note that this method is slow</b> when there are many images, so use it sparingly.
		/// </summary>
		/// <param name="objectIndex">General index of the image on the Object collection.</param>
		/// <returns>-1 if the object is not an image, else the index on the image collection.</returns>
		public abstract int ObjectIndexToImageIndex(int objectIndex);

		/// <summary>
        /// Returns the name of the object at objectIndex position.
		/// </summary>
		/// <param name="objectIndex">Object index. (1-based)</param>
		/// <returns></returns>
		public abstract string GetObjectName(int objectIndex);

        /// <summary>
        /// Returns the shape id of the object at objectIndex position. Shape Ids are internal identifiers for the shape, that you can use to uniquely identify a shape.
        /// Note that the shape id can change when you load the file, once it is loaded, it will remain the same for the shape lifetime.
        /// </summary>
        /// <param name="objectIndex">Object index. (1-based)</param>
        /// <returns></returns>
        public abstract long GetObjectShapeId(int objectIndex);

		/// <summary>
		/// Returns the object index for an existing name. Whenever possible you should prefer to use <see cref="FindObjectPath"/>
        /// instead of this method, since it is faster and finds also objects that are not in the root branch.
		/// </summary>
		/// <param name="objectName">Object name to search for. This is case insensitive.</param>
		/// <returns>-1 if the object is not found, or the object index otherwise.</returns>
        public int FindObject(string objectName)
        {
            for (int i = 1; i <= ObjectCount; i++)
                if (String.Equals(GetObjectName(i), objectName, StringComparison.CurrentCultureIgnoreCase)) return i;
            return -1;
        }

        /// <summary>
        /// Finds an object by its name, and returns the ObjectPath you need to use this object.
        /// Note that if there is more than an object with the same name in the sheet, this method 
        /// will return null
        /// </summary>
        /// <param name="objectName">Name of the object we are looking for.</param>
        /// <returns></returns>
        public abstract string FindObjectPath(string objectName);

        /// <summary>
        /// Finds an object given its internal shape id, and returns the object index you need to access the same object in FlexCel.
        /// </summary>
        /// <param name="ShapeId">Shape id of the object.</param>
        /// <returns></returns>
        public abstract int FindObjectByShapeId(long ShapeId);

        /// <summary>
        /// Returns true if the object name exists and it is unique in the sheet.
        /// You can usee <see cref="FindObjectPath"/> to find the object path you need for this object name.
        /// </summary>
        /// <param name="objectName"></param>
        /// <returns></returns>
        public bool IsValidObjectPathObjName(string objectName)
        {
            return FindObjectPath(objectName) != null;

        }

		/// <summary>
		/// Returns the placement of the object.
		/// </summary>
		/// <param name="objectIndex">Index of the object (1-based)</param>
		/// <returns>Coordinates of the object.</returns>
		public abstract TClientAnchor GetObjectAnchor(int objectIndex);

        /// <inheritdoc cref="SetObjectAnchor(int, string, TClientAnchor)" />
        public void SetObjectAnchor(int objectIndex, TClientAnchor objectAnchor)
        {
            SetObjectAnchor(objectIndex, null, objectAnchor);
        }

        /// <summary>
        /// Sets the object placement.
        /// </summary>
        /// <param name="objectIndex">Index of the object (1-based)</param>
        /// <param name="objectPath">Object path t the shape if this is a grouped shape.</param>
        /// <param name="objectAnchor">Coordinates of the object.</param>
        public abstract void SetObjectAnchor(int objectIndex, string objectPath, TClientAnchor objectAnchor);

        /// <summary>
        /// Returns information on an object and all of its children. If the shapeId doesn't exist, this method returns null.
        /// </summary>
        /// <param name="objectIndex">Index of the object (1-based)</param>
        /// <param name="getShapeOptions">When true, shape options will be retrieved. As this can be a slow operation,
        /// only specify true when you really need those options.</param>
        /// <returns></returns>
        public abstract TShapeProperties GetObjectProperties(int objectIndex, bool getShapeOptions);
        
        /// <summary>
        /// Returns information on an object and all of its children. 
        /// </summary>
        /// <param name="shapeId">Index of the object (1-based)</param>
        /// <param name="getShapeOptions">When true, shape options will be retrieved. As this can be a slow operation,
        /// only specify true when you really need those options.</param>
        /// <returns></returns>
        public abstract TShapeProperties GetObjectPropertiesByShapeId(long shapeId, bool getShapeOptions);

		/// <summary>
		/// Sets the text for an autoshape. If the object does not accept text, this method will do nothing.
		/// </summary>
		/// <param name="objectIndex">Index of the object (1-based)</param>
		/// <param name="objectPath">Index to the child object you want to change the text.
		/// If it is a simple object, you can use String.Empty here, if not you need to get the ObjectPath from <see cref="GetObjectProperties"/>
        /// <br></br>
        /// If it is "absolute"(it starts with "\\"), then the path includes the objectIndex, and the objectIndex is
        /// not used. An object path of "\\1\\2\\3" is exactly the same as using objectIndex = 1 and objectPath = "2\\3"</param>
		/// <param name="text">Text you want to use. Use null to delete text from an AutoShape.</param>
		public abstract void SetObjectText(int objectIndex, string objectPath, TRichString text);

        /// <summary>
        /// Sets the name for an autoshape.
        /// </summary>
        /// <param name="objectIndex">Index of the object (1-based)</param>
        /// <param name="objectPath">Index to the child object you want to change the text.
        /// If it is a simple object, you can use String.Empty here, if not you need to get the ObjectPath from <see cref="GetObjectProperties"/>
        /// <br></br>
        /// If it is "absolute"(it starts with "\\"), then the path includes the objectIndex, and the objectIndex is
        /// not used. An object path of "\\1\\2\\3" is exactly the same as using objectIndex = 1 and objectPath = "2\\3"</param>
        /// <param name="name">Name for the autoshape. Use null to remove the name from an AutoShape.</param>
        public abstract void SetObjectName(int objectIndex, string objectPath, string name);

		/// <summary>
		/// Sets a STRING property for an autoshape. Verify the property expects a STRING.
		/// This is an advanced method and should be used with care. For normal use, you should use one of the standard methods. (like SetObjectText)
		/// </summary>
		/// <param name="objectIndex">Index of the object (1-based)</param>
		/// <param name="objectPath">Index to the child object you want to change the property.
        /// If it is a simple object, you can use String.Empty here, if not you need to get the ObjectPath from <see cref="GetObjectProperties"/><br></br>
        /// If it is "absolute"(it starts with "\\"), then the path includes the objectIndex, and the objectIndex is
        /// not used. An object path of "\\1\\2\\3" is exactly the same as using objectIndex = 1 and objectPath = "2\\3"</param>
		/// <param name="shapeProperty">Property you want to change.</param>
		/// <param name="text">Text you want to use.</param>
		public abstract void SetObjectProperty(int objectIndex, string objectPath, TShapeOption shapeProperty, string text);

		/// <summary>
		/// Sets a LONG property for an autoshape. Verify the property expects a LONG.
		/// This is an advanced method and should be used with care. For normal use, you should use one of the standard methods.
		/// </summary>
		/// <param name="objectIndex">Index of the object (1-based)</param>
		/// <param name="objectPath">Index to the child object you want to change the property.
        /// If it is a simple object, you can use String.Empty here, if not you need to get the ObjectPath from <see cref="GetObjectProperties"/><br></br>
        /// If it is "absolute"(it starts with "\\"), then the path includes the objectIndex, and the objectIndex is
        /// not used. An object path of "\\1\\2\\3" is exactly the same as using objectIndex = 1 and objectPath = "2\\3"</param>
		/// <param name="shapeProperty">Property you want to change.</param>
		/// <param name="value">Value you want to use.</param>
		public abstract void SetObjectProperty(int objectIndex, string objectPath, TShapeOption shapeProperty, long value);

        /// <summary>
        /// Sets a DOUBLE (Encoded as 16.16) property for an autoshape. Verify the property expects a DOUBLE.
        /// This is an advanced method and should be used with care. For normal use, you should use one of the standard methods.
        /// </summary>
        /// <param name="objectIndex">Index of the object (1-based)</param>
        /// <param name="objectPath">Index to the child object you want to change the property.
        /// If it is a simple object, you can use String.Empty here, if not you need to get the ObjectPath from <see cref="GetObjectProperties"/><br></br>
        /// If it is "absolute"(it starts with "\\"), then the path includes the objectIndex, and the objectIndex is
        /// not used. An object path of "\\1\\2\\3" is exactly the same as using objectIndex = 1 and objectPath = "2\\3"</param>
        /// <param name="shapeProperty">Property you want to change.</param>
        /// <param name="value">Value you want to use.</param>
        public abstract void SetObjectProperty(int objectIndex, string objectPath, TShapeOption shapeProperty, double value);

		/// <summary>
		/// Sets a BOOLEAN property for an autoshape. Verify the property expects a BOOLEAN.
		/// This is an advanced method and should be used with care. For normal use, you should use one of the standard methods.
		/// Note that boolean properties are all stored in the same byte of the last property in the set.
		/// </summary>
		/// <param name="objectIndex">Index of the object (1-based)</param>
		/// <param name="objectPath">Index to the child object you want to change the property.
        /// If it is a simple object, you can use String.Empty here, if not you need to get the ObjectPath from <see cref="GetObjectProperties"/><br></br>
        /// If it is "absolute"(it starts with "\\"), then the path includes the objectIndex, and the objectIndex is
        /// not used. An object path of "\\1\\2\\3" is exactly the same as using objectIndex = 1 and objectPath = "2\\3"</param>
		/// <param name="shapeProperty">Property you want to change. MAKE SURE it is the LAST property in the set.</param>
		/// <param name="positionInSet">Boolean properties are grouped so all properties on one set are in only
		/// one value. So, the last bool property on the set is the first bit, and so on. ONLY THE LAST PROPERTY
		/// ON THE SET IS PRESENT.</param>
		/// <param name="value">Value you want to use.</param>
		public abstract void SetObjectProperty(int objectIndex, string objectPath, TShapeOption shapeProperty, int positionInSet, bool value);

        /// <summary>
        /// Sets an Hyperlink property for an autoshape. Verify the property expects a Hyperlink, currently only <see cref="TShapeOption.pihlShape"/> expects hyperlinks.
        /// This is an advanced method and should be used with care. For normal use, you should use one of the standard methods.
        /// </summary>
        /// <param name="objectIndex">Index of the object (1-based)</param>
        /// <param name="objectPath">Index to the child object you want to change the property.
        /// If it is a simple object, you can use String.Empty here, if not you need to get the ObjectPath from <see cref="GetObjectProperties"/><br></br>
        /// If it is "absolute"(it starts with "\\"), then the path includes the objectIndex, and the objectIndex is
        /// not used. An object path of "\\1\\2\\3" is exactly the same as using objectIndex = 1 and objectPath = "2\\3"</param>
        /// <param name="shapeProperty">Property you want to change. Fro hyperlinks it should be <see cref="TShapeOption.pihlShape"/></param>
        /// <param name="value">Value you want to use.</param>
        public abstract void SetObjectProperty(int objectIndex, string objectPath, TShapeOption shapeProperty, THyperLink value);

		/*		/// <summary>
				/// Adds an autoshape to the sheet.
				/// </summary>
				/// <param name="shapeType">Shape that we want to add.</param>
				/// <param name="clientAnchor">Coordinates where the shape will be inserted.</param>
				/// <param name="text">Text for the Autoshape. If you leave this null, no text will be added.</param>
				public abstract void AddAutoShape(TShapeType shapeType, TClientAnchor clientAnchor, TRichString text);
		*/

        /// <inheritdoc cref="DeleteObject(int, string)" />
        public void DeleteObject(int objectIndex)
        {
            DeleteObject(objectIndex, null);
        }

        /// <summary>
        /// Deletes the graphic object at objectIndex. Use it with care, there are some graphics objects you
        /// <b>don't</b> want to remove (like comment boxes when you don't delete the associated comment.)
        /// </summary>
        /// <param name="objectIndex">Index of the object (1 based).</param>
        /// <param name="objectPath">Index to the child object you want to change the property.
        /// If it is a simple object, you can use String.Empty here, if not you need to get the ObjectPath from <see cref="GetObjectProperties"/><br></br>
        /// If it is "absolute"(it starts with "\\"), then the path includes the objectIndex, and the objectIndex is
        /// not used. An object path of "\\1\\2\\3" is exactly the same as using objectIndex = 1 and objectPath = "2\\3"</param>
        public abstract void DeleteObject(int objectIndex, string objectPath);

        /// <summary>
        /// Returns a list with all the objects that are completely inside a range of cells.
        /// </summary>
        /// <param name="range">Range of cells where we want to find the objects.</param>
        /// <param name="objectsInRange">In this list we will add all the objects found. note that the objects will be added to the list, so if you want just the objects
        /// in this range, make sure you clear the list before calling this method.</param>
        public abstract void GetObjectsInRange(TXlsCellRange range, TExcelObjectList objectsInRange);

        #region Linked cells
        /// <summary>
        /// Returns the cell that is linked to the object. If the object isn't an object that can be linked or it isn't linked, this method
        /// will return null.
        /// Note that when you change the value in the cell linked to this object, 
        /// the value of the object will change.
        /// <br></br>The sheet returned in the TCellAddress might be null, in which case the reference is to a cell in the same sheet, or it might
        /// contain another sheet name.
        /// </summary>
        /// <param name="objectIndex">Index of the object (1 based)</param>
        /// <param name="objectPath">Index to the child object you want to change the property.
        /// If it is a simple object, you can use String.Empty here, if not you need to get the ObjectPath from <see cref="GetObjectProperties"/><br></br>
        /// If it is "absolute"(it starts with "\\"), then the path includes the objectIndex, and the objectIndex is
        /// not used. An object path of "\\1\\2\\3" is exactly the same as using objectIndex = 1 and objectPath = "2\\3"</param>
        /// <returns>The cell address this object is linked to, or null if it isn't linked. </returns>
        public abstract TCellAddress GetObjectLinkedCell(int objectIndex, string objectPath);

        /// <summary>
        /// Links the object to a cell, if the object can be linked. If the object is a radio button then all the other radio buttons in the group will be linked
        /// to the same cell, 
        /// so when the cell changes the radio buttons too, and vice-versa. To unlink the cell, make linkedCell null.
        /// </summary>
        /// <remarks>Note that when the object is a radio button, this affects all radio buttons in the same group, not just the first.</remarks>
        /// <param name="objectIndex">Index of the object (1 based)</param>
        /// <param name="objectPath">Index to the child object you want to change the property.
        /// If it is a simple object, you can use String.Empty here, if not you need to get the ObjectPath from <see cref="GetObjectProperties"/><br></br>
        /// If it is "absolute"(it starts with "\\"), then the path includes the objectIndex, and the objectIndex is
        /// not used. An object path of "\\1\\2\\3" is exactly the same as using objectIndex = 1 and objectPath = "2\\3"</param>
        /// <param name="linkedCell">Cell that will be linked to the radio button. To unlink the radio button, make this parameter null.</param>
        public abstract void SetObjectLinkedCell(int objectIndex, string objectPath, TCellAddress linkedCell);

        /// <summary>
        /// Returns the input range for the object. 
        /// If the object isn't a combobox or listbox, or it doesn't have an input range, this method
        /// will return null.
        /// Note that when you change the value in the cell linked to this object, 
        /// the value of the object will change.
        /// <br></br>The sheet in the TCellAddresses returned might be null, in which case the reference is to a cell in the same sheet, or it might
        /// contain another sheet name.
        /// </summary>
        /// <param name="objectIndex">Index of the object (1 based)</param>
        /// <param name="objectPath">Index to the child object you want to change the property.
        /// If it is a simple object, you can use String.Empty here, if not you need to get the ObjectPath from <see cref="GetObjectProperties"/><br></br>
        /// If it is "absolute"(it starts with "\\"), then the path includes the objectIndex, and the objectIndex is
        /// not used. An object path of "\\1\\2\\3" is exactly the same as using objectIndex = 1 and objectPath = "2\\3"</param>
        /// <returns>The input range, 
        /// or null if the object doesn't have an input range. </returns>
        public abstract TCellAddressRange GetObjectInputRange(int objectIndex, string objectPath);

        /// <summary>
        /// Sets the input range for a ListBox or a ComboBox. When applied to other objects, this method does nothing.
        /// </summary>
        /// <param name="objectIndex">Index of the object (1 based)</param>
        /// <param name="objectPath">Index to the child object you want to change the property.
        /// If it is a simple object, you can use String.Empty here, if not you need to get the ObjectPath from <see cref="GetObjectProperties"/><br></br>
        /// If it is "absolute"(it starts with "\\"), then the path includes the objectIndex, and the objectIndex is
        /// not used. An object path of "\\1\\2\\3" is exactly the same as using objectIndex = 1 and objectPath = "2\\3"</param>
        /// <param name="inputRange">Input range for the object.</param>
        public abstract void SetObjectInputRange(int objectIndex, string objectPath, TCellAddressRange inputRange);

        /// <summary>
        /// Returns the macro associated with an object, or null if there is no macro associated.
        /// </summary>
        /// <param name="objectIndex">Index of the object (1 based)</param>
        /// <param name="objectPath">Index to the child object you want to change the property.
        /// If it is a simple object, you can use String.Empty here, if not you need to get the ObjectPath from <see cref="GetObjectProperties"/><br></br>
        /// If it is "absolute"(it starts with "\\"), then the path includes the objectIndex, and the objectIndex is
        /// not used. An object path of "\\1\\2\\3" is exactly the same as using objectIndex = 1 and objectPath = "2\\3"</param>
        /// <returns>The macro assocaited with the object, or null if there isn't any macro associated.</returns>
        public abstract string GetObjectMacro(int objectIndex, string objectPath);

        /// <summary>
        /// Associates an object with a macro. While this will normally be used in buttons, you can associate macros to
        /// almost any object.
        /// </summary>
        /// <param name="objectIndex">Index of the object (1 based)</param>
        /// <param name="objectPath">Index to the child object you want to change the property.
        /// If it is a simple object, you can use String.Empty here, if not you need to get the ObjectPath from <see cref="GetObjectProperties"/><br></br>
        /// If it is "absolute"(it starts with "\\"), then the path includes the objectIndex, and the objectIndex is
        /// not used. An object path of "\\1\\2\\3" is exactly the same as using objectIndex = 1 and objectPath = "2\\3"</param>
        /// <param name="macro">Macro that will be associated with the object. Look at apimate to know the exat name you have to enter here.</param>
        public abstract void SetObjectMacro(int objectIndex, string objectPath, string macro);

        #endregion

        #region Checkboxes
        /// <summary>
        /// Gets the value of a checkbox in the active sheet. Note that this only works for <b>checkboxes added through the Forms toolbar.</b> 
        /// It won't return the values of ActiveX checkboxes.
        /// </summary>
        /// <remarks>Note that if the checkbox is linked to a cell and you changed the cell value, this method will return the cell value.</remarks>
        /// <param name="objectIndex">Index of the object (1 based)</param>
        /// <param name="objectPath">Index to the child object you want to change the property.
        /// If it is a simple object, you can use String.Empty here, if not you need to get the ObjectPath from <see cref="GetObjectProperties"/><br></br>
        /// If it is "absolute"(it starts with "\\"), then the path includes the objectIndex, and the objectIndex is
        /// not used. An object path of "\\1\\2\\3" is exactly the same as using objectIndex = 1 and objectPath = "2\\3"</param>
        /// <returns>If the checkbox is checked or not.</returns>
        /// <example>
        /// To get all the checkboxes in the active sheet, you can use this code:
        /// <code>
        /// XlsFile xls = new XlsFile("checkboxes.xls");
        /// for (int i = 1; i &lt;= xls.ObjectCount; i++)
        /// {
        ///     TShapeProperties shp = xls.GetObjectProperties(i, false);
        ///     if (shp.ObjectType == TObjectType.CheckBox)
        ///     {
        ///            MessageBox.Show(shp.Text + " :" + xls.GetCheckboxState(i, null).ToString());
        ///     }    
        /// }
        /// </code>
        /// </example>
        public abstract TCheckboxState GetCheckboxState(int objectIndex, string objectPath);

        /// <summary>
        /// Sets the value of a checkbox in the active sheet. Note that this only works for <b>checkboxes added through the Forms toolbar.</b> 
        /// It won't return the values of ActiveX checkboxes.
        /// </summary>
        /// <remarks>
        /// If the checkbox is linked to a cell, this method will change both the checkbox and the cell.
        /// </remarks>
        /// <param name="objectIndex">Index of the object (1 based)</param>
        /// <param name="objectPath">Index to the child object you want to change the property.
        /// If it is a simple object, you can use String.Empty here, if not you need to get the ObjectPath from <see cref="GetObjectProperties"/><br></br>
        /// If it is "absolute"(it starts with "\\"), then the path includes the objectIndex, and the objectIndex is
        /// not used. An object path of "\\1\\2\\3" is exactly the same as using objectIndex = 1 and objectPath = "2\\3"</param>
        /// <param name="value">Value to set.</param>
        /// <example>
        /// To set a checkbox given its text, use the code:
        /// <code>
        /// XlsFile xls = new XlsFile("checkboxes.xls");
        /// for (int i = 1; i &lt;= xls.ObjectCount; i++)
        /// {
        ///     TShapeProperties shp = xls.GetObjectProperties(i, false);
        ///     if (shp.ObjectType == TObjectType.CheckBox &amp;&amp; shp.Text == "MyText")
        ///     {
        ///            xls.SetCheckboxState(i, null, TCheckboxState.Checked);
        ///     }    
        /// }
        /// </code>
        /// To change a checkbox given its name instead of its text, replace <b>shp.Text == "MyText"</b> by  <b>shp.Name == "MyName"</b> in the code above.
        /// </example>
        public abstract void SetCheckboxState(int objectIndex, string objectPath, TCheckboxState value);

        /// <summary>
        /// Returns the cell that is linked to the checkbox. If the object isn't a checkbox or it isn't linked, this method
        /// will return null. Note that when you change the value in the cell linked to this checkbox, the value of the checkbox will change.
        /// <br></br>The sheet returned in the TCellAddress might be null, in which case the reference is to a cell in the same sheet, or it might
        /// contain another sheet name.
        /// </summary>
        /// <param name="objectIndex">Index of the object (1 based)</param>
        /// <param name="objectPath">Index to the child object you want to change the property.
        /// If it is a simple object, you can use String.Empty here, if not you need to get the ObjectPath from <see cref="GetObjectProperties"/><br></br>
        /// If it is "absolute"(it starts with "\\"), then the path includes the objectIndex, and the objectIndex is
        /// not used. An object path of "\\1\\2\\3" is exactly the same as using objectIndex = 1 and objectPath = "2\\3"</param>
        /// <returns>The cell address this checkbox is linked to, or null if it isn't linked. </returns>
        [Obsolete("Use GetObjectLinkedCell instead.")]
        public abstract TCellAddress GetCheckboxLinkedCell(int objectIndex, string objectPath);

        /// <summary>
        /// Links the checkbox to a cell, so when the cell changes the checkbox changes too, and vice-versa. To unlink the cell, make linkedCell null.
        /// </summary>
        /// <param name="objectIndex">Index of the object (1 based)</param>
        /// <param name="objectPath">Index to the child object you want to change the property.
        /// If it is a simple object, you can use String.Empty here, if not you need to get the ObjectPath from <see cref="GetObjectProperties"/><br></br>
        /// If it is "absolute"(it starts with "\\"), then the path includes the objectIndex, and the objectIndex is
        /// not used. An object path of "\\1\\2\\3" is exactly the same as using objectIndex = 1 and objectPath = "2\\3"</param>
        /// <param name="linkedCell">Cell that will be linked to the checkbox. To unlink the checkbox, make this parameter null.</param>
        [Obsolete("Use SetObjectLinkedCell instead.")]
        public abstract void SetCheckboxLinkedCell(int objectIndex, string objectPath, TCellAddress linkedCell);

        /// <inheritdoc cref="AddCheckbox(TClientAnchor, TRichString, TCheckboxState, TCellAddress, string)" />
        public int AddCheckbox(TClientAnchor anchor, TRichString text, TCheckboxState value, TCellAddress linkedCell)
        {
            return AddCheckbox(anchor, text, value, linkedCell, null);
        }

        /// <summary>
        /// Adds a checkbox to the active sheet. 
        /// </summary>
        /// <remarks>
        /// Excel supports 2 types of checkboxes: ActiveX and internal (In Excel you would add them from the "ActiveX" and "Forms" toolbars respectively)
        /// <br></br>The checkboxes added by this method are of type internal. ActiveX checkboxes are not supported.
        /// </remarks>
        /// <param name="anchor">Position for the checkbox.</param>
        /// <param name="text">Text for the checkbox.</param>
        /// <param name="value">Value of the checkbox.</param>
        /// <param name="linkedCell">Cell that will be linked to the checkbox. If you don't want to link the checkbox to a cell, make this parameter null.</param>
        /// <param name="name">Name that will be given to the checkbox.</param>
        /// <returns>Object Index of the inserted checkbox (1 based).</returns>
        public abstract int AddCheckbox(TClientAnchor anchor, TRichString text, TCheckboxState value, TCellAddress linkedCell, string name);
        #endregion

        #region Radio Buttons
        /// <summary>
        /// Gets if a radio button in the active sheet is selected or not. Note that this only works for <b>radio buttons added through the Forms toolbar.</b> 
        /// It won't return the values of ActiveX radio buttons.
        /// </summary>
        /// <remarks>Note that if the radio button is linked to a cell and you changed the cell value, 
        /// this method will return the cell value.</remarks>
        /// <param name="objectIndex">Index of the object (1 based)</param>
        /// <param name="objectPath">Index to the child object you want to change the property.
        /// If it is a simple object, you can use String.Empty here, if not you need to get the ObjectPath from <see cref="GetObjectProperties"/><br></br>
        /// If it is "absolute"(it starts with "\\"), then the path includes the objectIndex, and the objectIndex is
        /// not used. An object path of "\\1\\2\\3" is exactly the same as using objectIndex = 1 and objectPath = "2\\3"</param>
        /// <returns>If the radio button is selected or not.</returns>
        /// <example>
        /// To get all the radio buttons in the active sheet, you can use this code:
        /// <code>
        /// XlsFile xls = new XlsFile("radiobuttons.xls");
        /// for (int i = 1; i &lt;= xls.ObjectCount; i++)
        /// {
        ///     TShapeProperties shp = xls.GetObjectProperties(i, false);
        ///     if (shp.ObjectType == TObjectType.OptionButton)
        ///     {
        ///            MessageBox.Show(shp.Text + " :" + xls.GetRadioButtonState(i, null).ToString());
        ///     }    
        /// }
        /// </code>
        /// </example>
        public abstract bool GetRadioButtonState(int objectIndex, string objectPath);

        /// <summary>
        /// Sets the value of a radio button in the active sheet. Note that this only works for <b>radio buttons added through the Forms toolbar.</b> 
        /// It won't return the values of ActiveX radio buttons
        /// </summary>
        /// <remarks>
        /// If the radio button is linked to a cell, this method will change both the radio button and the cell.
        /// </remarks>
        /// <param name="objectIndex">Index of the object (1 based)</param>
        /// <param name="objectPath">Index to the child object you want to change the property.
        /// If it is a simple object, you can use String.Empty here, if not you need to get the ObjectPath from <see cref="GetObjectProperties"/><br></br>
        /// If it is "absolute"(it starts with "\\"), then the path includes the objectIndex, and the objectIndex is
        /// not used. An object path of "\\1\\2\\3" is exactly the same as using objectIndex = 1 and objectPath = "2\\3"</param>
        /// <param name="selected">If true, the option button will be set, and all other buttons in the range will be deselected.
        /// When false the radio button will be deselected.</param>
        /// <example>
        /// To set a radio button given its text, use the code:
        /// <code>
        /// XlsFile xls = new XlsFile("radiobuttons.xls");
        /// for (int i = 1; i &lt;= xls.ObjectCount; i++)
        /// {
        ///     TShapeProperties shp = xls.GetObjectProperties(i, false);
        ///     if (shp.ObjectType == TObjectType.OptionButton &amp;&amp; shp.Text == "MyText")
        ///     {
        ///            xls.SetRadioButtonState(i, null, true);
        ///     }    
        /// }
        /// </code>
        /// To change a radio button given its name instead of its text, replace <b>shp.Text == "MyText"</b> by  <b>shp.Name == "MyName"</b> in the code above.
        /// </example>
        public abstract void SetRadioButtonState(int objectIndex, string objectPath, bool selected);


        /// <inheritdoc cref="AddRadioButton(TClientAnchor, TRichString, string)" />
        public int AddRadioButton(TClientAnchor anchor, TRichString text)
        {
            return AddRadioButton(anchor, text, null);
        }

        /// <summary>
        /// Adds a radio button to the active sheet. Call <see cref="AddGroupBox(TClientAnchor, TRichString)"/> to insert a group box for grouping the radio buttons. 
        /// </summary>
        /// <remarks>
        /// Excel supports 2 types of radio buttons: ActiveX and internal (In Excel you would add them from the "ActiveX" and "Forms" toolbars respectively)
        /// <br></br>The radio buttons added by this method are of type internal. ActiveX radio buttons are not supported.
        /// </remarks>
        /// <param name="anchor">Position for the radio button.</param>
        /// <param name="text">Text for the radio button.</param>
        /// <param name="name">Name of the inserted radio button.</param>
        /// <returns>Object Index of the inserted radio button (1 based).</returns>
        public abstract int AddRadioButton(TClientAnchor anchor, TRichString text, string name);

        /// <inheritdoc cref="AddGroupBox(TClientAnchor, TRichString, string)" />
        public int AddGroupBox(TClientAnchor anchor, TRichString text)
        {
            return AddGroupBox(anchor, text, null);
        }

        /// <summary>
        /// Adds a Group box to the active sheet. Call <see cref="AddRadioButton(TClientAnchor, TRichString)"/> to insert radio buttons inside the group box. 
        /// </summary>
        /// <remarks>
        /// Excel determines if a radio button is inside a group box by looking at the coordinates of both objects.
        /// To include a radio button inside the group box, just make sure it is enterely inside the group box.
        /// Object hierarchy doesn't matter here, only the positions of the group box and the radio button.
        /// </remarks>
        /// <param name="anchor">Position for the group box.</param>
        /// <param name="text">Text for the group box.</param>
        /// <param name="name">Name for the inserted Group box</param>
        /// <returns>Object Index of the inserted group box (1 based).</returns>
        public abstract int AddGroupBox(TClientAnchor anchor, TRichString text, string name);

        #endregion

        #region Other objects
        /// <summary>
        /// Gets the selected item in an object from the "Forms" palette. It can be a combobox or a listbox. 
        /// 0 means no selection, 1 the first item in the list. Note that this only works for <b>objects added through the Forms toolbar.</b> 
        /// It won't return the values of ActiveX objects.
        /// </summary>
        /// <remarks>Note that if the object is linked to a cell and you changed the cell value, 
        /// this method will return the cell value.</remarks>
        /// <param name="objectIndex">Index of the object (1 based)</param>
        /// <param name="objectPath">Index to the child object you want to change the property.
        /// If it is a simple object, you can use String.Empty here, if not you need to get the ObjectPath from <see cref="GetObjectProperties"/><br></br>
        /// If it is "absolute"(it starts with "\\"), then the path includes the objectIndex, and the objectIndex is
        /// not used. An object path of "\\1\\2\\3" is exactly the same as using objectIndex = 1 and objectPath = "2\\3"</param>
        /// <returns>The position of the selected item in the object. 0 means no selection, 1 the first item is selected.</returns>
        /// <example>
        /// To get all the comboboxes in the active sheet, you can use this code:
        /// <code>
        /// XlsFile xls = new XlsFile("comboboxes.xls");
        /// for (int i = 1; i &lt;= xls.ObjectCount; i++)
        /// {
        ///     TShapeProperties shp = xls.GetObjectProperties(i, false);
        ///     if (shp.ObjectType == TObjectType.ComboBox)
        ///     {
        ///            MessageBox.Show(shp.Text + " :" + xls.GetObjectSelection(i, null).ToString());
        ///     }    
        /// }
        /// </code>
        /// </example>
        public abstract int GetObjectSelection(int objectIndex, string objectPath);

        /// <summary>
        /// Sets the selected item of an object from the "Forms" palette. It can be a combobox, a listbox, a spinbox or a scrollbar. Note that this only works for <b>objects added through the Forms toolbar.</b> 
        /// It won't return the values of ActiveX objects.
        /// </summary>
        /// <remarks>
        /// If the object is linked to a cell, this method will change both the object and the cell.
        /// </remarks>
        /// <param name="objectIndex">Index of the object (1 based)</param>
        /// <param name="objectPath">Index to the child object you want to change the property.
        /// If it is a simple object, you can use String.Empty here, if not you need to get the ObjectPath from <see cref="GetObjectProperties"/><br></br>
        /// If it is "absolute"(it starts with "\\"), then the path includes the objectIndex, and the objectIndex is
        /// not used. An object path of "\\1\\2\\3" is exactly the same as using objectIndex = 1 and objectPath = "2\\3"</param>
        /// <param name="selectedItem">Position of the selected item in the object. 0 means no selection, 1 means that the
        /// first item is selected.</param>
        /// <example>
        /// To set a combobox selection to be the 5th item, use the code:
        /// <code>
        /// XlsFile xls = new XlsFile("comboboxes.xls");
        /// for (int i = 1; i &lt;= xls.ObjectCount; i++)
        /// {
        ///     TShapeProperties shp = xls.GetObjectProperties(i, false);
        ///     if (shp.ObjectType == TObjectType.OptionButton &amp;&amp; shp.Text == "MyText")
        ///     {
        ///            xls.SetObjectSelection(i, null, 5);
        ///     }    
        /// }
        /// </code>
        /// To change a combo given its name instead of its text, replace <b>shp.Text == "MyText"</b> by  <b>shp.Name == "MyName"</b> in the code above.
        /// </example>
        public abstract void SetObjectSelection(int objectIndex, string objectPath, int selectedItem);

        /// <summary>
        /// Returns maximum, minimum and increment in any control that has a spin or dropdown, like a listbox, combobox, spinner or scrollbar.
        /// </summary>
        /// <param name="objectIndex">Index of the object (1 based)</param>
        /// <param name="objectPath">Index to the child object you want to change the property.
        /// If it is a simple object, you can use String.Empty here, if not you need to get the ObjectPath from <see cref="GetObjectProperties"/><br></br>
        /// If it is "absolute"(it starts with "\\"), then the path includes the objectIndex, and the objectIndex is
        /// not used. An object path of "\\1\\2\\3" is exactly the same as using objectIndex = 1 and objectPath = "2\\3"</param>
        /// <returns>The spin properties.</returns>
        public abstract TSpinProperties GetObjectSpinProperties(int objectIndex, string objectPath);

        /// <summary>
        /// Sets the spin properties of an object. You should apply this only to scrollbars and spinners.
        /// </summary>
        /// <param name="objectIndex">Index of the object (1 based)</param>
        /// <param name="objectPath">Index to the child object you want to change the property.
        /// If it is a simple object, you can use String.Empty here, if not you need to get the ObjectPath from <see cref="GetObjectProperties"/><br></br>
        /// If it is "absolute"(it starts with "\\"), then the path includes the objectIndex, and the objectIndex is
        /// not used. An object path of "\\1\\2\\3" is exactly the same as using objectIndex = 1 and objectPath = "2\\3"</param>
        /// <param name="spinProps">Properties of the spinner.</param>
        public abstract void SetObjectSpinProperties(int objectIndex, string objectPath, TSpinProperties spinProps);

        /// <summary>
        /// Returns the current selected value of a scrollbar.
        /// </summary>
        /// <param name="objectIndex">Index of the object (1 based)</param>
        /// <param name="objectPath">Index to the child object you want to change the property.
        /// If it is a simple object, you can use String.Empty here, if not you need to get the ObjectPath from <see cref="GetObjectProperties"/><br></br>
        /// If it is "absolute"(it starts with "\\"), then the path includes the objectIndex, and the objectIndex is
        /// not used. An object path of "\\1\\2\\3" is exactly the same as using objectIndex = 1 and objectPath = "2\\3"</param>
        /// <returns>The spin position.</returns>
        public abstract int GetObjectSpinValue(int objectIndex, string objectPath);

        /// <summary>
        /// Sets the positon in a scrollbar object. If the object is linked to a cell, the cell will be updated.
        /// </summary>
        /// <param name="objectIndex">Index of the object (1 based)</param>
        /// <param name="objectPath">Index to the child object you want to change the property.
        /// If it is a simple object, you can use String.Empty here, if not you need to get the ObjectPath from <see cref="GetObjectProperties"/><br></br>
        /// If it is "absolute"(it starts with "\\"), then the path includes the objectIndex, and the objectIndex is
        /// not used. An object path of "\\1\\2\\3" is exactly the same as using objectIndex = 1 and objectPath = "2\\3"</param>
        /// <param name="value">Position for the scrollbar.</param>
        public abstract void SetObjectSpinValue(int objectIndex, string objectPath, int value);


        /// <summary>
        /// Adds a ComboBox to the active sheet. 
        /// </summary>
        /// <remarks>
        /// Excel supports 2 types of comboboxes: ActiveX and internal (In Excel you would add them from the "ActiveX" and "Forms" toolbars respectively)
        /// <br></br>The comboboxes added by this method are of type internal. ActiveX comboboxes are not supported.
        /// </remarks>
        /// <param name="anchor">Position for the combobox.</param>
        /// <param name="name">Name of the inserted combobox.</param>
        /// <param name="linkedCell">Cell that will be linked to the combobox. Set this to null to not link any cell.</param>
        /// <param name="inputRange">Range of cells with the values that the combobox will display.</param>
        /// <param name="selectedItem">Item that will be selected, starting at 1. 0 means no selected item.</param>
        /// <returns>Object Index of the inserted combobox (1 based).</returns>
        public abstract int AddComboBox(TClientAnchor anchor, string name, TCellAddress linkedCell, TCellAddressRange inputRange, int selectedItem);

        /// <summary>
        /// Adds a ListBox to the active sheet. 
        /// </summary>
        /// <remarks>
        /// Excel supports 2 types of listboxes: ActiveX and internal (In Excel you would add them from the "ActiveX" and "Forms" toolbars respectively)
        /// <br></br>The listboxes added by this method are of type internal. ActiveX listboxes are not supported.
        /// </remarks>
        /// <param name="anchor">Position for the listbox.</param>
        /// <param name="name">Name of the inserted listbox.</param>
        /// <param name="linkedCell">Cell that will be linked to the listbox. Set this to null to not link any cell.</param>
        /// <param name="inputRange">Range of cells with the values that the listbox will display.</param>
        /// <param name="selectionType">How items are selected in the lisbox.</param>
        /// <param name="selectedItem">Item that will be selected, starting at 1. 0 means no selected item.</param>
        /// <returns>Object Index of the inserted listbox (1 based).</returns>
        public abstract int AddListBox(TClientAnchor anchor, string name, TCellAddress linkedCell, 
            TCellAddressRange inputRange, TListBoxSelectionType selectionType, int selectedItem);

        /// <summary>
        /// Adds a button to the sheet, with the associated macro.
        /// </summary>
        /// <param name="anchor">Position for the button.</param>
        /// <param name="text">Text that will be shown in the button.</param>
        /// <param name="name">Name for the inserted button.</param>
        /// <param name="macro">Macro associated with the button.</param>
        /// <returns></returns>
        public abstract int AddButton(TClientAnchor anchor, TRichString text, string name, string macro);


        /// <summary>
        /// Adds a Label to the active sheet. 
        /// </summary>
        /// <remarks>
        /// Excel supports 2 types of labels: ActiveX and internal (In Excel you would add them from the "ActiveX" and "Forms" toolbars respectively)
        /// <br></br>The labels added by this method are of type internal. ActiveX labels are not supported.
        /// </remarks>
        /// <param name="anchor">Position for the label.</param>
        /// <param name="text">Text for the label.</param>
        /// <param name="name">Name of the inserted label.</param>
        /// <returns>Object Index of the inserted listbox (1 based).</returns>
        public abstract int AddLabel(TClientAnchor anchor, TRichString text, string name);

        /// <summary>
        /// Adds a Spinner to the active sheet. 
        /// </summary>
        /// <remarks>
        /// Excel supports 2 types of spinners: ActiveX and internal (In Excel you would add them from the "ActiveX" and "Forms" toolbars respectively)
        /// <br></br>The spinners added by this method are of type internal. ActiveX spinners are not supported.
        /// </remarks>
        /// <param name="anchor">Position for the spinner.</param>
        /// <param name="name">Name of the inserted spinner.</param>
        /// <param name="linkedCell">Cell that will be linked to the spinner. Set this to null to not link any cell.</param>
        /// <param name="spinProps">Properties for the spinner.</param>
        /// <returns>Object Index of the inserted spinner (1 based).</returns>
        public abstract int AddSpinner(TClientAnchor anchor, string name, TCellAddress linkedCell,
            TSpinProperties spinProps);

        /// <summary>
        /// Adds a ScrollBar to the active sheet. 
        /// </summary>
        /// <remarks>
        /// Excel supports 2 types of ScrollBars: ActiveX and internal (In Excel you would add them from the "ActiveX" and "Forms" toolbars respectively)
        /// <br></br>The ScrollBars added by this method are of type internal. ActiveX ScrollBars are not supported.
        /// </remarks>
        /// <param name="anchor">Position for the ScrollBar.</param>
        /// <param name="name">Name of the inserted ScrollBar.</param>
        /// <param name="linkedCell">Cell that will be linked to the ScrollBar. Set this to null to not link any cell.</param>
        /// <param name="spinProps">Properties for the ScrollBar.</param>
        /// <returns>Object Index of the inserted ScrollBar (1 based).</returns>
        public abstract int AddScrollBar(TClientAnchor anchor, string name, TCellAddress linkedCell,
            TSpinProperties spinProps);

        #endregion

#if (!COMPACTFRAMEWORK && !MONOTOUCH && !SILVERLIGHT)
        /// <summary>
		/// !WARNING! Not CF compliant. 
		/// This method renders any object (chart, image, autoshape, etc) into an image, and returns it.
		/// <br/>Background of the image will be transparent.
		/// </summary>
		/// <param name="objectIndex">Index of the object (1 based).</param>
		/// <returns>Might return null if the image is not visible.</returns>
		public Image RenderObject(int objectIndex)
		{
			RectangleF ImageDimensions;
			PointF Origin;
			Size SizePixels;
			return RenderObject(objectIndex, 96, GetObjectProperties(objectIndex, true), 
				SmoothingMode.AntiAlias, InterpolationMode.HighQualityBicubic, true, true,
				out Origin, out ImageDimensions, out SizePixels);
		}

		/// <summary>
		/// !WARNING! Not CF compliant. 
		/// This method renders any object (chart, image, autoshape, etc) into an image, and returns it.
		/// </summary>
		/// <param name="objectIndex">Index of the object (1 based).</param>
		/// <param name="dpi">Resolution of the image to create in dots per inch. If creating the image for the screen, use 96 dpi.</param>
		/// <param name="aInterpolationMode">Interpolation mode used to render the object. For more information, see <see cref="System.Drawing.Drawing2D.InterpolationMode"/></param>
		/// <param name="aSmoothingMode">Smoothing mode used to render the object. For more information, see <see cref="System.Drawing.Drawing2D.SmoothingMode"/></param>
		/// <param name="antiAliased">If true text will be antialiased when rendering for example a chart.</param>
		/// <param name="returnImage">If false, this method will return null. Use it if you need to know the image dimensions, but do not care about the real image since it is faster and uses less resources.</param>
		/// <param name="shapeProperties">Properties of the shape you are about to render. You can get them by calling <see cref="GetObjectProperties"/>.</param>
		/// <param name="imageDimensions">Returns the image dimension of the rendered object in points. Note that this can be different from the image size reported by 
		/// <see cref="GetImageProperties(int)"/> because shadows or rotation of the image. You can get the image size in pixels just by looking at the image returned.</param>
		/// <param name="origin">Top-left coordinates of the image in points. While this is normally the same as the image coordinates you get in the properties, 
		/// if there is a shadow to the right or to the top it might change. Use it to properly position the image where you want it.</param>
		/// <param name="imageSizePixels">Size of the returned image in pixels. You only need to use this if returnImage is false, since the returned bitmap will be null. Otherwise, you can just read the bitmap size.</param>
		/// <returns>Might return null if the image is not visible.</returns>
		public Image RenderObject(int objectIndex, real dpi, TShapeProperties shapeProperties,
			SmoothingMode aSmoothingMode, InterpolationMode aInterpolationMode, bool antiAliased, bool returnImage,
			out PointF origin, out RectangleF imageDimensions, out Size imageSizePixels)
		{
			return RenderObject(objectIndex, dpi, shapeProperties, aSmoothingMode, aInterpolationMode, antiAliased, returnImage, ColorUtil.Empty,
				out origin, out imageDimensions, out imageSizePixels);
		}

		/// <summary>
		/// !WARNING! Not CF compliant. 
		/// This method renders any object (chart, image, autoshape, etc) into an image, and returns it.
		/// </summary>
		/// <param name="objectIndex">Index of the object (1 based).</param>
		/// <param name="dpi">Resolution of the image to create in dots per inch. If creating the image for the screen, use 96 dpi.</param>
		/// <param name="aInterpolationMode">Interpolation mode used to render the object. For more information, see <see cref="System.Drawing.Drawing2D.InterpolationMode"/></param>
		/// <param name="aSmoothingMode">Smoothing mode used to render the object. For more information, see <see cref="System.Drawing.Drawing2D.SmoothingMode"/></param>
		/// <param name="antiAliased">If true text will be antialiased when rendering for example a chart.</param>
		/// <param name="returnImage">If false, this method will return null. Use it if you need to know the image dimensions, but do not care about the real image since it is faster and uses less resources.</param>
		/// <param name="shapeProperties">Properties of the shape you are about to render. You can get them by calling <see cref="GetObjectProperties"/>.</param>
		/// <param name="BackgroundColor">Color for the background of the image. For a transparent background, use ColorUtil.Empty.</param>
		/// <param name="imageDimensions">Returns the image dimension of the rendered object in points. Note that this can be different from the image size reported by 
		/// <see cref="GetImageProperties(int)"/> because shadows or rotation of the image. You can get the image size in pixels just by looking at the image returned.</param>
		/// <param name="origin">Top-left coordinates of the image in points. While this is normally the same as the image coordinates you get in the properties, 
		/// if there is a shadow to the right or to the top it might change. Use it to properly position the image where you want it.</param>
		/// <param name="imageSizePixels">Size of the returned image in pixels. You only need to use this if returnImage is false, since the returned bitmap will be null. Otherwise, you can just read the bitmap size.</param>
		/// <returns>Might return null if the image is not visible.</returns>
		public abstract Image RenderObject(int objectIndex, real dpi, TShapeProperties shapeProperties,
			SmoothingMode aSmoothingMode, InterpolationMode aInterpolationMode, bool antiAliased, bool returnImage, Color BackgroundColor,
			out PointF origin, out RectangleF imageDimensions, out Size imageSizePixels);


		/// <summary>
		/// !WARNING! Not CF compliant. 
		/// This method renders a range of cells into an image, and returns it.<br/> 
		/// No objects will be drawn on the cells, you can use <see cref="RenderObject(System.Int32)"/> to draw those.<br/>
		/// <b>Important note:</b> This method will only render the text inside the cell, not the borders or anything else. If you want to 
		/// export an xls file to images, you probably want to use <see cref="FlexCel.Render.FlexCelImgExport"/>
		/// </summary>
        /// <param name="row1">Index of the first row to render. (1 based).</param>
        /// <param name="col1">Index of the first column to render (1 based).</param>
        /// <param name="row2">Index of the last row to render. (1 based).</param>
        /// <param name="col2">Index of the last column to render (1 based).</param>
        /// <param name="drawBackground">If true, the image will have a solid background with the color of the cells. If false, the image will have
        /// a transparent background.</param>
        /// <returns>An image with the rendered cells.</returns>
        public Image RenderCells(int row1, int col1, int row2, int col2, bool drawBackground)
		{
			return RenderCells(row1, col1, row2, col2, drawBackground, 96, SmoothingMode.AntiAlias, InterpolationMode.HighQualityBicubic, true);
		}

		/// <summary>
		/// !WARNING! Not CF compliant. 
		/// This method renders a range of cells into an image, and returns it.<br/> 
		/// No objects will be drawn on the cells, you can use <see cref="RenderObject(System.Int32)"/> to draw those.<br/>
		/// <b>Important note:</b> This method will only render the text inside the cell, not the borders or anything else. If you want to 
		/// export an xls file to images, you probably want to use <see cref="FlexCel.Render.FlexCelImgExport"/>
		/// </summary>
		/// <param name="row1">Index of the first row to render. (1 based).</param>
        /// <param name="col1">Index of the first column to render (1 based).</param>
        /// <param name="row2">Index of the last row to render. (1 based).</param>
        /// <param name="col2">Index of the last column to render (1 based).</param>
		/// <param name="drawBackground">If true, the image will have a solid background with the color of the cells. If false, the image will have
		/// a transparent background.</param>
		/// <param name="dpi">Resolution of the image to create in dots per inch. If creating the image for the screen, use 96 dpi.</param>
		/// <param name="aInterpolationMode">Interpolation mode used to render the object. For more information, see <see cref="System.Drawing.Drawing2D.InterpolationMode"/></param>
		/// <param name="aSmoothingMode">Smoothing mode used to render the object. For more information, see <see cref="System.Drawing.Drawing2D.SmoothingMode"/></param>
		/// <param name="antiAliased">If true text will be antialiased when rendering.</param>
		/// <returns>An image with the rendered cells.</returns>
		public abstract Image RenderCells(int row1, int col1, int row2, int col2, bool drawBackground, real dpi,
			SmoothingMode aSmoothingMode, InterpolationMode aInterpolationMode, bool antiAliased);
        
        /// <summary>
        /// Returns the height and width that would be used by a range of cells (in Points, or 1/72 inches). 
        /// </summary>
        /// <param name="row1">First row (1 based). If you use a value less or equal than 0 here, this method will return the full sheet dimensions.</param>
        /// <param name="col1">First column (1 based). If you use a value less or equal than 0 here, this method will return the full sheet dimensions.</param>
        /// <param name="row2">Last row (1 based). If you use a value less or equal than 0 here, this method will return the full sheet dimensions.</param>
        /// <param name="col2">Last colum (1 based). If you use a value less or equal than 0 here, this method will return the full sheet dimensions.</param>
        /// <param name="includeMargins">If true, the dimensions reported will include all margins in the sheet.</param>
        /// <returns></returns>
        public abstract RectangleF CellRangeDimensions(int row1, int col1, int row2, int col2, bool includeMargins);
#endif

        #endregion

        #region Object position
        ///<summary>
		/// Sends the graphical object to the bottom layer on the display (z-order) position. It will show below and will be covered by all other objects on the sheet. 
		/// </summary>
		/// <remarks>This will change the order of the array, 
		/// so after calling SendToBack(i), position i will have a new object.
		/// <seealso cref="SendBack"/><seealso cref="SendForward"/><seealso cref="BringToFront"/>
		/// </remarks>
		/// <param name="objectIndex">Index of the object to move. (1 based)</param>
		public abstract void SendToBack(int objectIndex);

		/// <summary>
		/// Sends the graphical object to the top layer on the display (z-order) position. It will show above and will cover all other objects on the sheet. 
		/// </summary>
		/// <remarks>This will change the order of the array, 
		/// so after calling BringToFront(i), position i will have a new object.
		/// <seealso cref="SendToBack"/><seealso cref="SendForward"/><seealso cref="SendBack"/>
		/// </remarks>
		/// <param name="objectIndex">Index of the object to move. (1 based)</param>
		public abstract void BringToFront(int objectIndex);

		/// <summary>
		/// Sends the graphical object one layer up on the display (z-order) position. It will show above and will cover the image at objectIndex+1. 
		/// </summary>
		/// <remarks>This will change the order of the array, 
		/// so after calling SendForward(i), position i will have a new object.
		/// 
		/// To move an object 2 steps down the correct code is:
		/// <code>
		///   SendForward(i);
		///   SendForward(i+1);
		/// </code>
		/// and not
		/// <code>
		///   SendForward(i);
		///   SendForward(i);
		/// </code>
		/// 
		/// The second code would actually leave the array unmodified.
		/// <seealso cref="SendToBack"/><seealso cref="SendBack"/><seealso cref="BringToFront"/>
		/// </remarks>
		/// <param name="objectIndex">Index of the object to move. (1 based)</param>
		public abstract void SendForward(int objectIndex);

		/// <summary>
		/// Sends the graphical object one layer down. It will show below and will be covered by image at objectIndex-1. 
		/// </summary>
		/// <remarks>This will change the order of the array, 
		/// so after calling SendToBack(i), position i will have a new object.
		/// 
		/// To move an object 2 steps down the correct code is:
		/// <code>
		///   SendBack(i);
		///   SendBack(i-1);
		/// </code>
		/// and not
		/// <code>
		///   SendBack(i);
		///   SendBack(i);
		/// </code>
		/// 
		/// The second code would actually leave the array unmodified.
		/// <seealso cref="SendToBack"/><seealso cref="SendForward"/><seealso cref="BringToFront"/>
		/// </remarks>
		/// <param name="objectIndex">Index of the object to move. (1 based)</param>
		public abstract void SendBack(int objectIndex);
		#endregion

		#endregion

		#region Comments
		/// <summary>
		/// Maximum row index including comments.
		/// </summary>
		/// <remarks>
		/// There are 2 groups of Comment related methods. One is used when you want to find out on a loop all existing comments on 
		/// a range, and the other is used to get/set comments at a specified position (row and col).
		/// 
		/// To loop among all existing comments, you can use CommentRowCount together with <see cref="CommentCountRow"/>.
		/// Do not confuse them, even when they have similar names. CommentRowCount returns the maximum row with a comment,
		/// and <see cref="CommentCountRow"/> returns the number of comments in a given row.
		/// </remarks>
		/// <example>
		/// You can loop on all the comments on a sheet with the following code:
		/// <code>
		/// for (int r=1; r&lt;= xls.CommentRowCount(); r++)
		///   for (int i=1; i&lt;= xls.CommentCountRow(r); i++)
		///   {
		///      //Do something(GetCommentRow(r,i);
		///   }
		/// </code>
		/// </example>
		/// <returns>The maximum row with a comment.</returns>
		/// <seealso cref="CommentCountRow"/> <seealso cref="GetCommentRow"/> <seealso cref="GetCommentRowCol"/> <seealso cref="GetComment"/> <seealso cref="SetCommentRow(int, int, string, TImageProperties)"/> <seealso cref="SetComment(int, int, TRichString)"/><seealso cref="GetCommentPropertiesRow"/> <seealso cref="SetCommentPropertiesRow"/><seealso cref="SetCommentProperties"/>
		public abstract int CommentRowCount();

		/// <summary>
		/// Number of comments on a given row.
		/// </summary>
		/// <param name="row">Row index (1 based)</param>
		/// <remarks>
		/// There are 2 groups of Comment related methods. One is used when you want to find out on a loop all existing comments on 
		/// a range, and the other is used to get/set comments at a specified position (row and col).
		/// 
		/// To loop among all existing comments, you can use <see cref="CommentRowCount"/>  together with CommentCountRow.
		/// Do not confuse them, even when they have similar names. <see cref="CommentRowCount"/> returns the maximum row with a comment,
		/// and CommentCountRow returns the number of comments in a given row.
		/// </remarks>
		/// <example>
		/// You can loop on all the comments on a sheet with the following code:
		/// <code>
		/// for (int r=1; r&lt;= xls.CommentRowCount(); r++)
		///   for (int i=1; i&lt;= xls.CommentCountRow(r); i++)
		///   {
		///      //Do something(GetCommentRow(r,i);
		///   }
		/// </code>
		/// </example>
		/// <returns>The number of comments on a given row.</returns>
		/// <seealso cref="CommentRowCount"/> <seealso cref="GetCommentRow"/> <seealso cref="GetCommentRowCol"/> <seealso cref="GetComment"/> <seealso cref="SetCommentRow(int, int, string, TImageProperties)"/> <seealso cref="SetComment(int, int, TRichString)"/><seealso cref="GetCommentPropertiesRow"/> <seealso cref="SetCommentPropertiesRow"/><seealso cref="SetCommentProperties"/>
		public abstract int CommentCountRow(int row);

		/// <summary>
		/// Returns the comment at position commentIndex on the specified row.
		/// </summary>
		/// <remarks>This method is used together with <see cref="CommentCountRow"/>. See the reference on it for an example.</remarks>
		/// <param name="row">Row index (1 based)</param>
		/// <param name="commentIndex">Comment index (1 based). See <see cref="CommentCountRow"/> </param>
		/// <returns>Comment at the specified position.</returns>
		/// <seealso cref="CommentRowCount"/> <seealso cref="CommentCountRow"/> <seealso cref="GetCommentRowCol"/> <seealso cref="GetComment"/> <seealso cref="SetCommentRow(int, int, string, TImageProperties)"/> <seealso cref="SetComment(int, int, TRichString)"/><seealso cref="GetCommentPropertiesRow"/> <seealso cref="SetCommentPropertiesRow"/><seealso cref="SetCommentProperties"/> 
		public abstract TRichString GetCommentRow(int row, int commentIndex);

		/// <summary>
		/// Returns the column for comment at position commentIndex
		/// </summary>
		/// <remarks>This method is used together with <see cref="CommentCountRow"/>. See the reference on it for an example.</remarks>
		/// <param name="row">Row with the comment. (1 based)</param>
		/// <param name="commentIndex">Index of the comment (1 based)</param>
		/// <returns>The column index corresponding to the comment.</returns>
		/// <seealso cref="CommentRowCount"/> <seealso cref="CommentCountRow"/> <seealso cref="GetCommentRow"/> <seealso cref="GetComment"/> <seealso cref="SetCommentRow(int, int, string, TImageProperties)"/> <seealso cref="SetComment(int, int, TRichString)"/><seealso cref="GetCommentPropertiesRow"/> <seealso cref="SetCommentPropertiesRow"/><seealso cref="SetCommentProperties"/>  
		public abstract int GetCommentRowCol(int row, int commentIndex);

		/// <summary>
		/// Returns the comment at the specified row and column, or an empty string if there is no comment on that cell.
		/// </summary>
		/// <remarks>Use this method when you are searching for a comment on a fixed position. To loop along all
		/// comments on a sheet, see <see cref="CommentRowCount"/></remarks>
		/// <param name="row">Row index (1 based)</param>
		/// <param name="col">Column index (1 based)</param>
		/// <returns>The comment on the specified cell, String.Empty if there is no comment on it.</returns>
		/// <seealso cref="CommentRowCount"/> <seealso cref="CommentCountRow"/> <seealso cref="GetCommentRow"/> <seealso cref="GetCommentRowCol"/> <seealso cref="SetCommentRow(int, int, string, TImageProperties)"/> <seealso cref="SetComment(int, int, TRichString)"/><seealso cref="GetCommentPropertiesRow"/> <seealso cref="SetCommentPropertiesRow"/><seealso cref="SetCommentProperties"/>  
		public abstract TRichString GetComment(int row, int col);

		/// <summary>
		/// Changes the properties (text and position of the popup) for an existing comment at commentIndex.
		/// To delete a comment, set a "new TRichString()" as the "value" param. To add a new comment, use <see cref="SetComment(int, int, TRichString)"/>. 
		/// </summary>
		/// <remarks>This method is used together with <see cref="CommentCountRow"/>. See the reference on it for an example.</remarks> 
		/// <param name="row">Row Index (1 based)</param>
		/// <param name="commentIndex">Comment index (1 based)</param>
		/// <param name="value">Text of the comment. Set it to "new TRichString()" to remove the comment.</param>
		/// <param name="commentProperties">Properties of the popup.</param>
		/// <seealso cref="CommentRowCount"/> <seealso cref="CommentCountRow"/> <seealso cref="GetCommentRow"/> <seealso cref="GetCommentRowCol"/> <seealso cref="GetComment"/> <seealso cref="SetComment(int, int, TRichString)"/><seealso cref="GetCommentPropertiesRow"/> <seealso cref="SetCommentPropertiesRow"/><seealso cref="SetCommentProperties"/>  
		public abstract void SetCommentRow(int row, int commentIndex, TRichString value, TImageProperties commentProperties);

		/// <summary>
		/// Changes the properties (text and position of the popup) for an existing comment at commentIndex.
		/// To delete a comment, set a String.Empty as the "value" param. To add a new comment, use <see cref="SetComment(int, int, TRichString)"/>. 
		/// </summary>
		/// <remarks>This method is used together with <see cref="CommentCountRow"/>. See the reference on it for an example.</remarks> 
		/// <param name="row">Row Index (1 based)</param>
		/// <param name="commentIndex">Comment index (1 based)</param>
		/// <param name="value">Text of the comment. Set it to String.Empty to remove the comment.</param>
		/// <param name="commentProperties">Properties of the popup.</param>
		/// <seealso cref="CommentRowCount"/> <seealso cref="CommentCountRow"/> <seealso cref="GetCommentRow"/> <seealso cref="GetCommentRowCol"/> <seealso cref="GetComment"/> <seealso cref="SetComment(int, int, TRichString)"/><seealso cref="GetCommentPropertiesRow"/> <seealso cref="SetCommentPropertiesRow"/><seealso cref="SetCommentProperties"/>  
		public void SetCommentRow(int row, int commentIndex, string value, TImageProperties commentProperties)
		{
			SetCommentRow(row, commentIndex, new TRichString(value), commentProperties);
		}

        /// <inheritdoc cref="SetComment(int, int, string, string, TImageProperties)" />
        public abstract void SetComment(int row, int col, TRichString value, string author, TImageProperties commentProperties);

        /// <summary>
        /// Sets or deletes a comment at the specified cell.
        /// </summary>
        /// <remarks>To delete a comment, set its value to String.Empty.</remarks>
        /// <param name="row">Row index (1 based)</param>
        /// <param name="col">Column index (1 based)</param>
        /// <param name="value">Text of the comment. Set it to String.Empty to delete the comment.</param>
        /// <param name="author">Author of the comment.</param>
		/// <param name="commentProperties">Properties of the popup.</param>
        /// <seealso cref="CommentRowCount"/> <seealso cref="CommentCountRow"/> <seealso cref="GetCommentRow"/> <seealso cref="GetCommentRowCol"/> <seealso cref="GetComment"/> <seealso cref="SetCommentRow(int, int, string, TImageProperties)"/><seealso cref="GetCommentPropertiesRow"/> <seealso cref="SetCommentPropertiesRow"/><seealso cref="SetCommentProperties"/>  
        public void SetComment(int row, int col, string value, string author, TImageProperties commentProperties)
		{
			SetComment(row, col, new TRichString(value), author, commentProperties);
        }

        /// <inheritdoc cref="SetComment(int, int, string, string, TImageProperties)" />
        public void SetComment(int row, int col, string value)
		{
			SetComment(row, col, value, String.Empty, null);
		}

        /// <inheritdoc cref="SetComment(int, int, string, string, TImageProperties)" />
        public void SetComment(int row, int col, TRichString value)
		{
			SetComment(row, col, value, String.Empty, null);
		}

		/// <summary>
		/// Returns the comment properties for the popup at position commentIndex
		/// </summary>
		/// <remarks>This method is used together with <see cref="CommentCountRow"/>. See the reference on it for an example.</remarks> 
		/// <param name="row">Row index (1 based)</param>
		/// <param name="commentIndex">Comment index (1 based)</param>
		/// <returns>The comment properties.</returns>
		/// <seealso cref="CommentRowCount"/> <seealso cref="CommentCountRow"/> <seealso cref="GetCommentRow"/> <seealso cref="GetCommentRowCol"/> <seealso cref="GetComment"/> <seealso cref="SetCommentRow(int, int, string, TImageProperties)"/><seealso cref="SetComment(int, int, TRichString)"/> <seealso cref="GetCommentPropertiesRow"/><seealso cref="SetCommentPropertiesRow"/><seealso cref="SetCommentProperties"/>  
		public abstract TCommentProperties GetCommentPropertiesRow(int row, int commentIndex);

		/// <summary>
		/// Gets the popup placement for an existing comment. If there is not a comment on cell (row,col), this will return null.
		/// </summary>
		/// <remarks>Note that you can change the size but not the placement of the popup.
		/// This placement you set here is the one you see when you right click the cell and choose "Show comment". 
		/// The yellow popup box is placed automatically by excel.</remarks>
		/// <param name="row">Row index (1 based).</param>
		/// <param name="col">Column index (1 based)</param>
		/// <returns>Placement of the comment popup.</returns>
		/// <seealso cref="CommentRowCount"/> <seealso cref="CommentCountRow"/> <seealso cref="GetCommentRow"/> <seealso cref="GetCommentRowCol"/> <seealso cref="GetComment"/> <seealso cref="SetCommentRow(int, int, string, TImageProperties)"/><seealso cref="SetComment(int, int, TRichString)"/> <seealso cref="GetCommentPropertiesRow"/> <seealso cref="SetCommentPropertiesRow"/><seealso cref="SetCommentProperties"/>  
		public abstract TCommentProperties GetCommentProperties(int row, int col);

		/// <summary>
		/// Sets the comment properties at the specified index.
		/// </summary>
		/// <remarks>This method is used together with <see cref="CommentCountRow"/>. See the reference on it for an example.</remarks> 
		/// <param name="row">Row index (1 based)</param>
		/// <param name="commentIndex">Comment index (1 based)</param>
		/// <param name="commentProperties">Comment properties. This parameter can be a TImageProperties, or the more complete derived class TCommentProperties, if you need to set extra information like the text alignment.</param>
		/// <seealso cref="CommentRowCount"/> <seealso cref="CommentCountRow"/> <seealso cref="GetCommentRow"/> <seealso cref="GetCommentRowCol"/> <seealso cref="GetComment"/> <seealso cref="SetCommentRow(int, int, string, TImageProperties)"/><seealso cref="SetComment(int, int, TRichString)"/> <seealso cref="GetCommentPropertiesRow"/> <seealso cref="GetCommentProperties"/><seealso cref="SetCommentProperties"/>  
		public abstract void SetCommentPropertiesRow(int row, int commentIndex, TImageProperties commentProperties);

		/// <summary>
		/// Sets the popup placement for an existing comment. If there is not a comment on cell (row,col), this will create an empty one.
		/// </summary>
		/// <remarks>Note that you can change the size but not the placement of the popup.
		/// This placement you set here is the one you see when you right click the cell and choose "Show comment". 
		/// The yellow popup box is placed automatically by excel.</remarks>
		/// <param name="row">Row index (1 based)</param>
		/// <param name="col">Column index (1 based)</param>
        /// <param name="commentProperties">Placement and properties of the comment popup. Null if there is no comment on the cell. This parameter can be a TImageProperties, or the more complete derived class TCommentProperties, if you need to set extra information like the text alignment.</param>
		/// <seealso cref="CommentRowCount"/> <seealso cref="CommentCountRow"/> <seealso cref="GetCommentRow"/> <seealso cref="GetCommentRowCol"/> <seealso cref="GetComment"/> <seealso cref="SetCommentRow(int, int, string, TImageProperties)"/><seealso cref="SetComment(int, int, TRichString)"/> <seealso cref="GetCommentPropertiesRow"/> <seealso cref="GetCommentProperties"/><seealso cref="SetCommentPropertiesRow"/>  
		public abstract void SetCommentProperties(int row, int col, TImageProperties commentProperties);
		#endregion

		#region Cell operations
        /// <inheritdoc cref="InsertAndCopyRange(TXlsCellRange, int, int, int, TFlxInsertMode, TRangeCopyMode, ExcelFile, int, TExcelObjectList)" />
        public void InsertAndCopyRange(TXlsCellRange sourceRange, int destRow, int destCol, int destCount, TFlxInsertMode insertMode)
		{
			InsertAndCopyRange(sourceRange, destRow, destCol, destCount, insertMode, TRangeCopyMode.All, null, 0);
		}

        /// <inheritdoc cref="InsertAndCopyRange(TXlsCellRange, int, int, int, TFlxInsertMode, TRangeCopyMode, ExcelFile, int, TExcelObjectList)" />
        ///<remarks>This overload version is useful for <b>inserting</b> only and not copying.</remarks>
        public void InsertAndCopyRange(TXlsCellRange sourceRange, int destRow, int destCol, int destCount, TFlxInsertMode insertMode, TRangeCopyMode copyMode)
		{
			InsertAndCopyRange(sourceRange, destRow, destCol, destCount, insertMode, copyMode, null, 0);
		}

        /// <inheritdoc cref="InsertAndCopyRange(TXlsCellRange, int, int, int, TFlxInsertMode, TRangeCopyMode, ExcelFile, int, TExcelObjectList)" />
        ///<remarks>This overload is useful for <b>copying from another file.</b> It is not as fast or complete as the other
        ///overloaded versions because it has to do a lot of transforms on the data. But it is very useful anyway.</remarks>
        public void InsertAndCopyRange(TXlsCellRange sourceRange, int destRow, int destCol, int destCount, TFlxInsertMode insertMode, TRangeCopyMode copyMode, ExcelFile sourceWorkbook, int sourceSheet)
        {
            InsertAndCopyRange(sourceRange, destRow, destCol, destCount, insertMode, copyMode, sourceWorkbook, sourceSheet, null);
        }

        /// <summary>
        /// Inserts and/or copies a range of cells from one place to another.<br/>
        /// This method is one of the most important on FlexCel API, and it allows you to copy ranges of cells from one place to another,
        /// adapting the formulas, images and everything as Excel would do it.
        /// </summary>
        /// <param name="sourceRange">The range of cells you want to copy. If you specify full rows, they will
        /// be copied with Row format information and size. If you copy just a part of a row, Row format will not be copied.
        /// The same applies to columns. The only way to copy <b>all</b> row and columns, is to specify the full (A:IV) range.</param>
        /// <param name="destRow">Destination row where the cells will be copied.</param>
        /// <param name="destCol">Destination column where the cells will be copied.</param>
        /// <param name="destCount">Number of times the sourceRange will be copied at (desRow, destCol).
        /// If you make for example destCount=2, sourceRange will be copied 2 times at (destRow, destCol)</param>
        /// <param name="insertMode">How the cells on destination will be inserted. They can shift down or left. 
        /// Specifying Row or Col as mode is equivalent to specify a sourceRange including full rows or columns respectively. </param>
        /// <param name="copyMode">Which cells on sourceRange will be copied. If you intend to replace values on the copied cells,
        /// you might specify OnlyFormulas. If you just want to <b>insert</b> cells and not copy, specify None.</param>
        ///<param name="sourceWorkbook">Workbook from where we are copying the cells. This might be the same workbook, and you would by copying from another sheet.</param>
        ///<param name="sourceSheet">Sheet index on the source workbook. If sourceWorkbook is the same instance as this, and sourceSheet is the active sheet on the instance, this method is equivalent to the simpler overloaded version.</param>
        ///<param name="ObjectsInRange">Returns the objects that are in the range to be copied this is an optimization so you don't have to find those objects again. Set it to null to not return any objects</param>
        ///<remarks>This overload is useful for <b>copying from another file.</b> It is not as fast or complete as the other
        ///overloaded versions because it has to do a lot of transforms on the data. But it is very useful anyway.</remarks>
        public abstract void InsertAndCopyRange(TXlsCellRange sourceRange, int destRow, int destCol, int destCount, TFlxInsertMode insertMode, TRangeCopyMode copyMode, ExcelFile sourceWorkbook, int sourceSheet, TExcelObjectList ObjectsInRange);

        /// <summary>
		/// Deletes a range of cells, and moves all cells below up or all cells to the right left, depending on the insert mode.
		/// </summary>
		/// <param name="cellRange">Range of cells to delete.</param>
		/// <param name="insertMode">Mode of deletion. Note that Row and Col are equivalent to ShiftRight and ShiftDown with a 
		/// cell range of full rows or cols respectively.</param>
		public abstract void DeleteRange(TXlsCellRange cellRange, TFlxInsertMode insertMode);

		/// <summary>
		/// Deletes a range of cells, and moves all cells below up or all cells to the right left, depending on the insert mode.
		/// </summary>
		/// <param name="sheet1">First sheet where to delete cells.</param>
		/// <param name="sheet2">Last sheet where to delete cells.</param>
		/// <param name="cellRange">Range of cells to delete.</param>
		/// <param name="insertMode">Mode of deletion. Note that Row and Col are equivalent to ShiftRight and ShiftDown with a 
		/// cell range of full rows or cols respectively.</param>
		public abstract void DeleteRange(int sheet1, int sheet2, TXlsCellRange cellRange, TFlxInsertMode insertMode);

		/// <summary>
		/// Moves a range of cells, the same way Excel does it. All references pointing to the old range will be relocated to the new, and all
		/// exisitng references to the new range will be relocated to #ref. 
		/// </summary>
		/// <param name="cellRange">Range you want to move.</param>
		/// <param name="newRow">Row where the range will be relocated.</param>
		/// <param name="newCol">Column where the range will be relocated.</param>
		/// <param name="insertMode">This parameter switches between 2 different working modes:
		/// <list type="bullet">
		/// <item>
		/// Modes NoneDown and NoneRight are the same, and work the same way as when you drag a range of cells in Excel to 
		/// a new location. Cells on the new range will be replaced by the old, and cells where the old range was located will be cleared(not deleted).
		/// No cells are insered or deleted.
		/// </item>
		/// <item>
		/// The other modes behave like when you select a range in Excel, cut it, and then right click and select "Insert Cut Cells...". The old range will be inserted
		/// where the new range goes, cells where the old range was will be deleted(not cleared). 
		/// Using for example MoveRange(CellRange, newRow, newCol, TFlxInsertMode.ShiftRight) is equivalent to:
		/// 1) InsertAndCopyRange(CellRange, newRow, newCol, 1, TFlxInsertMode.ShiftRight, TRangeCopyMode.None); 
		/// 2) MoveRange(CellRange, newRow, newCol, TFlxInsertMode.NoneRight);  
		/// 3) DeleteRange(CellRange, TFlxInsertMode.ShiftRight);
		/// </item>
		/// </list>
		/// </param>
		public abstract void MoveRange(TXlsCellRange cellRange, int newRow, int newCol, TFlxInsertMode insertMode);

		#endregion

		#region Data Validation
		/// <summary>
		/// Clears all data validation entries in the active sheet.
		/// </summary>
		public abstract void ClearDataValidation();

		/// <summary>
		/// Clears all data validation entries inside the specified range.
		/// </summary>
		/// <param name="range">Range of cells where data validation will be cleared.</param>
		public abstract void ClearDataValidation(TXlsCellRange range);

		/// <summary>
		/// Adds a new Data Validation to a specified range.
		/// </summary>
		/// <param name="range">Range of cells where we will apply the Data Validation.</param>
		/// <param name="validationInfo">Validation information.</param>
		public abstract void AddDataValidation(TXlsCellRange range, TDataValidationInfo validationInfo);

		/// <summary>
		/// Returns the validation information for an specific cell. If the cell has no Data Validation associated, this method returns null.
		/// </summary>
		/// <param name="row">Row of the cell (1 based)</param>
		/// <param name="col">Column of the cell (1 based)</param>
		/// <returns>The data validation for a cell, or null if the cell has no Data Validation associated.</returns>
		public abstract TDataValidationInfo GetDataValidation(int row, int col);

		#region Indexed Data Validation

		/// <summary>
		/// Returns the number of DataValidation structures in the active sheet. 
		/// There are 2 ways you can access the data validation
		/// information on a sheet: 
		/// <list type="number">
		/// <item>If you know the row and column where you want to look, you can use <see cref="GetDataValidation(int,int)"/> to return the data validation in the cell.</item>
		/// <item>If you want to find out all data validation structures in the sheet, you can use <see cref="DataValidationCount"/>, 
		/// <see cref="GetDataValidationInfo(int)"/> and <see cref="GetDataValidationRanges(int)"/> to loop over all existing data validations.</item>
		/// </list>
		/// </summary>
		public abstract int DataValidationCount {get;}

		/// <summary>
		/// Returns the data validation information for an entry of the index.
		/// There are 2 ways you can access the data validation
		/// information on a sheet: 
		/// <list type="number">
		/// <item>If you know the row and column where you want to look, you can use <see cref="GetDataValidation(int,int)"/> to return the data validation in the cell.</item>
		/// <item>If you want to find out all data validation structures in the sheet, you can use <see cref="DataValidationCount"/>, 
		/// <see cref="GetDataValidationInfo(int)"/> and <see cref="GetDataValidationRanges(int)"/> to loop over all existing data validations.</item>
		/// </list>
		/// </summary>
		/// <param name="index">Position in the list of data validations. (1 based)</param>
		/// <returns>Data validation information.</returns>
		public abstract TDataValidationInfo GetDataValidationInfo(int index);

		/// <summary>
		/// Returns a list of ranges for which a data validation definition applies.
		/// There are 2 ways you can access the data validation
		/// information on a sheet: 
		/// <list type="number">
		/// <item>If you know the row and column where you want to look, you can use <see cref="GetDataValidation(int,int)"/> to return the data validation in the cell.</item>
		/// <item>If you want to find out all data validation structures in the sheet, you can use <see cref="DataValidationCount"/>, 
		/// <see cref="GetDataValidationInfo(int)"/> and <see cref="GetDataValidationRanges(int)"/> to loop over all existing data validations.</item>
		/// </list>
		/// </summary>
		/// <param name="index">Position in the list of data validations. (1 based)</param>
		/// <returns>A list of cell ranges.</returns>
		public abstract TXlsCellRange[] GetDataValidationRanges(int index);

		#endregion

		#endregion

		#region HyperLinks
		/// <summary>
		/// The count of hyperlinks on the active sheet
		/// </summary>
		public abstract int HyperLinkCount{get;}

		/// <summary>
		/// Returns the hyperlink at position index on the list.
		/// </summary>
		/// <param name="hyperLinkIndex">Index of the hyperlink (1 based).</param>
		/// <returns>Hyperlink description.</returns>
		public abstract THyperLink GetHyperLink(int hyperLinkIndex);

		/// <summary>
		/// Modifies an existing Hyperlink. Use <see cref="AddHyperLink"/> to add a new one.
		/// </summary>
		/// <param name="hyperLinkIndex">Index of the hyperlink (1 based).</param>
		/// <param name="value">Hyperlink description.</param>
		public abstract void SetHyperLink(int hyperLinkIndex, THyperLink value);

		/// <summary>
		/// Returns the cell range a hyperlink refers to.
		/// </summary>
		/// <param name="hyperLinkIndex">Index of the hyperlink (1 based).</param>
		/// <returns>Range the hyperlink applies to.</returns>
		/// <remarks>While normally hyperlinks refer to a single cell, you can make them point to
		/// a range. This method will return the first and last cell of the range that the hyperlink applies to.</remarks>
		public abstract TXlsCellRange GetHyperLinkCellRange(int hyperLinkIndex);

		/// <summary>
		/// Changes the cells an hyperlink is linked to.
		/// </summary>
		/// <param name="hyperLinkIndex">Index of the hyperlink (1 based).</param>
		/// <param name="cellRange">Range of cells the hyperlink will refer to.</param>
		public abstract void SetHyperLinkCellRange(int hyperLinkIndex, TXlsCellRange cellRange);

		/// <summary>
		/// Adds a new hyperlink to the Active sheet. Use <see cref="SetHyperLink"/> to modify an existing one.
		/// </summary>
		/// <param name="cellRange">Range of cells the hyperlink will refer to.</param>
		/// <param name="value">Description of the hyperlink.</param>
		public abstract void AddHyperLink(TXlsCellRange cellRange, THyperLink value);

		/// <summary>
		/// Deletes an existing hyperlink.
		/// </summary>
		/// <param name="hyperLinkIndex">Index of the hyperlink (1 based).</param>
		public abstract void DeleteHyperLink(int hyperLinkIndex);
		#endregion

		#region Group and Outline
		/// <summary>
		/// Returns the Outline level for a row.
		/// </summary>
		/// <param name="row">Row index (1 based)</param>
		/// <returns>Outline level for a row. It is a number between 0 and 7.</returns>
		public abstract int GetRowOutlineLevel(int row);

		/// <summary>
		/// Sets the Outline level for a row.
		/// </summary>
		/// <param name="row">Row index (1 based)</param>
		/// <param name="level">Outline level. must be between 0 and 7.</param>
		public void SetRowOutlineLevel(int row, int level)
		{
			SetRowOutlineLevel(row, row, level);
		}

		/// <summary>
		/// Sets the Outline level for a row range.
		/// </summary>
		/// <param name="firstRow">Row index of the first row on the range. (1 based)</param>
		/// <param name="lastRow">Row index of the last row on the range. (1 based)</param>
		/// <param name="level">Outline level. must be between 0 and 7.</param>
		public abstract void SetRowOutlineLevel(int firstRow, int lastRow, int level);

		/// <summary>
		/// Returns the Outline level for a column.
		/// </summary>
		/// <param name="col">Column index (1 based)</param>
		/// <returns>Outline level for a column. It is a number between 0 and 7.</returns>
		public abstract int GetColOutlineLevel(int col);

		/// <summary>
		/// Sets the Outline level for a column.
		/// </summary>
		/// <param name="col">Column index (1 based)</param>
		/// <param name="level">Outline level. must be between 0 and 7.</param>
		public void SetColOutlineLevel(int col, int level)
		{
			SetColOutlineLevel(col, col, level);
		}
        
		/// <summary>
		/// Sets the Outline level for a column range.
		/// </summary>
		/// <param name="firstCol">Column index of the first column on the range. (1 based)</param>
		/// <param name="lastCol">Column index of the last column on the range. (1 based)</param>
		/// <param name="level">Outline level. must be between 0 and 7.</param>
		public abstract void SetColOutlineLevel(int firstCol, int lastCol, int level);

		/// <summary>
		/// Determines whether the summary rows should be below or above details on outline.
		/// </summary>
		public abstract bool OutlineSummaryRowsBelowDetail{get;set;}

		/// <summary>
		/// Determines whether the summary columns should be right to or left to the details on outline.
		/// </summary>
		public abstract bool OutlineSummaryColsRightToDetail{get;set;}

		/// <summary>
		/// This handles the setting of Automatic Styles inside the outline options.
		/// </summary>
		public abstract bool OutlineAutomaticStyles{get;set;}

		/// <summary>
		/// Collapses or expands the row outlines in a sheet to the specified level. It is equivalent to pressing the 
		/// numbers at the top of the outline gutter in Excel.
		/// </summary>
		/// <param name="level">Level that we want to show of the outline. (1 based). 
		/// For example, setting Level = 3 is the same as pressing the "3" number at the top of the outline gutter in Excel.
		/// Setting Level = 1 will collapse all groups, Level = 8 will expand all groups.</param>
		/// <param name="collapseChildren">Determines if the children of the collapsed nodes will be collapsed too.</param>
		public void CollapseOutlineRows(int level, TCollapseChildrenMode collapseChildren)
		{
			if (RowCount <= 0) return;
			CollapseOutlineRows(level, collapseChildren, 1, RowCount);
		}

		/// <summary>
		/// Collapses or expands the row outlines in a sheet to the specified level. It is equivalent to pressing the 
		/// numbers at the top of the outline gutter in Excel.
		/// </summary>
		/// <param name="level">Level that we want to show of the outline. (1 based). 
		/// For example, setting Level = 3 is the same as pressing the "3" number at the top of the outline gutter in Excel.
		/// Setting Level = 1 will collapse all groups, Level = 8 will expand all groups.</param>
		/// <param name="collapseChildren">Determines if the children of the collapsed nodes will be collapsed too.</param>
		/// <param name="firstRow">This defines the first row of the range to collapse/expand. Only rows inside that range will be modified.</param>
		/// <param name="lastRow">This defines the last row of the range to collapse/expand. Only rows inside that range will be modified.</param>
		public abstract void CollapseOutlineRows(int level, TCollapseChildrenMode collapseChildren, int firstRow, int lastRow);

		/// <summary>
		/// Collapses or expands the column outlines in a sheet to the specified level. It is equivalent to pressing the 
		/// numbers at the left of the outline gutter in Excel.
		/// </summary>
		/// <param name="level">Level that we want to display from the outline. (1 based). 
		/// For example, setting Level = 3 is the same as pressing the "3" number at the left of the outline gutter in Excel.
		/// Setting Level = 1 will collapse all groups, Level = 8 will expand all groups.</param>
		/// <param name="collapseChildren">Determines if the children of the collapsed nodes will be collapsed too.</param>
		public void CollapseOutlineCols(int level, TCollapseChildrenMode collapseChildren)
		{
			CollapseOutlineCols(level, collapseChildren, 1, FlxConsts.Max_Columns + 1);
		}

		/// <summary>
		/// Collapses or expands the column outlines in a sheet to the specified level. It is equivalent to pressing the 
		/// numbers at the top of the outline gutter in Excel.
		/// </summary>
		/// <param name="level">Level that we want to display from the outline. (1 based). 
		/// For example, setting Level = 3 is the same as pressing the "3" number at the left of the outline gutter in Excel.
		/// Setting Level = 1 will collapse all groups, Level = 8 will expand all groups.</param>		 
		/// <param name="collapseChildren">Determines if the children of the collapsed nodes will be collapsed too.</param>
		/// <param name="firstCol">This defines the first column of the range to collapse/expand. Only columns inside that range will be modified.</param>
		/// <param name="lastCol">This defines the last column of the range to collapse/expand. Only columns inside that range will be modified.</param>
		public abstract void CollapseOutlineCols(int level, TCollapseChildrenMode collapseChildren, int firstCol, int lastCol);

		/// <summary>
		/// Returns true when the row is the one that is used for collapsing an outline. (it has a "+" at the left).
		/// </summary>
		/// <param name="row">Row to test (1 based)</param>
		/// <returns>True if the node has a "+" mark.</returns>
		public abstract bool IsOutlineNodeRow(int row);
		
		/// <summary>
		/// Returns true when the column is the one that is used for collapsing an outline. (it has a "+" at the top).
		/// </summary>
		/// <param name="col">Column to test (1 based)</param>
		/// <returns>True if the node has a "+" mark.</returns>
		public abstract bool IsOutlineNodeCol(int col);

		/// <summary>
		/// Returns true when the row is an outline node (it has a "+" at the left) and it is closed (all children are hidden).
		/// </summary>
		/// <param name="row">Row to test (1 based)</param>
		/// <returns>True if the row contains a node and it is collapsed, false otherwise.</returns>
		public abstract bool IsOutlineNodeCollapsedRow(int row);
		
		/// <summary>
		/// Returns true when the column is an outline node (it has a "+" at the top) and it is closed (all children are hidden).
		/// </summary>
		/// <param name="col">Column to test (1 based)</param>
		/// <returns>True if the column contains a node and it is collapsed, false otherwise.</returns>
		public abstract bool IsOutlineNodeCollapsedCol(int col);

		/// <summary>
		/// Use this method to collapse a node of the outline. If the row is not a node (<see cref="IsOutlineNodeRow"/> is false) this method does nothing.
		/// While this method allows a better control of the rows expanded and collapsed, you will normally use <see cref="CollapseOutlineRows(int,TCollapseChildrenMode)"/> to collapse or
		/// expand all rows in a sheet.
		/// </summary>
		/// <param name="row">Row to expand or collapse (1 based)</param>
		/// <param name="collapse">If true, the node will be collapsed, else it will be expanded.</param>
		public abstract void CollapseOutlineNodeRow(int row, bool collapse);

		/// <summary>
		/// Use this method to collapse a node of the outline. If the column is not a node (<see cref="IsOutlineNodeCol"/> is false) this method does nothing.
        /// While this method allows a better control of the columns expanded and collapsed, you will normally use <see cref="CollapseOutlineCols(int,TCollapseChildrenMode)"/> to collapse or
		/// expand all columns in a sheet.
		/// </summary>
		/// <param name="col">Column to expand or collapse. (1 based)</param>
		/// <param name="collapse">If true, the node will be collapsed, else it will be expanded.</param>
		public abstract void CollapseOutlineNodeCol(int col, bool collapse);
		#endregion

		#region Protection
		/// <summary>
		/// Protection data for the file. Modify its properties to open and read encrypted files.
		/// </summary>
		public TProtection Protection {get {return FProtection;}}

		//Methods from here are internal, and work as a wrapper to the Protection object. To use them, call Protection.XXXX

		internal abstract void SetModifyPassword(string modifyPassword, bool recommendReadOnly, string reservingUser);

		internal abstract bool HasModifyPassword{get;}

		internal abstract bool RecommendReadOnly{get; set;}

		internal abstract void SetWorkbookProtection (string workbookPassword, TWorkbookProtectionOptions workbookProtectionOptions);

		internal abstract bool HasWorkbookPassword{get;}

		internal abstract TWorkbookProtectionOptions WorkbookProtectionOptions{get;set;}

        internal abstract void SetSharedWorkbookProtection(string sharedWorkbookPassword, TSharedWorkbookProtectionOptions sharedWorkbookProtectionOptions);

        internal abstract bool HasSharedWorkbookPassword { get; }

        internal abstract TSharedWorkbookProtectionOptions SharedWorkbookProtectionOptions { get; set; }

		internal abstract void SetSheetProtection (string sheetPassword, TSheetProtectionOptions sheetProtectionOptions);

		internal abstract bool HasSheetPassword{get;}

		internal abstract TSheetProtectionOptions SheetProtectionOptions{get;set;}

		internal abstract string WriteAccess {get;set;}

		#endregion

		#region Cell selection
		/// <summary>
		/// Selects a single cell. To select multiple cells,
		/// use <see cref="SelectCells"/>
		/// </summary>
		/// <param name="row">Row to select (1 based)</param>
		/// <param name="col">Column to select (1 based)</param>
		/// <param name="scrollWindow">When true, window will scroll so the selected cell is visible. This is equivalent to using <see cref="ScrollWindow(int, int)"/> method and is provided as a shortcut.</param>
		public void SelectCell(int row, int col, bool scrollWindow)
		{
			TXlsCellRange Cr1 = new TXlsCellRange(row, col, row, col);
			TXlsCellRange[] Crn = new TXlsCellRange[] {Cr1};
			SelectCells(Crn);
			if (scrollWindow) ScrollWindow(row, col);
		}

		/// <summary>
		/// Selects a group of cells on a given pane. If you just want to select just one cell, you can use the simpler method <see cref="SelectCell"/>
		/// </summary>
		/// <param name="cellRange">Cells to select.</param>
		public abstract void SelectCells(TXlsCellRange[] cellRange); 

		/// <summary>
		/// Returns the selected ranges on a sheet.
		/// </summary>
		public abstract TXlsCellRange[] GetSelectedCells(); 

		/// <summary>
		/// Scrolls the window to an specified place. If the window is split, it will move the left and top panels.
		/// </summary>
		/// <param name="row">First visible row.</param>
		/// <param name="col">First visible column.</param>
		public void ScrollWindow(int row, int col)
		{
			ScrollWindow(TPanePosition.UpperLeft, row, col);
		}

		/// <summary>
		/// Scrolls the window to an specified place. 
		/// </summary>
		/// <param name="panePosition">Pane to move. Note that if you move for example the left column of the upper left pane, you will also move the left column of the lower left pane.</param>
		/// <param name="row">First visible row.</param>
		/// <param name="col">First visible column.</param>
		public abstract void ScrollWindow(TPanePosition panePosition, int row, int col);

		/// <summary>
		/// Returns the window scroll for the main pane.
		/// </summary>
		public TCellAddress GetWindowScroll()
		{
			return GetWindowScroll(TPanePosition.UpperLeft);
		}

		/// <summary>
		/// Returns the window scroll for a specified pane.
		/// </summary>
		/// <param name="panePosition">Pane to return. </param>
		public abstract TCellAddress GetWindowScroll(TPanePosition panePosition);
		#endregion

		#region Freeze Panes
		/// <summary>
		/// This command is equivalent to Menu->Window->Freeze Panes. It will freeze
		/// the rows and columns above and to the left from cell. Note that because Excel
		/// works this way, when you <see cref="SplitWindow"/> the panes are suppressed and vice-versa
		/// See also <see cref="GetFrozenPanes"/>
		/// </summary>
		/// <param name="cell">All rows and columns above and to the left of this cell will be frozen. Set it to null or "A1" to unfreeze the panes.</param>
		public abstract void FreezePanes(TCellAddress cell);

		/// <summary>
		/// Returns the cell that is freezing the window, "A1" if no panes are frozen.
		/// See also <see cref="FreezePanes"/>
		/// </summary>
		/// <returns>The cell that is freezing the window, "A1" if no panes are frozen.</returns>
		public abstract TCellAddress GetFrozenPanes();

		/// <summary>
		/// This command is equivalent to Menu->Window->Split. It will split the
		/// window in 4 regions. Note that because Excel
		/// works this way, when you <see cref="FreezePanes"/> the windows are unsplitted and vice-versa
		/// See also <see cref="GetSplitWindow"/>
		/// </summary>
		/// <param name="xOffset">Offset from the left on 1/20 of a point. Zero for no vertical split.</param>
		/// <param name="yOffset">Offset from the top on 1/20 of a point. Zero for no horizontal split.</param>
		public abstract void SplitWindow(int xOffset, int yOffset);

		/// <summary>
		/// Returns the horizontal and vertical offsets for the split windows. Zero means no split.
		/// See also <see cref="SplitWindow"/>
		/// </summary>
		/// <returns>The horizontal and vertical offsets for the split windows. Zero means no split.</returns>
		public abstract TPoint GetSplitWindow();
		#endregion

		#region Document Properties
		/// <summary>
		/// Document properties for the file. With this object you can read the properties (Author, Title, etc.) of a file.
		/// </summary>
		public TDocumentProperties DocumentProperties {get {return FDocumentProperties;}}

		//Methods from here are internal, and work as a wrapper to the DocumentProperties object. To use them, call DocumentProperties.XXXX
		internal abstract object GetStandardProperty(TPropertyId PropertyId);

		#endregion

		#region Charts
		/// <summary>
		/// Returns a chart from an object position and path. If the object does not contain a chart, it returns null.
        /// Note that charts can be first-level objects (in chart sheets), or they can be embedded inside other objects, that can be
        /// themselves embedded inside other objects. So you need to recursively look inside all objects to see if there are charts anywhere. <br/>
        /// Look at the example in this topic to see how to get all charts in a sheet.
		/// </summary>
		/// <param name="objectIndex">Index of the object (1-based)</param>
		/// <param name="objectPath">Index to the child object where the chart is.
        /// If it is a simple object, you can use String.Empty here, if not you need to get the ObjectPath from <see cref="GetObjectProperties"/><br></br>
        /// If it is "absolute"(it starts with "\\"), then the path includes the objectIndex, and the objectIndex is
        /// not used. An object path of "\\1\\2\\3" is exactly the same as using objectIndex = 1 and objectPath = "2\\3"</param>
		/// <returns></returns>
        /// <example>
        /// The following example will retrieve all charts that are inserted as objects in a sheet, along with the chart sheets.
        /// <code>
        /// public void ProcessAllCharts()
		/// {
		/// 			XlsFile x = new XlsFile(true);
		/// 			x.Open("filewithcharts.xls");
		/// 			for (int iSheet = 1; iSheet &lt;= x.SheetCount; iSheet++)
		/// 			{
		/// 				x.ActiveSheet = iSheet;
		/// 
		/// 				//Process charts embedded as objects in worksheets.
        /// 				for (int iChart = 1; iChart &lt;= x.ObjectCount; iChart++)
		/// 				{
		/// 					TShapeProperties props = x.GetObjectProperties(iChart, true);
		/// 					ProcessChart(x,iChart,  props); //We need to process it even if it is not a chart, since it might be a group with a chart inside.
		/// 				}
		/// 
		/// 				//Process chart sheets.
		/// 				if (x.SheetType == TSheetType.Chart)
		/// 				{
		/// 					ExcelChart ch = x.GetChart(1, null); 
		/// 					DoSomething(x, ch);
		/// 				}
		/// 
		/// 			}
		/// 
		/// 			x.Save(zConsts.Write + "changedfilewithcharts.xls");
        /// 		}
        /// 
		/// 		private void ProcessChart(ExcelFile x, int iChart, TShapeProperties props)
		/// 		{
		/// 			if (props.ObjectType == TObjectType.Chart)
		/// 			{
		/// 				ExcelChart ch = x.GetChart(iChart, props.ObjectPath); 
		/// 				DoSomething(x, ch);
		/// 			}
		/// 
        /// 			for (int i = 1; i &lt;= props.ChildrenCount; i++)
		/// 			{
		/// 				TShapeProperties childProp = props.Children(i);
		/// 				ProcessChart(x, i, childProp);
		/// 			}
        /// 		}
        /// 
        /// </code>
        /// </example>
		public abstract ExcelChart GetChart(int objectIndex, string objectPath);

		/// <summary>
		/// Returns the count of charts on this sheet. Please take note that this method <b>will not return the number of embedded objects with charts inside in a sheet, but just the number of charts in the sheet.</b> 
        /// <br/>In simpler terms, this method will return 0 for all worksheets, and 1 for all chart sheets. 
        /// This is not a very useful method, but it needs to be this way to be consistent
        /// with <see cref="GetChart"/>. So, looping like this:
        /// <code>
        /// for (int i=1; i&lt;= xls.ChartCount; i++)
        /// {
        ///    xls.GetChart(i, null);
        /// }
        /// </code>
        /// will loop 0 times for worksheets and 1 time for workbooks.
        /// <br/> To see how to loop over all objects in a sheet, use the example in the example section of the <see cref="GetChart"/> topic.
		/// </summary>
		public abstract int ChartCount{get;}
		#endregion

		#region Sort, Search and Replace
		/// <summary>
		/// Finds a value inside a cell and returns the position for the cell, or null if nothing was found.
		/// </summary>
		/// <param name="value">Value we are searching.</param>
		/// <param name="Range">Range to Search. Null means the whole worksheet.</param>
		/// <param name="Start">Cell where to start searching. For the first time, use null. After this, you can use the result of this method to get the next cell.</param>
		/// <param name="ByRows">If true, the value will be searched down then left. If false, the search will go left then down. SEARCH IS FASTER WHEN BYROWS = FALSE</param>
		/// <param name="CaseInsensitive">If true, string searches will not be case sensitive, "a" = "A"</param>
		/// <param name="SearchInFormulas">If true, the search will cover formulas too.</param>
		/// <param name="WholeCellContents">If true, only whole cells will be replaced.</param>
		/// <returns>Cell where the string is found, or null if it is not found.</returns>
		/// <example>To find all cells on a sheet that contain the string "bolts":
		/// <code>
		///   do
		///   {
		///      Cell = xls.Find("bolts", null, Cell, false, true, true, false);
		///      if (Cell != null) MessageBox.Show(Cell.CellRef);
		///   }
		///      while (Cell != null);
		/// </code>
		/// </example>
		public abstract TCellAddress Find(object value, TXlsCellRange Range, TCellAddress Start, bool ByRows, bool CaseInsensitive, bool SearchInFormulas, bool WholeCellContents);
		
		/// <summary>
		/// Replaces the instances of oldValue by newValue in the active sheet.
		/// </summary>
		/// <param name="oldValue">Value we want to replace.</param>
		/// <param name="newValue">Value we want to use to replace oldValue.</param>
		/// <param name="Range">Range to Search. Null means the whole worksheet.</param>
		/// <param name="CaseInsensitive">If true, string searches will not be case sensitive, "a" = "A"</param>
		/// <param name="SearchInFormulas">If true, the search will cover formulas too.</param>
		/// <returns>The number of replacements done.</returns>
		/// <param name="WholeCellContents">If true, only whole cells will be replaced.</param>
		/// <example>To replace all cells on a sheet that contain "hello" with "hi":
		/// <code>
		///   xls.Replace("hello", "hi", null, true, true, false);
		/// </code>
		/// </example>
		public abstract int Replace(object oldValue, object newValue, TXlsCellRange Range, bool CaseInsensitive, bool SearchInFormulas, bool WholeCellContents);
		
		/// <summary>
		/// Sorts a range on the current sheet.
		/// </summary>
		/// <param name="Range">Range to sort. It must not include headers.</param>
		/// <param name="ByRows">If true, rows will be sorted. If false, columns will.</param>
		/// <param name="Keys">An array of integers indicating the columns or rows you want to use for sorting. Note that this number is absolute, for example column 1 always means column "A" no matter if the range we are sorting begins at column "B".
		/// A null array means sort by the first column or row, then by the second, etc. up to 8 entries.</param>
		/// <param name="SortOrder">An array of flags indicating whether to sort ascending or descending for each Key. If null, all sorts will be ascending. If not null and the array size is less than the size of the "Keys" parameter, all missing entries are assumed to be Ascending.</param>
		/// <param name="Comparer">Comparer to create a custom way to compare the different items. Set it to null to use default ordering.</param>
		public abstract void Sort (TXlsCellRange Range, bool ByRows, int[] Keys, TSortOrder[] SortOrder, IComparer Comparer);
		#endregion

		#region Options
		/// <summary>
		/// Excel has 2 different date systems. On windows systems it uses 1900 based dates, and on Macintosh systems it uses 1904 dates.
		/// You can change this on Excel under Options, and this property allows you to know and change which format is being used.
		/// </summary>
		public abstract bool OptionsDates1904{get;set;}
		
		/// <summary>
		/// Use this property to change the reference system used in the file. Note that this option <b>only changes how Excel and
        /// FlexCel will display the file.</b> Internally, the formulas will always be stored in A1 format, and converted by Excel
        /// to and from R1C1 if this property is true. FlexCel will also use this property to render the file when it is set to print formulas.
        /// <br></br><br></br>Also, this property doesn't change how FlexCel will parse or return the formula text in the cells or names. 
        /// By default, even if this property is true, you will need to enter the formulas in FlexCel in A1 mode. To change the
        /// entry mode in FlexCel, please use <see cref="FormulaReferenceStyle"/>
		/// </summary>
		public abstract bool OptionsR1C1{get;set;}

		/// <summary>
		/// This property has the value of the corresponding option on Excel options.
		/// </summary>
		public abstract bool OptionsSaveExternalLinkValues{get;set;}

		/// <summary>
		/// This property has the value of the corresponding option on Excel options.
		/// </summary>
		public abstract bool OptionsPrecisionAsDisplayed{get;set;}

        /// <summary>
        /// Number of threads that can be used at the same time by Excel when recalculating. Set it to 0 to disable multithread recalculation, and to -1
        /// to let Excel decide the best number of threads to use. The maximum value for this property is 1024.<br/>
        /// This option only aplies to Excel 2007 or newer.
        /// </summary>
        public abstract int OptionsMultithreadRecalc { get; set; }

        /// <summary>
        /// If true, the full workbook will be recalculated when open in Excel.
        /// This option only aplies to Excel 2007 or newer.
        /// </summary>
        public abstract bool OptionsForceFullRecalc { get; set; }

        /// <summary>
        /// If true, Excel will try to compress the pictures to keep sizes down.
        /// This option only aplies to Excel 2007 or newer.
        /// </summary>
        public abstract bool OptionsAutoCompressPictures { get; set; }

        /// <summary>
        /// Whether the "Check for compatibility" dialog will pop up when saving as xls in Excel 2007 or newer.
        /// This option only aplies to Excel 2007 or newer. 
        /// </summary>
        public abstract bool OptionsCheckCompatibility { get; set; }

        /// <summary>
        /// Defines whether to save a backup copy of the workbook or not.
        /// </summary>
        public abstract bool OptionsBackup { get; set; }

        /// <summary>
        /// Reads and writes the recalculating options in the file. 
        /// <b>Note that this only affects the file and how Excel will recalculate, not how FlexCel does its recalcuation.</b>
        /// FlexCel ignores this setting and uses <see cref="RecalcMode"/> instead.
        /// </summary>
        public abstract TSheetCalcMode OptionsRecalcMode { get; set; }

		
		#endregion

		#region AutoFilter
		/// <summary>
		/// Sets the AutoFilter in the Active sheet to point ot the range specified.
		/// </summary>
		/// <param name="row">Row where the AutoFilter will be placed (1 based).</param>
		/// <param name="col1">First column for the AutoFilter range (1 based).</param>
		/// <param name="col2">Last column for the AutoFilter range (1 based).</param>
		public abstract void SetAutoFilter(int row, int col1, int col2);

		/// <summary>
		/// Sets an AutoFilter in a cell range. Note that the bottom coordinate of the range will be ignored, since AutoFilters 
		/// use only one row.
		/// </summary>
		/// <param name="range">Range to set the AutoFilter. If range is null, this method does nothing. The bottom coordinate on range will not be used.</param>
		public void SetAutoFilter(TXlsCellRange range)
		{
			if (range == null) return;
			SetAutoFilter(range.Top, range.Left, range.Right);
		}

		/// <summary>
		/// Removes the AutoFilter from the active sheet. If there is no AutoFilter in the sheet, this method does nothing.
		/// </summary>
		public abstract void RemoveAutoFilter();

		/// <summary>
		/// Returns true if the active sheet has any AutoFilter defined.
		/// </summary>
		/// <returns>True if the active sheet has an AutoFilter, false otherwise.</returns>
		public abstract bool HasAutoFilter();

		/// <summary>
		/// Returns true if a cell has an AutoFilter.
		/// </summary>
		/// <param name="row">Row of the cell we want to find out. (1 based)</param>
		/// <param name="col">Column of the cell we want to find out. (1 based)</param>
		/// <returns>True if the cell has an AutoFilter, false otherwise.</returns>
		public abstract bool HasAutoFilter(int row, int col);

		/// <summary>
		/// Returns the range of cells with AutoFilter in the Active sheet, or null if there is not AutoFilter. The "Bottom"
		/// coordinate of the returned range has no meaning, only Top, left and right are used. (since AutoFilters only have one row).
		/// </summary>
		/// <returns></returns>
		public abstract TXlsCellRange GetAutoFilterRange();

		#endregion

		#region Recalculation
		/// <summary>
		/// Set this property to change how the file will be recalculated. Note that this affects only
        /// how FlexCel recalculates the file, but not how Excel will recalculate it. It doesn't change anything in 
        /// the generated file. To change the options for the file, use <see cref="OptionsRecalcMode"/> instead.
		/// </summary>
		/// <remarks>
		/// Setting RecalcMode = Manual might be a little faster for big spreadsheets with *lots* of formulas,
		/// but they won't preview ok on Excel viewers as FlexCelPrintDocument. See also <see cref="RecalcForced"/>
		/// When Manual, you can still call <see cref="Recalc(bool)"/> with "forced" = true to recalculate the sheet.
		/// <p>Not all functions are supported, those that are not will return a #NAME! error on a viewer, and will be recalculated by Excel
		/// when you open the file. See SupportedFunctions.xls in the documentation to get a list of supported functions.</p>
		///</remarks>
		public abstract TRecalcMode RecalcMode {get; set;}

		/// <summary>
		/// When <see cref="RecalcMode"/> is manual, use this method to force a recalculation of the spreadsheet.
		/// This specific version of the method will always perform a recalc, even if it is not needed.
		/// You can use <see cref="Recalc(System.Boolean)"/> to recalc only when is needed.
		/// </summary>
		public void Recalc()
		{
			Recalc(true);
		}

		/// <summary>
		/// When <see cref="RecalcMode"/> is manual, use this method to force a recalculation of the spreadsheet.
		/// </summary>
        /// <param name="forced">When true this method will always perform a recalc. When false, only if there has been a change on the spreadsheet.
        /// While for performance reasons you will normally want to keep this false, you might need to set it to true if the formulas refer to functions like "=NOW()" or "=RANDOM()"
        /// that change every time you recalculate.</param>
        public abstract void Recalc(bool forced);

        /// <summary>
        /// This method will recalculate a single cell and all of it's dependencies, but not the whole workbook.<br></br>
        /// <b>USE THIS METHOD WITH CARE!</b> You will normally want to simply call <see cref="Recalc()"/> or just save the file and let FlexCel calculate the workbook for you. 
        /// This method is for rare situations where you are making thousands of recalculations and the speed of Recalc is not enough, 
        /// and you have a big part of the spreadsheet that you know that didn't change.
        /// </summary>
        /// <param name="sheet">Sheet for the cell we want to recalculate. Use <see cref="ActiveSheet"/> here to refer to the active sheet.</param>
        /// <param name="row">Row for the cell we want to recalculate. (1 based)</param>
        /// <param name="col">Column for the cell we want to recalculate. (1 based)</param>
        /// <param name="forced">When true this method will always perform a recalc. When false, only if there has been a change on the spreadsheet.
        /// While for performance reasons you will normally want to keep this false, you might need to set it to true if the formulas refer to functions like "=NOW()" or "=RANDOM()"
        /// that change every time you recalculate.</param>
        /// <returns>The result of the formula at the cell, or null if there is no formula.</returns>
        /// <example>
        /// The following code will recalculate the value of cells A1 and A2, but not C7:
		/// <br/>
        /// <code>
        /// 	XlsFile xls = new XlsFile();
		/// 	xls.NewFile(1);
		/// 	xls.SetCellValue(1, 1, new TFormula("=A2 + 5"));
		/// 	xls.SetCellValue(2, 1, new TFormula("=A3 * 2"));
		/// 	xls.SetCellValue(3, 1, 7);
        /// 
		/// 	xls.SetCellValue(7, 3, new TFormula("=A3 * 2"));
        /// 
		/// 	object ResultValue = xls.RecalcCell(1, 1, 1, true);
		/// 	
		/// 	Debug.Assert((double)ResultValue == 19, "RecalcCell returns the value at the cell.");
		/// 	Debug.Assert((double)(xls.GetCellValue(1, 1) as TFormula).Result == 19, "Cell A1 was recalculated because we called RecalcCell on it.");
		/// 	Debug.Assert((double)(xls.GetCellValue(2, 1) as TFormula).Result == 14, "Cell A2 was recalculated because A1 depends on A2.");
        /// 	Debug.Assert((double)(xls.GetCellValue(7, 3) as TFormula).Result == 0,  "Cell C7 was NOT recalculated because A1 doesn't have a dependency with it. Call xlsFile.Recalc() to recalc all cells.");
        ///  </code>
		/// <br/>
        ///  <code lang = "vbnet">
        ///     Dim xls As New XlsFile()
		///     xls.NewFile(1)
		/// 	xls.SetCellValue(1, 1, New TFormula("=A2 + 5"))
		/// 	xls.SetCellValue(2, 1, New TFormula("=A3 * 2"))
		/// 	xls.SetCellValue(3, 1, 7)
        /// 
		/// 	xls.SetCellValue(7, 3, New TFormula("=A3 * 2"))
        /// 
		/// 	Dim ResultValue As Object = xls.RecalcCell(1, 1, 1, True)
        /// 
		/// 	Debug.Assert(CDbl(ResultValue) = 19, "RecalcCell returns the value at the cell.")
		/// 	Debug.Assert(CDbl(TryCast(xls.GetCellValue(1, 1), TFormula).Result) = 19, "Cell A1 was recalculated because we called RecalcCell on it.")
		/// 	Debug.Assert(CDbl(TryCast(xls.GetCellValue(2, 1), TFormula).Result) = 14, "Cell A2 was recalculated because A1 depends on A2.")
        /// 	Debug.Assert(CDbl(TryCast(xls.GetCellValue(7, 3), TFormula).Result) = 0, "Cell C7 was NOT recalculated because A1 doesn't have a dependency with it. Call xlsFile.Recalc() to recalc all cells.")
        ///  </code>
        /// </example>
        public abstract object RecalcCell(int sheet, int row, int col, bool forced);

        /// <inheritdoc cref="RecalcExpression(string, bool)" />
        public object RecalcExpression(string expression)
        {
            return RecalcExpression(expression, false);
        }

        /// <summary>
        /// Calculates the value of any formula and returns the result. The expression must be a valid Excel formula, it must start with "=",
        /// and cell references that don't specify a sheet (like for example "=A2") will refer to the active sheet. Cells used by the formula will be recalculated as needed too.<br/>
        /// You can use this method as a simple calculator, or to calculate things like the sum
        /// of a range of cells in the spreadsheet. Look at the example for more information on how to use it. <br/>
        /// Note that we will consider the expresion to be located in
        /// the cell A1 of the Active sheet. So for example "=ROW()" will return 1, and "=A2" will return the value of A2 in the active sheet.
        /// </summary>
        /// <param name="expression">Formula to evaluate. It must start with "=" and be a valid Excel formula.</param>
        /// <param name="forced">When true this method will always perform a recalc. When false, only if there has been a change on the spreadsheet.
        /// While for performance reasons you will normally want to keep this false, you might need to set it to true if the formulas refer to functions like "=NOW()" or "=RANDOM()"
        /// that change every time you recalculate.</param>
        /// <returns>The value of the calculated formula.</returns>
        /// <example>
        /// To calculate the sum of all the cells in column A of the sheet "Data", you can use the following code:
        /// <code>
        ///    XlsFile xls = new XlsFile("myfile.xls", true);            
        ///    xls.ActiveSheetByName = "Data";
        ///    Double Result = Convert.ToDouble(xls.RecalcExpression("=Sum(A:A)"));
        /// </code>
        /// To calculate a simple expression, you can use:
        /// <code>
        ///    return (double)xls.RecalcExpression("=1 + 2 * 3");
        /// </code>
        /// </example>
        public abstract object RecalcExpression(string expression, bool forced);

		/// <summary>
		/// Used by the framework to recalculate linked spreadsheets.
		/// </summary>
		internal abstract void InternalRecalc(bool forced, TUnsupportedFormulaList Ufl);

		/// <summary>
		/// Use this method to validate a file. FlexCel does not support all the range of functions from Excel
		/// when recalculating, so unknown functions will return "#NAME!" errors. Using this function you can validate
		/// your user worksheets and see if all the formulas they use are supported.
		/// </summary>
		/// <remarks>Note that you *can* use unsupported functions on FlexCel. When you open the generated file
		/// on Excel it will show ok. The only problem is if you need to natively print or export to pdf the file.
		/// <p>Also, take in account that RecalcAndVerify is slower than recalc, as it has to do more work
		/// to locate the errors. Do not use it as a replace for <see cref="Recalc()"/></p></remarks>
		/// <returns></returns>
		public abstract TUnsupportedFormulaList RecalcAndVerify();


		/// <summary>
		/// When false Excel will not recalculate the formulas when loading the generated file.
		/// </summary>
		/// <remarks>
		/// If this property is true, Excel will recalculate all formulas when you load the file, so it will
		/// modify the workbook and will ask for saving changes when closing. Even if you just open and close the file.
		/// Note that this will only happen if there are formulas on the sheet.
		///<p>
		///<b>It is strongly advised to leave this property = true</b>. If you set it to false and there is an error on a formula
		///calculated by FlexCel, it will show wrong when open on Excel too. For example, if you have the formula:
		///<code>"=GETPIVOTDATA("a";"b")"</code> 
		///FlexCel will return #NAME!, as it doesn't implement this function. If you open this file on Excel and 
		///RecalcForced was false when saving, Excel will not calculate it and will show also #NAME!.
		///If RecalcForced was true, Excel will show the right answer, but will ask for saving the file when closing.
		///<b>Only set RecalcForced=false if you are in control of the formulas used on the spreadsheet</b>,
		///so you can guarantee no unsupported formula will be used. If the final user can modify the templates, do not set it.
		///</p>
		///<p>
		///Note that if <see cref="RecalcMode"/> = Smart and no modification is done to the file, Autorecalc info on formulas
		///will not change, even if RecalcForced=true.
		///</p>
		///<example>
		///The following code:
		///<code>
		///XlsFile xls= new XlsFile();
		///xls.NewFile();
		///xls.RecalcMode = TRecalcMode.Manual; //So the sheet is not recalculated before saving.
		///xls.RecalcForced = false; //So Excel won't recalculate either.
		///xls.SetCellValue(1,1, new TFormula("=1+1",3));  //Basic math...
		///xls.Save(OutFileName);
		///</code>
		///will create a file with a formula "1+1" and result = 3 on cell A1. If RecalcForced were true, Excel would show the correct value when opening.
		///</example>
		///</remarks>
		public abstract bool RecalcForced{get;set;}

		/// <summary>
		/// True if the file has been modified after loading.
		/// </summary>
		internal abstract bool NeedsRecalc {get;set;}

		/// <summary>
		/// Returns true if the workbook is being recalculated.
		/// </summary>
		internal abstract bool Recalculating { get; }

        internal abstract bool ReorderCalcChain(int SheetBase1, IFormulaRecord i1, IFormulaRecord i2);


		/// <summary>
		/// Returns a recalculating supporting file for this spreadsheet. Supporting files are added using a Workspace object.
		/// </summary>
		/// <param name="fileName">File to return. It might be a full filename with path or only the name of the file.</param>
		/// <returns>The supporting file with the given filename.</returns>
		internal abstract ExcelFile GetSupportingFile(string fileName);

		internal abstract void SetRecalculating(bool value);
		internal abstract void CleanFlags();
		#endregion

        #region What-If Tables

        /// <summary>
        /// Returns a list of the upper cells of the What-if tables in the page. You can then use <see cref="GetWhatIfTable"/> to get the definition of each one.
        /// </summary>
        /// <returns>A list of the coordinates with the first row and column for every what-if table in the sheet.</returns>
        public abstract TCellAddress[] GetWhatIfTableList();

        /// <summary>
        /// Returns the range of cells that make the what-if table that starts at aRow and aCol.
        /// If there is no What-if table at aRow, aCol, this method retuns null.
        /// <br></br>If both the returned rowInputCell and colInputCell are null, this means this table points to deleted references.
        /// </summary>
        /// <param name="sheet">Sheet where the table is.</param>
        /// <param name="row">First cell from where we want to get a what-if table.</param>
        /// <param name="col">First cell from where we want to get a what-if table.</param>
        /// <param name="rowInputCell">Returns the row input cell for this table. If the table doesn't have a row input cell, this value is null.</param>
        /// <param name="colInputCell">Returns the column input cell for this table. If the table doesn't have a column input cell, this value is null.</param>
        /// <returns>The full range of the table, not incuding the formula headers. Only the cells where {=Table()} formulas are..</returns>
        public abstract TXlsCellRange GetWhatIfTable(int sheet, int row, int col, out TCellAddress rowInputCell, out TCellAddress colInputCell);

        /// <summary>
        /// Creates an Excel What-if table in the range of cells specified by Range. Calling this method is tha same as setting a cell value
        /// with a TFormula where TFormula.Span has more than one cell, and TFormula.Text is something like "{=TABLE(,A4)}". The parameters
        /// for the =TABLE function are rowInputCell and colInputCell, and they look the same a Excel will show them.
        /// </summary>
        /// <param name="range">Range for the table. This is the range of cells that will have "={TABLE()}" formulas.</param>
        /// <param name="rowInputCell">Row input cell for the table. Make it null if you don't want a row input cell. If both rowInputCell and colInputCell are null, a table with deleted references will be added.<br></br>
        /// Note that the sheet here is ignored, What-if tables need the input cells to be in the same sheet as the table.</param>
        /// <param name="colInputCell">Column input cell for the table. Make it null if you don't want a column input cell.  If both rowInputCell and colInputCell are null, a table with deleted references will be added.<br></br>
        /// Note that the sheet here is ignored, What-if tables need the input cells to be in the same sheet as the table.</param>
        public abstract void SetWhatIfTable(TXlsCellRange range, TCellAddress rowInputCell, TCellAddress colInputCell);
        #endregion

        #region External Links
        /// <summary>
        /// Returns the number of external links for the file. You can access those links with <see cref="GetLink"/> and <see cref="SetLink"/> 
        /// </summary>
        public abstract int LinkCount { get; }

        /// <summary>
        /// Gets the external link at position i. 
        /// </summary>
        /// <param name="index">Index of the link (1 based). i goes between 1 and <see cref="LinkCount"/> </param>
        /// <returns></returns>
        public abstract string GetLink(int index);

        /// <summary>
        /// Changes the external link at position i for a new value. Note that you can't add new links with this method, external links
        /// are added automatically when you add formulas that reference other worksheets. This method is only to change existing links to point to
        /// other place. All formulas pointing to the old link will point to the new.<br></br>
        /// Note that the replacing filename should have the same sheets as the original, or the formulas might break.
        /// </summary>
        /// <param name="index">Index of the link (1 based). i goes between 1 and <see cref="LinkCount"/> </param>
        /// <param name="value">Please make sure this is a VALID filename, or you are likely to crash Excel. Also, xls file format doesn't like
        /// paths starting with "..", so you might need to enter the full path here.</param>
        /// <returns></returns>
        public abstract void SetLink(int index, string value);

        #endregion

        #region Misc
        /// <summary>
		/// When this property is false, inserting and copying ranges will behave the same as it does in Excel.
		/// When this property is true, absolute references to cells inside the block being copied will be treated as relative.
		/// For example, if you have:
		/// <code>
		/// A1: 2
		/// B1: =$A$1 + $A$57
		/// and you copy the row 1 to row 2, in Excel or FlexCel when this property is false you will get:
		/// A2: 2
		/// B2: =$A$1 + $A$57
		/// When this property is true, you will get:
		/// A2: 2
		/// B2: =$A$2 + $A$57
		/// </code>
		/// In the second case, the first reference was updated because it was inside the range being copied, but the second was not.
		/// This property might be useful when you want to duplicate blocks of cells, but want the absolute references inside it to point to the newer block.
		/// </summary>
        public bool SemiAbsoluteReferences { get { return FSemiAbsoluteReferences; } set { FSemiAbsoluteReferences = value; } }

        /// <summary>
        /// Specifies which reference style to use when entering formulas: A1 or R1C1. Note that this property is different from
        /// <see cref="OptionsR1C1"/>. OptionsR1C1 modifies a property of the file, that handles how references will show in Excel.
        /// <br></br> This property modifies how FlexCel parses or returns the formulas, and has no effect at all in the file generated.
        /// <br></br><br></br>Also note that R1C1 and A1 modes are completely equivalent, and formulas are <b>always stored as A1</b> inside the
        /// generated files. This property only affects the parsing of the formulas, the file generated will be exactly the same no
        /// matter the value of this property. And Excel will show it in A1 or R1C1 mode depsnding only in <see cref="OptionsR1C1"/>
        /// </summary>
        public TReferenceStyle FormulaReferenceStyle { get { return FFormulaReferenceStyle; } set { FFormulaReferenceStyle = value; } }

        /// <summary>
        /// Determines if FlexCel will throw Exceptions or just ignore errors on specific situations.
        /// </summary>
        public TExcelFileErrorActions ErrorActions { get { return FErrorActions; } set { FErrorActions = value; } }

        /// <summary>
        /// Returns the file format that the file had when it was opened. If the file was created with <see cref="NewFile()"/>, the
        /// file format when opened is xls.
        /// </summary>
        public abstract TFileFormats FileFormatWhenOpened { get; }


		/// <summary>
		/// Use it to convert formulas to their values. It can be useful if for example you are copying the sheet to
		/// another workbook, and you don't want any references to it. NOTE: You will probably want to use <see cref="ConvertExternalNamesToRefErrors"/> too, to convert named ranges besides the formulas.<br/>
        /// Also note that if you want to convert a whole file, you need to call ConvertFormulasToValues in every sheet.
		/// </summary>
		/// <param name="onlyExternal">When true, it will only convert the formulas that do not refer to the same sheet.
		/// For example "=A1+Sheet2!A1" will be converted, but "=A2+A3" will not.
		/// </param>
		public abstract void ConvertFormulasToValues(bool onlyExternal);

		/// <summary>
		/// Use it to convert the external names in a sheet to #REF! . It can be useful when you need to remove all external links in a file. 
		/// NOTE: You will probably want to use <see cref="ConvertFormulasToValues"/> too.
		/// </summary>
		public abstract void ConvertExternalNamesToRefErrors();


		/// <summary>
		/// Factor to multiply default column widths. See remarks for a detailed explanation.
		/// </summary>
		/// <remarks>
		/// Excel does not print the same things to all printers (or even the same printer with another resolution).
		/// Depending on the driver and resolution, printed columns and rows will be a little smaller or bigger. 
		/// <p></p>FlexCel uses .NET framework to print, and so it is resolution independent, it will always
		/// print the same no matter where. So FlexCel is tuned to print like Excel on a 600 dpi generic printer.
		/// This will probably be ok for most uses, but in some cases might cause FlexCel to print an Extra
		/// page with only one column or row.
		/// <p></p>You can use this property and <see cref="HeightCorrection"/> to manually fine tune FlexCel printing,
		/// previewing and exporting to pdf to make it smaller or larger. A normal value for this property might be 1.05 or 0.99,
		/// but don't use it unless you really need to.
		/// </remarks>
		/// <example>
		/// To calculate the exact HeightCorrection and WidthCorrection for your printer create a new Excel sheet,
		/// make column A wide (almost a sheet wide) and row 1 larger (almost a sheet tall). 
		/// Set the borders around cell A1 to a solid line.
		/// Now, print this file from Excel and from FlexCel (you can use the PrintPreview demo for this) and compare
		/// the 2 resulting boxes. If for example the Excel printed box is 1.2 cm wide and Flexcel is 1.4, 
		/// WidthCorrection should be 1.4f/1.2f.  (A larger WidthCorrection means a smaller box).
		/// </example>
		public abstract real WidthCorrection{get;set;}

		/// <summary>
		/// Factor to multiply default row heights. See remarks for a detailed explanation.
		/// </summary>
		/// <remarks>
		/// Excel does not print the same things to all printers (or even the same printer with another resolution).
		/// Depending on the driver and resolution, printed columns and rows will be a little smaller or bigger. 
		/// <p></p>FlexCel uses .NET framework to print, and so it is resolution independent, it will always
		/// print the same no matter where. So FlexCel is tuned to print like Excel on a 600 dpi generic printer.
		/// This will probably be ok for most uses, but in some cases might cause FlexCel to print an Extra
		/// page with only one column or row.
		/// <p></p>You can use this property and <see cref="WidthCorrection"/> to manually fine tune FlexCel printing,
		/// previewing and exporting to pdf to make it smaller or larger. A normal value for this property might be 1.05 or 0.99,
		/// but don't use it unless you really need to.
		/// </remarks>
		/// <example>
		/// To calculate the exact HeightCorrection and WidthCorrection for your printer create a new Excel sheet,
		/// make column A wide (almost a sheet wide) and row 1 larger (almost a sheet tall). 
		/// Set the borders around cell A1 to a solid line.
		/// Now, print this file from Excel and from FlexCel (you can use the PrintPreview demo for this) and compare
		/// the 2 resulting boxes. If for example the Excel printed box is 1.2 cm wide and Flexcel is 1.4, 
		/// WidthCorrection should be 1.4f/1.2f.  (A larger WidthCorrection means a smaller box)
		/// </example>
		public abstract real HeightCorrection{get;set;}

        /// <summary>
        /// A Linespacing of 1 means use the standard GDI+ linespace when a cell has more than one line. A linespace of 2 would mean
        /// double linespacing, and 0.5 would mean half linespacing. Normally linespacing in Excel is a little bigger than linespacing in GDI+,
        /// so you can use this property to fine tune what you need.<br/> 
        /// <b>This property doesn't alter the Excel file in any way. It is only used when rendering.</b> 
        /// </summary>
        public abstract double Linespacing { get; set; }

        /// <summary>
        /// This is an optimization property. If you set it to true, methods like GetCellValue or GetNamedRange won't return the
        /// formula text, just the formula results. If you don't care about formula texts, setting this property to true can speed up
        /// the processing of huge files.
        /// </summary>
        public abstract bool IgnoreFormulaText { get; set; }


		#endregion

        #region Custom Formulas

        /// <summary>
        /// Adds a custom formula function to the FlexCel recalculation engine. Note that this formulas are only valid for Excel custom formulas, not for internal ones.
        /// For example, you could define "EDATE" since it is a custom formula defined in the Analisis Addin, but you cannot redefine "SUM". 
        /// Note that if a custom formula with the name already exists, it will be replaced. Names are Case insensitive ("Date" is the same as "DATE").
        /// <br></br>Also note that some user defined functions come already built in in FlexCel, so you might not need to define them.
        /// For more information on adding Custom Formulas make sure you read the PDF documentation and take a look at the demo.
        /// </summary>
        /// <param name="scope">Defines if the custom function will be available globally to all ExcelFile instances or only to the ExcelFile instance where
        /// it was added. It is recommended to add functions globally, unless you have different xls files with functions that might have the same name but could be implemented different.</param>
        /// <param name="location">Defines if the function will be inserted as a reference to a macro in the local sheet or in an external book or addin.
        /// This parameter is used only when adding formulas with user defined functions to a sheet. It is not needed or used when recalculating those functions or when
        /// reading the text of a formula.</param>
        /// <param name="userFunction">Formula function we want to add.</param>
        public abstract void AddUserDefinedFunction(TUserDefinedFunctionScope scope, TUserDefinedFunctionLocation location, TUserDefinedFunction userFunction);

        /// <summary>
        /// Evaluates a custom function you have added earlier with <see cref="AddUserDefinedFunction"/>. You will not normally need to call this method, but it could be used for testing.
        /// If the function has not been added with <see cref="AddUserDefinedFunction"/>, this method will return <see cref="TFlxFormulaErrorValue.ErrName"/>.
        /// </summary>
        /// <param name="functionName">Function you want to evaluate.</param>
        /// <param name="arguments">Extra arguments you can use to evaluate the formula.</param>
        /// <param name="parameters">Parameters for the formula.</param>
        /// <returns>The result of evaluating the formula. It might be a string, a double, a boolean, a TFlxFormulaError or an Array.</returns>
        public abstract object EvaluateUserDefinedFunction(string functionName, TUdfEventArgs arguments, object[] parameters);

        /// <summary>
        /// Removes all the custom formula functions from the FlexCel recalculation engine.
        /// </summary>
        /// <param name="scope"></param>
        public abstract void ClearUserDefinedFunctions(TUserDefinedFunctionScope scope);

		/// <summary>
		/// Returns true if the Custom formula function has been added to the FlexCel recalculating engine.
        /// Note that internal functions are not returned by this method, but user defined functions pre-defined in FlexCel will be.
        /// </summary>
		/// <param name="functionName">Name of the function. Case insensitive.</param>
		/// <returns>True if the name has been added, false if not.</returns>
        public bool IsDefinedFunction(string functionName)
        {
            TUserDefinedFunctionLocation location;
            return IsDefinedFunction(functionName, out location);
        }

        /// <summary>
        /// Returns true if the Custom formula function has been added to the FlexCel recalculating engine.
        /// Note that internal functions are not returned by this method, but user defined functions pre-defined in FlexCel will be.
        /// </summary>
        /// <param name="functionName">Name of the function. Case insensitive.</param>
        /// <param name="location">Returns if the function is defined as an internal or external function.</param>
        /// <returns>True if the name has been added, false if not.</returns>
        public abstract bool IsDefinedFunction(string functionName, out TUserDefinedFunctionLocation location);

        internal abstract TUserDefinedFunctionContainer GetUserDefinedFunction(string functionName);

        internal abstract TUserDefinedFunctionContainer GetUserDefinedFunctionFromDisplayName(string internalFunctionName);


		/// <summary>
		/// Creates and addin external name or returns an existing one.
		/// </summary>
		/// <param name="functionName"></param>
		/// <param name="externSheet"></param>
		/// <param name="externName"></param>
        internal abstract void EnsureAddInExternalName(string functionName, out int externSheet, out int externName);
        
        internal abstract void EnsureAddInInternalName(string functionName, bool AddErrorDataToFormula, out int nameIndex);

        internal abstract int EnsureExternName(int ExternSheet, string Name);

        internal abstract void AddUnsupported(TUnsupportedFormulaErrorType ErrorType, string FuncName);

        internal abstract void SetUnsupportedFormulaList(TUnsupportedFormulaList Ufl);

        internal abstract void SetUnsupportedFormulaCellAddress(TCellAddress aCellAddress);

        /// <summary>
        /// Returns the externsheetindex for a record.
        /// </summary>
        internal abstract int GetExternSheet(string ExternSheet, bool ReadingXlsx);

        internal abstract int GetExternSheet(string ExternSheet, bool IsCellReference, bool ReadingXlsx, out bool IsLocal, out int Sheet1);


        #endregion

        #region Macros
        /// <summary>
        /// If the file has macros, this method will remove them.
        /// </summary>
        public abstract void RemoveMacros();

        /// <summary>
        /// Returns true if the file has any macros.
        /// </summary>
        /// <returns>True if the file has macros.</returns>
        public abstract bool HasMacros();

        internal abstract byte[] GetMacroData();
        internal abstract bool HasMacroXlsm();
        internal abstract void SetMacrodata(byte[] MacroData);

        #endregion

        #region IEnumerable Members

        /// <summary>
        /// Returns an enumerator that allows you to loop on all cells in the active worksheet. Once you start the foreach loop, you might cahnge the active sheet and it
        /// won't change inside the enumerator.
        /// </summary>
        /// <returns></returns>
        public IEnumerator<CellValue> GetEnumerator()
        {
            int RCount = RowCount;
            int ASheet = ActiveSheet; //it might change, so we want to make the calls here not depend on ActiveSheet.
            for (int r = 1; r <= RCount; r++)
            {
                int CCount = ColCountInRow(ASheet, r);
                for (int cIndex = 1; cIndex <= CCount; cIndex++)
                {
                    int Col = ColFromIndex(ASheet, r, cIndex);

                    int XF = 0;
                    object val = GetCellValueIndexed(ASheet, r, cIndex, ref XF);

                    yield return new CellValue(ASheet, r, Col, val, XF);
                }
            }
        }



        /// <summary>
        /// Returns an enumerator that allows you to loop on all cells in the active worksheet. Once you start the foreach loop, you might cahnge the active sheet and it
        /// won't change inside the enumerator.
        /// </summary>
        /// <returns></returns>
        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        #endregion

    }
}
