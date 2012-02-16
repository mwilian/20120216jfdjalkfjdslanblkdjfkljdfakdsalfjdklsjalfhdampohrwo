using System;

namespace FlexCel.Core
{
    #region Misc
    /// <summary>
	/// Possible image types on an excel sheet.
	/// </summary>
	public enum TXlsImgType
	{
		/// <summary>
		/// Enhanced Windows Metafile. This is a Vectorial image format.
		/// </summary>
		Emf, 
		
        /// <summary>
		/// Windows Metafile. This is a Vectorial image format.
		/// </summary>
		Wmf, 
		
        /// <summary>
		/// JPEG Image. This is a losely compressed bitmap, best suited for photos.
		/// </summary>
		Jpeg, 
		
        /// <summary>
		/// Portable Network Graphics. This is a lossless compressed bitmap, best suited for text.
		/// </summary>
		Png, 
		
        /// <summary>
		/// Windows Bitmap. As this is not compressed, don't use it except for really small images.
		/// </summary>
		Bmp, 

        /// <summary>
        /// Macintosh PICT. This won't be rendered by FlexCel, and you shouldn't add images in this format.
        /// </summary>
        Pict,

        /// <summary>
        /// Tiff Image: http://partners.adobe.com/public/developer/tiff/index.html#spec 
        /// This is NOT supported in xls file format, only xlsx.
        /// </summary>
        Tiff,

        /// <summary>
        /// Gif Image.
        /// This is NOT supported in xls file format, only xlsx.
        /// </summary>
        Gif,

		/// <summary>
		/// Unsupported image format.
		/// </summary>
		Unknown
	}

	/// <summary>
	/// General options for a sheet. In Excel, this settings are located in Tools->Options->View in the "Window options" box.
	/// Most of this options can be set with dedicated methods in <see cref="ExcelFile"/>, but this type allows to set them all at once, 
	/// or copy them from other file. This options apply only to the active sheet. For options that apply to all the sheets see <see cref="TSheetWindowOptions"/>
	/// </summary>
	[Flags]
	public enum TSheetOptions
	{
		/// <summary>
		/// No option is selected.
		/// </summary>
		None = 0x0000,

		/// <summary>
		/// If 1, Excel will show formula text instead of formula results. This is the same as calling <see cref="ExcelFile.ShowFormulaText"/>
		/// </summary>
		ShowFormulaText = 0x0001,

		/// <summary>
		/// If 0 Excel will hide the gridlines, if 1 it will show them. This is the same as calling <see cref="ExcelFile.ShowGridLines"/>
		/// </summary>
		ShowGridLines = 0x0002,

		/// <summary>
		/// if 0 Excel will hide the "A", "B"... column headers and "1", "2" ... row Headers.
		/// </summary>
		ShowRowAndColumnHeaders = 0x0004,

		/// <summary>
		/// If 1 Excel will hide 0 values. This is the same as calling <see cref="ExcelFile.HideZeroValues"/>
		/// </summary>
		ZeroValues = 0x0010,

		/// <summary>
		/// If 1, the color of gridlines will be gray. If 0 it will be the color speciifed at <see cref="ExcelFile.GridLinesColor"/>.
		/// Note that calling <see cref="ExcelFile.GridLinesColor"/> will automatically set this property to 0.
		/// </summary>
		AutomaticGridLineColors = 0x0020,

        /// <summary>
        /// If 1, the sheet is Right to left.
        /// </summary>
        RightToLeft = 0x0040,

		/// <summary>
		/// If 1, Excel will show the outline symbols (+/-) when there are gouped rows or columns.
		/// </summary>
		OutlineSymbols = 0x0080,

		/// <summary>
		/// If 1 Excel will show the page breaks.
		/// </summary>
		PageBreakView = 0x0800
	}

	/// <summary>
	/// General options for how sheets display. In Excel, this settings are located in Tools->Options->View in the "Window options" box.
    /// This options apply only to all the sheets. For options that apply only to the active sheet see <see cref="TSheetOptions"/>
	/// </summary>
	[Flags]
	public enum TSheetWindowOptions
	{
		/// <summary>
		/// No option is selected.
		/// </summary>
		None = 0x0000,

		/// <summary>
		/// If true, the workbook window will be hidden. You will need to go to Window->Unhide to see it.
		/// </summary>
		HideWindow = 0x0001,

		/// <summary>
		/// If true, the workbok window will be minimized.
		/// </summary>
		MinimizeWindow = 0x0002,

		/// <summary>
		/// If false, the horizontal scroll bar will be hidden.
		/// </summary>
		ShowHorizontalScrollBar = 0x0008,

		/// <summary>
		/// If false, the vartical scroll bar will be hidden.
		/// </summary>
		ShowVerticalScrollBar = 0x0010,

		/// <summary>
		/// If false, the bar with sheet names at the bottom will be hidden.
		/// </summary>
		ShowSheetTabBar = 0x0020
    }

    /// <summary>
    /// Specifies how a merged cell will be autofitted. For example, if you have a merged cell from row 1 to 4,
    /// You might want to increase the size of the first row, the second, the last, or every row a little.
    /// </summary>
    public enum TAutofitMerged
    {
        /// <summary>
        /// Merged cells with more than one row will not be autofitted when autofitting rows, and merged
        /// cells with more than one column will not be autofitted when autofitting rows.
        /// </summary>
        None = 0,

        /// <summary>
        /// Autofit will change the size of the last row of the merged cell when autofitting rows, or the
        /// last column when autofitting columns.
        /// </summary>
        OnLastCell = 1,

        /// <summary>
        /// Autofit will change the size of the row before the last row of the merged cell when autofitting rows, or the
        /// column before the last column when autofitting columns.
        /// </summary>
        OnLastCellMinusOne = 2,

        /// <summary>
        /// Autofit will change the size of 2 rows before the last row of the merged cell when autofitting rows, or
        /// 2 columns before the last column when autofitting columns.
        /// </summary>
        OnLastCellMinusTwo = 3,

        /// <summary>
        /// Autofit will change the size of 3 rows before the last row of the merged cell when autofitting rows, or
        /// 3 columns before the last column when autofitting columns.
        /// </summary>
        OnLastCellMinusThree = 4,

        /// <summary>
        /// Autofit will change the size of 4 rows before the last row of the merged cell when autofitting rows, or
        /// 4 columns before the last column when autofitting columns.
        /// </summary>
        OnLastCellMinusFour = 5,

        /// <summary>
        /// Autofit will change the size of the first row of the merged cell when autofitting rows, or the
        /// first column when autofitting columns.
        /// </summary>
        OnFirstCell = 6,

        /// <summary>
        /// Autofit will change the size of the second row of the merged cell when autofitting rows, or the
        /// second column when autofitting columns.
        /// </summary>
        OnSecondCell = 7,

        /// <summary>
        /// Autofit will change the size of the third row of the merged cell when autofitting rows, or the
        /// third column when autofitting columns.
        /// </summary>
        OnThirdCell = 8,

        /// <summary>
        /// Autofit will change the size of the fourth row of the merged cell when autofitting rows, or the
        /// fourth column when autofitting columns.
        /// </summary>
        OnFourthCell = 9,

        /// <summary>
        /// Autofit will change the size of the fifth row of the merged cell when autofitting rows, or the
        /// fifth column when autofitting columns.
        /// </summary>
        OnFifthCell = 10,

        /// <summary>
        /// Autofit will change the height every row in the merged cell by the same amount.
        /// </summary>
        Balanced = 11
    }

    /// <summary>
    /// Position on a pane when window is split
    /// </summary>
    public enum TPanePosition
    {
        /// <summary>
        /// Lower-right corner.
        /// </summary>
        LowerRight = 0,

        /// <summary>
        /// Upper-right corner.
        /// </summary>
        UpperRight = 1,

        /// <summary>
        /// Lower-left corner.
        /// </summary>
        LowerLeft = 2,

        /// <summary>
        /// Upper left corner. This is the default when you have only one pane.
        /// </summary>
        UpperLeft = 3
    }

	/// <summary>
	/// How an image behaves when inserting/copying rows/columns
	/// </summary>
	public enum TFlxAnchorType
	{
		/// <summary>
		/// Move and resize the image with the sheet.
		/// </summary>
		MoveAndResize = 0, 

		/// <summary>
		/// Move the image when inserting/copying, but keep its size.
		/// </summary>
		MoveAndDontResize = 2, 

		/// <summary>
		/// Keep the image fixed on the sheet.
		/// </summary>
		DontMoveAndDontResize = 3
	}

    /// <summary>
    /// Text direction for objects like comments, that allow rotation, but only in 90 degrees.
    /// </summary>
    public enum TTextRotation
    {
        /// <summary>
        /// Text is not rotated.
        /// </summary>
        Normal,

        /// <summary>
        /// Text is rotated 90 degrees counterclockwise.
        /// </summary>
        Rotated90Degrees,

        /// <summary>
        /// Text is rotated 90 degrees clockwise.
        /// </summary>
        RotatedMinus90Degrees,

        /// <summary>
        /// Text is written in vertical layout, a character below the other.
        /// </summary>
        Vertical
    }

	/// <summary>
	/// Sheet visibility.
	/// </summary>
	public enum TXlsSheetVisible
	{
		/// <summary>Sheet is hidden, can be shown by the user with excel.</summary>
		Hidden,
		/// <summary>Sheet is hidden, only way to show it is with a macro. (user can't see it with excel)</summary>
		VeryHidden,
		/// <summary>Sheet is visible to the user.</summary>
		Visible
	}

    /// <summary>
    /// Parameter names that can go into an "invalid params" error message.
    /// </summary>
    public enum FlxParam
    {
        /// <summary>ActiveSheet</summary>
        ActiveSheet,
        ///<summary>SheetFrom</summary>
        SheetFrom,
        ///<summary>SheetDest</summary>
        SheetDest,
        ///<summary>SheetCount</summary>
        SheetCount,
        ///<summary>CellMergedIndex</summary>
        CellMergedIndex,
        ///<summary>NamedRangeIndex</summary>
        NamedRangeIndex,
        ///<summary>ImageIndex</summary>
        ImageIndex,
        ///<summary>ObjectIndex</summary>
        ObjectIndex,
        ///<summary>CommentIndex</summary>
        CommentIndex,
        ///<summary>SourceSheet</summary>
        SourceSheet,
        ///<summary>HyperLinkIndex</summary>
        HyperLinkIndex,
        /// <summary>OutlineLevel</summary>
        OutlineLevel,
        /// <summary>IgnoreCase</summary>
        IgnoreCase,
        /// <summary>AutoShapeIndex</summary>
        AutoShapeIndex,
		///<summary>ChartIndex</summary>
		ChartIndex,
		///<summary>SeriesIndex</summary>
		SeriesIndex,
		///<summary>ConditionalFormatIndex</summary>
		ConditionalFormatIndex,
        ///<summary>PercentOfUsedSheet</summary>
        PercentOfUsedSheet,
        ///<summary>Level</summary>
        Level,
        ///<summary>PageScale</summary>
        PageScale,
		///<summary>DataValidationIndex</summary>
		DataValidationIndex,
		///<summary>FormatIndex</summary>
		FormatIndex,
        ///<summary>NumberOfThreads</summary>
        NumberOfThreads,
        ///<summary>Index</summary>
        Index,
        ///<summary>FirstSheetVisible</summary>
        FirstSheetVisible
}

    /// <summary>
    /// Sheet types
    /// </summary>
    public enum TSheetType
    {
        /// <summary>
        /// Normal WorkSheet.
        /// </summary>
        Worksheet,
        /// <summary>
        /// Chart Sheet
        /// </summary>
        Chart,

        /// <summary>
        /// An Excel 5.0 dialog sheet.
        /// </summary>
        Dialog,

        /// <summary>
        /// An Excel 4.0 Macro sheet.
        /// </summary>
        Macro,

        /// <summary>
        /// Something we don't support. It shouldn't happen.
        /// </summary>
        Other
    }

    /// <summary>
    /// A list of records that might not be saved into a file when using <see cref="ExcelFile.SaveForHashing(System.IO.Stream)"/>
    /// </summary>
    [Flags]
    public enum TExcludedRecords
    {
        /// <summary>
        /// This includes all records, including the write access. It is <b>not</b> recommended to use this setting, because
        /// write access will change every time you save a file with a different user.
        /// </summary>
        None = 0,

        /// <summary>
        /// This is a stamp that is saved to the file each time it is saved, identifying the current user. You normally won't want to save this record, so you should specify this value.
        /// </summary>
        WriteAccess = 1,

        /// <summary>
        /// If you specify CellSelection, the selected cells will be ignored.
        /// </summary>
        CellSelection = 2,

        /// <summary>
        /// If you specify SheetSelected, the active sheet will be ignored.
        /// </summary>
        SheetSelected = 4,

        /// <summary>
        /// If you specify Version, the version of Excel used to save the file will be ignored. (Note that version might change in an Excel Service Pack).
        /// </summary>
        Version = 8,

        /// <summary>
        /// This excludes all records in this enumeration. This is the most recomended option.
        /// </summary>
        All = WriteAccess | CellSelection | SheetSelected | Version
    }

    /// <summary>
    /// Handles how to convert a column from text when importing a text file.
    /// </summary>
    public enum ColumnImportType
    {
        /// <summary>
        /// Try to convert it to a number, a date, etc.
        /// </summary>
        General,

		/// <summary>
        /// Keep the column as text, even if it can be converted to a number or other things.
        /// </summary>
        Text,
        
		/// <summary>
        /// Do not import this column.
        /// </summary>
        Skip
    }

    /// <summary>
    /// List of internal range names.
    /// On Excel, internal range names like "Print_Area" are stored as a 1 character string.
    /// This is the list of the names and their value.
    /// You can convert an InternalNameRange into a string by casting it to a char, or by calling <see cref="TXlsNamedRange.GetInternalName(InternalNameRange)"/><br>See the example.</br>
    /// </summary>
    /// <example>
    /// To get the print range on the ActiveSheet, use:
    /// <code>
    /// xlsFile.GetNamedRange(TXlsNamedRange.GetInternalName(InternalNameRange.Print_Area), xlsFile.ActiveSheet);
    /// </code>
    /// </example>
    public enum InternalNameRange
    {
        ///<summary>Consolidate_Area  </summary>
        Consolidate_Area  = 0x00,
 
        ///<summary>Auto_Open         </summary>
        Auto_Open         = 0x01,
 
        ///<summary>Auto_Close        </summary>
        Auto_Close        = 0x02,
 
        ///<summary>Extract           </summary>
        Extract           = 0x03,
 
        ///<summary>Database          </summary>
        Database          = 0x04,
 
        ///<summary>Criteria          </summary>
        Criteria          = 0x05,
 
        ///<summary>Print_Area        </summary>
        Print_Area        = 0x06,
 
        ///<summary>Print_Titles      </summary>
        Print_Titles      = 0x07,
 
        ///<summary>Recorder          </summary>
        Recorder          = 0x08,
 
        ///<summary>Data_Form         </summary>
        Data_Form         = 0x09,
 
        ///<summary>Auto_Activate     </summary>
        Auto_Activate     = 0x0A,
 
        ///<summary>Auto_Deactivate   </summary>
        Auto_Deactivate   = 0x0B,
 
        ///<summary>Sheet_Title       </summary>
        Sheet_Title       = 0x0C,

		///<summary>Used in AutoFilters. </summary>
		Filter_DataBase   = 0x0D
    }

    /// <summary>
    /// Determines how the children of the node of an outline will be when the node  is collapsed or expanded.
    /// </summary>
    public enum TCollapseChildrenMode
    {
        /// <summary>
        /// Children nodes will be kept in the state they already were. If they were collapsed they will still be collapsed after
        /// collpasing or expanding the parent. Same if they were expanded.
        /// </summary>
        DontModify,

        /// <summary>
        /// All children nodes of collapsed parents will be collapsed, and will show collapsed when you expand the parent.
        /// </summary>
        Collapsed,

        /// <summary>
        /// All children nodes of collapsed parents will be expanded, and will show expanded when you expand the parent.
        /// </summary>
        Expanded
    }

    /// <summary>
    /// The sort order for a sort operation.
    /// </summary>
    public enum TSortOrder
    {
        /// <summary>
        /// Sort ascending.
        /// </summary>
        Ascending,

        /// <summary>
        /// Sort descending.
        /// </summary>
        Descending
    }
    #endregion

    #region Print
    /// <summary>
    /// Pre-defined Page sizes. For Printer specific page-sizes, see <see cref="TPrinterDriverSettings"/>
    /// Note that a printer specific page size might have a value that is <i>not</i> on this enumeration.
    /// </summary>
    public enum TPaperSize
    {
        /// <summary>Not defined.</summary>
        Undefined=0,
        ///<summary>Letter - 81/2"" x 11""</summary>
        Letter=1,
        ///<summary>Letter small - 81/2"" x 11""</summary>
        Lettersmall=2,
        ///<summary>Tabloid - 11"" x 17""</summary>
        Tabloid=3,
        ///<summary>Ledger - 17"" x 11""</summary>
        Ledger=4,
        ///<summary>Legal - 81/2"" x 14""</summary>
        Legal=5,
        ///<summary>Statement - 51/2"" x 81/2""</summary>
        Statement=6,
        ///<summary>Executive - 71/4"" x 101/2""</summary>
        Executive=7,
        ///<summary>A3 - 297mm x 420mm</summary>
        A3=8,
        ///<summary>A4 - 210mm x 297mm</summary>
        A4=9,
        ///<summary>A4 small - 210mm x 297mm</summary>
        A4small=10,
        ///<summary>A5 - 148mm x 210mm</summary>
        A5=11,
        ///<summary>B4 (JIS) - 257mm x 364mm</summary>
        B4_JIS=12,
        ///<summary>B5 (JIS) - 182mm x 257mm</summary>
        B5_JIS=13,
        ///<summary>Folio - 81/2"" x 13""</summary>
        Folio=14,
        ///<summary>Quarto - 215mm x 275mm</summary>
        Quarto=15,
        ///<summary>10x14 - 10"" x 14""</summary>
        s10x14=16,
        ///<summary>11x17 - 11"" x 17""</summary>
        s11x17=17,
        ///<summary>Note - 81/2"" x 11""</summary>
        Note=18,
        ///<summary>Envelope #9 - 37/8"" x 87/8""</summary>
        Envelope9=19,
        ///<summary>Envelope #10 - 41/8"" x 91/2""</summary>
        Envelope10=20,
        ///<summary>Envelope #11 - 41/2"" x 103/8""</summary>
        Envelope11=21,
        ///<summary>Envelope #12 - 43/4"" x 11""</summary>
        Envelope12=22,
        ///<summary>Envelope #14 - 5"" x 111/2""</summary>
        Envelope14=23,
        ///<summary>C - 17"" x 22""</summary>
        C=24,
        ///<summary>D - 22"" x 34""</summary>
        D=25,
        ///<summary>E - 34"" x 44""</summary>
        E=26,
        ///<summary>Envelope DL - 110mm x 220mm</summary>
        EnvelopeDL=27,
        ///<summary>Envelope C5 - 162mm x 229mm</summary>
        EnvelopeC5=28,
        ///<summary>Envelope C3 - 324mm x 458mm</summary>
        EnvelopeC3=29,
        ///<summary>Envelope C4 - 229mm x 324mm</summary>
        EnvelopeC4=30,
        ///<summary>Envelope C6 - 114mm x 162mm</summary>
        EnvelopeC6=31,
        ///<summary>Envelope C6/C5 - 114mm x 229mm</summary>
        EnvelopeC6_C5=32,
        ///<summary>B4 (ISO) - 250mm x 353mm</summary>
        B4_ISO=33,
        ///<summary>B5 (ISO) - 176mm x 250mm</summary>
        B5_ISO=34,
        ///<summary>B6 (ISO) - 125mm x 176mm</summary>
        B6_ISO=35,
        ///<summary>Envelope Italy - 110mm x 230mm</summary>
        EnvelopeItaly=36,
        ///<summary>Envelope Monarch - 37/8"" x 71/2""</summary>
        EnvelopeMonarch=37,
        ///<summary>63/4 Envelope - 35/8"" x 61/2""</summary>
        s63_4Envelope=38,
        ///<summary>US Standard Fanfold - 147/8"" x 11""</summary>
        USStandardFanfold=39,
        ///<summary>German Std. Fanfold - 81/2"" x 12""</summary>
        GermanStdFanfold=40,
        ///<summary>German Legal Fanfold - 81/2"" x 13""</summary>
        GermanLegalFanfold=41,
        ///<summary>B4 (ISO) - 250mm x 353mm</summary>
        B4_ISO_2=42,
        ///<summary>Japanese Postcard - 100mm x 148mm</summary>
        JapanesePostcard=43,
        ///<summary>9x11 - 9"" x 11""</summary>
        s9x11=44,
        ///<summary>10x11 - 10"" x 11""</summary>
        s10x11=45,
        ///<summary>15x11 - 15"" x 11""</summary>
        s15x11=46,
        ///<summary>Envelope Invite - 220mm x 220mm</summary>
        EnvelopeInvite=47,
        ///<summary>Letter Extra - 91/2"" x 12""</summary>
        LetterExtra=50,
        ///<summary>Legal Extra - 91/2"" x 15""</summary>
        LegalExtra=51,
        ///<summary>Tabloid Extra - 1111/16"" x 18""</summary>
        TabloidExtra=52,
        ///<summary>A4 Extra - 235mm x 322mm</summary>
        A4Extra=53,
        ///<summary>Letter Transverse - 81/2"" x 11""</summary>
        LetterTransverse=54,
        ///<summary>A4 Transverse - 210mm x 297mm</summary>
        A4Transverse=55,
        ///<summary>Letter Extra Transv. - 91/2"" x 12""</summary>
        LetterExtraTransv=56,
        ///<summary>Super A/A4 - 227mm x 356mm</summary>
        SuperA_A4=57,
        ///<summary>Super B/A3 - 305mm x 487mm</summary>
        SuperB_A3=58,
        ///<summary>Letter Plus - 812"" x 1211/16""</summary>
        LetterPlus=59,
        ///<summary>A4 Plus - 210mm x 330mm</summary>
        A4Plus=60,
        ///<summary>A5 Transverse - 148mm x 210mm</summary>
        A5Transverse=61,
        ///<summary>B5 (JIS) Transverse - 182mm x 257mm</summary>
        B5_JIS_Transverse=62,
        ///<summary>A3 Extra - 322mm x 445mm</summary>
        A3Extra=63,
        ///<summary>A5 Extra - 174mm x 235mm</summary>
        A5Extra=64,
        ///<summary>B5 (ISO) Extra - 201mm x 276mm</summary>
        B5_ISO_Extra=65,
        ///<summary>A2 - 420mm x 594mm</summary>
        A2=66,
        ///<summary>A3 Transverse - 297mm x 420mm</summary>
        A3Transverse=67,
        ///<summary>A3 Extra Transverse - 322mm x 445mm</summary>
        A3ExtraTransverse=68,
        ///<summary>Dbl. Japanese Postcard - 200mm x 148mm</summary>
        DblJapanesePostcard=69,
        ///<summary>A6 - 105mm x 148mm</summary>
        A6=70,
        ///<summary>Letter Rotated - 11"" x 81/2""</summary>
        LetterRotated=75,
        ///<summary>A3 Rotated - 420mm x 297mm</summary>
        A3Rotated=76,
        ///<summary>A4 Rotated - 297mm x 210mm</summary>
        A4Rotated=77,
        ///<summary>A5 Rotated - 210mm x 148mm</summary>
        A5Rotated=78,
        ///<summary>B4 (JIS) Rotated - 364mm x 257mm</summary>
        B4_JIS_Rotated=79,
        ///<summary>B5 (JIS) Rotated - 257mm x 182mm</summary>
        B5_JIS_Rotated=80,
        ///<summary>Japanese Postcard Rot. - 148mm x 100mm</summary>
        JapanesePostcardRot=81,
        ///<summary>Dbl. Jap. Postcard Rot. - 148mm x 200mm</summary>
        DblJapPostcardRot=82,
        ///<summary>A6 Rotated - 148mm x 105mm</summary>
        A6Rotated=83,
        ///<summary>B6 (JIS) - 128mm x 182mm</summary>
        B6_JIS=88,
        ///<summary>B6 (JIS) Rotated - 182mm x 128mm</summary>
        B6_JIS_Rotated=89,
        ///<summary>12x11 - 12"" x 11""</summary>
        s12x11=90
    }

    /// <summary>
    /// How the sheet should be printed. You can mix value together by and'ing and or'ing the flags. 
    /// See the example to see how to set or clear one specific value of the enumeration.
    /// </summary>
    /// <example>
    /// Here we will show how to set the page orientation to landscape or portrait.
    /// <br/>
    /// <code>
    /// if (Landscape)  
    ///     Xls.PrintOptions &amp;= ~(TPrintOptions.Orientation | TPrintOptions.NoPls); 
    /// else
    /// {
    ///     //ALWAYS SET NOPLS TO 0 BEFORE CHANGING THE OTHER OPTIONS.
    ///     Xls.PrintOptions &amp;= ~ TPrintOptions.NoPls; 
    ///     Xls.PrintOptions |= TPrintOptions.Orientation; 
    /// }
    /// </code>
    /// <br/>
    /// <code lang = "vbnet">
    ///  If Landscape Then
    ///		Xls.PrintOptions = Xls.PrintOptions And (Not (TPrintOptions.Orientation Or TPrintOptions.NoPls))
    ///  Else
    ///		'ALWAYS SET NOPLS TO 0 BEFORE CHANGING THE OTHER OPTIONS.
    ///		Xls.PrintOptions = Xls.PrintOptions And (Not TPrintOptions.NoPls)
    ///		Xls.PrintOptions = Xls.PrintOptions Or (TPrintOptions.Orientation)
    ///  End If
    /// </code>
    /// <br/>
    /// <code lang = "Delphi .NET" title = "Delphi .NET">
    /// if Landscape then
    /// begin
    ///    Xls.PrintOptions := TPrintOptions(integer(Xls.PrintOptions) and (not (integer(TPrintOptions.Orientation) or integer(TPrintOptions.NoPls)));
    /// end
    /// else
    /// begin
    ///    //ALWAYS SET NOPLS TO 0 BEFORE CHANGING THE OTHER OPTIONS.
    ///    Xls.PrintOptions := TPrintOptions(integer(Xls.PrintOptions) and (not integer(TPrintOptions.NoPls)));
    ///    Xls.PrintOptions := TPrintOptions(integer(Xls.PrintOptions) or integer(TPrintOptions.Orientation));
    ///  end;
    /// </code>
    /// </example>
    [Flags]
    public enum TPrintOptions
    {
        /// <summary>
        /// All options cleared.
        /// </summary>
        None,

        /// <summary>
        /// Print over, then down
        /// </summary>
        LeftToRight = 0x01,

        /// <summary>
        /// 0= landscape, 1=portrait
        /// </summary>
        Orientation = 0x02,

        /// <summary>
        /// if 1, then PaperSize, Scale, Res, VRes, Copies, and Landscape data have not been obtained from the printer, so they are not valid.
        /// MAKE SURE YOU MAKE THIS BIT = 0 *BEFORE* CHANGING ANY OTHER OPTION. THEY WILL NOT CHANGE IF THIS IS NOT SET.
        /// </summary>
        NoPls = 0x04,

        /// <summary>
        /// 1= Black and white
        /// </summary>
        NoColor = 0x08,

        /// <summary>
        /// 1= Draft quality
        /// </summary>
        Draft = 0x10,

        /// <summary>
        /// 1= Print Notes
        /// </summary>
        Notes = 0x20,

        /// <summary>
        /// 1=orientation not set
        /// </summary>
        NoOrient = 0x40,

        /// <summary>
        /// 1=use custom starting page number.
        /// </summary>
        UsePage = 0x80
    }

    /// <summary>
    /// Enumeration defining which objects should not be printed or exported to pdf. You can 'or' more than one option together.
    /// For example, to not print images and not comments, you should specify: THidePrintObjects.Images | THidePrintOption.Comments
    /// </summary>
    [Flags]
    public enum THidePrintObjects
    {
        /// <summary>
        /// Print and Export everything.
        /// </summary>
        None = 0,

        /// <summary>
        /// Do not print or export images to pdf.
        /// </summary>
        Images = 1,

        /// <summary>
        /// Do not export Comments to pdf.
        /// </summary>
        Comments = 2,

        /// <summary>
        /// Do not export Hyperlinks to pdf.
        /// </summary>
        Hyperlynks = 4,

        /// <summary>
        /// Do not print or export to pdf the Headers.
        /// </summary>
        Headers = 8,

        /// <summary>
        /// Do not print or export to pdf the Footers.
        /// </summary>
        Footers = 16,

        /// <summary>
        /// Do not export headers nor footers. This is the same as specifying TPrintOptions.Headers and TPrintOptions.Footers.
        /// </summary>
        HeadersAndFooters = 16 | 8,

        /// <summary>
        /// Do not export page breaks. This property only applies to html exports.
        /// </summary>
        PageBreaks = 32
    }

    /// <summary>
    /// Different kinds of headers and footers depending on which pages they apply.
    /// </summary>
    public enum THeaderAndFooterKind
    {
        /// <summary>
        /// This applies to all headers and footers that don't have a dedicated definition.
        /// </summary>
        Default,

        /// <summary>
        /// This applies to the header and footer of the first pages, when it is defined different than the others.
        /// </summary>
        FirstPage,

        /// <summary>
        /// This applies to the headers and footes of even pages, when they are different than odd pages.
        /// </summary>
        EvenPages

    }

    /// <summary>
    /// Different sections on a header or footer.
    /// </summary>
    public enum THeaderAndFooterPos
    {
        /// <summary>
        /// Left section on the header.
        /// </summary>
        HeaderLeft = 0,

        /// <summary>
        /// Center section on the header.
        /// </summary>
        HeaderCenter = 1,

        /// <summary>
        /// Right section on the header.
        /// </summary>
        HeaderRight = 2,

        /// <summary>
        /// Left section on the footer.
        /// </summary>
        FooterLeft = 3,

        /// <summary>
        /// Center section on the footer.
        /// </summary>
        FooterCenter = 4,

        /// <summary>
        /// Right section on the footer.
        /// </summary>
        FooterRight = 5
    }

    #endregion

    #region Excel Version
    /// <summary>
    /// Defines which Excel version FlexCel is targeting. Note that while on v2007 you still can make xls 97 spreadsheets.
    /// </summary>
    public enum TExcelVersion
    {
        /// <summary>
        /// Versions from Excel 97 to Excel 2003. Those versions have a grid limited by 65536 rows x 255 columns.
        /// </summary>
        v97_2003 = 1,

        /// <summary>
        /// Excel 2007 and up. This version has a grid of 1048576 rows x 16384 columns.
        /// </summary>
        v2007 = 0  //0 so it is the default.
    }

    /// <summary>
    /// Different Excel versions create different empty xls/xlsx files. For example an empty xls file created by Excel 2003 will have "Arial"
    /// as its default font, and one created by Excel 2007 will have "Calibri". By default, when you call <see cref="ExcelFile.NewFile()"/> FlexCel will
    /// create a file that is similar to what Excel 2003 would create. This is fine, but if you want to start for example from an 
    /// empty Excel 2007 file, you can do so by calling NewFile() with this enumeration. 
    /// </summary>
    public enum TExcelFileFormat
    {
        /// <summary>
        /// Empty files will be created as if they were created by Excel 2003. The default font is Arial.
        /// </summary>
        v2003,

        /// <summary>
        /// Empty files will be created as if they were created by Excel 2007. The default font is Calibri.
        /// </summary>
        v2007,

        /// <summary>
        /// Empty files will be created as if they were created by Excel 2010. The default font is Calibri.
        /// </summary>
        v2010
    }

    /// <summary>
    /// The specific Excel version that FlexCel will emulate when reading and saving files.
    /// </summary>
    public enum TXlsBiffVersion
    {
        /// <summary>
        /// FlexCel will identify itself as Excel 2007 in the generated xls files.
        /// </summary>
        Excel2007,

        /// <summary>
        /// FlexCel will identify itself as Excel 2003 in the generated xls files. It will also ignore the
        /// extra information written by Excel 2007 when reading xls files created with it.
        /// </summary>
        Excel2003
    }

    /// <summary>
    /// Supported file formats to read and write files.
    /// </summary>
    public enum TFileFormats
    {
        /// <summary>
        /// Automatically detect the type of the file when opening files. If used when saving, FlexCel will choose whether to use xls or xlsx depending on the
        /// file extension (when saving to a file) or the value of <see cref="ExcelFile.DefaultFileFormat"/> when saving to a stream or when the format can't be
        /// determined from the extension.
        /// </summary>
        Automatic = 0,

        /// <summary>
        /// Excel 97-2000-XP-2003
        /// </summary>
        Xls = 1,

        /// <summary>
        /// Delimiter separated values. Depending on the delimiter, this can be csv, tab delimited text, etc.
        /// </summary>
        Text = 2,

        /// <summary>
        /// Pocket Excel 1.0 or 2.0
        /// </summary>
        Pxl = 3,

        /// <summary>
        /// Excel 2007 standard file format. Note that this is *not* a macro enabled file format. If you want to save a file with macros,
        /// you need to use Xlsm instead. 
        /// </summary>
        Xlsx = 4,

        /// <summary>
        /// Excel 2007 macro enabled file format. 
        /// </summary>
        Xlsm = 5
    }

    #endregion

    #region Cells
    /// <summary>
    /// Values we can write on a cell
    /// </summary>
    public enum TCellType
    {
        /// <summary>
        /// Double precision number.
        /// </summary>
        Number,

        /// <summary>
        /// Not a real type (it is a number).
        /// </summary>
        DateTime,

        /// <summary>
        /// An unicode string. Might be formatted.
        /// </summary>
        String,

        /// <summary>
        /// Boolean.
        /// </summary>
        Bool,

        /// <summary>
        /// Error code 
        /// </summary>
        Error,

        /// <summary>
        /// Blank cell.
        /// </summary>
        Empty,

        /// <summary>
        /// A formula.
        /// </summary>
        Formula,

        /// <summary>
        /// The .NET type does not map to any Excel type. (for example a class)
        /// </summary>
        Unknown

    }

    /// <summary>
    /// Sets how the excel file will be recalculated. Normally FlexCel calculates a file only before saving and when you explicitly call
    /// <see cref="ExcelFile.Recalc(bool)"/>. With this enum you can change that behavior.
    /// </summary>
    public enum TRecalcMode
    {
        /// <summary>
        /// The file will be recalculated before saving, *only* if there are any changes on the file. This mode
        /// is the recommended, because it allows you to open and save a file (for example to remove a password)
        /// without modifying the formula results. If you modify any value on the sheet, a recalculation will be 
        /// done to get the new values.
        /// </summary>
        Smart,

        /// <summary>
        /// The file will *always* be recalculated before saving. Use it if you are loading files that might not be
        /// recalculated, to make sure they will. If you open a not calculated file 
        /// on Smart mode and then save it without modifying anything, it will remain not calculated.
        /// </summary>
        Forced,

        /// <summary>
        /// The file will *never* be recalculated by FlexCel (Except if you do a forced recalc: XlsFile.Recalc(true). All formula results will be set to null and formulas will be
        /// modified so Excel recalculates them when you open the file. If you are only using Excel to open the generated files
        /// the result will be the same as with a recalculated file, but if you try to open it with a viewer
        /// you won't see the formula results.
        /// You should not need to use manual recalculation. The only case where it might be useful is if you are setting the formula results
        /// yourself and don't want FlexCel to change them when it saves.
        /// </summary>
        /// <remarks>Note that this mode is is not related with "Manual Recalculation" in Excel. FlexCel always does "Manual" recalculation 
        /// except when you use <see cref="TRecalcMode.OnEveryChange"/>. It will only recalc when saving or explicitly calling Recalc(). In this
        /// "Manual" mode it will not even recalculate when you save or do a not-forced recalc.</remarks>
        Manual,

        /// <summary>
        /// The file will be recalculated each time a value changes on the sheet. Do not use this mode on normal files,
        /// as it can be really slow!!  This option could be of use for an interactive viewer.
        /// </summary>
        OnEveryChange
    }

    /// <summary>
    /// How the file will be calculated by Excel. This enum doesn't affect FlexCel recalculation.
    /// </summary>
    public enum TSheetCalcMode
    {
        /// <summary>
        /// Manual recalculation.
        /// </summary>
        Manual = 0,

        /// <summary>
        /// Automatic recalculation.
        /// </summary>
        Automatic = 1,

        /// <summary>
        /// Automatic recalculation without tables.
        /// </summary>
        AutomaticExceptTables = 2
    }


    /// <summary>
    /// Error codes for cells on excel
    /// </summary>
    public enum TFlxFormulaErrorValue
    {
        /// <summary>Null Value</summary>
        ErrNull = 0x00,
        /// <summary>Division by 0</summary>
        ErrDiv0 = 0x07,
        /// <summary>Invalid Value</summary>
        ErrValue = 0x0F,
        /// <summary>Invalid or deleted cell reference</summary>
        ErrRef = 0x17,
        /// <summary>Invalid name</summary>
        ErrName = 0x1D,
        /// <summary>Invalid number</summary>
        ErrNum = 0x24,
        /// <summary>Not available</summary>
        ErrNA = 0x2A,
    }

    /// <summary>
    /// Category to which a cell style belongs. This is only valid for Excel 2007 or newer.
    /// </summary>
    public enum TStyleCategory
    {
        /// <summary>
        /// Custom style.
        /// </summary>
        Custom = 0x00,

        /// <summary>
        /// Good, bad, neutral style.
        /// </summary>
        GoodBadNeutral = 0x01,

        /// <summary>
        /// Data model style.
        /// </summary>
        DataModel = 0x02,

        /// <summary>
        /// Title and heading style.
        /// </summary>
        TitleHeading = 0x03,

        /// <summary>
        /// Themed cell style. (ACCENT styles)
        /// </summary>
        ThemedCell = 0x04,

        /// <summary>
        /// Number format style.
        /// </summary>
        NumberFormat = 0x05
    }

    /// <summary>
    /// Possible types of cell hyperlinks.
    /// </summary>
    public enum THyperLinkType
    {
        /// <summary>
        /// Web, file or mail URL. (like http://, file://, mailto://, ftp://) 
        /// </summary>
        URL = 0,

        /// <summary>
        /// A file on the local disk. Not an url or unc file.
        /// </summary>
        LocalFile = 1,

        /// <summary>
        /// Universal Naming convention. For example: \\server\path\file.ext
        /// </summary>
        UNC = 2,

        /// <summary>
        /// An hyperlink inside the current file.
        /// </summary>
        CurrentWorkbook = 3
    }

    
    /// <summary>
    /// Specifies the function group index if the defined name refers to a function. The function 
    /// group defines the general category for the function. This attribute is used when there is 
    /// an add-in or other code project associated with the file. 
    /// </summary>
    public enum TFunctionGroup
    {
        /// <summary>
        /// Not defined. Don't use.
        /// </summary>
        None = 0,

        /// <summary>
        /// Financial
        /// </summary> 
        Financial = 1,
     
        /// <summary>
        /// Date and Time 
        /// </summary> 
        DateAndTime = 2,

        ///Math and Trig
        MathAndTrig = 3,

        ///Statistical
        Statistical = 4,

        ///Lookup and Reference
        LookupAndReference = 5,

        ///Database
        Database = 6,

        ///Text
        Text = 7,

        ///Logical
        Logical = 8,

        ///Information
        Information = 9,

        ///Commands
        Commands = 10,

        ///Customizing
        Customizing = 11,

        ///Macro Control
        MacroControl = 12,

        ///DDE / External
        DDE_External = 13,

        ///User Defined
        UserDefined = 14,

        ///Engineering
        Engineering = 15,

        ///Cube
        Cube = 16
    }

    /// <summary>
    /// Types of error that might happen while recalculating.
    /// </summary>
    public enum TUnsupportedFormulaErrorType
    {
        /// <summary>
        /// FlexCel was not able to parse the formula.
        /// </summary>
        FormulaTooComplex,

        /// <summary>
        /// There is a function on the formula that is not implemented by FlexCel.
        /// </summary>
        MissingFunction,

        /// <summary>
        /// The function is supported, but not with those arguments.
        /// </summary>
        FunctionalityNotImplemented,

        /// <summary>
        /// There is a circular reference on this cell.
        /// </summary>
        CircularReference,

        /// <summary>
        /// The file in the external reference was not found.
        /// </summary>
        ExternalReference
    }

    #endregion

    #region Copy/Insert modes
    /// <summary>
    /// Inserting mode.
    /// </summary>
    public enum TFlxInsertMode
    {
        /// <summary>
        /// Cells will not be inserted but just overwrite existing ones. If count is &gt;0, additional ranges will be copied down.
        /// <br></br>When deleting, this mode will clear the cells and not move anything.
        /// </summary>
        NoneDown,

        /// <summary>
        /// Cells will not be inserted but just overwrite existing ones. If count is &gt;0, additional ranges will be copied right.
        /// <br></br>When deleting, this mode will clear the cells and not move anything.
        /// </summary>
        NoneRight,

        /// <summary>
        /// Inserts whole rows. Moves all destination rows down.
        /// <br></br>When deleting, moves cells up. 
        /// </summary>
        ShiftRowDown,

        /// <summary>
        /// Inserts whole columns. Moves all destination columns to the right.
        /// <br></br>When deleting, moves columns to the right of the deleted columns to the left. 
        /// </summary>
        ShiftColRight,

        /// <summary>
        /// Moves all destination cells down. This WON'T move the whole row, only cells on the range. 
        /// <br></br>When deleting, moves cells up. 
        /// </summary>
        ShiftRangeDown,

        /// <summary>
        /// Moves all destination cells to the right. This WON'T move the whole column, only cells on the range.
        /// <br></br>When deleting, moves cells to the left. 
        /// </summary>
        ShiftRangeRight
    }

    /// <summary>
    /// What we do with the cells when we call InsertAndCopyRange.
    /// </summary>
    public enum TRangeCopyMode
    {
        /// <summary>
        /// Will copy all (values, ranges, images, formulas, etc).
        /// </summary>
        All,

        /// <summary>
        /// Will copy all except values. Useful when you are going to replace those values anyway.
        /// </summary>
        OnlyFormulas,

        /// <summary>
        /// Won't copy anything. Only inserts.
        /// </summary>
        None,

        /// <summary>
        /// Will copy all except values and objects (like images). Useful when you are going to replace those values anyway.
        /// </summary>
        OnlyFormulasAndNoObjects,

        /// <summary>
        /// Will copy all objects, including those that are marked as "do not copy" in the file. You will normally not want to use this.
        /// </summary>
        AllIncludingDontMoveAndSizeObjects,

        /// <summary>
        /// Will copy cell formatting, but not cell contents. Images and objects will be copied too.
        /// </summary>
        Formats
    }
    #endregion

    #region Formula references
    /// <summary>
    /// Use this enumerator in the property <see cref="ExcelFile.FormulaReferenceStyle"/> to specify the reference
    /// mode that FlexCel will use when you enter formulas as text or when it returns the formula text.
    /// </summary>
    public enum TReferenceStyle
    {
        /// <summary>
        /// Standard A1 mode. Uses letters for the columns and numbers for the rows.
        /// </summary>
        A1,

        /// <summary>
        /// R1C1 style. Uses numbers for both columns and rows.
        /// </summary>
        R1C1
    }
    #endregion

    #region Object Selection
    /// <summary>
    /// Types of selection allowed in a listbox.
    /// </summary>
    public enum TListBoxSelectionType
    {
        /// <summary>
        /// The list control is only allowed to have one selected item. 
        /// </summary>
        Single,

        /// <summary>
        /// The list control is allowed to have multiple items selected by clicking on each item. 
        /// </summary>
        Multi,

        /// <summary>
        /// The list control is allowed to have multiple items selected by holding the CTRL key and clicking on each item. 
        /// </summary>
        Extend
    }
    #endregion
}
