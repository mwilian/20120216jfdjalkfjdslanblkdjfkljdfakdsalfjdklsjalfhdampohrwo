using System;
using System.Reflection;
using System.Globalization;
using System.Text;
using System.Collections.Generic;

#if (MONOTOUCH)
    using Color = MonoTouch.UIKit.UIColor;
#endif

#if (WPF)
using System.Windows.Media;
using real = System.Double;
#else
using System.Drawing;
using real = System.Single;
#endif

namespace FlexCel.Core
{
	/// <summary>
	/// Utility methods and constants usable anywhere.
	/// </summary>
	public sealed class FlxConsts
	{
		private FlxConsts(){}

		/// <summary>
		/// OBSOLETE: Use <see cref="ExcelMetrics.DefColWidthAdapt(int, ExcelFile)"/> instead.
		/// Multiply by this number to convert the DEFAULT column width to pixels. This is different from <see cref="ColMult"/>, that goes in a column by column basis.
		/// </summary>
		[Obsolete("This constant will not work when the user modifies the default font on a file. Use ExcelMetrics.DefColWidthAdapt instead.")]
		public static readonly int DefColWidthAdapt = 293; //(256*8)/7;  //font used here is 8 pixels wide, not 7
        
		/// <summary>
		/// Defines the Excel mode used in this thread.
		/// Note that while on v2007 (the default) you still can make xls 97 spreadsheets, so the only reason to change this setting
		/// is if you have any compatibility issues (for example your code depends on a sheet having 65536 rows).
		/// IMPORTANT: Do NOT change this value after reading a workbook. Also, remember that the value is changed for all the reports in all threads.
		/// </summary>
		internal static TExcelVersion ExcelVersion; //STATIC*
        
        /// <summary>
        /// Defines what FlexCel will do when it finds a reference to the last row or column in an Excel 97-2003 spreadsheet, and it is upgrading to Excel 2007.
        /// If false (the default) row 65536 will be updated to row 1048576, and column 256 to column 16384.
        /// If true, references will stay the same. <b>Note: </b> This is a global property, so it affects all threads running.
        /// </summary>
        internal static bool KeepMaxRowsAndColumsWhenUpdating; //STATIC* 

        /// <summary>
		/// Maximum column on an xls (Excel 97 - 2003) spreadsheet. (0 based, that is 255)
		/// </summary>
		public const int Max_Columns97_2003       = 255;

		/// <summary>
		/// Maximum column on an xlsx (Excel 2007 and up) spreadsheet. (0 based, that is 16383)
		/// </summary>
		public const int Max_Columns2007          = 16383;

		/// <summary>
		/// Maximum column in the spreadsheet. (0 based). This number might be 255 if <see cref="FlexCel.Core.ExcelFile.ExcelVersion"/> is TExcelVersion.v97_2003 or 16383 otherwise.
		/// </summary>
		public static int Max_Columns {get { return ExcelVersion == TExcelVersion.v2007? Max_Columns2007: Max_Columns97_2003;}}

		/// <summary>
		/// Maximum row on an xls (Excel 97 - 2003) spreadsheet. (0 based, that is 65535)
		/// </summary>
		public const int Max_Rows97_2003          = 65535; //0 based

		/// <summary>
		/// Maximum row on an xlsx (Excel 2007 and up) spreadsheet. (0 based, that is 1048575)
		/// </summary>
		public const int Max_Rows2007            = 1048575; //0 based

		/// <summary>
        /// Maximum row in the spreadsheet. (0 based). This number might be 65535 if <see cref="FlexCel.Core.ExcelFile.ExcelVersion"/> is TExcelVersion.v97_2003 or 1048575 otherwise.
		/// </summary>
		public static int Max_Rows {get { return ExcelVersion == TExcelVersion.v2007? Max_Rows2007: Max_Rows97_2003;}}

		/// <summary>
		/// Maximum sheet on a spreadsheet. (0 based, that is 65530)
		/// </summary>
		public const int Max_Sheets         = 65530;

		/// <summary>
		/// Maximum column on Pocket Excel a spreadsheet. (0 based, that is 255)
		/// </summary>
		public const int Max_PxlColumns       = 255;

		/// <summary>
		/// Maximum row on a Pocket Excel spreadsheet. (0 based, that is 16383)
		/// </summary>
		public const int Max_PxlRows          = 0x3FFF; //0 based

		/// <summary>
		/// Maximum sheet on a Pocket Excel spreadsheet. (0 based, that is 255)
		/// </summary>
		public const int Max_PxlSheets         = 255;
        
		/// <summary>
		/// Number of letters in a column name. This is 2 in xls97 (columns go up to IV) and 3 in xls2007 (columns go up to XFD)
		/// </summary>
		public static int Max_LettersInColumnName {get{ return ExcelVersion == TExcelVersion.v2007? 3: 2;}}

		/// <summary>
		/// Maximun number of characters in a Formula
		/// </summary>
		public static int Max_FormulaLen {get{ return ExcelVersion == TExcelVersion.v2007? Max_FormulaLen2007: Max_FormulaLen97_2003;}}

        /// <summary>
        /// Maximum number of arguments for a formula in xls file format.
        /// </summary>
        public const int Max_FormulaArguments2003 = 30;

        /// <summary>
        /// Maximum number of arguments for a formula in xlsx file format.
        /// </summary>
        public const int Max_FormulaArguments2007 = 255;

        /// <summary>
        /// Maximum length of a direct string inside a formula, as in ' = "my long string..."
        /// </summary>
        public static int Max_FormulaStringConstant { get { return 255; } }

        /// <summary>
        /// Maximum length of a string in a cell.
        /// </summary>
        public const int Max_StringLenInCell = 0x7FFF;

        /// <summary>
        /// Maximun number of characters in a Formula for an Excel 97 to 2003 spreadsheet.
        /// </summary>
        public static int Max_FormulaLen97_2003 { get { return 1024; } }

        /// <summary>
        /// Maximun number of characters in a Formula for an Excel 2007 or newer spreadsheet.
        /// </summary>
        public static int Max_FormulaLen2007 { get { return 8 * 1024; } }

        /// <summary>
        /// Maximum number of characters in an Error title for a Data Validation.
        /// </summary>
        public static int Max_DvErrorTitleLen { get { return 32; } }

        /// <summary>
        /// Maximum number of characters in an Error text for a Data Validation.
        /// </summary>
        public static int Max_DvErrorTextLen { get { return 225; } }

        /// <summary>
        /// Maximum number of characters in an Input title for a Data Validation.
        /// </summary>
        public static int Max_DvInputTitleLen { get { return 32; } }

        /// <summary>
        /// Maximum number of characters in an Input text for a Data Validation.
        /// </summary>
        public static int Max_DvInputTextLen { get { return 255; } }

        /// <summary>
        /// Maximum number of characters allowed in the author of a comment.
        /// </summary>
        public static int Max_CommentAuthor { get { return 54; } }

        /// <summary>
        /// OBSOLETE: Use <see cref="ExcelMetrics.ColMult"/> instead.
        /// Multiply by this number to convert the width of a column from pixels to excel internal units. 
        /// Note that the default column width is different, you need to multiply by <see cref="DefColWidthAdapt"/>
        /// </summary>
        [Obsolete("This constant will not work when the user modifies the default font on a file. Use ExcelMetrics.ColMult instead.")]
        public static readonly float ColMult=256F/7; //36.6;

        /// <summary>
        /// Multiply by this number to convert pixels to excel row height units.
        /// </summary>
        /// <remarks>
        /// 1 Height unit= 1/20 pt. 1pt=1/72 inch.  At 96ppi, 1 Height Unit= 96/(72*20)pixels -> 1 pix=(72*20)/96 = 15 Height units. 
        /// </remarks>
        public static readonly float RowMult=15;  // 

        internal const float DispMul = 72F; //WE ARE USING POINTS ON RENDER.
        //100F;  //Display should be 75 acording to doc, but it is 100. We are using Points now.

        internal const double PixToPoints = 1 / 0.75; //Pixels in silverlight are 1/96 of an inch, points 1/72

        /// <summary>
        /// OBSOLETE: Use <see cref="ExcelMetrics.ColMultDisplay"/> instead.
        /// Multiply by this number to convert the width of a column from GraphicsUnit.Display units (1/100 inch) 
        /// to Excel internal units. Note that the default column width is different, you need to multiply by <see cref="DefColWidthAdapt"/>
        /// </summary>
        [Obsolete("This constant will not work when the user modifies the default font on a file. Use ExcelMetrics.ColMultDisplay instead.")]
        public static readonly float ColMultDisplay=33.45F;

        /// <summary>
        /// OBSOLETE: Use <see cref="ExcelMetrics.RowMultDisplay"/> instead.
        /// Multiply by this number to convert the height of a row from GraphicsUnit.Display units (1/100 inch) 
        /// to Excel internal units.
        /// </summary>
        /// <remarks>
        /// 1 Height unit=1/20 pt. 1pt=1/72 inch. -> 1 Height unit=1/(72*20) inch. -> 1inch/100= 72*20/100= 14.4
        /// </remarks>
        [Obsolete("Use ExcelMetrics.RowMultDisplay instead.")]
        public static readonly float RowMultDisplay= 14.72F;  //14.72F;

        /// <summary>
        /// Brightness to keep the image unchanged.
        /// </summary>
        public const int DefaultBrightness = 0;

        /// <summary>
        /// Contrast to keep the image unchanged.
        /// </summary>
        public const int DefaultContrast = 1<<16;

        /// <summary>
        /// Gamma to keep the image unchanged.
        /// </summary>
        public const int DefaultGamma = 0;

        /// <summary>
        /// Zero rotation.
        /// </summary>
        public const int DefaultRotation = 0;
            
        /// <summary>
        /// Constant meaning there is no transparent color defined on the image.
        /// </summary>
        public const long NoTransparentColor = ~0L;

        /// <summary>
        /// The default XF for a file. You can also access this value with <see cref="ExcelFile.DefaultFormatId"/>
        /// </summary>
        public const int DefaultFormatId = 0x00;

        internal const int DefaultFormatIdBiff8 = 0x0F;

        internal const string NormalStyleName = "Normal";

        /// <summary>
        /// Returns "A" for column 1, "B"  for 2 and so on.
        /// </summary>
        /// <param name="C">Index of the column, 1 based.</param>
        /// <returns></returns>
        [Obsolete("Use TCellAddress.EncodeColumn instead.")]
        public static string EncodeColumn(int C)
        {
            return TCellAddress.EncodeColumn(C);
        }

        /// <summary>
        /// String used to separate 2 objects on an object path.
        /// </summary>
        public const string ObjectPathSeparator = @"\";

        /// <summary>
        /// When an objpath starts with this character, it is an absolute path that includes the object index.
        /// If it doesn't start with it, then the ObjPath doesn't include the original object.
        /// </summary>
        public const string ObjectPathAbsolute = @"\";

        /// <summary>
        /// When an objpath starts with this character, it is a path that goes directly to the name of an object.
        /// Note that when more than an object have the same name in the same sheet, this path won't work and you will
        /// have to use absolute or relative ones.
        /// </summary>
        public const string ObjectPathObjName = "@";

        /// <summary>
        /// When an objpath starts with this character, what follows is a single shape id that identifies the object.
        /// </summary>
        public const string ObjectPathSpId = "|";

    }
}
