using System.Text;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

#if(WPF)
using real = System.Double;
using System.Windows.Media;

#else
using real = System.Single;
using System.Drawing;
#endif

using System.Security.Permissions;
using System.Security;

namespace FlexCel.Core
{
    /// <summary>
    /// PointF is not supported in CF, so this is an equivalent
    /// </summary>
    public struct TPointF
    {
        private real FX;
        private real FY;

        /// <summary>
        /// Creates a new Point.
        /// </summary>
        /// <param name="aX">X coord.</param>
        /// <param name="aY">Y coord.</param>
        public TPointF(real aX, real aY)
        {
            FX = aX;
            FY = aY;
        }

        /// <summary>
        /// X coord.
        /// </summary>
        public real X { get { return FX; } set { FX = value; } }

        /// <summary>
        /// Y coord.
        /// </summary>
        public real Y { get { return FY; } set { FY = value; } }

        /// <summary>
        /// True if both objects are equal.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            if (!(obj is TPointF)) return false;
            TPointF o2 = (TPointF)obj;
            return (o2.X == X && o2.Y == Y);
        }

        /// <summary></summary>
        public static bool operator ==(TPointF b1, TPointF b2)
        {
            return b1.Equals(b2);
        }

        /// <summary></summary>
        public static bool operator !=(TPointF b1, TPointF b2)
        {
            return !(b1 == b2);
        }

        /// <summary>
        /// Hash code for the point.
        /// </summary>
        /// <returns>hashcode.</returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHash(FX.GetHashCode(), FY.GetHashCode());

        }

        /// <summary>
        /// String with both coordinates.
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return string.Format("X: {0}, Y: {1}", FX, FY); 
        }


    }

    /// <summary>
    /// An encapsulation of an Excel hyperlink.
    /// </summary>
    public class THyperLink : ICloneable
    {
        private THyperLinkType FLinkType;

        private string FDescription;
        private string FTargetFrame;
        private string FTextMark;
        private string FText;

        private string FHint;

        /// <summary>
        /// Creates a new instance of an Hyperlink class.
        /// </summary>
        public THyperLink()
        {
        }


        /// <summary>
        /// Creates a new instance of an Hyperlink class.
        /// </summary>
        /// <param name="aLinkType">The type of hyperlink: to a local file, to an url, to a cell or to a networked file.</param>
        /// <param name="aText">Text of the HyperLink. This is empty when linking to a cell.</param>
        /// <param name="aDescription">Description of the HyperLink.</param>
        /// <param name="aTargetFrame">When entering an URL on excel, you can enter additional text following the url with a "#" character (for example www.your_url.com#myurl") The text Mark is the text after the "#" char. When entering an address to a cell, the address goes here too.</param>
        /// <param name="aTextMark">This parameter is not documented. You can leave it empty.</param>
        public THyperLink(THyperLinkType aLinkType, string aText, string aDescription, string aTargetFrame, string aTextMark)
        {
            LinkType = aLinkType;
            Text = aText;
            Description = aDescription;
            TargetFrame = aTargetFrame;
            TextMark = aTextMark;
        }

        /// <summary>
        /// The type of hyperlink: to a local file, to an url, to a cell or to a networked file.
        /// </summary>
        public THyperLinkType LinkType { get { return FLinkType; } set { FLinkType = value; } }

        /// <summary>
        /// Text of the HyperLink. This is empty when linking to a cell.
        /// </summary>
        public string Text { get { return FText; } set { FText = value; } }

        /// <summary>
        /// Description of the HyperLink.
        /// </summary>
        public string Description { get { return FDescription; } set { FDescription = value; } }

        /// <summary>
        /// This parameter is not documented. You will probably leave it empty.
        /// </summary>
        public string TargetFrame { get { return FTargetFrame; } set { FTargetFrame = value; } }

        /// <summary>
        /// When entering an URL, you can enter additional text following the url with a "#" character (for example www.your_url.com#myurl") The text Mark is the text after the "#" char. When entering an address to a cell, the address goes here too.
        /// </summary>
        public string TextMark { get { return FTextMark; } set { FTextMark = value; } }

        /// <summary>
        /// Hint when the mouse hovers over the hyperlink.
        /// </summary>
        public string Hint { get { return FHint; } set { FHint = value; } }
        #region ICloneable Members

        /// <summary>
        /// Returns a copy of the hyperlink.
        /// </summary>
        /// <returns></returns>
        public object Clone()
        {
            return MemberwiseClone();
        }

        #endregion
    }

    /// <summary>
    /// Manages converting from/to objects/Native excel types.
    /// </summary>
    public sealed class TExcelTypes
    {
        private TExcelTypes() { }

        /// <summary>
        /// Converts an object on a CellType representation.
        /// </summary>
        /// <param name="o">Object with the value.</param>
        /// <returns>Cell Type</returns>
        public static TCellType ObjectToCellType(object o)
        {
            if (o == null) return TCellType.Empty;
            if (o is TFlxFormulaErrorValue) return TCellType.Error;
            if (o is TFormula) return TCellType.Formula;
            if (o is TRichString) return TCellType.String;

            switch (Type.GetTypeCode(o.GetType()))
            {
                case TypeCode.Byte:
                case TypeCode.Decimal:
                case TypeCode.Double:
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                case TypeCode.SByte:
                case TypeCode.Single:
                case TypeCode.UInt16:
                case TypeCode.UInt32:
                case TypeCode.UInt64:
                    {
                        return TCellType.Number;
                    }
                case TypeCode.DateTime:
                    {
                        return TCellType.DateTime;
                    }
                case TypeCode.Char:
                case TypeCode.String:
                    {
                        return TCellType.String;
                    }
                case TypeCode.Boolean:
                    {
                        return TCellType.Bool;
                    }
                case TypeCode.DBNull:
                case TypeCode.Empty:
                    {
                        return TCellType.Empty;
                    }
            }
            return TCellType.Unknown;
        }

        /// <summary>
        /// Converts an object to a native Excel datatype, that is:
        /// Number, String, Null, bool or Error.
        /// </summary>
        /// <param name="o">Object to convert.</param>
        /// <param name="Dates1904">True if using 1904 as start date. Excel for windows normally uses 1900.</param>
        /// <returns></returns>
        public static object ConvertToAllowedObject(object o, bool Dates1904)
        {
            if (o == null) return o;
            TRichString rstr = o as TRichString; if (rstr != null) return rstr.ToString();
            if (o is TFlxFormulaErrorValue) return o;
            TFormula fmla = o as TFormula; if (fmla != null) return ConvertToAllowedObject(fmla.Result, Dates1904);
            if (o is TMissingArg) return (double)0;

            object[,] ArrResult = o as object[,];
            if (ArrResult != null && ArrResult.GetLength(0) > 0 && ArrResult.GetLength(1) > 0)
            {
                return ConvertToAllowedObject(ArrResult[0, 0], Dates1904);
            }

            switch (Type.GetTypeCode(o.GetType()))
            {
                case TypeCode.Byte:
                case TypeCode.Decimal:
                case TypeCode.Double:
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                case TypeCode.SByte:
                case TypeCode.Single:
                case TypeCode.UInt16:
                case TypeCode.UInt32:
                case TypeCode.UInt64:
                    {
                        return Convert.ToDouble(o);
                    }
                case TypeCode.DateTime:
                    {
                        return FlxDateTime.ToOADate((DateTime)o, Dates1904);
                    }
                case TypeCode.Char:
                    {
                        return Convert.ToString(o);
                    }
                case TypeCode.DBNull:
                case TypeCode.Empty:
                    {
                        return null;
                    }
                case TypeCode.Boolean:
                    {
                        return o;
                    }
            }
            return Convert.ToString(o);
        }
    }

    /// <summary>
    /// Use this class to convert between a Date expressed on Excel format (a double)
    /// and a TDateTime.
    /// </summary>
    public sealed class FlxDateTime
    {
        private FlxDateTime() { }

        private const double OADateMinAsDouble = -657435.0;
        private const double OADateMaxAsDouble = 2958466.0;
        private const int MillisPerSecond = 1000;
        private const int MillisPerMinute = MillisPerSecond * 60;
        private const int MillisPerHour = MillisPerMinute * 60;
        private const int MillisPerDay = MillisPerHour * 24;
        private const long TicksPerMillisecond = 10000;
        private const long TicksPerSecond = TicksPerMillisecond * 1000;
        private const long TicksPerMinute = TicksPerSecond * 60;
        private const long TicksPerHour = TicksPerMinute * 60;
        private const long TicksPerDay = TicksPerHour * 24;
        private const int DaysTo1899 = DaysPer400Years * 4 + DaysPer100Years * 3 - 367;
        private const int DaysTo10000 = DaysPer400Years * 25 - 366;
        private const long DoubleDateOffset = DaysTo1899 * TicksPerDay;
        private const int DaysPerYear = 365;
        private const int DaysPer4Years = DaysPerYear * 4 + 1;
        private const int DaysPer100Years = DaysPer4Years * 25 - 1;
        private const int DaysPer400Years = DaysPer100Years * 4 + 1;
        private const long MaxMillis = (long)DaysTo10000 * MillisPerDay;
        private const long OADateMinAsTicks = (DaysPer100Years - DaysPerYear) * TicksPerDay;

        private const long Date1904Diff = 4 * 365 + 2;

        // Converts an OLE Date to a tick count.
        // This function is equivalent to the one on the framework, but it isn't implemented on CF.
        private static bool DoubleDateToTicks(double value, out long ResultDate)
        {
            ResultDate = 0;
            if (value >= OADateMaxAsDouble || value <= OADateMinAsDouble) return false;
            double Ofs = value >= 0 ? 0.5 : -0.5;
            long millis = (long)(value * MillisPerDay + Ofs);
            if (millis < 0)
            {
                millis -= (millis % MillisPerDay) * 2;
            }

            millis += DoubleDateOffset / TicksPerMillisecond;

            if (millis < 0 || millis >= MaxMillis) return false;
            ResultDate = millis * TicksPerMillisecond;
            return true;
        }


        private static bool TicksToOADate(long value, out double ResultDate)
        {
            if (value == 0)
            {
                ResultDate = 0.0;  // Returns OleAut's zero'ed date value.
                return true;
            }
            if (value < TicksPerDay) // This is a fix for VB. They want the default day to be 1/1/0001 rather then 12/30/1899.
                value += DoubleDateOffset; // We could have moved this fix down but we would like to keep the bounds check.
            if (value < OADateMinAsTicks)
            {
                ResultDate = 0;
                return false;
            }
            long millis = (value - DoubleDateOffset) / TicksPerMillisecond;
            if (millis < 0)
            {
                long frac = millis % MillisPerDay;
                if (frac != 0) millis -= (MillisPerDay + frac) * 2;
            }
            ResultDate = (double)millis / MillisPerDay;
            return true;
        }

        /// <summary>
        /// Converts a DateTime into a Double on Excel format for dates (Ole Automation Format).
        /// </summary>
        /// <param name="value">DateTime you want to convert.</param>
        /// <param name="Dates1904">When true dates start at 1904 (Macs) instead of 1900 (Windows)</param>
        /// <returns>The value as a double on Excel format.</returns>
        public static double ToOADate(DateTime value, bool Dates1904)
        {
            double Result;

            if (!TicksToOADate(value.Ticks, out Result))
                FlxMessages.ThrowException(FlxErr.ErrInvalidValue, "Date", value, OADateMinAsTicks, "x");

            if (Dates1904) Result -= Date1904Diff;

            return Result;
        }

        /// <summary>
        /// Converts a DateTime into a Double on Excel format for dates (Ole Automation Format).
        /// </summary>
        /// <param name="value">DateTime you want to convert.</param>
        /// <param name="ResultDate">Returns the value as double on Excel format.</param>
        /// <param name="Dates1904">When true dates start at 1904 (Macs) instead of 1900 (Windows)</param>
        /// <returns>True if conversion was successful, false otherwise.</returns>
        public static bool TryToOADate(DateTime value, bool Dates1904, out double ResultDate)
        {
            if (!TicksToOADate(value.Ticks, out ResultDate)) return false;
            if (Dates1904) ResultDate -= Date1904Diff;
            return true;
        }

        /// <summary>
        /// Converts a Double on Excel format for dates (Ole Automation Format) into a DateTime.
        /// </summary>
        /// <param name="value">Double you want to convert.</param>
        /// <param name="Dates1904">When true dates start at 1904 (Macs) instead of 1900 (Windows)</param>
        /// <returns>The value as DateTime.</returns>
        public static DateTime FromOADate(double value, bool Dates1904)
        {
            long Ticks;
            if (Dates1904) value += Date1904Diff;
            if (!DoubleDateToTicks(value, out Ticks))
                FlxMessages.ThrowException(FlxErr.ErrInvalidValue, "Date", value, OADateMinAsDouble, OADateMaxAsDouble);
            return new DateTime(Ticks);
        }

        /// <summary>
        /// Returns true is the double value can be converted into and Excel date.
        /// </summary>
        /// <param name="value">Double you want to convert.</param>
        /// <param name="Dates1904">When true dates start at 1904 (Macs) instead of 1900 (Windows)</param>
        /// <returns></returns>
        public static bool IsValidDate(double value, bool Dates1904)
        {
            if (Dates1904) value += Date1904Diff;
            if (value >= OADateMaxAsDouble || value <= OADateMinAsDouble) return false;
            return true;
        }

        /// <summary>
        /// Converts a Double on Excel format for dates (Ole Automation Format) into a DateTime.
        /// </summary>
        /// <param name="value">Double you want to convert.</param>
        /// <param name="Dates1904">When true dates start at 1904 (Macs) instead of 1900 (Windows)</param>
        /// <param name="ResultDate">Returns the value as DateTime.</param>
        /// <returns>True if conversion was successful, false otherwise.</returns>
        public static bool TryFromOADate(double value, bool Dates1904, out DateTime ResultDate)
        {
            ResultDate = DateTime.MinValue;
            long Ticks;
            if (Dates1904) value += Date1904Diff;
            if (!DoubleDateToTicks(value, out Ticks)) return false;
            ResultDate = new DateTime(Ticks);
            return true;
        }

    }

    /// <summary>
    /// A list of unsupported formulas on a sheet.
    /// </summary>
    public class TUnsupportedFormulaList
    {
        private List<TUnsupportedFormula> FList = new List<TUnsupportedFormula>();

        /// <summary>
        /// Use this one to know on which cell we are working.
        /// </summary>
        internal TCellAddress CellAddress;

        /// <summary>
        /// The number of errors on the list.
        /// </summary>
        public int Count { get { return FList.Count; } }

        /// <summary>
        /// Adds a new unsupported formula to the list.
        /// </summary>
        /// <param name="fmla">Unsupported formula.</param>
        public void Add(TUnsupportedFormula fmla)
        {
            FList.Add(fmla);
        }

        /// <summary>
        /// Returns the items at position index.
        /// </summary>
        public TUnsupportedFormula this[int index]
        {
            get
            {
                return (TUnsupportedFormula)FList[index];
            }

            set
            {
                FList[index] = value;
            }
        }
    }

    /// <summary>
    /// An unsupported formula, the cell it is in, and the reason why it is not supported.
    /// </summary>
    public class TUnsupportedFormula
    {
        private TUnsupportedFormulaErrorType FErrorType;
        private TCellAddress FCell;
        private string FFunctionName;
        private string FFileName;

        /// <summary>
        /// Creates a new empty instance.
        /// </summary>
        public TUnsupportedFormula()
        {
        }

        /// <summary>
        /// Creates a new instance of a TUnsupported formula class.
        /// </summary>
        /// <param name="aErrorType">See <see cref="ErrorType"/></param>
        /// <param name="aCell">See <see cref="Cell"/></param>
        /// <param name="aFunctionName">See <see cref="FunctionName"/></param>
        /// <param name="aFileName">See <see cref="FileName"/></param>
        public TUnsupportedFormula(TUnsupportedFormulaErrorType aErrorType, TCellAddress aCell, string aFunctionName, string aFileName)
        {
            FErrorType = aErrorType;
            FCell = aCell;
            FFunctionName = aFunctionName;
            FFileName = aFileName;
        }

        /// <summary>
        /// Type of error.
        /// </summary>
        public TUnsupportedFormulaErrorType ErrorType { get { return FErrorType; } set { FErrorType = value; } }

        /// <summary>
        /// Cell where the formula is (1 based)
        /// </summary>
        public TCellAddress Cell { get { return FCell; } set { FCell = value; } }

        /// <summary>
        /// If the error is <see cref="TUnsupportedFormulaErrorType.MissingFunction"/> then this is the name of the missing function.  
        /// If the error is <see cref="TUnsupportedFormulaErrorType.ExternalReference"/> then this is the name of the file not found.
        /// </summary>
        public string FunctionName { get { return FFunctionName; } set { FFunctionName = value; } }

        /// <summary>
        /// This property has the name of the physical file being evaluated, and can be of use when evaluating linked files. If the files are opened
        /// from a stream or not from a physical place, it will be null.
        /// </summary>
        public string FileName { get { return FFileName; } set { FFileName = value; } }

    }

    #region HashTables
    internal sealed class TCaseInsensitiveHashtableStrInt : Dictionary<string, int>
    {
        public TCaseInsensitiveHashtableStrInt()
            : base(StringComparer.InvariantCultureIgnoreCase)
        {
        }

        public int GetValue(string Key)
        {
            int Result = 0;
            if (!TryGetValue(Key, out Result))
                FlxMessages.ThrowException(FlxErr.ErrInvalidFormat, Key);
            return Result;

        }
    }
    
    internal sealed class StringIntHashtable : Dictionary<string, int>
    {
        public StringIntHashtable()
            : base(StringComparer.InvariantCulture)
        {
        }
    }

    internal sealed class IntStringHashtable : Dictionary<int, string>
    {
        public IntStringHashtable()
            : base()
        {
        }
    }

    internal sealed class StringStringHashtable : Dictionary<string, string>
    {
        public StringStringHashtable()
            : base(StringComparer.InvariantCulture)
        {
        }
    }
    #endregion

    #region Lists
    internal class TDoubleList
    {
        private List<Double> FList;
        public TDoubleList()
        {
            FList = new List<Double>();
        }

        public Double[] ToArray()
        {
            return FList.ToArray();
        }

        public void Add(Double d)
        {
            FList.Add(d);
        }

        public void AddRange(TDoubleList d)
        {
            FList.AddRange(d.FList);
        }

        public double this[int index]
        {
            get
            {
                return (double)FList[index];
            }
        }

        public int Count { get { return FList.Count; } }

        public void Sort()
        {
            FList.Sort();
        }

    }

    #endregion

    #region Data Validation
    /// <summary>
    /// Contains the information to define a data validation in a range of cells.
    /// </summary>
    public class TDataValidationInfo : IComparable
    {
        #region Privates
        private TDataValidationDataType FValidationType;
        private TDataValidationConditionType FCondition;
        private string FFirstFormula;
        private string FSecondFormula;
        private bool FIgnoreEmptyCells;
        private bool FInCellDropDown;
        private bool FExplicitList;


        private bool FShowErrorBox;
        private string FErrorBoxCaption;
        private string FErrorBoxText;

        private bool FShowInputBox;
        private string FInputBoxCaption;
        private string FInputBoxText;

        private TDataValidationIcon FErrorIcon;
        private TDataValidationImeMode FImeMode;

        #endregion

        #region Constructors
        /// <summary>
        /// Empty constructor. Creates a new instance of TDataValidationInfo without assigning any value.
        /// </summary>
        public TDataValidationInfo()
        {
            FImeMode = TDataValidationImeMode.NoControl;
        }

        /// <summary>
        /// Creates a new Data Validation condition with all parameters.
        /// </summary>
        /// <param name="aValidationType">See <see cref="ValidationType"/></param>
        /// <param name="aCondition">See <see cref="Condition"/></param>
        /// <param name="aFirstFormula">See <see cref="FirstFormula"/></param>
        /// <param name="aSecondFormula">See <see cref="SecondFormula"/></param>
        /// <param name="aIgnoreEmptyCells">See <see cref="IgnoreEmptyCells"/></param>
        /// <param name="aInCellDropDown">See <see cref="InCellDropDown"/></param>
        /// <param name="aErrorBoxCaption">See <see cref="ErrorBoxCaption"/></param>
        /// <param name="aErrorBoxText">See <see cref="ErrorBoxText"/></param>
        /// <param name="aInputBoxCaption">See <see cref="InputBoxCaption"/></param>
        /// <param name="aInputBoxText">See <see cref="InputBoxText"/></param>
        /// <param name="aErrorIcon">See <see cref="ErrorIcon"/></param>
        /// <param name="aExplicitList">See <see cref="ExplicitList"/></param>
        /// <param name="aShowErrorBox">See <see cref="ShowErrorBox"/></param>
        /// <param name="aShowInputBox"></param>See <see cref="ShowInputBox"/>
        public TDataValidationInfo(TDataValidationDataType aValidationType, TDataValidationConditionType aCondition, string aFirstFormula, string aSecondFormula,
            bool aIgnoreEmptyCells, bool aInCellDropDown, bool aExplicitList, bool aShowErrorBox, string aErrorBoxCaption, string aErrorBoxText,
            bool aShowInputBox, string aInputBoxCaption, string aInputBoxText,
            TDataValidationIcon aErrorIcon)
        {
            ValidationType = aValidationType;
            Condition = aCondition;
            FirstFormula = aFirstFormula;
            SecondFormula = aSecondFormula;
            IgnoreEmptyCells = aIgnoreEmptyCells;
            InCellDropDown = aInCellDropDown;
            ExplicitList = aExplicitList;
            ShowErrorBox = aShowErrorBox;
            ErrorBoxCaption = aErrorBoxCaption;
            ErrorBoxText = aErrorBoxText;
            ShowInputBox = aShowInputBox;
            InputBoxCaption = aInputBoxCaption;
            InputBoxText = aInputBoxText;
            ErrorIcon = aErrorIcon;
            ImeMode = TDataValidationImeMode.NoControl;
        }


        #endregion

        #region Properties

        /// <summary>
        /// Type of validation we will be doing.
        /// </summary>
        public TDataValidationDataType ValidationType { get { return FValidationType; } set { FValidationType = value; } }

        /// <summary>
        /// Condition used to apply the data validation.
        /// </summary>
        public TDataValidationConditionType Condition { get { return FCondition; } set { FCondition = value; } }

        /// <summary>
        /// Formula for the first condition of the data validation. The text of the formula is limited to 255 characters.
        /// If <see cref="ExplicitList"/> is true, this formula can contain a list of values. 
        /// <br/>Note that with <b>relative</b> references, we always consider "A1" to be the cell where the data validation is. This means that the formula:
        /// "=$A$1 + A1" when evaluated in Cell B8, will read "=$A$1 + B8". To provide a negative offset, you need to wrap the formula.
        /// For example "=A1048575" will evaluate to B7 when evaluated in B8.
        /// </summary>
        public string FirstFormula { get { return FFirstFormula; } set { FFirstFormula = value; } }

        /// <summary>
        /// Formula for the second condition of the data validation, if it has two conditions. The text of the formula is limited to 255 characters.
        /// <br/>Note that with <b>relative</b> references, we always consider "A1" to be the cell where the data validation is. This means that the formula:
        /// "=$A$1 + A1" when evaluated in Cell B8, will read "=$A$1 + B8". To provide a negative offset, you need to wrap the formula.
        /// For example "=A1048575" will evaluate to B7 when evaluated in B8.
        /// </summary>
        public string SecondFormula { get { return FSecondFormula; } set { FSecondFormula = value; } }

        /// <summary>
        /// If true Empty cells will not trigger data validation errors.
        /// </summary>
        public bool IgnoreEmptyCells { get { return FIgnoreEmptyCells; } set { FIgnoreEmptyCells = value; } }

        /// <summary>
        /// When the <see cref="ValidationType"/> parameter is a list, this property indicates whether to display a drop down box or not.
        /// </summary>
        public bool InCellDropDown { get { return FInCellDropDown; } set { FInCellDropDown = value; } }

        /// <summary>
        /// If true, <see cref="FirstFormula"/> contains a list of values. 
        /// In this case, Formula1 <b>must</b> be a formula of the type: ="string", where string is a list of values separated by Character(0).
        /// For example, in C# Formula1 could be: <i>="Apples\0Lemmons\0Melons</i>  
        /// In Delphi.NET, Formula1 could be: <i>'="Apples' + #0 + 'Lemmons' + #0 + 'Melons'</i>
        /// </summary>
        public bool ExplicitList { get { return FExplicitList; } set { FExplicitList = value; } }

        /// <summary>
        /// If true, an error box dialog will be shown when the user enters an invalid value.
        /// </summary>
        public bool ShowErrorBox { get { return FShowErrorBox; } set { FShowErrorBox = value; } }

        /// <summary>
        /// Caption of the Error Alert box. Note that this text cannot be longer than 32 characters.
        /// Extra caracters will be truncated. If this parameter is null, the default Error alert will be displayed.
        /// If <see cref="ShowErrorBox"/> is false, this parameter does nothing.
        /// </summary>
        public string ErrorBoxCaption { get { return FErrorBoxCaption; } set { FErrorBoxCaption = value; } }

        /// <summary>
        /// Text on the Error Alert box. Note that this text cannot be longer than 225 characters. 
        /// Extra characters will be truncated. If this parameter is null, the default Error alert will be displayed.
        /// If <see cref="ShowErrorBox"/> is false, this parameter does nothing.
        /// </summary>
        public string ErrorBoxText { get { return FErrorBoxText; } set { FErrorBoxText = value; } }

        /// <summary>
        /// If true, a box showing a message will be shown when the user selecte the cell.
        /// </summary>
        public bool ShowInputBox { get { return FShowInputBox; } set { FShowInputBox = value; } }

        /// <summary>
        /// Caption of the Input Message box. Note that this text cannot be longer than 32 characters.
        /// Extra caracters will be truncated. If this parameter is null, the Input box will display the default message.
        /// if <see cref="ShowInputBox"/> is false, this parameter does nothing.
        /// </summary>
        public string InputBoxCaption { get { return FInputBoxCaption; } set { FInputBoxCaption = value; } }

        /// <summary>
        /// Text on the Input Message box. Note that this text cannot be longer than 255 characters. 
        /// Extra characters will be truncated. If this parameter is null, the Input box will display the default message.
        /// if <see cref="ShowInputBox"/> is false, this parameter does nothing.
        /// </summary>
        public string InputBoxText { get { return FInputBoxText; } set { FInputBoxText = value; } }


        /// <summary>
        /// Icon to display in the error box.
        /// </summary>
        public TDataValidationIcon ErrorIcon { get { return FErrorIcon; } set { FErrorIcon = value; } }

        /// <summary>
        /// The IME (input method editor) mode enforced by this data validation.
        /// </summary>
        public TDataValidationImeMode ImeMode { get { return FImeMode; } set { FImeMode = value; } }

        #endregion

        #region IComparable Members

        private static int TextCompare(string s1, string s2, int MaxLen)
        {
            if (s1 == null || s2 == null || s1.Length == 0 || s2.Length == 0) //This will raise an exception when called with the full form
            {
                return String.Compare(s1, s2, StringComparison.InvariantCulture);
            }
            return String.Compare(s1, 0, s2, 0, MaxLen, StringComparison.InvariantCulture);
        }

        /// <summary>
        /// Compares the object with other.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns>-1 is o2 is less than this, 0 if they are equal and 1 if o2 is bigger.</returns>
        public int CompareTo(object obj)
        {
            TDataValidationInfo o2 = obj as TDataValidationInfo;
            if (o2 == null) return -1;

# if (COMPACTFRAMEWORK && !FRAMEWORK20)
			int Result = ((int)FValidationType).CompareTo((int)o2.FValidationType); if (Result != 0) return Result;
			Result = ((int)FCondition).CompareTo((int)o2.FCondition); if (Result != 0) return Result;
			Result = ((int)FErrorIcon).CompareTo((int)o2.FErrorIcon); if (Result != 0) return Result;
			Result = ((int)FImeMode).CompareTo((int)o2.FImeMode); if (Result != 0) return Result;
#else
            int Result = FValidationType.CompareTo(o2.FValidationType); if (Result != 0) return Result;
            Result = FCondition.CompareTo(o2.FCondition); if (Result != 0) return Result;
            Result = FErrorIcon.CompareTo(o2.FErrorIcon); if (Result != 0) return Result;
            Result = FImeMode.CompareTo(o2.FImeMode); if (Result != 0) return Result;
#endif

            Result = String.Compare(FFirstFormula, o2.FFirstFormula); if (Result != 0) return Result;
            Result = String.Compare(FSecondFormula, o2.FSecondFormula); if (Result != 0) return Result;
            Result = FIgnoreEmptyCells.CompareTo(o2.FIgnoreEmptyCells); if (Result != 0) return Result;
            Result = FInCellDropDown.CompareTo(o2.FInCellDropDown); if (Result != 0) return Result;
            Result = FExplicitList.CompareTo(o2.FExplicitList); if (Result != 0) return Result;


            Result = FShowErrorBox.CompareTo(o2.FShowErrorBox); if (Result != 0) return Result;
            Result = TextCompare(FErrorBoxCaption, o2.FErrorBoxCaption, 32); if (Result != 0) return Result;
            Result = TextCompare(FErrorBoxText, o2.FErrorBoxText, 225); if (Result != 0) return Result;

            Result = FShowInputBox.CompareTo(o2.FShowInputBox); if (Result != 0) return Result;
            Result = TextCompare(FInputBoxCaption, o2.FInputBoxCaption, 32); if (Result != 0) return Result;
            Result = TextCompare(FInputBoxText, o2.FInputBoxText, 255); if (Result != 0) return Result;


            return Result;
        }

        #endregion

        #region Compare
        /// <summary>
        /// Returns true if both objects are equal.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            return CompareTo(obj) == 0;
        }

        /// <summary>
        /// Returns true if both objects are equal.
        /// Note this is for backwards compatibility, this is a class and not immutable,
        /// so this method should return true if references are different. But that would break old code.
        /// </summary>
        /// <param name="o1"></param>
        /// <param name="o2"></param>
        /// <returns></returns>
        public static bool operator ==(TDataValidationInfo o1, TDataValidationInfo o2)
        {
            if ((object)o1 == null) return (object)o2 == null;
            return o1.Equals(o2);
        }

        /// <summary>
        /// Returns true if both objects are not equal.
        /// Note this is for backwards compatibility, this is a class and not immutable,
        /// so this method should return true if references are different. But that would break old code.
        /// </summary>
        /// <param name="o1"></param>
        /// <param name="o2"></param>
        /// <returns></returns>
        public static bool operator !=(TDataValidationInfo o1, TDataValidationInfo o2)
        {
            if ((object)o1 == null) return (object)o2 != null;
            return !(o1.Equals(o2));
        }

        /// <summary>
        /// Returns true if o1 is bigger than o2.
        /// </summary>
        /// <param name="o1"></param>
        /// <param name="o2"></param>
        /// <returns></returns>
        public static bool operator >(TDataValidationInfo o1, TDataValidationInfo o2)
        {
            if ((object)o1 == null)
            {
                if ((object)o2 == null) return false;
                return true;
            }
            if ((object)o2 == null) return true;
            return o1.CompareTo(o2) > 0;
        }

        /// <summary>
        /// Returns true if o1 is less than o2.
        /// </summary>
        /// <param name="o1"></param>
        /// <param name="o2"></param>
        /// <returns></returns>
        public static bool operator <(TDataValidationInfo o1, TDataValidationInfo o2)
        {
            if ((object)o1 == null)
            {
                if ((object)o2 == null) return true;
                return true;
            }
            if ((object)o2 == null) return false;
            return o1.CompareTo(o2) < 0;
        }

        /// <summary>
        /// Returns a hashcode for the object.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

        #endregion

    }

    /// <summary>
    /// Icon to be displayed in the error box of a data validation action. Note that this not only affects the icon used, but the possible values.
    /// An information icon will allow you to enter an invalid value in a cell, a stop icon will not.
    /// </summary>
    public enum TDataValidationIcon
    {
        /// <summary>
        /// Stop icon.  ( a red circle with a cross). When selected, invalid values cannot be entered into the cell.
        /// </summary>
        Stop = 0,

        /// <summary>
        /// Warning icon. ( a yellow triangle with an exclamation sign). When selected and there is an invalid entry, you get an error
        /// dialog allowing to cancel the operation, enter the invalid value anyway or re edit the cell.
        /// </summary>
        Warning = 1,

        /// <summary>
        /// Information icon. (a text ballon with an "i" inside). When selected and there is an invalid entry, a waning will be shown but the invalid data
        /// can be entered anyway.
        /// </summary>
        Information = 2
    }

    /// <summary>
    /// The IME (input method editor) mode enforced by a data validation.
    /// </summary>
    public enum TDataValidationImeMode
    {
                /// <summary>
        /// IME Mode Not Controlled. <br/>
        /// Data validation does not control the IME control's mode. 
        /// </summary>
        NoControl = 0x00,

        /// <summary>
        /// IME Off. <br/>
        /// Forces the IME control to be off when first selecting the cell (goes to direct cell input mode). 
        /// </summary>
        Off = 0x01,

        /// <summary>
        /// IME On. <br/>
        /// Forces the IME control to be on when first selecting the cell. 
        /// </summary>
        On = 0x02,

        /// <summary>
        /// IME mode is disabled. <br/>
        /// Forces the IME control to be disabled when this cell is selected. 
        /// </summary>
        Disabled = 0x03,

        /// <summary>
        /// Hiragana IME Mode. <br/>
        /// Forces the IME control to be on and in Hiragana input mode when first selecting the cell. Applies when the application's language is Japanese and a Japanese IME control is selected 
        /// </summary>
        Hiragana = 0x04,

        /// <summary>
        /// Full Katakana IME Mode. <br/>
        /// Forces the IME control to be on and in full-width Katakana input mode when first selecting the cell. Applies when the application's language is Japanese and a Japanese IME control is selected. 
        /// </summary>
        FullKatakana = 0x05,

        /// <summary>
        /// Half-Width Katakana. <br/>
        /// Forces the IME control to be on and in half-width Katakana input mode when first selecting the cell. Applies when the application's language is Japanese and a Japanese IME control is selected. 
        /// </summary>
        HalfKatakana = 0x06,

        /// <summary>
        /// Full-Width Alpha-Numeric IME Mode.  <br/>
        /// Forces the IME control to be on and in full-width alpha-numeric input mode when the cell is first selected. 
        /// </summary>
        FullAlpha = 0x07,

        /// <summary>
        /// Half Alpha IME. <br/>
        /// Forces the IME control to be on and in half-width alpha-numeric input mode when the cell is first selected. 
        /// </summary>
        HalfAlpha = 0x08,

        /// <summary>
        /// Full Width Hangul. <br/>
        /// Forces the IME control to be on and in full-width Hangul input mode when first selecting the cell. Applies when the application's language is Korean and a Korean IME control is selected. 
        /// </summary>
        FullHangul = 0x09,

        /// <summary>
        /// Half-Width Hangul IME Mode. <br/>
        /// Forces the IME control to be on and in half-width Hangul input mode when first selecting the cell. Applies when the application's language is Korean and a Korean IME control is selected. 
        /// </summary>
        HalfHangul = 0x0A

    }

    /// <summary>
    /// Possible types of data validation.
    /// </summary>
    public enum TDataValidationDataType
    {
        /// <summary>
        /// All values are allowed.
        /// </summary>
        AnyValue = 0,

        /// <summary>
        /// The value must be an integer.
        /// </summary>
        WholeNumber = 1,

        /// <summary>
        /// The value must be a decimal.
        /// </summary>
        Decimal = 2,

        /// <summary>
        /// The value must be in the list. 
        /// </summary>
        List = 3,

        /// <summary>
        /// The value must be a Date.
        /// </summary>
        Date = 4,

        /// <summary>
        /// The value must be a Time.
        /// </summary>
        Time = 5,

        /// <summary>
        /// The value will be validated depending on the text lenght.
        /// </summary>
        TextLenght = 6,

        /// <summary>
        /// The value will be validated depending on the results of a formula.
        /// </summary>
        Custom = 7
    }

    /// <summary>
    /// Defines the condition used in the data validation box.
    /// </summary>
    public enum TDataValidationConditionType
    {
        /// <summary>
        /// Value must be bewtween <see cref="TDataValidationInfo.FirstFormula"/> and <see cref="TDataValidationInfo.SecondFormula"/>.
        /// </summary>
        Between = 0,

        /// <summary>
        /// Value must be not bewtween <see cref="TDataValidationInfo.FirstFormula"/> and <see cref="TDataValidationInfo.SecondFormula"/>.
        /// </summary>
        NotBetween = 1,

        /// <summary>
        /// Value must be equal to <see cref="TDataValidationInfo.FirstFormula"/>.
        /// </summary>
        EqualTo = 2,

        /// <summary>
        /// Value must be different from <see cref="TDataValidationInfo.FirstFormula"/>.
        /// </summary>
        NotEqualTo = 3,

        /// <summary>
        /// Value must be greater than <see cref="TDataValidationInfo.FirstFormula"/>.
        /// </summary>
        GreaterThan = 4,

        /// <summary>
        /// Value must be less than <see cref="TDataValidationInfo.FirstFormula"/>.
        /// </summary>
        LessThan = 5,

        /// Value must be greater than or equal to <see cref="TDataValidationInfo.FirstFormula"/>.
        GreaterThanOrEqualTo = 6,

        /// Value must be less than or equal to <see cref="TDataValidationInfo.FirstFormula"/>.
        LessThanOrEqualTo = 7
    }
    #endregion

    #region Utils
    internal sealed class FlxUtils
    {
        private FlxUtils() { }

        /// <summary>
        /// Returns true if running under MONO. this method is used to workaround known bugs (or functionality) in the MONO framework,
        /// and calls to this method should be rechecked from time to time to verify if the bugs have not ben fixed in the latest version.
        /// </summary>
        /// <returns>Whether we are running under MONO or not.</returns>
        public static bool IsMonoRunning()
        {
            return Type.GetType("Mono.Runtime") != null;
        }

        //Repeated in BitOps, should be moved to a neutral place.
        internal static bool CompareMem(byte[] a1, byte[] a2)
        {
            if (a1 == null)
            {
                if ((a2 == null) || (a2.Length == 0)) return true; else return false;
            }

            if (a2 == null)
            {
                if (a1.Length == 0) return true; else return false;
            }

            if (a1.Length != a2.Length) return false;
            for (int i = 0; i < a1.Length; i++)
                if (a1[i] != a2[i]) return false;
            return true;
        }

        internal static bool CompareMem(byte[] a1, byte[] a2, int a2Pos)
        {
            for (int i = 0; i < a1.Length; i++)
                if (a1[i] != a2[i + a2Pos]) return false;
            return true;
        }

        internal static bool CompareMem(byte[] a1, int a1Pos, byte[] a2, int a2Pos, int len)
        {
            if (a1Pos + len > a1.Length) return false;
            if (a2Pos + len > a2.Length) return false;
            for (int i = 0; i < len; i++)
                if (a1[i + a1Pos] != a2[i + a2Pos]) return false;
            return true;
        }


        internal static int CompareObjects(IComparable a1, IComparable a2)
        {
            if (Object.ReferenceEquals(a1, null))
            {
                if (Object.ReferenceEquals(a2, null)) return 0;
                return -1;
            }

            if (Object.ReferenceEquals(a2, null)) return 1;
            return a1.CompareTo(a2);
        }
        internal static int CompareArray<T>(T[] a1, T[] a2) where T: IComparable
        {
            if (a1 == null)
            {
                if (a2 == null) return 0;
                return -1;
            }
            if (a2 == null) return 1;

            int Result = a1.Length.CompareTo(a2.Length);
            if (Result != 0) return Result;

            for (int i = 0; i < a1.Length; i++)
            {
                if (a1[i] == null)
                {
                    if (a2[i] != null) return -1;
                }
                else
                {
                    Result = a1[i].CompareTo(a2[i]);
                    if (Result != 0) return Result;
                }
            }

            return 0;
        }

        internal static int CompareArray(int[] a1, int[] a2)
        {
            if (a1 == null)
            {
                if (a2 == null) return 0;
                return -1;
            }
            if (a2 == null) return 1;

            int Result = a1.Length.CompareTo(a2.Length);
            if (Result != 0) return Result;

            for (int i = 0; i < a1.Length; i++)
            {
                Result = a1[i].CompareTo(a2[i]);
                if (Result != 0) return Result;
            }

            return 0;
        }

        internal static int CompareArray(byte[] a1, byte[] a2)
        {
            if (a1 == null)
            {
                if (a2 == null) return 0;
                return -1;
            }
            if (a2 == null) return 1;

            int Result = a1.Length.CompareTo(a2.Length);
            if (Result != 0) return Result;

            for (int i = 0; i < a1.Length; i++)
            {
                Result = a1[i].CompareTo(a2[i]);
                if (Result != 0) return Result;
            }

            return 0;
        }

        internal static int CompareArray(double[] a1, double[] a2)
        {
            if (a1 == null)
            {
                if (a2 == null) return 0;
                return -1;
            }
            if (a2 == null) return 1;

            int Result = a1.Length.CompareTo(a2.Length);
            if (Result != 0) return Result;

            for (int i = 0; i < a1.Length; i++)
            {
                Result = a1[i].CompareTo(a2[i]);
                if (Result != 0) return Result;
            }

            return 0;
        }

#if (FRAMEWORK40)
        [SecuritySafeCritical]
#else
#endif
        internal static bool HasUnamanagedPermissions()
        {
#if(COMPACTFRAMEWORK || FULLYMANAGED || SILVERLIGHT)
            return false;
#else
            try
            {
                new SecurityPermission(SecurityPermissionFlag.UnmanagedCode).Demand();
            }
            catch (SecurityException)
            {
                return false;
            }
            return true;
#endif
        }

        internal static void TryDelete(string fileName)
        {
            try
            {
                if (fileName != null) File.Delete(fileName);
            }
            catch (IOException)
            {
                //If we can't delete it, ignore it..
            }
        }
    }

    internal sealed class CharUtils
    {
        private CharUtils() { }

        public static void SameOrLess(string s, int start, ref int len)
        {
            int posi = start + len - 1;  //if start is 0 and len = 1, we want to test char 0.
            if (s == null || posi < 1 || posi >= s.Length) return;
            if (PreviousIsSurrogatePair(s, posi)) len--;
        }

        public static void SameOrMore(string s, int start, ref int len)
        {
            int posi = start + len - 1;  //if start is 0 and len = 1, we want to test char 0.
            if (s == null || posi < 0 || posi + 1 >= s.Length) return;
            if (IsSurrogatePair(s, posi)) len++;
        }

        public static bool PreviousIsSurrogatePair(string s, int index)
        {
            if (s == null || index < 1 || index >= s.Length) return false;
            return IsSurrogatePair(s, index - 1);
        }

        public static bool IsSurrogatePair(string s, int index)
        {
#if(FRAMEWORK20)
            return Char.IsSurrogatePair(s, index);
#else
			if (s == null)
			{
				throw new ArgumentNullException("s");
			}
			if ((index < 0) || (index >= s.Length))
			{
				throw new ArgumentOutOfRangeException("index");
			}

			if (s[index] < 0xD800 || s[index] > 0xDBFF)	return false;
			
			if (index + 1 >= s.Length) return false;

			return (s[index+1] >= 0xDC00 && s[index+1] <= 0xDFFF);
#endif
        }


        // from http://www.unicode.org/Public/MAPPINGS/VENDORS/MICSFT/WINDOWS/CP1252.TXT
        public static bool IsWin1252(int chuni)
        {
            if (chuni <= 0x7F) return true;
            if (chuni >= 0xA0 && chuni <= 0xFF) return true;

            switch (chuni)
            {
                case 0x20AC:
                case 0x201A:
                case 0x0192:
                case 0x201E:
                case 0x2026:
                case 0x2020:
                case 0x2021:
                case 0x02C6:
                case 0x2030:
                case 0x0160:
                case 0x2039:
                case 0x0152:
                case 0x017D:
                case 0x2018:
                case 0x2019:
                case 0x201C:
                case 0x201D:
                case 0x2022:
                case 0x2013:
                case 0x2014:
                case 0x02DC:
                case 0x2122:
                case 0x0161:
                case 0x203A:
                case 0x0153:
                case 0x017E:
                case 0x0178:
                    return true;
            }
            return false;

        }

        public static bool IsWin1252(string s)
        {
            foreach (char c in s)
            {
                if (!IsWin1252((int)c)) return true;
            }
            return false;
        }

        public static byte GetWin1252Bytes_PDF(int c)
        {
            switch (c)
            {
                case 0x7F: return 31;  //this is the same as unicode, but it is a control character in both.

                case 0x20AC: return 0x80;
                case /*not defined*/0x81: return 31;
                case 0x201A: return 0x82;
                case 0x0192: return 0x83;
                case 0x201E: return 0x84;
                case 0x2026: return 0x85;
                case 0x2020: return 0x86;
                case 0x2021: return 0x87;
                case 0x02C6: return 0x88;
                case 0x2030: return 0x89;
                case 0x0160: return 0x8A;
                case 0x2039: return 0x8B;
                case 0x0152: return 0x8C;
                case /*not defined*/0x8D: return 31;
                case 0x017D: return 0x8E;
                case /*not defined*/0x8F: return 31;
                case /*not defined*/0x90: return 31;
                case 0x2018: return 0x91;
                case 0x2019: return 0x92;
                case 0x201C: return 0x93;
                case 0x201D: return 0x94;
                case 0x2022: return 0x95;
                case 0x2013: return 0x96;
                case 0x2014: return 0x97;
                case 0x02DC: return 0x98;
                case 0x2122: return 0x99;
                case 0x0161: return 0x9A;
                case 0x203A: return 0x9B;
                case 0x0153: return 0x9C;
                case /*not defined*/0x9D: return 31;
                case 0x017E: return 0x9E;
                case 0x0178: return 0x9F;

                default:
                    if (c < 32 || c > 255) return 31;
                    else
                        return (byte)c;
            }
        }

        public static byte[] GetWin1252Bytes_PDF(string s)
        {
            byte[] Result = new byte[s.Length];
            for (int i = 0; i < s.Length; i++)
            {
                int c = s[i];
                {
                    Result[i] = GetWin1252Bytes_PDF(c);
                }
            }

            return Result;
        }

        static readonly int[] W1252ToUni =
			{
				/* 0x80 */	0x20AC,
				/* 0x81 */	/*not defined*/0x81,
				/* 0x82 */	0x201A,
				/* 0x83 */	0x0192,
				/* 0x84 */	0x201E,
				/* 0x85 */	0x2026,
				/* 0x86 */	0x2020,
				/* 0x87 */	0x2021,
				/* 0x88 */	0x02C6,
				/* 0x89 */	0x2030,
				/* 0x8A */	0x0160,
				/* 0x8B */	0x2039,
				/* 0x8C */	0x0152,
				/* 0x8D */	/*not defined*/0x8D,
				/* 0x8E */	0x017D,
				/* 0x8F */	/*not defined*/0x8F,
				/* 0x90 */	/*not defined*/0x90,
				/* 0x91 */	0x2018,
				/* 0x92 */	0x2019,
				/* 0x93 */	0x201C,
				/* 0x94 */	0x201D,
				/* 0x95 */	0x2022,
				/* 0x96 */	0x2013,
				/* 0x97 */	0x2014,
				/* 0x98 */	0x02DC,
				/* 0x99 */	0x2122,
				/* 0x9A */	0x0161,
				/* 0x9B */	0x203A,
				/* 0x9C */	0x0153,
				/* 0x9D */	/*not defined*/0x9D,
				/* 0x9E */	0x017E,
				/* 0x9F */	0x0178
			};

        public static char GetUniFromWin1252(byte b)
        {
            if (b >= 0x80 && b - 0x80 < W1252ToUni.Length) return (char)W1252ToUni[b - 0x80];
            return (char)b;
        }

        static readonly int[] W1252ToUniPDF =
			{
				/* 0x80 */	0x20AC,
				/* 0x81 */	/*not defined*/31,
				/* 0x82 */	0x201A,
				/* 0x83 */	0x0192,
				/* 0x84 */	0x201E,
				/* 0x85 */	0x2026,
				/* 0x86 */	0x2020,
				/* 0x87 */	0x2021,
				/* 0x88 */	0x02C6,
				/* 0x89 */	0x2030,
				/* 0x8A */	0x0160,
				/* 0x8B */	0x2039,
				/* 0x8C */	0x0152,
				/* 0x8D */	/*not defined*/31,
				/* 0x8E */	0x017D,
				/* 0x8F */	/*not defined*/31,
				/* 0x90 */	/*not defined*/31,
				/* 0x91 */	0x2018,
				/* 0x92 */	0x2019,
				/* 0x93 */	0x201C,
				/* 0x94 */	0x201D,
				/* 0x95 */	0x2022,
				/* 0x96 */	0x2013,
				/* 0x97 */	0x2014,
				/* 0x98 */	0x02DC,
				/* 0x99 */	0x2122,
				/* 0x9A */	0x0161,
				/* 0x9B */	0x203A,
				/* 0x9C */	0x0153,
				/* 0x9D */	/*not defined*/31,
				/* 0x9E */	0x017E,
				/* 0x9F */	0x0178
			};

        public static char GetUniFromWin1252_PDF(byte b)
        {
            if (b < 31) return (char)31;
            if (b >= 0x80 && b - 0x80 < W1252ToUniPDF.Length) return (char)W1252ToUniPDF[b - 0x80];
            return (char)b;
        }

    }
    #endregion

    #region Custom Formula Functions

    internal class TUserDefinedFunctionContainer
    {
        internal TUserDefinedFunctionLocation Location;
        internal TUserDefinedFunction Function;

        internal TUserDefinedFunctionContainer(TUserDefinedFunctionLocation aLocation, TUserDefinedFunction aFunction)
        {
            Location = aLocation;
            Function = aFunction;
        }
    }

    internal sealed class TCustomFormulaList
    {
        Dictionary<string, TUserDefinedFunctionContainer> FList;
        Dictionary<string, TUserDefinedFunctionContainer> FListDisplay;

        public TCustomFormulaList()            
        {
            FList = new Dictionary<string, TUserDefinedFunctionContainer>(StringComparer.InvariantCultureIgnoreCase);
            FListDisplay = new Dictionary<string, TUserDefinedFunctionContainer>(StringComparer.InvariantCultureIgnoreCase);
        }

        public void Add(TUserDefinedFunctionContainer fn)
        {
            FList.Add(fn.Function.InternalName, fn);
            if (fn.Function.Name != fn.Function.InternalName) FListDisplay.Add(fn.Function.Name, fn);
        }

        public TUserDefinedFunctionContainer GetValue(string Key)
        {
            TUserDefinedFunctionContainer Result;
            if (!FList.TryGetValue(Key, out Result))
                return null;
            return Result;
        }

        public TUserDefinedFunctionContainer GetValueFromDisplayName(string Key)
        {
            TUserDefinedFunctionContainer Result;
            if (!FListDisplay.TryGetValue(Key, out Result))
            {
                if (!FList.TryGetValue(Key, out Result))
                    return null;

                if (Result.Function.InternalName != Result.Function.InternalName) return null;
            }
            return Result;
        }

        public void Clear()
        {
            FList.Clear();
            FListDisplay.Clear();
        }
    }


    /// <summary>
    /// This interface is passed to methods that will process the workbook. Implement your own custom decendant to create new functions.
    /// </summary>
    public interface IUserDefinedFunctionAggregator
    {
        /// <summary>
        /// Implement this method to do something for every value in the range. <br>
        /// </br>Note: You can abort the processing (for example if there is an error) by returning false in this function, and the error value in the "error" parameter.
        /// But, Excel normally doesn't behave this way. Excel will normally first check all values to see if there is an error, and only then
        /// if no cell was a #Err! value, do other checks (for example negative parameters). So, for this to be 100% like Excel, you always need to return true in 
        /// this function, and check for other errors only after all values have been processed. If you don't care about returning the exact error message Excel returns, you can return 
        /// the error directly here while you are processing the values and speed up things (since other values won't be processed after you know the first error).
        /// </summary>
        /// <param name="value">Value that will be processed.</param>
        /// <param name="error">Return an error here when the method returns false. If the method returns true, this parameter is undefined.</param>
        /// <returns>False if you want to abort, true to continue.</returns>
        bool Process(double value, out TFlxFormulaErrorValue error);
    }

    /// <summary>
    /// Defines how custom functions are added to the recalculation engine.
    /// If a function is defined in both Global and Local scope, Local scope will be used.
    /// </summary>
    public enum TUserDefinedFunctionScope
    {
        /// <summary>
        /// Function will be added for all instances of FlexCel. You will normally use this setting only once at the beginning of your application.
        /// </summary>
        Global,

        /// <summary>
        /// Functions will be available only for the instance where they were added. Use this option if you might have different custom functions with the same name in different 
        /// spreadsheets, and adding the function globally would clash. You will normally use this setting when adding the formulas after creating the ExcelFile
        /// instances.
        /// </summary>
        Local
    }

    /// <summary>
    /// Defines where the custom function is located, if inside a macro in the same file, or inside a macro in an external file.
    /// </summary>
    public enum TUserDefinedFunctionLocation
    {
        /// <summary>
        /// The custom function is defined in an external file or addin. Whenever you add this function to a workbook, references will be created to an external function.
        /// </summary>
        External,

        /// <summary>
        /// The custom function is defined inside the same file where the formula is. Whenever you add this function to a workbook, references will be created to a
        /// macro in the same file.
        /// </summary>
        Internal
    }

    /// <summary>
    /// Encapsulates the parameters to send to an User Defined Function for evaluation.
    /// </summary>
    public class TUdfEventArgs
    {
        #region Privates
        private ExcelFile FXls;
        private int FSheet;
        private int FRow;
        private int FCol;
        #endregion

        /// <summary>
        /// Creates a new instance.
        /// </summary>
        /// <param name="aXls">See <see cref="Xls"/></param>
        /// <param name="aSheet">See <see cref="Sheet"/></param>
        /// <param name="aRow">See <see cref="Row"/></param>
        /// <param name="aCol">See <see cref="Col"/></param>
        public TUdfEventArgs(ExcelFile aXls, int aSheet, int aRow, int aCol)
        {
            FXls = aXls;
            FSheet = aSheet;
            FRow = aRow;
            FCol = aCol;
        }

        /// <summary>
        /// ExcelFile that has the formula being evaluated. You might change its ActiveSheet propery inside this method and there is no
        /// need to restore it back.
        /// </summary>
        public ExcelFile Xls { get { return FXls; } }

        /// <summary>
        /// Index of the sheet where the formula is located. This value only has meaning when evaluating formulas in cells. (Not when for example evaluating formulas inside named ranges)
        /// </summary>
        public int Sheet { get { return FSheet; } }

        /// <summary>
        /// Row index where the formula is located. This value only has meaning when evaluating formulas in cells. (Not when for example evaluating formulas inside named ranges)
        /// </summary>
        public int Row { get { return FRow; } }

        /// <summary>
        /// Column index where the formula is located. This value only has meaning when evaluating formulas in cells. (Not when for example evaluating formulas inside named ranges)
        /// </summary>
        public int Col { get { return FCol; } }

    }

    /// <summary>
    /// Inherit from this class to create your own user defined functions. Make sure you read the pdf documentation to get more information on what user defined functions are
    /// and how they are created.
    /// </summary>
    public abstract class TUserDefinedFunction
    {
        #region Privates
        private string FName;
        private string FInternalName;
        #endregion

        /// <summary>
        /// Initializes the name of the user defined function.
        /// </summary>
        /// <param name="aName">Name to be used in the user defined function. This is the same name that should be in the xls file.</param>
        protected TUserDefinedFunction(string aName)
        {
            FName = aName;
            FInternalName = aName;
        }

        /// <summary>
        /// Initializes the name of the user defined function, with an special name for older Excel versions.
        /// </summary>
        /// <param name="aName">Name to be used in the user defined function. This is the same name that should be in the xls file.</param>
        /// <param name="aInternalName">Name that will be used when saving xls (biff8) files. Some functions are saved by Excel 2010 as .xlfn_Name when
        /// saving xls (not xlsx). This is the name that should be saved in the xls file, not the real name of the function.</param>
        protected TUserDefinedFunction(string aName, string aInternalName)
        {
            FName = aName;
            FInternalName = aInternalName;
        }

        /// <summary>
        /// Override this method to provide your own implementation on the function.<br/>
        /// If this method throws an exception, it will not be handled and the recalculation will be aborted. So if you want to return an error, return a <see cref="TFlxFormulaErrorValue"/> value.
        /// <br/>
        /// <b>Do not use any global variable in this method</b>, it must be stateless and always return the same value when called with the same arguments.
        /// <br/>
        /// See the PDF documentation for more information.
        /// </summary>
        /// <param name="arguments">Extra objects you can use to evaluate the function.</param>
        /// <param name="parameters">Parameters for the function. When this method is called by FlexCel, this parameter will never be null, but might be an array
        /// of zero length if there are no parameters.
        /// <br/> Each parameter in the array will always be one of the following objects:
        /// <list type="bullet">
        /// <item>Null. (Nothing in VB). If the parameter is empty or missing.</item>
        /// <item>A <see cref="System.Boolean"/></item>
        /// <item>A <see cref="System.String"/></item>
        /// <item>A <see cref="System.Double"/></item>
        /// <item>A <see cref="FlexCel.Core.TXls3DRange"/> This will be returned when the argument is a cell reference.</item>
        /// <item>A <see cref="FlexCel.Core.TFlxFormulaErrorValue"/>. Except in very special cases (like an IsError function), the expected
        /// behavior in Excel is that whenever you get an Error parameter your function should return the same error and exit. You can use the method <see cref="CheckParameters"/>
        /// to implement that.</item>
        /// <item>A 2-dimensional array of objects, where each object in the array will be of any of the types mentioned here again.</item> 
        /// </list>
        ///  This class provides utility methods like <see cref="TryGetDouble"/> that will help you get an specific type of object from a parameter.
        /// </param>
        /// <returns>Return any object you want here. Normally a double, a string, a boolean a TFlxFormulaErrorValue or a null. If this method returns a class, it will be converted to an allowed value, normally to a string. </returns>
        public abstract object Evaluate(TUdfEventArgs arguments, object[] parameters);

        /// <summary>
        /// Name that will be assigned to the function.
        /// </summary>
        public string Name { get { return FName; } set { FName = value; } }

        /// <summary>
        /// Name that will be used when saving xls (biff8) files. Some functions are saved by Excel 2010 as .xlfn_Name when
        /// saving xls (not xlsx). This is the name that should be saved in the xls file, not the real name of the function.       
        /// </summary>
        public string InternalName { get { return FInternalName; } set { FInternalName = value; } }


        #region Utility functions for retrieving parameters.

        /// <summary>
        /// Checks that the parameter array has the expected number of arguments, and that no one is an Error. If any argument is an error
        /// it is returned in ResultError, since the default in Excel is to stop processing arguments in a function when one is an error.
        /// </summary>
        /// <param name="parameters">Array of parameters to check.</param>
        /// <param name="expectedCount">Number of parameters expected. If this number is variable, specify -1 here.</param>
        /// <param name="ResultError">Returns the error in the parameters. This parameter si only valid if this function returns false.</param>
        /// <returns>True if all parameters are correct, false otherwise.</returns>
        public static bool CheckParameters(object[] parameters, int expectedCount, out TFlxFormulaErrorValue ResultError)
        {
            //parameters is never null, but it is not bad practice to check it anyway.
            if (parameters == null || (expectedCount >= 0 && parameters.Length != expectedCount))
            {
                ResultError = TFlxFormulaErrorValue.ErrValue; //ErrValue is the standard Excel error with invalid parameters.
                return false;
            }

            //check that all the parameters are valid. If any of them is an error, we should return that error.
            foreach (object p in parameters)
            {
                if (p is TFlxFormulaErrorValue)
                {
                    ResultError = (TFlxFormulaErrorValue)p;
                    return false;
                }
            }

            ResultError = TFlxFormulaErrorValue.ErrNA;
            return true;

        }


        /// <summary>
        /// Returns a single value from a parameter. 
        /// If the parameter is a cell range and the cell range has only one cell, this method will return the
        /// value of the cell, else it will return an error.
        /// </summary>
        /// <param name="xls">ExcelFile that will be used to read the value if param is a cell reference.</param>
        /// <param name="param">One of the parameters passed to <see cref="Evaluate"/></param>
        /// <returns></returns>
        public static object GetSingleParameter(ExcelFile xls, object param)
        {
            //if it is a range, it must have only one cell.
            TXls3DRange NameRange = param as TXls3DRange;
            if (NameRange == null) return param;

            if (!NameRange.IsOneCell) return TFlxFormulaErrorValue.ErrValue;

            return xls.GetCellValueAndRecalc(NameRange.Sheet1, NameRange.Top, NameRange.Left, new TCalcState(), new TCalcStack());
        }

        /// <summary>
        /// Tries to retrieve a double from a parameter, and return it if it can be converted or an error if not.
        /// </summary>
        /// <param name="xls">ExcelFile used to read the parameter when it is a cell reference.</param>
        /// <param name="param">One of the parameters passed to <see cref="Evaluate"/></param>
        /// <param name="ResultValue">Value of the parameter as double. If TryGetDouble returns false, this value is undefined.</param>
        /// <param name="ResultError">Value of the error when converting the parameter. If TryGetDouble returns true (there was no error), this value is undefined.</param>
        /// <returns>True if the parameter can be converted to a double, false if there was an error.</returns>
        public static bool TryGetDouble(ExcelFile xls, object param, out double ResultValue, out TFlxFormulaErrorValue ResultError)
        {
            object r = GetSingleParameter(xls, param);
            if (r is TFlxFormulaErrorValue)
            {
                ResultValue = 0;
                ResultError = (TFlxFormulaErrorValue)r;
                return false;
            }

            ResultError = TFlxFormulaErrorValue.ErrValue;
            return TBaseParsedToken.ExtToDouble(r, out ResultValue);
        }

        /// <summary>
        /// Tries to retrieve a date/time from a parameter, and return it if it can be converted or an error if not.
        /// </summary>
        /// <param name="xls">ExcelFile used to read the parameter when it is a cell reference.</param>
        /// <param name="param">One of the parameters passed to <see cref="Evaluate"/></param>
        /// <param name="SuppressTime">If true, the date returned will have the time set to 0:00 no matter the real value.</param>
        /// <param name="ResultValue">Value of the parameter as double. If TryGetDate returns false, this value is undefined.</param>
        /// <param name="ResultError">Value of the error when converting the parameter. If TryGetDate returns true (there was no error), this value is undefined.</param>
        /// <returns>True if the parameter can be converted to a datetime, false if there was an error.</returns>
        public static bool TryGetDate(ExcelFile xls, object param, bool SuppressTime, out DateTime ResultValue, out TFlxFormulaErrorValue ResultError)
        {
            double dResult;
            ResultValue = DateTime.MinValue;
            ResultError = TFlxFormulaErrorValue.ErrNA;

            if (!TryGetDouble(xls, param, out dResult, out ResultError)) return false;
            if (SuppressTime) dResult = Math.Floor(dResult);
            if (dResult < 0 || !FlxDateTime.TryFromOADate(dResult, xls.OptionsDates1904, out ResultValue))
            {
                ResultError = TFlxFormulaErrorValue.ErrValue;
                return false;
            }

            return true;
        }

        /// <summary>
        /// Tries to retrieve a list of double arguments from the parameters, starting at parameter startParam.
        /// Use this method for functions that accept a range of numeric values as an entry. (for example =Sum(a1:a10))
        /// </summary>
        /// <param name="Xls">ExcelFile used to read the parameter when it is a cell reference.</param>
        /// <param name="parameters">The parametes passed to the function.</param>
        /// <param name="startParam">First parameter we want to evaluate.</param>
        /// <param name="endParam">Last parameter we want to evaluate. If &lt; 0, it will evaluate all parameters from startParam to parameters.Length</param>
        /// <param name="agg">A class decending from TUserDefinedFunctionAggregator that will process the values for every entry in the range.</param>
        /// <param name="Err">Value of the error when converting the parameter. If TryGetDouble returns true (there was no error), this value is undefined.</param>
        /// <returns></returns>
        public bool TryGetDoubleList(ExcelFile Xls, object[] parameters, int startParam, int endParam, IUserDefinedFunctionAggregator agg, out TFlxFormulaErrorValue Err)
        {
            if (endParam < 0) endParam = parameters.Length;
            for (int i = startParam; i < endParam; i++)
            {
                TXls3DRange ResultRange;
                if (TryGetCellRange(parameters[i], out ResultRange, out Err))
                {
                    if (!ProcessRange(Xls, ResultRange, agg, out Err))
                    {
                        return false;
                    }
                    continue;
                }

                object[,] ObjArray;
                if (TryGetArray(Xls, parameters[i], out ObjArray, out Err))
                {
                    for (int r = 0; r < ObjArray.GetLength(0); r++)
                    {
                        for (int c = 0; c < ObjArray.GetLength(1); c++)
                        {
                            object val = ObjArray[r, c];
                            if (val is TFlxFormulaErrorValue)
                            {
                                Err = (TFlxFormulaErrorValue)val;
                                return false;
                            }

                            double ResultValue;
                            if (TBaseParsedToken.ExtToDouble(val, out ResultValue))
                            {
                                if (!agg.Process(ResultValue, out Err)) return false;
                            }
                            else
                            {
                                Err = TFlxFormulaErrorValue.ErrValue;
                                return false;
                            }
                        }
                    }
                    continue;
                }



                double dNum0;
                if (!TryGetDouble(Xls, parameters[i], out dNum0, out Err)) return false;
                if (!agg.Process(dNum0, out Err)) return false;

            }

            Err = TFlxFormulaErrorValue.ErrNA;
            return true;
        }

        private bool ProcessRange(ExcelFile Xls, TXls3DRange ResultRange, IUserDefinedFunctionAggregator agg, out TFlxFormulaErrorValue Err)
        {
            for (int s = ResultRange.Sheet1; s <= ResultRange.Sheet2; s++)
            {
                int RowCount = Xls.GetRowCount(s);

                for (int r = ResultRange.Top; r <= ResultRange.Bottom; r++)
                {
                    if (r > RowCount) continue;
                    for (int c = ResultRange.Left; c <= ResultRange.Right; c++)
                    {
                        object val = Xls.GetCellValueAndRecalc(s, r, c, new TCalcState(), new TCalcStack());

                        if (val is TFlxFormulaErrorValue)
                        {
                            Err = (TFlxFormulaErrorValue)val;
                            return false;
                        }

                        if (val is double) //we will only process numeric values.
                        {
                            if (!agg.Process((double)val, out Err)) return false;
                        }

                        else if (val == null)
                        {
                            if (!agg.Process(0, out Err)) return false;

                        }
                        else
                        {
                            Err = TFlxFormulaErrorValue.ErrValue;
                            return false;
                        }
                    }
                }
            }

            Err = TFlxFormulaErrorValue.ErrNA;
            return true;
        }


        /// <summary>
        /// Tries to retrieve a string from a parameter, and return it if it can be converted or an error if not.
        /// </summary>
        /// <param name="xls">ExcelFile used to read the parameter when it is a cell reference.</param>
        /// <param name="param">One of the parameters passed to <see cref="Evaluate"/></param>
        /// <param name="ResultValue">Value of the parameter as string. If TryGetString returns false, this value is undefined.</param>
        /// <param name="ResultError">Value of the error when converting the parameter. If TryGetString returns true (there was no error), this value is undefined.</param>
        /// <returns>True if the parameter can be converted to a string, false if there was an error.</returns>
        public static bool TryGetString(ExcelFile xls, object param, out string ResultValue, out TFlxFormulaErrorValue ResultError)
        {
            object r = GetSingleParameter(xls, param);
            if (r is TFlxFormulaErrorValue)
            {
                ResultError = (TFlxFormulaErrorValue)r;
                ResultValue = null;
                return false;
            }
            else ResultError = TFlxFormulaErrorValue.ErrValue;

            ResultValue = Convert.ToString(r);
            if (ResultValue != null)
            {
                ResultError = TFlxFormulaErrorValue.ErrNA;
                return true;
            }
            return false;
        }

        /// <summary>
        /// Tries to retrieve a boolean from a parameter, and return it if it can be converted or an error if not.
        /// </summary>
        /// <param name="xls">ExcelFile used to read the parameter when it is a cell reference.</param>
        /// <param name="param">One of the parameters passed to <see cref="Evaluate"/></param>
        /// <param name="ResultValue">Value of the parameter as a boolean. If TryGetBoolean returns false, this value is undefined.</param>
        /// <param name="ResultError">Value of the error when converting the parameter. If TryGetBoolean returns true (there was no error), this value is undefined.</param>
        /// <returns>True if the parameter can be converted to a boolean, false if there was an error.</returns>
        public static bool TryGetBoolean(ExcelFile xls, object param, out bool ResultValue, out TFlxFormulaErrorValue ResultError)
        {
            object r = GetSingleParameter(xls, param);
            if (r is TFlxFormulaErrorValue)
            {
                ResultValue = false;
                ResultError = (TFlxFormulaErrorValue)r;
                return false;
            }

            ResultError = TFlxFormulaErrorValue.ErrValue;
            return TBaseParsedToken.ExtToBool(r, out ResultValue);
        }

        /// <summary>
        /// Tries to retrieve a cell range from a parameter, and return it if it can be converted or an error if not.
        /// </summary>
        /// <param name="param">One of the parameters passed to <see cref="Evaluate"/></param>
        /// <param name="ResultValue">Value of the parameter as a TXls3DRange. If TryGetRange returns false, this value is undefined.</param>
        /// <param name="ResultError">Value of the error when converting the parameter. If TryGetCellRange returns true (there was no error), this value is undefined.</param>
        /// <returns>True if the parameter can be converted to a cell range, false if there was an error.</returns>
        public static bool TryGetCellRange(object param, out TXls3DRange ResultValue, out TFlxFormulaErrorValue ResultError)
        {
            ResultValue = param as TXls3DRange;
            if (ResultValue != null)
            {
                ResultError = TFlxFormulaErrorValue.ErrNA;
                return true;
            }

            ResultError = TFlxFormulaErrorValue.ErrValue;
            return false;
        }

        /// <summary>
        /// Tries to retrieve an array from a parameter, and return it if it can be converted or an error if not.
        /// </summary>
        /// <param name="xls">ExcelFile used to read the parameter when it is a cell reference.</param>
        /// <param name="param">One of the parameters passed to <see cref="Evaluate"/></param>
        /// <param name="ResultValue">Value of the parameter as an array. If TryGetArray returns false, this value is undefined.</param>
        /// <param name="ResultError">Value of the error when converting the parameter. If TryGetArray returns true (there was no error), this value is undefined.</param>
        /// <returns>True if the parameter can be converted to an array, false if there was an error.</returns>
        public static bool TryGetArray(ExcelFile xls, object param, out object[,] ResultValue, out TFlxFormulaErrorValue ResultError)
        {
            ResultValue = param as object[,];
            if (ResultValue != null)
            {
                ResultError = TFlxFormulaErrorValue.ErrNA;
                return true;
            }

            ResultError = TFlxFormulaErrorValue.ErrValue;
            return false;
        }


        #endregion

    }
    #endregion

    #region Exceptions Options
    /// <summary>
    /// Enumerates what to do on different FlexCel error situations.
    /// </summary>
    [Flags]
    public enum TExcelFileErrorActions
    {
        /// <summary>
        /// FlexCel will try to recover from most errors.
        /// </summary>
        None = 0,

        /// <summary>
        /// When true and the number of manual pagebreaks is bigger than the maximum Excel allows,
        /// an Exception will be raised. When false, the page break will be silently ommited.
        /// Note that This exception is raised when saving the file as xls, when you are exporting your report to 
        /// PDF or images, all page breaks will be used.
        /// </summary>
        ErrorOnTooManyPageBreaks = 1,

        /// <summary>
        /// When true, FlexCel will complain when you try to set a formula that has a string constant bigger than 255 characters.
        /// <br/>For example, the formula: '="very long string that has more than 255 characters...." &amp; "other string" '  would raise an Exception,
        /// since Excel won't allow it. Note that you can still use ' =a1 &amp; "other string" ' where the cell A1 has the value:
        /// "very long string that has more than 255 characters....". this restriction applies only to inline strings.<br/>
        /// Note that when this property is false you will still get the error, but only when saving to xls, xlsx or other file formats that don't support longer strings. (this error is too important to be ignored)
        /// </summary>
        ErrorOnFormulaConstantTooLong = 2,

        /// <summary>
        /// When true and the row height is bigger than the maximum allowed by Excel, you will get an Exception.
        /// </summary>
        ErrorOnRowHeightTooBig = 4,

        /// <summary>
        /// If this is true and the xlsx file contains an invalid name, an exception will be thrown.
        /// </summary>
        ErrorOnXlsxInvalidName = 8,

        /// <summary>
        /// If this is true and the xlsx file contains a missing part (like an image), an exception will be thrown.
        /// </summary>
        ErrorOnXlsxMissingPart = 16,

        /// <summary>
        /// Sets all error actions together.
        /// </summary>
        All = ErrorOnTooManyPageBreaks | ErrorOnFormulaConstantTooLong | ErrorOnRowHeightTooBig | ErrorOnXlsxInvalidName 
             | ErrorOnXlsxMissingPart

    }
    #endregion

    #region UndisposableStream
    /// <summary>
    /// A stream that will not be closed when you call close. Useful to pass to a StreamReader/StreamWriter and not have the main stream closed.
    /// </summary>
    internal class TUndisposableStream : Stream
    {
        private Stream FStream;

        internal TUndisposableStream(Stream aStream)
        {
            FStream = aStream;
        }

        public override bool CanRead
        {
            get
            {
                return FStream.CanRead;
            }
        }

        public override bool CanSeek
        {
            get
            {
                return FStream.CanSeek;
            }
        }

        public override bool CanWrite
        {
            get
            {
                return FStream.CanWrite;
            }
        }

        public override void Flush()
        {
            FStream.Flush();
        }

        public override long Length
        {
            get
            {
                return FStream.Length;
            }
        }

        public override long Position
        {
            get
            {
                return FStream.Position;
            }
            set
            {
                FStream.Position = value;
            }
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            return FStream.Read(buffer, offset, count);
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            return FStream.Seek(offset, origin);
        }

        public override void SetLength(long value)
        {
            FStream.SetLength(value);
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            FStream.Write(buffer, offset, count);
        }

        public override void Close()
        {
            base.Close();
            FStream.Flush();
            //Do not close the stream!
        }
    }
    #endregion

    #region Workspace
    /// <summary>
    /// This class links together a group of spreadsheets, so you can recalculate among linked spreadsheets. 
    /// In order to use it, just define an object of this class and add all the files you need for the linked recalculation.
    /// If you don't know in advance which files you will need, you can use the <see cref="LoadLinkedFile"/> event.
    /// <br/>Note that whenever you recalculate any file in the workspace, all files will be recalculated, so you don't need to calculate them twice.
    /// </summary>
    /// <remarks>Files are case insensitive, even if running in mono. "a.xls" is the same as "A.XLS"</remarks>
    /// <example>
    /// If you have 3 files, xls1, xls2 and xls3, you can recalculate them together with the following code:
    /// <br/>
    /// <code>
    /// Workspace work = new Workspace();
    /// work.Add("xls1", xls1);
    /// work.Add("xls2", xls2);
    /// work.Add("xls3", xls3);
    /// xls1.Recalc();  //Either xls1.Recalc, xls2.Recalc or xls3.Recalc will recalculate all the files in the workspace.
    /// </code>
    /// <br/>
    /// <code lang = "vbnet">
    /// Dim work as Workspace = new Workspace
    /// work.Add("xls1", xls1)
    /// work.Add("xls2", xls2)
    /// work.Add("xls3", xls3)
    /// xls1.Recalc  'Either xls1.Recalc, xls2.Recalc or xls3.Recalc will recalculate all the files in the workspace.
    /// </code>
    /// <br/>
    /// <code lang = "Delphi .NET" title = "Delphi .NET">
    /// work := Workspace.Create;
    /// work.Add('xls1', xls1);
    /// work.Add('xls2', xls2);
    /// work.Add('xls3', xls3);
    /// xls1.Recalc;  //Either xls1.Recalc, xls2.Recalc or xls3.Recalc will recalculate all the files in the workspace.
    /// </code>
    /// </example>
    public class TWorkspace : IEnumerable, IEnumerable<ExcelFile>

    {
        #region privates
        Dictionary<string, ExcelFile> WorkbookSearch;
        List<ExcelFile> WorkbookList;

        #endregion

        /// <summary>
        /// Creates a new workspace.
        /// </summary>
        public TWorkspace()
        {
            WorkbookList = new List<ExcelFile>();
            WorkbookSearch = new Dictionary<string, ExcelFile>(StringComparer.InvariantCultureIgnoreCase);
        }

        /// <summary>
        /// Adds a file to the workspace. Whenever you recalculate any file in this workspace, all linked files will be recalculated too.
        /// <b>Note that you can't add two files with the same name or same reference twice to this collection.</b>
        /// </summary>
        /// <param name="FileName">Name that this file will have in the workspace. When recalculating a linked formula, this name will be used.
        /// If you want to include paths in the name you can too, but it is normally not needed since FlexCel will search for the simple filename anyway.
        /// You only need to include paths if you have 2 files with the same name in different folders used in the recalculation.</param>
        /// <param name="xls">Excel file to add.</param>
        public void Add(string FileName, ExcelFile xls)
        {
            if (WorkbookSearch.ContainsKey(FileName)) FlxMessages.ThrowException(FlxErr.ErrDuplicatedLinkedFile, FileName);

            //This is a bit slow, but there shouldn't be that much files anyway, and it is important to make sure you don't add the same reference twice.
            foreach (ExcelFile xl in WorkbookList)
            {
                if ((object)xl == (object)xls) FlxMessages.ThrowException(FlxErr.ErrDuplicatedLinkedFile, FileName);
            }

            WorkbookSearch[FileName] = xls;
            WorkbookList.Add(xls);
            xls.Workspace = this;
        }

        /// <summary>
        /// Number of linked files in this workspace.
        /// </summary>
        public int Count
        {
            get
            {
                return WorkbookList.Count;
            }
        }

        /// <summary>
        /// Returns the file at index. 
        /// </summary>
        /// <param name="index">Index of the file in the workspace. (1 based)</param>
        /// <returns></returns>
        public ExcelFile GetFile(int index)
        {
            return WorkbookList[index - 1];
        }

        /// <summary>
        /// Removes all files from the workspace.
        /// </summary>
        public void Clear()
        {
            foreach (ExcelFile xls in WorkbookList)
            {
                xls.Workspace = null;
            }

            WorkbookList.Clear();
            WorkbookSearch.Clear();
        }

        /// <summary>
        /// Use this method to force a recalculation of all the spreadsheets in the workspace. This is the same as calling Recalc() in any of the files in the workspace.
        /// </summary>
        /// <param name="forced">When true this method will perform a recalc in all files. When false, it will only recalculate the files where there has been a change. </param>
        public void Recalc(bool forced)
        {
            InternalRecalc(forced, null);
        }

        /// <summary>
        /// This method will do the same as <see cref="ExcelFile.RecalcAndVerify()"/>, but for a workspace of files.
        /// </summary>
        /// <returns>A list with the unsupported functions in the workspace.</returns>
        public TUnsupportedFormulaList RecalcAndVerify()
        {
            TUnsupportedFormulaList UnsupportedFormulas = new TUnsupportedFormulaList();
            InternalRecalc(true, UnsupportedFormulas);
            return UnsupportedFormulas;
        }

        private static void PrepareFile(ExcelFile xls)
        {
            xls.CleanFlags();
            xls.SetRecalculating(true);
        }

        private void InternalRecalc(bool forced, TUnsupportedFormulaList UnsupportedFormulas)
        {
            try
            {
                //First clean the recalculated flag in *all* sheets.
                for (int i = 0; i < WorkbookList.Count; i++)
                {
                    ExcelFile xls = WorkbookList[i];
                    PrepareFile(xls);
                }

                int index = 0;
                while (index < WorkbookList.Count)  //count might increase while recalculating because the LoadLinkedFile event.
                {
                    WorkbookList[index].InternalRecalc(forced, UnsupportedFormulas);
                    index++;
                }
            }
            finally
            {
                for (int i = 0; i < WorkbookList.Count; i++)
                {
                    ExcelFile xls = WorkbookList[i];
                    xls.SetRecalculating(false);
                }
            }
        }

        /// <summary>
        /// Returns the Excel file with the given name. To get the file at a given position, use <see cref="GetFile(int)"/>
        /// </summary>
        public ExcelFile this[string fileName]
        {
            get
            {
                ExcelFile Result;
                if (!WorkbookSearch.TryGetValue(fileName, out Result)) return null;
                return Result;
            }
        }

        #region IEnumerable Members

        /// <summary>
        /// Returns an enumerator with all the files in the Workspace.
        /// </summary>
        /// <returns></returns>
        public System.Collections.IEnumerator GetEnumerator()
        {
            return WorkbookList.GetEnumerator();
        }

        #endregion


        /// <summary>
        /// Use this event to load files to recalculate on demand, if you don't know a priori which linked files you need.
        /// Note that this event will add the new file to the workspace. 
        /// It will only be called once for each file, even if the file is used many times.
        /// </summary>
        public event LoadLinkedFileEventHandler LoadLinkedFile;

        /// <summary>
        /// Replace this event when creating a custom descendant of TWorkspace. See also <see cref="LoadLinkedFile"/>
        /// </summary>
        /// <param name="e"></param>
        protected virtual void OnLoadLinkedFile(LoadLinkedFileEventArgs e)
        {
            if (LoadLinkedFile != null) LoadLinkedFile(this, e);
        }

        internal ExcelFile GetLinkedFile(string aFileName)
        {
            LoadLinkedFileEventArgs e = new LoadLinkedFileEventArgs(aFileName);
            OnLoadLinkedFile(e);
            if (e.Xls != null)
            {
                Add(aFileName, e.Xls);
                PrepareFile(e.Xls);
            }
            return e.Xls;
        }

        #region IEnumerable<ExcelFile> Members
        /// <summary>
        /// Returns an enumerator with the ExcelFile objects in the Workspace.
        /// </summary>
        /// <returns></returns>
        IEnumerator<ExcelFile> IEnumerable<ExcelFile>.GetEnumerator()
        {
            return WorkbookList.GetEnumerator();
        }
        #endregion
    }

    /// <summary>
    /// Arguments passed on <see cref="TWorkspace.LoadLinkedFile"/>
    /// </summary>
    public class LoadLinkedFileEventArgs : EventArgs
    {
        private readonly string FFileName;
        private ExcelFile FXls;

        /// <summary>
        /// Creates a new Argument.
        /// </summary>
        public LoadLinkedFileEventArgs(string aFileName)
        {
            FFileName = aFileName;
            FXls = null;
        }

        /// <summary>
        /// The filename of the file we need. <b>Note:</b> The path of this filename is relative to where the parent file is.
        /// you might need to add the main path to it in order to load the files.
        /// </summary>
        public string FileName
        {
            get { return FFileName; }
        }

        /// <summary>
        /// Use this parameter to return the ExcelFile that corresponds with <see cref="FileName"/>.  If you return null here, it means
        /// that the file was not found and it will result in #REF errors in the formulas that reference that file.
        /// </summary>
        public ExcelFile Xls
        {
            get { return FXls; }
            set { FXls = value; }
        }
    }

    /// <summary>
    /// Delegate for LoadLinkedFile event.
    /// </summary>
    public delegate void LoadLinkedFileEventHandler(object sender, LoadLinkedFileEventArgs e);


    #endregion

    #region HashCode
    internal sealed class HashCoder
    {
        private HashCoder() { }

        private static readonly int[] Primes = { 23, 47, 53, 59, 61, 67, 71, 73, 79, 83, 89, 97, 101, 103, 107, 109, 113, 127, 131 };

        internal static int GetHash(params int[] p)
        {
            unchecked
            {
                int Result = 17;
                for (int i = 0; i < p.Length; i++)
                {
                    {
                        Result = Result * Primes[i % Primes.Length] + p[i];
                    }
                }

                return Result;
            }
        }

        internal static int GetHashObj(params object[] p)
        {
            int[] p1 = new int[p.Length];
            for (int i = 0; i < p1.Length; i++)
            {
                if (p[i] != null) p1[i] = p[i].GetHashCode();
            }

            return GetHash(p1);
        }
    }
    #endregion

    #region Headers and footers
    /// <summary>
    /// Contains all information about headers and footers in an Excel sheet.
    /// </summary>
    public struct THeaderAndFooter
    {
        #region Variables
        string FDefaultHeader;
        string FDefaultFooter;
        string FEvenHeader;
        string FEvenFooter;
        string FFirstHeader;
        string FFirstFooter;

        bool FDiffEvenPages;
        bool FDiffFirstPage;

        bool FScaleWithDoc;
        bool FAlignMargins;
        #endregion

        private static string NonNull(string s)
        {
            if (s == null) return string.Empty;
            return s;
        }

        /// <summary>
        /// Sets the headers for all the pages to a given string. <b>Note that setting this value will  set <see cref="DiffEvenPages"/> and <see cref="DiffFirstPage"/> to false.</b>
        /// For a description of the possible values of this string, see <see cref="ExcelFile.PageHeader"/>
        /// </summary>
        /// <param name="value">String with the header.</param>
        public void SetAllHeaders(string value)
        {
            FDefaultHeader = value;
            FDiffEvenPages = false;
            FDiffFirstPage = false;
        }


        /// <summary>
        /// Sets the footers for all the pages to a given string. <b>Note that setting this value will set <see cref="DiffEvenPages"/> and <see cref="DiffFirstPage"/> to false.</b>
        /// For a description of the possible values of this string, see <see cref="ExcelFile.PageHeader"/>
        /// </summary>
        /// <param name="value">Sting with the footer.</param>
        public void SetAllFooters(string value)
        {
            FDefaultFooter = value;
            FDiffEvenPages = false;
            FDiffFirstPage = false;
        }


        /// <summary>
        /// Returns or sets the header for all pages that are not even or the first page. If <see cref="DiffFirstPage"/> is false, then this
        /// string also applies to the first page. If <see cref="DiffEvenPages"/> is false, this string also applies for even pages.
        /// For a description of the possible values of this string, see <see cref="ExcelFile.PageHeader"/>
        /// <br/>To set a header for all pages, use <see cref="SetAllHeaders(string)"/>
        /// </summary>
        public string DefaultHeader { get { return NonNull(FDefaultHeader); } set { FDefaultHeader = value; } }

        /// <summary>
        /// Returns or sets the footer for all pages that are not even or the first page. If <see cref="DiffFirstPage"/> is false, then this
        /// string also applies to the first page. If <see cref="DiffEvenPages"/> is false, this string also applies for even pages.
        /// For a description of the possible values of this string, see <see cref="ExcelFile.PageHeader"/>
        /// <br/>To set a footer for all pages, use <see cref="SetAllFooters(string)"/>
        /// </summary>
        public string DefaultFooter { get { return NonNull(FDefaultFooter); } set { FDefaultFooter = value; } }

        /// <summary>
        /// Header for even pages. <b>Note that this value is valid if and only if <see cref="DiffEvenPages"/> is true.</b>        
        /// For a description of the possible values of this string, see <see cref="ExcelFile.PageHeader"/>
        /// <br/>If you don't want a different header for even pages, set <see cref="DiffEvenPages"/> to false or
        /// call <see cref="SetAllHeaders"/>.
        /// </summary>
        public string EvenHeader { get { return NonNull(FEvenHeader); } set { FEvenHeader = value; } }

        /// <summary>
        /// Footer for even pages. <b>Note that this value is valid if and only if <see cref="DiffEvenPages"/> is true.</b>
        /// For a description of the possible values of this string, see <see cref="ExcelFile.PageHeader"/>
        /// <br/>If you don't want a different footer for even pages, set <see cref="DiffEvenPages"/> to false or
        /// call <see cref="SetAllFooters"/>.
        /// </summary>
        public string EvenFooter { get { return NonNull(FEvenFooter); } set { FEvenFooter = value; } }

        /// <summary>
        /// Header for the first page. <b>Note that this value is valid if and only if <see cref="DiffFirstPage"/> is true.</b>        
        /// For a description of the possible values of this string, see <see cref="ExcelFile.PageHeader"/>
        /// <br/>If you don't want a different header for the first page, set <see cref="DiffFirstPage"/> to false or
        /// call <see cref="SetAllHeaders"/>.
        /// </summary>
        public string FirstHeader { get { return NonNull(FFirstHeader); } set { FFirstHeader = value; } }

        /// <summary>
        /// Footer for the first page.<b>Note that this value is valid if and only if <see cref="DiffFirstPage"/> is true.</b>        
        /// For a description of the possible values of this string, see <see cref="ExcelFile.PageHeader"/>
        /// <br/>If you don't want a different footer for the first page, set <see cref="DiffFirstPage"/> to false or
        /// call <see cref="SetAllFooters"/>.
        /// </summary>
        public string FirstFooter { get { return NonNull(FFirstFooter); } set { FFirstFooter = value; } }

        /// <summary>
        /// When true the first page will have a different header and footer from the rest, and it will be specified in <see cref="FirstHeader"/>
        /// and <see cref="FirstFooter"/>. When false, FirstHeader and FirstFooter have no meaning.
        /// </summary>
        public bool DiffFirstPage { get { return FDiffFirstPage; } set { FDiffFirstPage = value; } }

        /// <summary>
        /// When true even pages will have different headers and footers from odd pages, and headers/footer for even pages will be specified in <see cref="EvenHeader"/>
        /// and <see cref="EvenFooter"/>. When false, EvenHeader and EvenFooter have no meaning.
        /// </summary>
        public bool DiffEvenPages { get { return FDiffEvenPages; } set { FDiffEvenPages = value; } }

        /// <summary>
        /// Determines if to scale header and footer with document scaling or not.
        /// </summary>
        public bool ScaleWithDoc { get { return FScaleWithDoc; } set { FScaleWithDoc = value; } }

        /// <summary>
        /// Align header footer margins with page margins. When true, as left/right margins grow 
        /// and shrink, the header and footer edges stay aligned with the margins. When false, 
        /// headers and footers are aligned on the paper edges, regardless of margins. 
        /// </summary>
        public bool AlignMargins { get { return FAlignMargins; } set { FAlignMargins = value; } }


        /// <summary>
        /// Returns the header for a given page, considering if there are differences in even/odd pages or the first page.
        /// </summary>
        /// <param name="currentPage">Page for which you want the headers (1 based)</param>
        /// <returns></returns>
        public string GetHeader(int currentPage)
        {
            switch (GetHeaderAndFooterKind(currentPage))
            {
                case THeaderAndFooterKind.FirstPage:
                    return FirstHeader;

                case THeaderAndFooterKind.EvenPages:
                    return EvenHeader;

                default:
                    return DefaultHeader;
            }
        }

        /// <summary>
        /// Returns the footer for a given page, considering if there are differences in even/odd pages or the first page.
        /// </summary>
        /// <param name="currentPage">Page for which you want the footers (1 based)</param>
        /// <returns></returns>
        public string GetFooter(int currentPage)
        {
            switch (GetHeaderAndFooterKind(currentPage))
            {
                case THeaderAndFooterKind.FirstPage:
                    return FirstFooter;

                case THeaderAndFooterKind.EvenPages:
                    return EvenFooter;

                default:
                    return DefaultFooter;
            }
        }

        /// <summary>
        /// Returns the kind of footer image for a given page. This method is normally useful to get the correct image
        /// for an specific page.
        /// </summary>
        /// <param name="currentPage">Page for which you want the headers and footers (1 based)</param>
        /// <returns></returns>
        public THeaderAndFooterKind GetHeaderAndFooterKind(int currentPage)
        {
            if (currentPage == 1)
            {
                if (DiffFirstPage) return THeaderAndFooterKind.FirstPage;
                return THeaderAndFooterKind.Default;
            }

            if (currentPage % 2 == 0)
            {
                if (DiffEvenPages) return THeaderAndFooterKind.EvenPages;
                return THeaderAndFooterKind.Default;
            }

            return THeaderAndFooterKind.Default;
        }


        /// <summary>
        /// Returns true if both objects are the same.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            if (!(obj is THeaderAndFooter)) return false;
            THeaderAndFooter o2 = (THeaderAndFooter)obj;
            return o2 == this;
        }

        /// <summary>
        /// Returns true if both structures are the same.
        /// </summary>
        /// <param name="o1"></param>
        /// <param name="o2"></param>
        /// <returns></returns>
        public static bool operator ==(THeaderAndFooter o1, THeaderAndFooter o2)
        {
            return o2.AlignMargins == o1.AlignMargins
                && o2.DiffEvenPages == o1.DiffEvenPages
                && o2.DiffFirstPage == o1.DiffFirstPage
                && o2.ScaleWithDoc == o1.ScaleWithDoc
                && o2.DefaultHeader == o1.DefaultHeader
                && o2.DefaultFooter == o1.DefaultFooter
                && o2.FirstHeader == o1.FirstHeader
                && o2.FirstFooter == o1.FirstFooter
                && o2.EvenHeader == o1.EvenHeader
                && o2.EvenFooter == o1.EvenFooter;

        }

        /// <summary>
        /// Returns true if both structures are different.
        /// </summary>
        /// <param name="o1"></param>
        /// <param name="o2"></param>
        /// <returns></returns>
        public static bool operator !=(THeaderAndFooter o1, THeaderAndFooter o2)
        {
            return !(o1 == o2);
        }

        /// <summary>
        /// Hashcode for the obeject.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHash(
                AlignMargins.GetHashCode(),
                DiffEvenPages.GetHashCode(),
                DiffFirstPage.GetHashCode(),
                ScaleWithDoc.GetHashCode(),
                DefaultHeader.GetHashCode(),
                DefaultFooter.GetHashCode(),
                FirstHeader.GetHashCode(),
                FirstFooter.GetHashCode(),
                EvenHeader.GetHashCode(),
                EvenFooter.GetHashCode());
        }
    }
    #endregion

    #region Copied objects
    struct TIdAndShapeId
    {
        public int Id;
        public long ShapeId;

        public TIdAndShapeId(int aId, long aShapeId)
        {
            Id = aId;
            ShapeId = aShapeId;
        }
    }

    /// <summary>
    /// A list that contains the Ids and positions of the Excel objects
    /// </summary>
    public class TExcelObjectList
    {
        private List<TIdAndShapeId> FList;
        private List<long[]> FCopies;
        private List<long> FLastCopy;
        private int FLastCopiedRow;

        /// <summary>
        /// Creates a new instance. IncludeCopies is false.
        /// </summary>
        public TExcelObjectList()
        {
            FList = new List<TIdAndShapeId>();
        }

        /// <summary>
        /// Creates a new TObjectList instance.
        /// </summary>
        /// <param name="includeCopies">If true, all shape ids of copied shapes will be included in the Copies property.</param>
        public TExcelObjectList(bool includeCopies): this()
        {
            if (includeCopies)
            {
                FCopies = new List<long[]>();
                FLastCopy = new List<long>();
            }
        }

        /// <summary>
        /// If this property is true, all shape ids from the copies made will be stored in the Copies property.
        /// </summary>
        public bool IncludeCopies { get { return FCopies != null; } }

        /// <summary>
        /// Number of objects in the list.
        /// </summary>
        public int Count
        {
            get
            {
                return FList.Count;
            }
        }

        /// <summary>
        /// Returns position i in the list.
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public int this[int index] { get { return FList[index].Id; } }

        /// <summary>
        /// Returns the shape id of the object in the list.
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public long ShapeId(int index)
        {
            return FList[index].ShapeId;
        }

        internal void Add(int i, long ShapeId)
        {
            FList.Add(new TIdAndShapeId(i, ShapeId));
        }

        internal void Reverse()
        {
            FList.Reverse();
        }

        internal void AddCopy(int CopyPos, long ShapeId)
        {
            if (!IncludeCopies) return;
            if (CopyPos == FLastCopiedRow)
            {
                FLastCopy.Add(ShapeId);
            }
            else if (CopyPos == FLastCopiedRow + 1)
            {
                FCopies.Add(FLastCopy.ToArray());
                FLastCopy.Clear();
                FLastCopiedRow = CopyPos;
                FLastCopy.Add(ShapeId);
            }
            else FlxMessages.ThrowException(FlxErr.ErrInternal);
        }

        internal long[] GetObjects(int RecordPos)
        {
            if (RecordPos == 0) //Original object
            {
                long[] Result = new long[FList.Count];
                for (int i = 0; i < FList.Count; i++)
                {
                    Result[i] = FList[i].ShapeId;
                }

                return Result;
            }

            if (RecordPos - 1 == FCopies.Count) return FLastCopy.ToArray();
            return FCopies[RecordPos - 1];
        }
    }
    #endregion

}
