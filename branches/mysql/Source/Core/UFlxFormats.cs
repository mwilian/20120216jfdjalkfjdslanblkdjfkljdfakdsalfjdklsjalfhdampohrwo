using System;

namespace FlexCel.Core
{
	#region Enums

    /// <summary>
    /// Horizontal Alignment on a cell.
    /// </summary>
    public enum THFlxAlignment
    {
        /// <summary>
        /// General Alignment. (Text to the left, numbers to the right and Errors and booleans centered)
        /// </summary>
        general, 
        /// <summary>
        /// Aligned to the left.
        /// </summary>
        left, 
        /// <summary>
        /// Horizontally centered on the cell.
        /// </summary>
        center, 
        /// <summary>
        /// Aligned to the right.
        /// </summary>
        right, 
        /// <summary>
        /// Repeat the text to fill the cell width.
        /// </summary>
        fill,
        /// <summary>
        /// Justify with spaces the text so it fills the cell width.
        /// </summary>
        justify, 
        /// <summary>
        /// Centered on a group of cells.
        /// </summary>
        center_across_selection,
        /// <summary>
        /// Justified.
        /// </summary>
        distributed
    }

    /// <summary>
    /// Vertical Alignment on a cell.
    /// </summary>
    public enum TVFlxAlignment
    {
        /// <summary>
        /// Aligned to the top.
        /// </summary>
        top, 
        /// <summary>
        /// Vertically centered on the cell.
        /// </summary>
        center, 
        /// <summary>
        /// Aligned to the bottom.
        /// </summary>
        bottom, 
        /// <summary>
        /// Justified on the cell.
        /// </summary>
        justify,
        /// <summary>
        /// Distributed on the cell.
        /// </summary>
        distributed
    }

    /// <summary>
    /// Cell border style.
    /// </summary>
    public enum TFlxBorderStyle 
    {
        ///<summary>None</summary>
        None, 
        ///<summary>Thin</summary>
        Thin, 
        ///<summary>Medium</summary>
        Medium, 
        ///<summary>Dashed</summary>
        Dashed, 
        ///<summary>Dotted</summary>
        Dotted, 
        ///<summary>Thick</summary>
        Thick,
        ///<summary>Double</summary>
        Double, 
        ///<summary>Hair</summary>
        Hair, 
        ///<summary>Medium_dashed</summary>
        Medium_dashed, 
        ///<summary>Dash_dot</summary>
        Dash_dot, 
        ///<summary>Medium_dash_dot</summary>
        Medium_dash_dot,
        ///<summary>Dash_dot_dot</summary>
        Dash_dot_dot, 
        ///<summary>Medium_dash_dot_dot</summary>
        Medium_dash_dot_dot, 
        ///<summary>Slanted_dash_dot</summary>
        Slanted_dash_dot
    };

    /// <summary>
    /// Pattern style.
    /// </summary>
    public enum TFlxPatternStyle
    {
        ///<summary>Automatic </summary>
        Automatic = 0,
        ///<summary>None </summary>
        None = 1,
        ///<summary>Solid </summary>
        Solid = 2,
        ///<summary>Gray50 </summary>
        Gray50 = 3,
        ///<summary>Gray75 </summary>
        Gray75 = 4,
        ///<summary>Gray25 </summary>
        Gray25 = 5,
        ///<summary>Horizontal </summary>
        Horizontal = 6,
        ///<summary>Vertical </summary>
        Vertical = 7,
        ///<summary>Down </summary>
        Down = 8,
        ///<summary>Up </summary>
        Up = 9,
        ///<summary>Diagonal hatch.</summary>
        Checker = 10,
        ///<summary>bold diagonal.</summary>
        SemiGray75 = 11,
        ///<summary>thin horz lines </summary>
        LightHorizontal = 12,
        ///<summary>thin vert lines</summary>
        LightVertical = 13,
        ///<summary>thin \ lines</summary>
        LightDown = 14,
        ///<summary>thin / lines</summary>
        LightUp = 15,
        ///<summary>thin horz hatch</summary>
        Grid = 16,  
        ///<summary>thin diag</summary>
        CrissCross = 17,
        ///<summary>12.5 % gray</summary>
        Gray16 = 18,
        ///<summary>6.25 % gray</summary>
        Gray8 = 19,

        /// <summary>
        /// The fill style will be a <see cref="TExcelGradient"/>.
        /// </summary>
        Gradient = 40
    }

    /// <summary>
    /// Diagonal border style.
    /// </summary>
    public enum TFlxDiagonalBorder 
    {
        /// <summary>
        /// No diagonal line.
        /// </summary>
        None, 

        /// <summary>
        /// A line going from left-top to right-bottom.
        /// </summary>
        DiagDown,
 
        /// <summary>
        /// A line going from left-bottom to right-top.
        /// </summary>
        DiagUp, 

        /// <summary>
        /// A diagonal cross.
        /// </summary>
        Both
    }

    /// <summary>
    /// Font style. You can "or" on "and" it to get the actual styles.
    /// For example, to set style to bold+italic,  you should use TFlxFontStyles.Bold | TFlxFontStyles.Italic.
    ///              to check if style includes italic, use ((Style &amp; TFlxFontStyles.Italic)!=0)  
    /// </summary>
    [Flags]
    public enum TFlxFontStyles 
    {
        /// <summary>Normal font.</summary>
        None=0, 

        /// <summary>Bold font.</summary>
        Bold=1,
 
        /// <summary>Italic font.</summary>
        Italic=2, 
        
        /// <summary>Striked out font.</summary>
        StrikeOut=4, 
        
        /// <summary>Superscript font.</summary>
        Superscript=8, 
        
        /// <summary>Subscript font.</summary>
        Subscript=16,

        /// <summary>Outlined font. Excel currently ignores this setting.</summary>
        Outline=32,

        /// <summary>Font has a shadow. Excel currently ignores this setting.</summary>
        Shadow = 32,

        /// <summary>
        /// Condensed font, for backwards compatibility. Excel ignores this setting.
        /// </summary>
        Condense = 64,

        /// <summary>
        /// Extended font, for backwards compatibility. Excel ignores this setting.
        /// </summary>
        Extend = 128
    };

    /// <summary>
    /// Underline type.
    /// </summary>
    public enum TFlxUnderline 
    {
        /// <summary>
        /// No underline.
        /// </summary>
        None, 
        /// <summary>
        /// Simple underline.
        /// </summary>
        Single, 
        /// <summary>
        /// Double underline.
        /// </summary>
        Double, 
        /// <summary>
        /// Underlines at the bottom of the cell.
        /// </summary>
        SingleAccounting, 
        /// <summary>
        /// Double underline at the bottom of the cell.
        /// </summary>
        DoubleAccounting
    };
	#endregion

	#region TFlxFont
    /// <summary>
    /// Specifies the scheme to which a font belongs. This attribute is only valid in Excel 2007.
    /// </summary>
    public enum TFontScheme
    {
        /// <summary>
        /// Font is not a theme font.
        /// </summary>
        None,

        /// <summary>
        /// The font is a minor font for the scheme.
        /// </summary>
        Minor, 

        /// <summary>
        /// The font is a major font for the scheme.
        /// </summary>
        Major
    }

    /// <summary>
    /// Encapsulation of an Excel Font.
    /// </summary>
    public class TFlxFont: ICloneable
    {
        #region Privates
        private string FName;
        private int FSize20;
        private TExcelColor FColor;
        private TFlxFontStyles FStyle;
        private TFlxUnderline  FUnderline;
        private byte FFamily;
        private byte FCharSet;
        private TFontScheme FScheme;
        #endregion

        #region Constructor
        /// <summary>
        /// Creates a font with size 10 and name "Arial".
        /// </summary>
        public TFlxFont()
        {
            FName = "Arial";
            FSize20 = 200;
        }

        internal TFlxFont(
        string aName,
        int aSize20,
        TExcelColor aColor,
        TFlxFontStyles aStyle,
        TFlxUnderline aUnderline,
        byte aFamily,
        byte aCharSet,
        TFontScheme aScheme)
        {
            FName = aName;
            FSize20 = aSize20;
            FColor = aColor;
            FStyle = aStyle;
            FUnderline = aUnderline;
            FFamily = aFamily;
            FCharSet = aCharSet;
            FScheme = aScheme;
        }
        #endregion

        #region Public properties
        /// <summary>
        /// Font name. (For example, "Arial")
        /// </summary>
        public string Name {get {return FName;} set {FName=value;}}

        /// <summary>
        /// Height of the font (in units of 1/20th of a point). A Size20=200 means 10 points.
        /// </summary>
        public int Size20 {get {return FSize20;} set {FSize20=value;}}

        /// <summary>
        /// Color of the font. 
        /// </summary>
        public TExcelColor Color { get { return FColor; } set { FColor = value; } }

        /// <summary>
        /// Style of the font, such as bold or italics. Underline is a different option.
        /// </summary>
        public TFlxFontStyles Style {get {return FStyle;} set {FStyle=value;}}

        /// <summary>
        /// Underline type.
        /// </summary>
        public TFlxUnderline  Underline {get {return FUnderline;} set {FUnderline=value;}}

        /// <summary>
        /// Font family, (see Windows API LOGFONT structure).
        /// </summary>
        public byte Family {get {return FFamily;} set {FFamily=value;}}

        /// <summary>
        /// Character set. (see Windows API LOGFONT structure)
        /// </summary>
        public byte CharSet {get {return FCharSet;} set {FCharSet=value;}}

        /// <summary>
        /// Font scheme. This only applies to Excel 2007.
        /// </summary>
        public TFontScheme Scheme { get { return FScheme; } set { FScheme = value; } }

        #endregion

        #region Public Methods
        /// <summary>
        /// Copies this font information to other font object.
        /// </summary>
        /// <param name="Dest">Existing Font object where new data will be copied.</param>
        public void CopyTo(TFlxFont Dest)
        {
            Dest.FName = FName;
            Dest.FSize20 = FSize20;
            Dest.Color = Color;
            Dest.FStyle = FStyle;
            Dest.FUnderline = FUnderline;
            Dest.FFamily = FFamily;
            Dest.FCharSet = FCharSet;
            Dest.FScheme = FScheme;
        }
        #endregion
        
        #region ICloneable Members

        /// <summary>
        /// Returns a deep copy of the font.
        /// </summary>
        /// <returns>A copy of the font.</returns>
        public object Clone()
        {
            return new TFlxFont(
            FName,
            FSize20,
            FColor,
            FStyle,
            FUnderline,
            FFamily,
            FCharSet,
            FScheme);
        }

        #endregion

        #region Equals
        /// <summary>
        /// Returns true if a font has is the same as the current.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            TFlxFont fnt = obj as TFlxFont;
            if (fnt == null) return false;

            return
                FName == fnt.FName &&
                FSize20 == fnt.FSize20 &&
                FColor == fnt.FColor &&
                FStyle == fnt.FStyle &&
                FUnderline == fnt.FUnderline &&
                FFamily == fnt.FFamily &&
                FCharSet == fnt.FCharSet &&
                FScheme == fnt.FScheme;

        }

        /// <summary>
        /// Hash code of the font.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHashObj(
            FName,
            FSize20,
            FColor,
            (int)FStyle,
            (int)FUnderline,
            FFamily,
            FCharSet,
            (int)FScheme);
        }


        #endregion
    }
    
	#endregion

	#region Borders and Patterns
    /// <summary>
    /// Border style and color for one of the 4 sides of a cell.
    /// </summary>
    public struct TFlxOneBorder
    {
        #region Privates
        private TFlxBorderStyle FStyle;
        private TExcelColor FColor;
        #endregion

		/// <summary>
		/// Initializes the structure to its default values.
		/// </summary>
		/// <param name="aBorderStyle">See <see cref="Style"/></param>
		/// <param name="aColor">See <see cref="Color"/></param>
		public TFlxOneBorder(TFlxBorderStyle aBorderStyle, TExcelColor aColor)
		{
			FStyle = aBorderStyle;
			FColor = aColor;
		}

        /// <summary>
        /// Border style.
        /// </summary>
        public TFlxBorderStyle Style {get {return FStyle;} set {FStyle=value;}}

        /// <summary>
        /// Color of the border.
        /// </summary>
        public TExcelColor Color { get { return FColor; } set { FColor = value; } }

        /// <summary>
        /// Returns true when 2 borders are the same.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            if (obj == null || GetType() != obj.GetType()) return false;
            TFlxOneBorder o2= (TFlxOneBorder)obj;
            if (Style == TFlxBorderStyle.None) return o2.Style == Style; //When border is none, color doesn't matter.
            return (o2.Style==Style) && (o2.Color==Color);

        }

        /// <summary></summary>
        public static bool operator== (TFlxOneBorder b1, TFlxOneBorder b2)
        {
            return b1.Equals(b2);
        }

        /// <summary></summary>
        public static bool operator!= (TFlxOneBorder b1, TFlxOneBorder b2)
        {
            return !(b1 == b2);
        }

        /// <summary>
        /// Hash code for the border.
        /// </summary>
        /// <returns>hashcode.</returns>
        public override int GetHashCode()
        {
            if (Style == TFlxBorderStyle.None) return HashCoder.GetHash(((int)Style).GetHashCode());
            return HashCoder.GetHash(((int)Style).GetHashCode(), Color.GetHashCode());
        }
    }

    /// <summary>
    /// Border style for a cell.
    /// </summary>
    public class TFlxBorders: ICloneable
    {
        /// <summary>
        /// Left border.
        /// </summary>
        public TFlxOneBorder Left;
        
        /// <summary>
        /// Right border.
        /// </summary>
        public TFlxOneBorder Right;
        
        /// <summary>
        /// Top border.
        /// </summary>
        public TFlxOneBorder Top;
        
        /// <summary>
        /// Bottom border.
        /// </summary>
        public TFlxOneBorder Bottom;
        
        /// <summary>
        /// Diagonal border.
        /// </summary>
        public TFlxOneBorder Diagonal;

        /// <summary>
        /// When defined, there will be one or two diagonal lines across the cell.
        /// </summary>
        public TFlxDiagonalBorder DiagonalStyle;

        /// <summary>
        /// Sets all borders to a linestyle and color. Diagonal borders are not changed.
        /// </summary>
        /// <param name="borderStyle">Border style to apply.</param>
        /// <param name="color">Color to apply</param>
        public void SetAllBorders(TFlxBorderStyle borderStyle, TExcelColor color)
        {
            Left.Color = color;
            Right.Color = color;
            Top.Color = color;
            Bottom.Color = color;

            Left.Style = borderStyle;
            Right.Style = borderStyle;
            Top.Style = borderStyle;
            Bottom.Style = borderStyle;
        }

        #region ICloneable Members

        /// <summary>
        /// Returns a deep copy of the border.
        /// </summary>
        /// <returns></returns>
        public object Clone()
        {
            return MemberwiseClone();
        }

        #endregion

        /// <summary>
        /// Returns true if both borders are the same.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            TFlxBorders Borders = obj as TFlxBorders;
            if (Borders == null) return false;

            return
            Left == Borders.Left
            && Right == Borders.Right
            && Top == Borders.Top
            && Bottom == Borders.Bottom
            && Diagonal == Borders.Diagonal
            && DiagonalStyle == Borders.DiagonalStyle;
        }

        /// <summary>
        /// Returns the hashcode for the border.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHash(
                Left.GetHashCode(),
                Right.GetHashCode(),
                Top.GetHashCode(),
                Bottom.GetHashCode(),
                Diagonal.GetHashCode(),
                DiagonalStyle.GetHashCode());
        }
    }

    /// <summary>
    /// Fill pattern and color for the background of a cell.
    /// </summary>
    public struct TFlxFillPattern
    {
       private TExcelGradient FGradient;
        private const TExcelGradient EmptyGradient = null;

        #region Members
        /// <summary>
        /// Fill style.
        /// </summary>
        public TFlxPatternStyle Pattern;

        /// <summary>
        /// Color for the foreground of the pattern. It is used when the pattern is solid, but not when it is automatic.
        /// </summary>
        public TExcelColor FgColor;
        
        /// <summary>
        /// Color for the background of the pattern.  If the pattern is solid it has no effect, but it is used when pattern is automatic.
        /// </summary>
        public TExcelColor BgColor;

        /// <summary>
        /// Gradient definition. This is only valid if <see cref="Pattern"/> is TFlxPatternStyle.Gradient.
        /// </summary>
        public TExcelGradient Gradient 
        { 
            get { if (Pattern == TFlxPatternStyle.Gradient) return FGradient; else return EmptyGradient; }
            set 
            {
                if (value == null) FGradient = null;
                else
                {
                    FGradient = value.Clone();
                    Pattern = TFlxPatternStyle.Gradient;
                }
            }
        }
        #endregion

        /// <summary></summary>
        public override bool Equals(object obj)
        {
            if (obj == null || GetType() != obj.GetType()) return false;
            TFlxFillPattern o2=(TFlxFillPattern)obj;
            if (Pattern == TFlxPatternStyle.None) return o2.Pattern == TFlxPatternStyle.None; //rest doesn't matter here.
            bool SkipBgColor = Pattern == TFlxPatternStyle.Solid && o2.Pattern == TFlxPatternStyle.Solid;
            return (o2.Pattern==Pattern) && (o2.FgColor==FgColor) && (SkipBgColor || o2.BgColor==BgColor) && (o2.Gradient == Gradient);
        }

        /// <summary></summary>
        public static bool operator== (TFlxFillPattern f1, TFlxFillPattern f2)
        {
            return f1.Equals(f2);
        }

        /// <summary></summary>
        public static bool operator!= (TFlxFillPattern f1, TFlxFillPattern f2)
        {
            return !(f1 == f2);
        }

        /// <summary></summary>
        public override int GetHashCode()
        {
            if (Pattern == TFlxPatternStyle.None) return Pattern.GetHashCode();
            int GradientHashCode = 0;
            if (Gradient != null) GradientHashCode = Gradient.GetHashCode();
            return HashCoder.GetHash(Pattern.GetHashCode(), FgColor.GetHashCode(), BgColor.GetHashCode(), GradientHashCode);
        }

        /// <summary>
        /// Creates a deep copy of this object.
        /// </summary>
        /// <returns></returns>
        public TFlxFillPattern Clone()
        {
            TFlxFillPattern Result = (TFlxFillPattern)MemberwiseClone();
            Result.Gradient = Gradient; //setting it will clone it
            Result.Pattern = Pattern;
            return Result;
        }
    }

	#endregion

	#region TFlxFormat
    /// <summary>
    /// Format for one cell or named style.
    /// Cell formats are used to format cells, Named styles to create styles. A Cell format can have a parent style format, even when normally this is null (parent is normal format).
    /// Named styles will have a non-null Style property. Cell formats will have style = null.
    /// </summary>
    public class TFlxFormat: ICloneable
    {
		#region Private variables
		private TFlxFont        FFont;
		private TFlxBorders     FBorders;
		private string          FFormat;
		private THFlxAlignment  FHAlignment;
		private TVFlxAlignment  FVAlignment; 
		private bool            FLocked;
		private bool			FHidden;
		private bool            FWrapText;
		private bool            FShrinkToFit;
		private byte            FRotation;
		private byte            FIndent;
		private string          FParentStyle;
		private TLinkedStyle    FLinkedStyle = new TLinkedStyle();
        private bool            FLotus123Prefix;
		#endregion

        /// <summary>
        /// Creates an empty Format class. Don't use this to get TFlxFormat instances, use XlsFile.GetDefaultFormat instead.
        /// </summary>
        internal TFlxFormat()
        {
        }

        /// <summary>
        /// Returns a standard TFlxFormat for Excel 2007. (Font name is Calibri, etc). You will normally want to use <see cref="ExcelFile.GetDefaultFormat"/> instead of this,
        /// since it returns the default format for an specific file, and not a generic format like this. 
        /// </summary>
        public static TFlxFormat CreateStandard2007()
        {
            TFlxFormat Result = new TFlxFormat();
            Result.FFont = new TFlxFont();
            Result.Font.Name = "Calibri";
            Result.Font.Size20 = 220;
            Result.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground1);
            Result.Font.Family = 2;
            Result.Font.Scheme = TFontScheme.Minor;

            Result.FillPattern.Pattern = TFlxPatternStyle.None;
            
            Result.FBorders = new TFlxBorders();
            Result.FFormat = String.Empty;
            Result.FHAlignment = THFlxAlignment.general;
            Result.FVAlignment = TVFlxAlignment.bottom;
            Result.FLocked = true;
            Result.FHidden = false;
            Result.FWrapText = false;
            Result.FShrinkToFit = false;
            Result.FRotation = 0;
            Result.FIndent = 0;
            Result.FParentStyle = null;
            Result.FLinkedStyle = new TLinkedStyle();
            Result.FLotus123Prefix = false;

            Result.IsStyle = false;

            return Result;
        }

        /// <summary>
        /// Cell Font.
        /// </summary>
		public TFlxFont Font {get {return FFont;} set{FFont = value;}}
     
        /// <summary>
        /// Cell borders.
        /// </summary>
		public TFlxBorders     Borders {get {return FBorders;} set{FBorders = value;}}

        /// <summary>
        /// Format string.  (For example, "yyyy-mm-dd" for a date format, or "#.00" for a numeric 2 decimal format)
        /// <br/>This format string is the same you use in Excel under "Custom" format when formatting a cell, and it is documented
        /// in Excel documentation. Under <b>"Finding out what format string to use in TFlxFormat.Format"</b> section in <b>UsingFlexCelAPI.pdf</b>
        /// you can find more detailed information on how to create this string.
        /// </summary>
		public string          Format {get {return FFormat;} set{FFormat = value;}}

        /// <summary>
        /// Fill pattern.
        /// </summary>
		public TFlxFillPattern FillPattern;

        /// <summary>
        /// Horizontal alignment on the cell.
        /// </summary>
		public THFlxAlignment  HAlignment {get {return FHAlignment;} set{FHAlignment = value;}}

        /// <summary>
        /// Vertical alignment on the cell.
        /// </summary>
		public TVFlxAlignment  VAlignment {get {return FVAlignment;} set{FVAlignment = value;}}

        /// <summary>
        /// Cell is locked.
        /// </summary>
		public bool           Locked {get {return FLocked;} set{FLocked = value;}}

        /// <summary>
        /// Cell is Hidden.
        /// </summary>
		public bool			  Hidden {get {return FHidden;} set{FHidden = value;}}

        /// <summary>
        /// Cell wrap.
        /// </summary>
		public bool           WrapText {get {return FWrapText;} set{FWrapText = value;}}

        /// <summary>
        /// Shrink to fit.
        /// </summary>
		public bool           ShrinkToFit {get {return FShrinkToFit;} set{FShrinkToFit = value;}}

        /// <summary>
        /// Text Rotation in degrees. <br/>
        /// 0 - 90 is up, <br/>
        /// 91 - 180 is down, <br/>
        /// 255 is vertical.
        /// </summary>
		public byte            Rotation {get {return FRotation;} set{FRotation = value;}}

        /// <summary>
        /// Indent value. (in characters). This value can't be bigger than 15 in Excel 2003 or earlier, and no bigger than 250 in Excel 2007 or newer.
        /// </summary>
		public byte            Indent {get {return FIndent;} set{FIndent = value;}}


		/// <summary>
		/// When true this format is a named style, when false a cell format.
		/// </summary>
		public bool IsStyle;

        /// <summary>
        /// If true the prefix for the cell is compatible with Lotus 123.
        /// </summary>
        public bool Lotus123Prefix { get { return FLotus123Prefix; } set { FLotus123Prefix = value; } }

		/// <summary>
		/// Name of the Parent style. Normally you will want to keep it at null (parent is normal style), but you can write an existing style here.
		/// If <see cref="IsStyle"/> is true this property is not used.
		/// </summary>
		public string		   ParentStyle {get {return FParentStyle;} set{FParentStyle = value;}}

        /// <summary>
        /// This is similar to <see cref="ParentStyle"/> but will return "Normal" when the parent is null.
        /// </summary>
        public string NotNullParentStyle { get { if (FParentStyle == null) return FlxConsts.NormalStyleName; else return FParentStyle; } }

		/// <summary>
		/// If this object holds a Cell format, LinkedStyle specifies which properties of the cell format are linked to its parent style.
		/// If this object holds a Style formt, LinkedStyle specifies the default set of properties that will be applied when you use this style from Excel.
		/// </summary>
		public TLinkedStyle   LinkedStyle {get {return FLinkedStyle;}}

        #region ICloneable Members

        /// <summary>
        /// Returns a deep copy of the format.
        /// </summary>
        /// <returns></returns>
        public object Clone()
        {
            TFlxFormat Result = (TFlxFormat) MemberwiseClone();
            Result.Font = (TFlxFont) Font.Clone();
            Result.Borders = (TFlxBorders) Borders.Clone();
            Result.FillPattern = FillPattern.Clone();
			Result.FLinkedStyle = (TLinkedStyle) FLinkedStyle.Clone();
            return Result;
        }

        #endregion

        #region Internal Methods
        internal void FixColors(IFlexCelPalette Source, IFlexCelPalette Dest)
        {
            Font.Color = TExcelColor.Copy(Font.Color, Source, Dest);
            FillPattern.FgColor = TExcelColor.Copy(FillPattern.FgColor, Source, Dest);
            FillPattern.BgColor = TExcelColor.Copy(FillPattern.BgColor, Source, Dest);
            Borders.Left.Color = TExcelColor.Copy(Borders.Left.Color, Source, Dest);
            Borders.Top.Color = TExcelColor.Copy(Borders.Top.Color, Source, Dest);
            Borders.Right.Color = TExcelColor.Copy(Borders.Right.Color, Source, Dest);
            Borders.Bottom.Color = TExcelColor.Copy(Borders.Bottom.Color, Source, Dest);
            Borders.Diagonal.Color = TExcelColor.Copy(Borders.Diagonal.Color, Source, Dest);
        }
        #endregion
    }
	#endregion

    #region ApplyFormat
    /// <summary>
    /// Encapsulation of an Excel Font.
    /// </summary>
    public class TFlxApplyFont: ICloneable
    {
        #region Privates
        private bool FName;
        private bool FSize20;
        private bool FColor;
        private bool FStyle;
        private bool FUnderline;
        private bool FFamily;
        private bool FCharSet;
        #endregion

        #region Public properties
        /// <summary>
        /// Font name. (For example, "Arial")
        /// </summary>
        public bool Name {get {return FName;} set {FName=value;}}

        /// <summary>
        /// Height of the font (in units of 1/20th of a point). A Size20=200 means 10 points.
        /// </summary>
        public bool Size20 {get {return FSize20;} set {FSize20=value;}}

        /// <summary>
        /// Color of the font. 
        /// </summary>
        public bool Color {get {return FColor;} set {FColor=value;}}

        /// <summary>
        /// Style of the font, such as bold or italics. Underline is a different option.
        /// </summary>
        public bool Style {get {return FStyle;} set {FStyle=value;}}

        /// <summary>
        /// Underline type.
        /// </summary>
        public bool Underline {get {return FUnderline;} set {FUnderline=value;}}

        /// <summary>
        /// Font family, (see Windows API LOGFONT structure).
        /// </summary>
        public bool Family {get {return FFamily;} set {FFamily=value;}}

        /// <summary>
        /// Character set. (see Windows API LOGFONT structure)
        /// </summary>
        public bool CharSet {get {return FCharSet;} set {FCharSet=value;}}
        #endregion

        #region Public Methods
        /// <summary>
        /// Sets all members to true or false
        /// </summary>
        public void SetAllMembers(bool Value)
        {
            FName = Value;
            FSize20 = Value;
            FColor = Value;
            FStyle = Value;
            FUnderline = Value;
            FFamily = Value;
            FCharSet = Value;
        }

		/// <summary>
		/// Returns true if the format definition does not apply any setting.
		/// </summary>
		public bool IsEmpty
		{
			get
			{
				return
					FName == false &&
					FSize20 == false &&
					FColor == false &&
					FStyle == false &&
					FUnderline == false &&
					FFamily == false &&
					FCharSet == false;
			}
		}


        /// <summary>
        /// This method will modify existingFormat with the properties from newFormat that are specified on this class
        /// </summary>
        /// <param name="existingFormat">Existing format that will be updated with the properties of newFormat specified.</param>
        /// <param name="newFormat">New format to apply</param>
        /// <returns>True if there was any change on existingFormat, false otherwise.</returns>
        public bool Apply(TFlxFont existingFormat, TFlxFont newFormat)
        {
            bool Result = false;
            if (FName && existingFormat.Name != newFormat.Name) {Result = true; existingFormat.Name = newFormat.Name;}
            if (FSize20 && existingFormat.Size20 != newFormat.Size20) {Result = true; existingFormat.Size20 = newFormat.Size20;}

            if (FColor && existingFormat.Color != newFormat.Color) {Result = true; existingFormat.Color = newFormat.Color;}
            if (FStyle && existingFormat.Style != newFormat.Style) {Result = true; existingFormat.Style = newFormat.Style;}
            if (FUnderline && existingFormat.Underline != newFormat.Underline) {Result = true; existingFormat.Underline = newFormat.Underline;}
            if (FFamily && existingFormat.Family != newFormat.Family) {Result = true; existingFormat.Family = newFormat.Family;}
            if (FCharSet && existingFormat.CharSet != newFormat.CharSet) {Result = true; existingFormat.CharSet = newFormat.CharSet;}

            return Result;
        }


        #endregion
        
        #region ICloneable Members

        /// <summary>
        /// Returns a deep copy of the font.
        /// </summary>
        /// <returns>A copy of the font.</returns>
        public object Clone()
        {
            return MemberwiseClone();
        }

        #endregion
    }
    
    /// <summary>
    /// Border style for a cell.
    /// </summary>
    public class TFlxApplyBorders: ICloneable
    {
        #region Private members
        private bool FLeft;
        private bool FRight;
        private bool FTop;
        private bool FBottom;
        private bool FDiagonal;
        private bool FDiagonalStyle;
        #endregion
        /// <summary>
        /// Left border.
        /// </summary>
        public bool Left {get {return FLeft;} set {FLeft=value;}}

        /// <summary>
        /// Right border.
        /// </summary>
        public bool Right {get {return FRight;} set {FRight=value;}}
        
        /// <summary>
        /// Top border.
        /// </summary>
        public bool Top {get {return FTop;} set {FTop=value;}}
        
        /// <summary>
        /// Bottom border.
        /// </summary>
        public bool Bottom {get {return FBottom;} set {FBottom=value;}}
        
        /// <summary>
        /// Diagonal border.
        /// </summary>
        public bool Diagonal {get {return FDiagonal;} set {FDiagonal=value;}}

        /// <summary>
        /// When defined, there will be one or two diagonal lines across the cell.
        /// </summary>
        public bool DiagonalStyle {get {return FDiagonalStyle;} set {FDiagonalStyle=value;}}

        /// <summary>
        /// Sets all members to true or false
        /// </summary>
        public void SetAllMembers(bool Value)
        {
            FLeft = Value;
            FRight = Value;
            FTop = Value;
            FBottom = Value;
            FDiagonal = Value;
            FDiagonalStyle = Value;
        }

		/// <summary>
		/// Returns true if the format does not apply any setting.
		/// </summary>
		public bool IsEmpty
		{
			get
			{
				return
					FLeft == false &&
					FRight == false &&
					FTop == false &&
					FBottom == false &&
					FDiagonal == false &&
					FDiagonalStyle == false;
			}
		}


        /// <summary>
        /// This method will modify existingFormat with the properties from newFormat that are specified on this class
        /// </summary>
        /// <param name="existingFormat">Existing format that will be updated with the properties of newFormat specified.</param>
        /// <param name="newFormat">New format to apply</param>
        /// <returns>True if there was any change on existingFormat, false otherwise.</returns>
        public bool Apply(TFlxBorders existingFormat, TFlxBorders newFormat)
        {
            bool Result = false;
            if (FLeft && existingFormat.Left != newFormat.Left) {Result = true; existingFormat.Left = newFormat.Left;}
            if (FRight && existingFormat.Right != newFormat.Right) {Result = true; existingFormat.Right = newFormat.Right;}

            if (FTop && existingFormat.Top != newFormat.Top) {Result = true; existingFormat.Top = newFormat.Top;}
            if (FBottom && existingFormat.Bottom != newFormat.Bottom) {Result = true; existingFormat.Bottom = newFormat.Bottom;}
            if (FDiagonal && existingFormat.Diagonal != newFormat.Diagonal) {Result = true; existingFormat.Diagonal = newFormat.Diagonal;}
            if (FDiagonalStyle && existingFormat.DiagonalStyle != newFormat.DiagonalStyle) {Result = true; existingFormat.DiagonalStyle = newFormat.DiagonalStyle;}

            return Result;
        }

        #region ICloneable Members

        /// <summary>
        /// Returns a deep copy of the border.
        /// </summary>
        /// <returns></returns>
        public object Clone()
        {
            return MemberwiseClone();
        }

        #endregion
    }

    /// <summary>
    /// Fill pattern and color for the background of a cell.
    /// </summary>
    public struct TFlxApplyFillPattern
    {
        #region Privates
        private bool FPattern;
        private bool FFgColor;
        private bool FBgColor;
        private bool FGradient;
        #endregion

        #region Members
        /// <summary>
        /// Fill style.
        /// </summary>
        public bool Pattern {get {return FPattern;} set {FPattern=value;}}
        
        /// <summary>
        /// Color for the foreground of the pattern.
        /// </summary>
        public bool FgColor {get {return FFgColor;} set {FFgColor=value;}}

        /// <summary>
        /// Color for the background of the pattern.  If the pattern is solid, has no effect.
        /// </summary>
        public bool BgColor {get {return FBgColor;} set {FBgColor=value;}}

        /// <summary>
        /// Defines if to apply a gradient to a cell. Only valid in Excel 2007 or newer.
        /// </summary>
        public bool Gradient { get { return FGradient; } set { FGradient = value; } }

        /// <summary>
        /// Sets all members to true or false
        /// </summary>
        public void SetAllMembers(bool Value)
        {
            FPattern = Value;
            FFgColor = Value;
            FBgColor = Value;
            FGradient = Value;
        }

		/// <summary>
		/// Returns true if the format does not apply any setting.
		/// </summary>
		public bool IsEmpty
		{
			get
			{
				return
					FPattern == false &&
					FFgColor == false &&
					FBgColor == false &&
                    FGradient == false;
			}
		}


        /// <summary>
        /// This method will modify existingFormat with the properties from newFormat that are specified on this class
        /// </summary>
        /// <param name="existingFormat">Existing format that will be updated with the properties of newFormat specified.</param>
        /// <param name="newFormat">New format to apply</param>
        /// <returns>True if there was any change on existingFormat, false otherwise.</returns>
        public bool Apply(ref TFlxFillPattern existingFormat, ref TFlxFillPattern newFormat)
        {
            bool Result = false;
            if (FPattern && existingFormat.Pattern != newFormat.Pattern) { Result = true; existingFormat.Pattern = newFormat.Pattern; }
            if (FFgColor && existingFormat.FgColor != newFormat.FgColor) { Result = true; existingFormat.FgColor = newFormat.FgColor; }
            if (FBgColor && existingFormat.BgColor != newFormat.BgColor) { Result = true; existingFormat.BgColor = newFormat.BgColor; }
            if (FGradient && !TExcelGradient.Equals(existingFormat.Gradient, newFormat.Gradient))
            {
                Result = true;
                existingFormat.Gradient = newFormat.Gradient; //no need to clone, it is done by the setter in Gradient.
            }

            return Result;
        }


        #endregion

        /// <summary></summary>
        public override bool Equals(object obj)
        {
            if (obj == null || GetType() != obj.GetType()) return false;
            TFlxApplyFillPattern o2=(TFlxApplyFillPattern)obj;
            return (o2.Pattern==Pattern) && (o2.FgColor==FgColor) && (o2.BgColor==BgColor) && (o2.Gradient==Gradient);
        }

        /// <summary></summary>
        public static bool operator== (TFlxApplyFillPattern f1, TFlxApplyFillPattern f2)
        {
            return (f1.Pattern==f2.Pattern) && (f1.FgColor==f2.FgColor) && (f1.BgColor==f2.BgColor) && (f1.Gradient==f2.Gradient);
        }

        /// <summary></summary>
        public static bool operator!= (TFlxApplyFillPattern f1, TFlxApplyFillPattern f2)
        {
            return !(f1 == f2);
        }

        /// <summary></summary>
        public override int GetHashCode()
        {
#if(DELPHIWIN32)
            int Result = 0;
            if (Pattern) Result ++;
            if (FgColor) Result +=2;
            if (BgColor) Result += 4;
            return Result;
#else
            return (Pattern?1:0) + (FgColor?2:0) + (BgColor?4:0) + (Gradient?8:0);
#endif
        }

    }

    /// <summary>
    /// Defines which attributes of a <see cref="TFlxFormat"/> will be applied for one cell. 
    /// Whatever member is set to false, it will not apply this member property to the cell.
    /// </summary>
    public class TFlxApplyFormat: ICloneable
    {
        #region privates
        private TFlxApplyFont        FFont;
        private TFlxApplyBorders     FBorders;
        private bool  FFormat;
        private bool  FHAlignment;
        private bool  FVAlignment;

        private bool  FLocked;
        private bool  FHidden;
        private bool  FParentStyle;
        private bool  FWrapText;
        private bool  FShrinkToFit;
        private bool  FRotation;
        private bool  FIndent;
        private bool  FLotus123Prefix;
        #endregion
        /// <summary>
        /// Creates an empty Format class.
        /// </summary>
        public TFlxApplyFormat()
        {
            FFont = new TFlxApplyFont();
            FBorders = new TFlxApplyBorders();
            FillPattern = new TFlxApplyFillPattern();
        }

        /// <summary>
        /// Cell Font.
        /// </summary>
        public TFlxApplyFont        Font {get {return FFont;}}
     
        /// <summary>
        /// Cell borders.
        /// </summary>
        public TFlxApplyBorders     Borders {get {return FBorders;}}

		/// Format string.  (For example, "yyyy-mm-dd" for a date format, or "#.00" for a numeric 2 decimal format)
		/// <br/>This format string is the same you use in Excel unde "Custom" format when formatting a cell, and it is documented
		/// in Excel documentation. Under <b>"Finding out what format string to use in TFlxFormat.Format"</b> section in <b>UsingFlexCelAPI.pdf</b>
		/// you can find more detailed information on how to create this string.
		public bool         Format {get {return FFormat;} set {FFormat=value;}}

        /// <summary>
        /// Fill pattern.
        /// </summary>
        public TFlxApplyFillPattern FillPattern;

        /// <summary>
        /// Horizontal align on the cell.
        /// </summary>
        public bool  HAlignment {get {return FHAlignment;} set {FHAlignment=value;}}

        /// <summary>
        /// Vertical align on the cell.
        /// </summary>
        public bool  VAlignment {get {return FVAlignment;} set {FVAlignment=value;}}

        /// <summary>
        /// Cell is locked.
        /// </summary>
        public bool            Locked {get {return FLocked;} set {FLocked=value;}}

        /// <summary>
        /// Cell is Hidden.
        /// </summary>
        public bool            Hidden {get {return FHidden;} set {FHidden=value;}}

        /// <summary>
        /// Parent style. This is the parent style name and all the properties that are linked to it.
        /// </summary>
        public bool             ParentStyle {get {return FParentStyle;} set {FParentStyle=value;}}

        /// <summary>
        /// Cell wrap.
        /// </summary>
        public bool            WrapText {get {return FWrapText;} set {FWrapText=value;}}

        /// <summary>
        /// Shrink to fit.
        /// </summary>
        public bool            ShrinkToFit {get {return FShrinkToFit;} set {FShrinkToFit=value;}}

        /// <summary>
        /// Text Rotation on degrees. 
        /// 0 - 90 is up, 
        /// 91 - 180 is down, 
        /// 255 is vertical.
        /// </summary>
        public bool            Rotation {get {return FRotation;} set {FRotation=value;}}

        /// <summary>
        /// Indent value. (on characters)
        /// </summary>
        public bool            Indent {get {return FIndent;} set {FIndent=value;}}

        /// <summary>
        /// Lotus 123 compatibility prefixes.
        /// </summary>
        public bool Lotus123Prefix { get { return FLotus123Prefix; } set { FLotus123Prefix = value; } }


        #region Utility methods
        /// <summary>
        /// Sets all members to true or false
        /// </summary>
        public void SetAllMembers(bool Value)
        {
            FFont.SetAllMembers(Value);
            FBorders.SetAllMembers(Value);
            FFormat = Value;
            FillPattern.SetAllMembers(Value);
            FHAlignment = Value;
            FVAlignment = Value;

            FLocked = Value;
            FHidden = Value;
            FParentStyle = Value;
            FWrapText = Value;
            FShrinkToFit = Value;
            FRotation = Value;
            FIndent = Value;
            FLotus123Prefix = Value;
        }

		/// <summary>
		/// Returns true if the format does not apply any setting.
		/// </summary>
		public bool IsEmpty
		{
			get
			{
				return IsEmptyWithBorders(true);
			}
		}

		/// <summary>
		/// Returns true if the format definition contains only borders.
		/// </summary>
		public bool HasOnlyBorders
		{
			get
			{
				return IsEmptyWithBorders(false);
			}
		}

		private bool IsEmptyWithBorders(bool IncludeBorders)
		{
			return
				FFont.IsEmpty &&
				FFormat == false &&
				FillPattern.IsEmpty &&
				(!IncludeBorders || Borders.IsEmpty) &&

				FHAlignment == false &&
				FVAlignment == false &&

				FLocked == false &&
				FHidden == false &&
				FParentStyle == false &&
				FWrapText == false &&
				FShrinkToFit == false &&
				FRotation == false &&
				FIndent == false &&
                FLotus123Prefix == false;
		}

        /// <summary>
        /// This method will modify existingFormat with the properties from newFormat that are specified on this class
        /// </summary>
        /// <param name="existingFormat">Existing format that will be updated with the properties of newFormat specified.</param>
        /// <param name="newFormat">New format to apply</param>
        /// <returns>True if there was any change on existingFormat, false otherwise.</returns>
        public bool Apply(TFlxFormat existingFormat, TFlxFormat newFormat)
        {
            bool Result = false;
            if (FFont.Apply(existingFormat.Font, newFormat.Font)) Result = true;
            if (FBorders.Apply(existingFormat.Borders, newFormat.Borders)) Result = true;
            if (FFormat && existingFormat.Format != newFormat.Format) {Result = true; existingFormat.Format = newFormat.Format;}
            if (FillPattern.Apply(ref existingFormat.FillPattern, ref newFormat.FillPattern)) Result = true;
            if (FHAlignment && existingFormat.HAlignment != newFormat.HAlignment) {Result = true; existingFormat.HAlignment = newFormat.HAlignment;}
            if (FVAlignment && existingFormat.VAlignment != newFormat.VAlignment) {Result = true; existingFormat.VAlignment = newFormat.VAlignment;}

            if (FLocked && existingFormat.Locked != newFormat.Locked) {Result = true; existingFormat.Locked = newFormat.Locked;}
            if (FHidden && existingFormat.Hidden != newFormat.Hidden) {Result = true; existingFormat.Hidden = newFormat.Hidden;}
            
			if (FParentStyle)
			{
				if (existingFormat.NotNullParentStyle != newFormat.NotNullParentStyle || existingFormat.LinkedStyle.SameData(newFormat.LinkedStyle)) 
				{Result = true; existingFormat.ParentStyle = newFormat.ParentStyle; existingFormat.LinkedStyle.Assign(newFormat.LinkedStyle);}
			}

            if (FWrapText && existingFormat.WrapText != newFormat.WrapText) {Result = true; existingFormat.WrapText = newFormat.WrapText;}
            if (FShrinkToFit && existingFormat.ShrinkToFit != newFormat.ShrinkToFit) {Result = true; existingFormat.ShrinkToFit = newFormat.ShrinkToFit;}
            if (FRotation && existingFormat.Rotation != newFormat.Rotation) {Result = true; existingFormat.Rotation = newFormat.Rotation;}
            if (FIndent && existingFormat.Indent != newFormat.Indent) { Result = true; existingFormat.Indent = newFormat.Indent; }
            if (FLotus123Prefix && existingFormat.Lotus123Prefix != newFormat.Lotus123Prefix) { Result = true; existingFormat.Lotus123Prefix = newFormat.Lotus123Prefix; }
            
            return Result;
        }
        #endregion
        #region ICloneable Members

        /// <summary>
        /// Returns a deep copy of the format.
        /// </summary>
        /// <returns></returns>
        public object Clone()
        {
            TFlxApplyFormat Result = (TFlxApplyFormat) MemberwiseClone();
            Result.FFont = (TFlxApplyFont) Font.Clone();
            Result.FBorders = (TFlxApplyBorders) Borders.Clone();
            return Result;
        }

        #endregion
    }
    #endregion

	#region Styles

	/// <summary>
	/// This class is used as a part of a <see cref="TFlxFormat"/> class, and stores how a cell format is linked to a style.
	/// </summary>
    public class TLinkedStyle : ICloneable
    {
        #region Privates
        private bool FAutomaticChoose = true;
        private bool FLinkedNumericFormat = true;
        private bool FLinkedFont = true;
        private bool FLinkedAlignment = true;
        private bool FLinkedBorder = true;
        private bool FLinkedFill = true;
        private bool FLinkedProtection = true;

        #endregion

        #region Properties

        /// <summary>
        /// When this property is true (the default) FlexCel will automatically choose which linked
        /// properties to apply depending on what changes from the base style. For example, if this style
        /// has a different font than the basic style, the font will be not linked, and when you change the base style it will keep the same. 
        /// Excel behaves this way when it adds styles. To manually choose what the format will affect, set this to none.
        /// This property doesn't correspond with any Excel property, and it is not stored in the file.
        /// </summary>
        public bool AutomaticChoose { get { return FAutomaticChoose; } set { FAutomaticChoose = value; } }

        /// <summary>
        /// If true, the numeric format will be linked to the parent style, and it will change when you change the style.
        /// If false the numeric format will not change even if you change it in the base style.
        /// <b>Note that this property has no effect unless <see cref="AutomaticChoose"/> is false.</b>
        /// </summary>
        public bool LinkedNumericFormat { get { return FLinkedNumericFormat; } set { FLinkedNumericFormat = value; } }

        /// <summary>
        /// If true, the font will be linked to the parent style, and it will change when you change the style.
        /// If false the font will not change even if you change it in the base style.
        /// <b>Note that this property has no effect unless <see cref="AutomaticChoose"/> is false.</b>
        /// </summary>
        public bool LinkedFont { get { return FLinkedFont; } set { FLinkedFont = value; } }

        /// <summary>
        /// If true, the alignment will be linked to the parent style, and it will change when you change the style.
        /// If false the alignment not change even if you change it in the base style.
        /// <b>Note that this property has no effect unless <see cref="AutomaticChoose"/> is false.</b>
        /// </summary>
        public bool LinkedAlignment { get { return FLinkedAlignment; } set { FLinkedAlignment = value; } }

        /// <summary>
        /// If true, the border will be linked to the parent style, and it will change when you change the style.
        /// If false the border will not change even if you change it in the base style.
        /// <b>Note that this property has no effect unless <see cref="AutomaticChoose"/> is false.</b>
        /// </summary>
        public bool LinkedBorder { get { return FLinkedBorder; } set { FLinkedBorder = value; } }

        /// <summary>
        /// If true, the fill pattern will be linked to the parent style, and it will change when you change the style.
        /// If false the fill pattern will not change even if you change it in the base style.
        /// <b>Note that this property has no effect unless <see cref="AutomaticChoose"/> is false.</b>
        /// </summary>
        public bool LinkedFill { get { return FLinkedFill; } set { FLinkedFill = value; } }

        /// <summary>
        /// If true, the protection will be linked to the parent style, and it will change when you change the style.
        /// If false the protection will not change even if you change it in the base style.
        /// <b>Note that this property has no effect unless <see cref="AutomaticChoose"/> is false.</b>
        /// </summary>
        public bool LinkedProtection { get { return FLinkedProtection; } set { FLinkedProtection = value; } }


        /// <summary>
        /// Copies a new style into this object.
        /// </summary>
        /// <param name="newStyle">Style with the values to copy.</param>
        public void Assign(TLinkedStyle newStyle)
        {
            FLinkedNumericFormat = newStyle.LinkedNumericFormat;
            FLinkedFont = newStyle.FLinkedFont;
            FLinkedAlignment = newStyle.FLinkedAlignment;
            FLinkedBorder = newStyle.FLinkedBorder;
            FLinkedFill = newStyle.FLinkedFill;
            FLinkedProtection = newStyle.LinkedProtection;
        }

        /// <summary>
        /// Returns true if the 2 instances have the same data.
        /// </summary>
        /// <param name="otherStyle">Style with which we want to compare.</param>
        /// <returns>True if otherStyle has the same values as this object.</returns>
        public bool SameData(TLinkedStyle otherStyle)
        {
            return
                FLinkedNumericFormat == otherStyle.LinkedNumericFormat &&
                FLinkedFont == otherStyle.FLinkedFont &&
                FLinkedAlignment == otherStyle.FLinkedAlignment &&
                FLinkedBorder == otherStyle.FLinkedBorder &&
                FLinkedFill == otherStyle.FLinkedFill &&
                FLinkedProtection == otherStyle.LinkedProtection;
        }



        #endregion

        #region ICloneable Members

        /// <summary>
        /// Returns a deep copy of the object.
        /// </summary>
        /// <returns>A deep copy of the style. (without any references in common with the original object)</returns>
        public object Clone()
        {
            return MemberwiseClone();
        }

        #endregion

        /// <summary>
        /// Returns true if obj and this object have the same value.
        /// </summary>
        /// <param name="obj">Object to compare.</param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            TLinkedStyle o = obj as TLinkedStyle;
            if (o == null) return false;

            if (FAutomaticChoose != o.FAutomaticChoose) return false;
            if (FLinkedNumericFormat != o.FLinkedNumericFormat) return false;
            if (FLinkedFont != o.FLinkedFont) return false;
            if (FLinkedAlignment != o.FLinkedAlignment) return false;
            if (FLinkedBorder != o.FLinkedBorder) return false;
            if (FLinkedFill != o.FLinkedFill) return false;
            if (FLinkedProtection != o.FLinkedProtection) return false;

            return true;
        }

        /// <summary>
        /// Returns a hashcode for this object.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHash(
                FAutomaticChoose.GetHashCode(),
                FLinkedNumericFormat.GetHashCode(),
                FLinkedFont.GetHashCode(),
                FLinkedAlignment.GetHashCode(),
                FLinkedBorder.GetHashCode(),
                FLinkedFill.GetHashCode(),
                FLinkedProtection.GetHashCode());

        }
    }

	/// <summary>
	/// Enumerator with all built-in styles in Excel.
	/// </summary>
	public enum TBuiltInStyle
	{
		/// <summary>Normal style, used in all non formatted cells.</summary>
		Normal = 0,
		
		/// <summary>Style used in row outlines.</summary>
		RowLevel = 1,

		/// <summary>Style used in column outlines.</summary>
		ColLevel = 2,

		/// <summary>Built-in Style.</summary>
		Comma = 3,

		/// <summary>Built-in Style.</summary>
		Currency = 4,

		/// <summary>Built-in Style.</summary>
		Percent = 5,

		/// <summary>Built-in Style.</summary>
		Comma0 = 6,

		/// <summary>Built-in Style.</summary>
		Currency0 = 7,

		/// <summary>Built-in Style.</summary>
		Hyperlink = 8,

		/// <summary>Built-in Style.</summary>
		Followed_Hyperlink = 9,

		/// <summary>Standard style (not actually built in).</summary>
		Note = 10,

		/// <summary>Standard style (not actually built in).</summary>
		Warning_Text = 11,

		/// <summary>Standard style (not actually built in).</summary>
		Emphasis_1 = 12,

		/// <summary>Standard style (not actually built in).</summary>
		Emphasis_2 = 13,

		/// <summary>Standard style (not actually built in).</summary>
		Emphasis_3 = 14,

		/// <summary>Standard style (not actually built in).</summary>
		Title = 15,

		/// <summary>Standard style (not actually built in).</summary>
		Heading_1 = 16,

		/// <summary>Standard style (not actually built in).</summary>
		Heading_2 = 17,

		/// <summary>Standard style (not actually built in).</summary>
		Heading_3 = 18,

		/// <summary>Standard style (not actually built in).</summary>
		Heading_4 = 19,

		/// <summary>Standard style (not actually built in).</summary>
		Input = 20,

		/// <summary>Standard style (not actually built in).</summary>
		Output = 21,

		/// <summary>Standard style (not actually built in).</summary>
		Calculation = 22,

		/// <summary>Standard style (not actually built in).</summary>
		Check_Cell = 23,

		/// <summary>Standard style (not actually built in).</summary>
		Linked_Cell = 24,

		/// <summary>Standard style (not actually built in).</summary>
		Total = 25,

		/// <summary>Standard style (not actually built in).</summary>
		Good = 26,

		/// <summary>Standard style (not actually built in).</summary>
		Bad = 27,

		/// <summary>Standard style (not actually built in).</summary>
		Neutral = 28,

		/// <summary>Standard style (not actually built in).</summary>
		Accent1 = 29,

		/// <summary>Standard style (not actually built in).</summary>
		Accent1_20_percent = 30,

		/// <summary>Standard style (not actually built in).</summary>
		Accent1_40_percent = 31,

		/// <summary>Standard style (not actually built in).</summary>
		Accent1_60_percent = 32,

		/// <summary>Standard style (not actually built in).</summary>
		Accent2 = 33,

		/// <summary>Standard style (not actually built in).</summary>
		Accent2_20_percent = 34,

		/// <summary>Standard style (not actually built in).</summary>
		Accent2_40_percent = 35,

		/// <summary>Standard style (not actually built in).</summary>
		Accent2_60_percent = 36,

		/// <summary>Standard style (not actually built in).</summary>
		Accent3 = 37,

		/// <summary>Standard style (not actually built in).</summary>
		Accent3_20_percent = 38,

		/// <summary>Standard style (not actually built in).</summary>
		Accent3_40_percent = 39,

		/// <summary>Standard style (not actually built in).</summary>
		Accent3_60_percent = 40,

		/// <summary>Standard style (not actually built in).</summary>
		Accent4 = 41,

		/// <summary>Standard style (not actually built in).</summary>
		Accent4_20_percent = 42,

		/// <summary>Standard style (not actually built in).</summary>
		Accent4_40_percent = 43,

		/// <summary>Standard style (not actually built in).</summary>
		Accent4_60_percent = 44,

		/// <summary>Standard style (not actually built in).</summary>
		Accent5 = 45,

		/// <summary>Standard style (not actually built in).</summary>
		Accent5_20_percent = 46,

		/// <summary>Standard style (not actually built in).</summary>
		Accent5_40_percent = 47,

		/// <summary>Standard style (not actually built in).</summary>
		Accent5_60_percent = 48,

		/// <summary>Standard style (not actually built in).</summary>
		Accent6 = 49,

		/// <summary>Standard style (not actually built in).</summary>
		Accent6_20_percent = 50,

		/// <summary>Standard style (not actually built in).</summary>
		Accent6_40_percent = 51,

		/// <summary>Standard style (not actually built in).</summary>
		Accent6_60_percent = 52,

		/// <summary>Standard style (not actually built in).</summary>
		Explanatory_Text = 53
	}
	#endregion
}
