using System.Text;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

#if (MONOTOUCH)
	using real = System.Single;
	using System.Drawing;
    using Color = MonoTouch.UIKit.UIColor;
#else
	#if(WPF)
	using real = System.Double;
	using System.Windows;
	using System.Windows.Media;
	#else
	using real = System.Single;
	using System.Drawing;
	using Colors = System.Drawing.Color;
	#endif
#endif

namespace FlexCel.Core
{
    #region Enums
    /// <summary>
    /// A list of special colors that depend on the user settings in windows.
    /// </summary>
    internal enum TAutomaticColor
    {
        /// <summary>
        /// No Automatic color is applied.
        /// </summary>
        None,

        /// <summary>
        /// Font automatic color. This is the window text color.  
        /// </summary>
        Font,

        /// <summary>
        /// Default foreground color. This is the window text color in the sheet display
        /// </summary>
        DefaultForeground,

        /// <summary>
        /// Default background color.  This is the window background color in the sheet display and is the default background color for a cell.
        /// </summary>
        DefaultBackground,

        /// <summary>
        /// Default chart foreground color. This is the window text color in the chart display
        /// </summary>
        DefaultChartForeground,

        /// <summary>
        /// Default chart background color.  This is the window background color in the chart display.
        /// </summary>
        DefaultChartBackground,

        /// <summary>
        /// Chart neutral color (Black)
        /// </summary>
        ChartNeutralColor,

        /// <summary>
        /// ToolTip text color.  This is the automatic font color for comments
        /// </summary>
        TooltipTextColor,

        /// <summary>
        /// Sheet tab color.
        /// </summary>
        SheetTab

    }

    /// <summary>
    /// Defines which one of the values stored in a <see cref="TExcelColor"/> structure is the one that must be used.
    /// </summary>
    public enum TColorType
    {
        /// <summary>
        /// Color is Automatic.
        /// </summary>
        Automatic, //keep it at position 0

        /// <summary>
        /// The structure contains an indexed color.
        /// </summary>
        RGB,

        /// <summary>
        /// The structure contains a themed color.
        /// </summary>
        Theme,

        /// <summary>
        /// The structure contains an indexed color.
        /// </summary>
        Indexed
    }

    /// <summary>
    /// Specifies one of the 12 theme colors.
    /// </summary>
    public enum TThemeColor
    {
        /// <summary>
        /// No theme color.
        /// </summary>
        None = -1,

        /// <summary>
        /// Main color for backgrounds in Theme. This theme can also be called "Light1" by Excel.
        /// </summary>
        Background1 = 0x00,  // Excel Documentation is wrong, 0 is light1, not dark1

        /// <summary>
        /// Main color for foregrounds in Theme. This theme can also be called "Dark1" by Excel.
        /// </summary>
        Foreground1 = 0x01, // Excel Documentation is wrong, 1 is dark1, not light1

        /// <summary>
        /// Secondary color for backgrounds in Theme. This theme can also be called "Light2" by Excel.
        /// </summary>
        Background2 = 0x02,  //Again this is wrong in docs, switched light2 with dark2

        /// <summary>
        /// Secondary color for foregrounds in Theme. This theme can also be called "Dark2" by Excel.
        /// </summary>
        Foreground2 = 0x03,   //Again this is wrong in docs, switched light2 with dark2

        /// <summary>
        /// Accent1 Theme.
        /// </summary>
        Accent1 = 0x04,

        /// <summary>
        /// Accent2 Theme.
        /// </summary>
        Accent2 = 0x05,

        /// <summary>
        /// Accent3 Theme.
        /// </summary>
        Accent3 = 0x06,

        /// <summary>
        /// Accent4 Theme.
        /// </summary>
        Accent4 = 0x07,

        /// <summary>
        /// Accent5 Theme.
        /// </summary>
        Accent5 = 0x08,

        /// <summary>
        /// Accent6 Theme.
        /// </summary>
        Accent6 = 0x09,

        /// <summary>
        /// HyperLink Theme.
        /// </summary>
        HyperLink = 0x0A,

        /// <summary>
        /// FollowedHyperLink Theme.
        /// </summary>
        FollowedHyperLink = 0x0B
    }
    #endregion

    #region TExcelColor
    /// <summary>
    /// Represents an Excel color. Colors in Excel can be defined in four ways: Automatic Colors, Indexed Colors (for compatibility with Excel 2003 or older),
    /// Palette colors, and RGB colors. This Structure is immutable, once you create it you cannot change its members. You need to create a new struct to modify it.
    /// </summary>
    public struct TExcelColor : IComparable
    {
        #region Variables
        private readonly TColorType FColorType;
        private readonly TAutomaticColor FAutomaticType;
        private readonly long FRGB;
        private readonly int FIndex; //1 based. It is the biff8 index - 7, and contains no info in automatic colors. The property "Index" will show the usual 1..56 palette.
        private readonly TThemeColor FTheme;
        private readonly double FTint;
        #endregion

        #region Properties
        /// <summary>
        /// Identifies which kind of color is the one to apply in this structure.
        /// </summary>
        public TColorType ColorType { get { return FColorType; } }

        /// <summary>
        /// Returns the color when this structure has an RGB color, as a 0xRRGGBB integer. This property is fully functional with Excel 2007 or newer,
        /// older versions will be converted to Indexed color before saving as xls.
        /// <para></para><b>Note:</b> When reading a color, the value here might not be the final one, since <see cref="Tint"/> is applied to get the final color. Use 
        /// <see cref="ToColor(IFlexCelPalette)"/> method to find out the RGB color stored in this struct.
        /// <br/><br/>If you try to read the value of this property and <see cref="ColorType"/> is not the right kind, an Exception will be raised.
        /// </summary>
        public long RGB
        {
            get
            {
                CheckColorType("RGB", TColorType.RGB);
                return FRGB;
            }
        }

        /// <summary>
        /// Returns the color when this structure contains an indexed color (1 based). This property is for compatibility with xls files (Excel 2003 or older),
        /// but if you are not changing the color palette, even for older files, it is preferred to use <see cref="RGB"/> or <see cref="Theme"/> instead.
        /// <br/><br/>If you try to read the value of this property and <see cref="ColorType"/> is not the right kind, an Exception will be raised.
        /// </summary>
        public int Index
        {
            get
            {
                CheckColorType("Index", TColorType.Indexed);
                return FIndex < 1 ? FIndex + 8 : FIndex;
            }
        }

        internal int InternalIndex
        {
            get
            {
                return FIndex;
            }
        }

        /// <summary>
        /// Returns the color if it is one of the entries in the theme palette (1 based). 
        /// <br/><br/>If you try to read the value of this property and <see cref="ColorType"/> is not the right kind, an Exception will be raised.
        /// </summary>
        public TThemeColor Theme
        {
            get
            {
                CheckColorType("Theme", TColorType.Theme);
                return FTheme;
            }
        }

        /// <summary>
        /// Returns the tint value applied to the color. <br></br>
        /// If tint is supplied, then it is applied to the RGB value of the color to determine the final color applied. 
        /// <br></br>The tint value is stored as a double from -1.0 .. 1.0, where -1.0 means 100% darken and 1.0 means 100% lighten. Also, 0.0 means no change.
        /// </summary>
        public double Tint { get { return FTint; } }
        #endregion

        #region Setters

        private void CheckColorType(string PropName, TColorType aColorType)
        {
            if (aColorType != ColorType) FlxMessages.ThrowException(FlxErr.ErrInvalidColorType, "TExcelColor." + PropName, TCompactFramework.EnumGetName(typeof(TColorType), ColorType), TCompactFramework.EnumGetName(typeof(TColorType), aColorType));
        }

        private TExcelColor(TColorType aColorType, TAutomaticColor aAutomatic, long aRGB, TThemeColor aTheme, int aIndex, double aTint)
            : this(aColorType, aAutomatic, aRGB, aTheme, aIndex, aTint, false)
        { }

        private TExcelColor(TColorType aColorType, TAutomaticColor aAutomatic, long aRGB, TThemeColor aTheme, int aIndex, double aTint, bool AllowBiff8Indexed)
        {
            FColorType = aColorType;
            FTint = aTint < -1 ? -1 : aTint > 1 ? 1 : aTint;

            //if (aColorType == TColorType.Automatic && (!Enum.IsDefined(typeof(TAutomaticColor), aAutomatic) || aAutomatic == TAutomaticColor.None)) 
            //    FlxMessages.ThrowException(FlxErr.ErrInvalidColorEnum, "Automatic");
            FAutomaticType = aAutomatic;

            FRGB = aRGB & 0xFFFFFFFF;

            if (aColorType == TColorType.Theme && (!Enum.IsDefined(typeof(TThemeColor), aTheme) || aTheme == TThemeColor.None))
                FlxMessages.ThrowException(FlxErr.ErrInvalidColorEnum, "Theme");
            FTheme = aTheme;

            if (!AllowBiff8Indexed && aColorType == TColorType.Indexed && (aIndex < 1 || aIndex > 56))
            {
                FColorType = TColorType.Automatic;
                FAutomaticType = TAutomaticColor.None; //we won't raise an exception here, to avoid issues with old code.
                FIndex = -9;
                return;
            }
            FIndex = aIndex;
        }

        #endregion

        #region Compare

        /// <summary>
        /// Returns -1 if obj is more than color, 0 if both colors are the same, and 1 if obj is less than color.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public int CompareTo(object obj)
        {
            if (!(obj is TExcelColor)) return -1;
            TExcelColor Color2 = (TExcelColor)obj;

#if (COMPACTFRAMEWORK && !FRAMEWORK20)
			int Result = ((int)ColorType).CompareTo((int)Color2.ColorType);
#else
            int Result = ColorType.CompareTo(Color2.ColorType);
#endif
            if (Result != 0) return Result;

            int TintCompare = Tint.CompareTo(Color2.Tint);
            if (TintCompare != 0 && Math.Abs(Tint - Color2.Tint) > 1e-6) return TintCompare;

            switch (ColorType)
            {
                case TColorType.RGB:
                    return RGB.CompareTo(Color2.RGB);
                case TColorType.Automatic:
#if (COMPACTFRAMEWORK && !FRAMEWORK20)
					return ((int)FAutomaticType).CompareTo((int)Color2.FAutomaticType);
#else
					return FAutomaticType.CompareTo(Color2.FAutomaticType);
#endif
                case TColorType.Theme:
#if (COMPACTFRAMEWORK && !FRAMEWORK20)
					return ((int)Theme).CompareTo((int)Color2.Theme);
#else
                    return Theme.CompareTo(Color2.Theme);
#endif
                case TColorType.Indexed:
                    return FIndex.CompareTo(Color2.FIndex);
            }

            return 0;
        }

        /// <summary>
        /// Returns the hashcode of the object.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            switch (ColorType)
            {
                case TColorType.RGB:
                    return HashCoder.GetHash(((int)ColorType).GetHashCode(), RGB.GetHashCode(), Tint.GetHashCode());

                case TColorType.Automatic:
                    return HashCoder.GetHash(((int)ColorType).GetHashCode(), ((int)FAutomaticType).GetHashCode(), Tint.GetHashCode());

                case TColorType.Theme:
                    return HashCoder.GetHash(((int)ColorType).GetHashCode(), ((int)Theme).GetHashCode(), Tint.GetHashCode());

                case TColorType.Indexed:
                    return HashCoder.GetHash(((int)ColorType).GetHashCode(), FIndex.GetHashCode(), Tint.GetHashCode());
            }

            return ((int)ColorType).GetHashCode();
        }

        /// <summary>
        /// Returns true if both instances have the same color.
        /// </summary>
        /// <param name="obj">Object to compare.</param>
        /// <returns>True if both colors are the same.</returns>
        public override bool Equals(object obj)
        {
            return CompareTo(obj) == 0;
        }

        /// <summary>
        /// Returns true if both colors are equal.
        /// </summary>
        /// <param name="o1">First color to compare.</param>
        /// <param name="o2">Second color to compare.</param>
        /// <returns></returns>
        public static bool operator ==(TExcelColor o1, TExcelColor o2)
        {
            return o1.Equals(o2);
        }

        /// <summary>
        /// Returns true if both colors do not have the same value.
        /// </summary>
        /// <param name="o1">First color to compare.</param>
        /// <param name="o2">Second color to compare.</param>
        /// <returns></returns>
        public static bool operator !=(TExcelColor o1, TExcelColor o2)
        {
            return !(o1.Equals(o2));
        }

        /// <summary>
        /// Returns true is a color is less than the other.
        /// </summary>
        /// <param name="o1">First color to compare.</param>
        /// <param name="o2">Second color to compare.</param>
        /// <returns></returns>
        public static bool operator <(TExcelColor o1, TExcelColor o2)
        {
            return o1.CompareTo(o2) < 0;
        }

        /// <summary>
        /// Returns true is a color is bigger than the other.
        /// </summary>
        /// <param name="o1">First color to compare.</param>
        /// <param name="o2">Second color to compare.</param>
        /// <returns></returns>
        public static bool operator >(TExcelColor o1, TExcelColor o2)
        {
            return o1.CompareTo(o2) > 0;
        }

        #endregion

        #region Conversion

        private Color ApplyTint(Color aColor)
        {
            if (Tint == 0) return aColor; //so we don't create a new one.
            THSLColor Hsl = aColor;
            double NewBrightness = THSLColor.ApplyTint(Hsl.Lum, Tint);
            return new THSLColor(Hsl.Hue, Hsl.Sat, NewBrightness);
        }

        /// <summary>
        /// Returns a Color class with the specified rgb color.
        /// </summary>
        /// <param name="argb">Color to set.</param>
        /// <returns></returns>
        public static TExcelColor FromArgb(int argb)
        {
            return FromArgb(argb, 0);
        }

        /// <summary>
        /// Returns a Color class with the specified rgb color and with the specified tint.
        /// </summary>
        /// <param name="argb">Color to set.</param>
        /// <param name="tint">Tint for the color.
        /// <para></para>If you try to set a value less than -1 it will be stored as -1, and values bigger than 1 as 1. No exceptions will be raised.
        /// </param>
        /// <returns></returns>
        public static TExcelColor FromArgb(int argb, double tint)
        {
            return new TExcelColor(TColorType.RGB, TAutomaticColor.None, argb, TThemeColor.None, -9, tint);
        }

        /// <summary>
        /// Returns a Color class with the specified rgb color and with the specified tint.
        /// </summary>
        /// <param name="r">Red component of the color.</param>
        /// <param name="g">Green component of the color.</param>
        /// <param name="b">Blue component of the color.</param>
        /// <param name="tint">Tint for the color.
        /// <para></para>If you try to set a value less than -1 it will be stored as -1, and values bigger than 1 as 1. No exceptions will be raised.
        /// </param>
        /// <returns></returns>
        public static TExcelColor FromArgb(byte r, byte g, byte b, double tint)
        {
            unchecked
            {
                return FromArgb((int)((UInt32)0xFF000000 | (UInt32)(r << 16) | (UInt32)(g << 8) | b), tint);
            }
        }

        /// <summary>
        /// Returns a color class with an specified color index.
        /// <b>For compatibility with old code, you can enter any index here.</b> If the value is less than 1 or more than 56, it will assume automatic color.
        /// </summary>
        /// <param name="index">Index to the color palette. (1 based)</param>
        /// <returns></returns>
        public static TExcelColor FromIndex(int index)
        {
            return FromIndex(index, 0);
        }

        /// <summary>
        /// Returns a color class with an specified color index.
        /// <b>For compatibility with old code, you can enter any index here.</b> If the value is less than 1 or more than 56, it will assume automatic color.
        /// </summary>
        /// <param name="index">Index to the color palette. (1 based)</param>
        /// <param name="tint">Tint for the color.
        /// <para></para>If you try to set a value less than -1 it will be stored as -1, and values bigger than 1 as 1. No exceptions will be raised.
        /// </param>
        /// <returns></returns>
        public static TExcelColor FromIndex(int index, double tint)
        {
            return new TExcelColor(TColorType.Indexed, TAutomaticColor.None, -1, TThemeColor.None, index, tint);
        }

        /// <summary>
        /// Returns a color class with an specified theme color and tint.
        /// </summary>
        /// <param name="themeColor">Theme color index.</param>
        /// <returns></returns>
        public static TExcelColor FromTheme(TThemeColor themeColor)
        {
            return new TExcelColor(TColorType.Theme, TAutomaticColor.None, -1, themeColor, -9, 0);
        }

        /// <summary>
        /// Returns a color class with an specified theme color and tint.
        /// </summary>
        /// <param name="themeColor">Theme color index.</param>
        /// <param name="tint">Tint for the color.
        /// <para></para>If you try to set a value less than -1 it will be stored as -1, and values bigger than 1 as 1. No exceptions will be raised.
        /// </param>
        /// <returns></returns>
        public static TExcelColor FromTheme(TThemeColor themeColor, double tint)
        {
            return new TExcelColor(TColorType.Theme, TAutomaticColor.None, -1, themeColor, -9, tint);
        }

        /// <summary>
        /// Returns an standard Automatic color.
        /// </summary>
        public static TExcelColor Automatic
        {
            get
            {
                return new TExcelColor(TColorType.Automatic, TAutomaticColor.None, -1, TThemeColor.None, -9, 0);
            }
        }

        /// <summary>
        /// Returns true if this instance has an automatic color.
        /// </summary>
        /// <returns>True if this structure has an automatic color.</returns>
        public bool IsAutomatic
        {
            get
            {
                return ColorType == TColorType.Automatic;
            }
        }

        /// <summary>
        /// Assigns a system color to this instance.
        /// </summary>
        /// <param name="aColor"></param>
        /// <returns></returns>
        public static implicit operator TExcelColor(Color aColor)
        {
            return TExcelColor.FromArgb(aColor.ToArgb());
        }

        /// <summary>
        /// Returns the value of this class as a system color.
        /// </summary>
        /// <param name="xls">Excel file containing the themes and palettes for the color indexes.</param>
        /// <param name="automaticColor">Color to be returned if this structure has an automatic color.</param>
        /// <returns></returns>
        public Color ToColor(IFlexCelPalette xls, Color automaticColor)
        {
            if (automaticColor != ColorUtil.Empty && IsAutomatic) return ApplyTint(automaticColor);

            if (xls == null) xls = TDummyFlexCelPalette.Instance;
            Color Result = automaticColor;
            switch (ColorType)
            {
                case TColorType.RGB:
                    unchecked
                    {
                        Result = ColorUtil.FromArgb((int)((uint)(0xFF000000 | FRGB)));
                    }
                    break;
                case TColorType.Automatic:
                    Result = GetAutomaticRGB();
                    break;
                case TColorType.Theme:
                    Result = xls.GetColorTheme(FTheme).ToColor(xls);
                    break;
                case TColorType.Indexed:
                    Result = xls.GetColorPalette(Index);
                    break;
            }

            return ApplyTint(Result);
        }

        private Color GetAutomaticRGB()
        {
            switch (FAutomaticType)
            {
                case TAutomaticColor.Font:
                    return SystemColors.WindowText;
                case TAutomaticColor.DefaultForeground:
                    return SystemColors.WindowText;
                case TAutomaticColor.DefaultBackground:
                    return SystemColors.WindowFrame;
                case TAutomaticColor.DefaultChartForeground:
                    return SystemColors.WindowText;
                case TAutomaticColor.DefaultChartBackground:
                    return SystemColors.WindowFrame;
                case TAutomaticColor.ChartNeutralColor:
                    return Colors.Black;
                case TAutomaticColor.TooltipTextColor:
                    return SystemColors.InfoText;

                case TAutomaticColor.SheetTab:
                    return SystemColors.WindowFrame;

                default: return Colors.Black;
            }
        }

        /// <summary>
        /// Returns the value of this class as a system color.
        /// </summary>
        /// <param name="Xls">Excel file containing the themes and palettes for the color indexes.</param>
        public Color ToColor(IFlexCelPalette Xls)
        {
            return ToColor(Xls, ColorUtil.Empty);
        }

        #endregion

        #region Biff8

        private static TExcelColor FromAutomatic(TAutomaticColor AutColor)
        {
            return new TExcelColor(TColorType.Automatic, AutColor, -1, TThemeColor.None, -9, 0);
        }

        internal static TExcelColor FromBiff8ColorIndex(long aBiff8ColorIndex)
        {
            int BiffIndex = (int)(aBiff8ColorIndex & 0x7FFF);

            if (BiffIndex >= 0x00 && BiffIndex <= 0x03F)
            {
                //We will enter the real color in here. Note that we can't do this from the public interface.
                //So we can't use FromBiffIndex here.
                return new TExcelColor(TColorType.Indexed, TAutomaticColor.None, -1, TThemeColor.None, BiffIndex - 0x08 + 1, 0, true);
            }

            switch (BiffIndex)
            {
                case 0x040: return FromAutomatic(TAutomaticColor.DefaultForeground);
                case 0x041: return FromAutomatic(TAutomaticColor.DefaultBackground);
                case 0x04D: return FromAutomatic(TAutomaticColor.DefaultChartForeground);
                case 0x04E: return FromAutomatic(TAutomaticColor.DefaultChartBackground);
                case 0x04F: return FromAutomatic(TAutomaticColor.ChartNeutralColor);
                case 0x051: return FromAutomatic(TAutomaticColor.TooltipTextColor);
                case 0x07F: return FromAutomatic(TAutomaticColor.SheetTab);
            }

            return Automatic;
        }

        internal byte GetPxlColorIndex(IFlexCelPalette xls, TAutomaticColor autoType)
        {
            if (xls == null) xls = TDummyFlexCelPalette.Instance;
            switch (ColorType)
            {
                case TColorType.Automatic:
                    return 0xFF;
                case TColorType.RGB:
                    unchecked
                    {
                        Color cRGB = ColorUtil.FromArgb((int)((uint)(0xFF000000 | FRGB)));
                        return (byte)(xls.NearestColorIndex(ApplyTint(cRGB)) - 1);   // nearestcolorindex returns 1-based
                    }
                case TColorType.Theme:
                    return (byte)(xls.NearestColorIndex(ApplyTint(xls.GetColorTheme(FTheme).ToColor(xls))) - 1);
                case TColorType.Indexed:
                    return (byte)(Index - 1);  //index is 1-based too.
            }
            return 0xFF; //automatic
        }

        internal int GetBiff8ColorIndex(IFlexCelPalette xls, TAutomaticColor autoType, ref TColorIndexCache IndexCache)
        {
            if (xls == null) xls = TDummyFlexCelPalette.Instance;
            if (IndexCache.ColorIsValid(this, xls)) return IndexCache.Index - 1 + 8;

            int ColorIndex = GetBiff8ColorIndex(xls, autoType);

            IndexCache.LastColorStored = this;
            IndexCache.LastColorInPalette = ToColor(xls).ToArgb();
            IndexCache.Index = ColorIndex + 1 - 8; //IndexCache is the same base as index.
            if (ColorType == TColorType.Theme)
            {
                IndexCache.LastColorInTheme = xls.GetColorTheme(Theme);
            }

            return ColorIndex;
        }

        internal int GetBiff8ColorIndex(IFlexCelPalette xls, TAutomaticColor autoType)
        {
            if (xls == null) xls = TDummyFlexCelPalette.Instance;
            switch (ColorType)
            {
                case TColorType.RGB:
                    unchecked
                    {
                        Color cRGB = ColorUtil.FromArgb((int)((uint)(0xFF000000 | FRGB)));
                        return xls.NearestColorIndex(ApplyTint(cRGB)) - 1 + 8;   // nearestcolorindex returns 1-based
                    }
                case TColorType.Automatic:
                    return GetAutomaticColorIndex(autoType);
                case TColorType.Theme:
                    return xls.NearestColorIndex(ApplyTint(xls.GetColorTheme(FTheme).ToColor(xls))) - 1 + 8;
                case TColorType.Indexed:
                    return FIndex - 1 + 8;  //findex is 1-based too.
            }

            return 1;
        }


        private int GetAutomaticColorIndex(TAutomaticColor autoType)
        {
            if (FAutomaticType != TAutomaticColor.None) autoType = FAutomaticType; //if we have something stored here, we will use that better.
            switch (autoType)
            {
                case TAutomaticColor.DefaultForeground: return 0x040;
                case TAutomaticColor.DefaultBackground: return 0x041;
                case TAutomaticColor.DefaultChartForeground: return 0x04d;
                case TAutomaticColor.DefaultChartBackground: return 0x04e;
                case TAutomaticColor.ChartNeutralColor: return 0x04f;
                case TAutomaticColor.TooltipTextColor: return 0x051;
                case TAutomaticColor.SheetTab: return 0x07F;
                default: return 0x7FFF;
            }
        }

        internal static TExcelColor FromBiff8(byte[] Data)
        {
            return FromBiff8(Data, 0, 2, 4, true);
        }

        internal static TExcelColor FromBiff8(byte[] Data, int CTypePos, int TintPos, int DataPos, bool TintIsInt)
        {
            int ct = BitConverter.ToUInt16(Data, CTypePos);
            double Tint;

            if (TintIsInt)
            {
                int nTint = BitConverter.ToInt16(Data, TintPos);
                Tint = nTint / (double)Int16.MaxValue;
            }
            else
            {
                Tint = BitConverter.ToDouble(Data, TintPos);
            }

            if (Tint < -1) Tint = -1; // Just in case nTint was Int16.MinValue. Int16.MinValue < -Int16.MaxValue
            switch (ct)
            {
                case 0: //Automatic
                    return TExcelColor.Automatic;

                case 1: //Indexed
                    return TExcelColor.FromBiff8ColorIndex(BitConverter.ToUInt32(Data, DataPos));

                case 2: // RGB
                    unchecked
                    {
                        UInt32 BGR = BitConverter.ToUInt32(Data, DataPos);
                        return TExcelColor.FromArgb((byte)(BGR), (byte)(BGR >> 8), (byte)(BGR >> 16), Tint);
                    }

                case 3: //Theme
                    return TExcelColor.FromTheme((TThemeColor)(BitConverter.ToUInt32(Data, DataPos)), Tint);

            }

            return TExcelColor.Automatic;
        }


        #endregion

        #region Copy
        internal static TExcelColor Copy(TExcelColor aColor, IFlexCelPalette Source, IFlexCelPalette Dest)
        {
            if (aColor.ColorType != TColorType.Indexed) return aColor; //Themed colors are not converted when copying. Color Acc1 in source will be Color Acc2 in Dest, no matter if Acc and Dest are different.

            if (Source.GetColorPalette(aColor.Index).ToArgb() == Dest.GetColorPalette(aColor.Index).ToArgb()) return aColor; //Most common case, palettes are the same.

            return Source.GetColorPalette(aColor.Index);
        }
        #endregion
    }

    internal class TDummyFlexCelPalette : IFlexCelPalette
    {
        public static TDummyFlexCelPalette Instance = new TDummyFlexCelPalette();

        private TDummyFlexCelPalette()
        {
        }

        #region IFlexCelPalette Members

        public TThemeColor NearestColorTheme(Color value, out double tint)
        {
            tint = 0;
            return TThemeColor.None;
        }

        public int NearestColorIndex(Color value)
        {
            return -1;
        }

        public TDrawingColor GetColorTheme(TThemeColor themeIndex)
        {
            return ColorUtil.Empty;
        }

        public Color GetColorPalette(int index)
        {
            return ColorUtil.Empty;
        }

        public bool PaletteContainsColor(TExcelColor value)
        {
            return false;
        }

        public TTheme GetTheme()
        {
            return new TTheme();
        }

        #endregion
    }
    #endregion

    #region ColorCache
    /// <summary>
    /// Will store an index asociated with a color, so we save the same thing we loaded (except if the color or the palette changed).
    /// This will also speed up saving more than one time.
    /// </summary>
    struct TColorIndexCache
    {
        internal TExcelColor LastColorStored;
        internal int LastColorInPalette;
        internal TDrawingColor LastColorInTheme;
        internal int Index; //stores at 1 the first entry, that is biff8 entry 8.

        internal bool ColorIsValid(TExcelColor aColor, IFlexCelPalette xls)
        {
            int RealIndex = Index < 1 ? Index + 8 : Index;
            if (RealIndex <= 0 || RealIndex > 56) return false;
            if (LastColorStored != aColor) return false;
            Color c = xls.GetColorPalette(RealIndex);

            //If the color is Theme, it could have changed if the theme palette changed.
            if (aColor.ColorType == TColorType.Theme)
            {
                    if (xls.GetColorTheme(aColor.Theme) != LastColorInTheme) return false;

            }

            return c.ToArgb() == LastColorInPalette;
        }
    }
#endregion

    #region TExcelGradient

    /// <summary>
    /// The type of gradient stored inside a <see cref="TExcelGradient"/> object.
    /// </summary>
    public enum TGradientType
    {
        /// <summary>
        /// Linear gradient.
        /// </summary>
        Linear,

        /// <summary>
        /// Rectangular gradient.
        /// </summary>
        Rectangular
    }

#if(FRAMEWORK30)
    /// <summary>
    /// Represents one of the points in a Gradient definition for an Excel cell. Note that drawings (autoshapes, charts, etc)
    /// use a different Gradient definition: <see cref="TDrawingGradientStop"/>
    /// </summary>
#else
    /// <summary>
    /// Represents one of the points in a Gradient definition for an Excel cell.
    /// </summary>
#endif
    public struct TGradientStop: IComparable
    {
        private double FPosition;
        private TExcelColor FColor;

        /// <summary>
        /// This value must be between 0 and 1, and represents the position in the gradient where the <see cref="Color"/> in this structure is pure.
        /// </summary>
        public double Position { get { return FPosition; } set { if (value < 0 || value > 1) FlxMessages.ThrowException(FlxErr.ErrInvalidValue, "Position", value, 0, 1); FPosition = value; } }

        /// <summary>
        /// Color for this definition.
        /// </summary>
        public TExcelColor Color { get { return FColor; } set { FColor = value; } }

        /// <summary>
        /// Creates a new Gradient stop.
        /// </summary>
        /// <param name="aPosition">Position for the stop.</param>
        /// <param name="aColor">Color for the stop.</param>
        public TGradientStop(double aPosition, TExcelColor aColor)
        {
            FPosition = 0;
            FColor = aColor;// to compile

            Position = aPosition; //to set the real values
            Color = aColor;
        }

        #region IComparable Members

        /// <summary>
        /// Compares 2 instances of this struct.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public int CompareTo(object obj)
        {
            if (!(obj is TGradientStop)) return -1;
            TGradientStop o2 = (TGradientStop)obj;

            int Result = Position.CompareTo(o2.Position);
            if (Result != 0) return Result;
            Result = Color.CompareTo(o2.Color);
            if (Result != 0) return Result;

            return 0;
        }

        /// <summary>
        /// Returns if this struct has the same values as other one.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            return CompareTo(obj) == 0;
        }

        /// <summary>
        /// Returns the hashcode for this struct.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHash(Position.GetHashCode(), Color.GetHashCode());
        }

        /// <summary>
        /// Returns true if both gradient stops are equal.
        /// </summary>
        /// <param name="o1">First stop to compare.</param>
        /// <param name="o2">Second stop to compare.</param>
        /// <returns></returns>
        public static bool operator ==(TGradientStop o1, TGradientStop o2)
        {
            return o1.Equals(o2);
        }

        /// <summary>
        /// Returns true if both gradient stops are different.
        /// </summary>
        /// <param name="o1">First stop to compare.</param>
        /// <param name="o2">Second stop to compare.</param>
        /// <returns></returns>
        public static bool operator !=(TGradientStop o1, TGradientStop o2)
        {
            return !o1.Equals(o2);
        }

        /// <summary>
        /// Returns true if o1 is bigger than o2.
        /// </summary>
        /// <param name="o1"></param>
        /// <param name="o2"></param>
        /// <returns></returns>
        public static bool operator >(TGradientStop o1, TGradientStop o2)
        {
            return o1.CompareTo(o2) > 0;
        }

        /// <summary>
        /// Returns true if o1 is less than o2.
        /// </summary>
        /// <param name="o1"></param>
        /// <param name="o2"></param>
        /// <returns></returns>
        public static bool operator <(TGradientStop o1, TGradientStop o2)
        {
            return o1.CompareTo(o2) < 0;
        }


        #endregion
    }

    /// <summary>
    /// Represents a gradient fill for a background cell. This class is abstract, you need to use its children: <see cref="TExcelLinearGradient"/> and <see cref="TExcelRectangularGradient"/>
    /// </summary>
    public abstract class TExcelGradient
    {
        #region Privates
        /// <summary>
        /// Type of gradient.
        /// </summary>
        protected TGradientType FGradientType;
        private TGradientStop[] FStops;
        #endregion

        /// <summary>
        /// Type of gradient stored inside this object.
        /// </summary>
        public TGradientType GradientType { get { return FGradientType; } }

        /// <summary>
        /// Different colors used in the gradient. This array must have at least one stop, and no more than 256.
        /// </summary>
        public TGradientStop[] Stops { get { return FStops; } set { FStops = value; } }

        /// <summary>
        /// Creates a deep copy ot this object.
        /// </summary>
        /// <returns></returns>
        public TExcelGradient Clone()
        {
            TExcelGradient Result = (TExcelGradient)MemberwiseClone();
            if (Stops != null) Result.Stops = (TGradientStop[]) Stops.Clone();
            return Result;
        }

        /// <summary>
        /// Returns true if both classes contain the same gradient.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            TExcelGradient o2 = obj as TExcelGradient;
            if (o2 == null) return false;
            if (GradientType != o2.GradientType) return false;
            
            if (Stops == null)
            {
                return o2.Stops == null;
            }
            if (o2.Stops == null) return false;

            if (Stops.Length != o2.Stops.Length) return false;
            for (int i = 0; i < Stops.Length; i++)
            {
                if (Stops[i] != o2.Stops[i]) return false;
            }
            return true;
        }

        /// <summary>
        /// Returns the hashcode for this object.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            int StopsHashCode = Stops == null ? 0 : Stops.GetHashCode();
            return HashCoder.GetHash(GradientType.GetHashCode(), StopsHashCode);
        }

		#if (COMPACTFRAMEWORK && !FRAMEWORK20)
		public static bool Equals(Object o1, Object o2)
		{
			return (o1 != null && o1.Equals(o2)) || (o1 == null && o2 == null);
		}
		#endif
	}

    /// <summary>
    /// A linear gradient used for filling a background.
    /// </summary>
    public class TExcelLinearGradient: TExcelGradient
    {
        #region Privates
        private double FRotationAngle;
        #endregion

        /// <summary>
        /// Creates a new TExcelLinearGradient class.
        /// </summary>
        public TExcelLinearGradient()
        {
            FGradientType = TGradientType.Linear;
        }

        /// <summary>
        /// Creates a new TExcelLinearGradient instance.
        /// </summary>
        /// <param name="aStops">Gradient stops.</param>
        /// <param name="aRotationAngle">Rotation angle in degrees.</param>
        public TExcelLinearGradient(TGradientStop[] aStops, double aRotationAngle): this()
        {
            if (aStops != null) Stops = (TGradientStop[])aStops.Clone();
            FRotationAngle = aRotationAngle;
        }

        /// <summary>
        /// Rotation angle of the gradient in degrees.
        /// </summary>
        public double RotationAngle { get { return FRotationAngle; } set { FRotationAngle = value; } }

        /// <summary>
        /// Returns true if both classes contain the same gradient.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            if (!base.Equals(obj)) return false;
            TExcelLinearGradient o2 = obj as TExcelLinearGradient;
            if (o2 == null) return false;
            if (RotationAngle != o2.RotationAngle) return false;
            return true;
        }

        /// <summary>
        /// Returns the hashcode for this object.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHash(base.GetHashCode(), RotationAngle.GetHashCode());
        }        
    }

    /// <summary>
    /// A rectangular gradient used for filling a background.
    /// </summary>
    public class TExcelRectangularGradient : TExcelGradient
    {
        #region Privates
        private double FTop;
        private double FLeft;
        private double FBottom;
        private double FRight;
        #endregion

        /// <summary>
        /// Creates a new TExcelRectangularGradient class.
        /// </summary>
        public TExcelRectangularGradient()
        {
            FGradientType = TGradientType.Rectangular;
        }


        /// <summary>
        /// Creates a new rectangular gradient.
        /// </summary>
        /// <param name="aStops">Gradient stops.</param>
        /// <param name="aTop">Top coordinate for the gradient. This value must be between 0 and 1, and specifies in percent where the first color of the gradient will be placed.</param>
        /// <param name="aLeft">Left coordinate for the gradient. This value must be between 0 and 1, and specifies in percent where the first color of the gradient will be placed.</param>
        /// <param name="aBottom">Bottom coordinate for the gradient. This value must be between 0 and 1, and specifies in percent where the last color of the gradient will be placed.</param>
        /// <param name="aRight">Right coordinate for the gradient. This value must be between 0 and 1, and specifies in percent where the last color of the gradient will be placed.</param>
        public TExcelRectangularGradient(TGradientStop[] aStops, double aTop, double aLeft, double aBottom, double aRight): this()
        {
            if (aStops != null) Stops = (TGradientStop[])aStops.Clone();
            Top = aTop;
            Left = aLeft;
            Bottom = aBottom;
            Right = aRight;
        }

        private void Set01Value(double value, ref double prop, string propname)
        {
            if (value < 0 || value > 1) FlxMessages.ThrowException(FlxErr.ErrInvalidValue, propname, value, 0, 1);
            prop = value;
        }

        /// <summary>
        /// Top coordinate for the gradient. This value must be between 0 and 1, and specifies in percent where the first color of the gradient will be placed.
        /// </summary>
        public double Top { get { return FTop; } set { Set01Value(value, ref FTop, "Top"); } }

        /// <summary>
        /// Left coordinate for the gradient. This value must be between 0 and 1, and specifies in percent where the first color of the gradient will be placed.
        /// </summary>
        public double Left { get { return FLeft; } set { Set01Value(value, ref FLeft, "Left"); } }

        /// <summary>
        /// Bottom coordinate for the gradient. This value must be between 0 and 1, and specifies in percent where the last color of the gradient will be placed.
        /// </summary>
        public double Bottom { get { return FBottom; } set { Set01Value(value, ref FBottom, "Bottom"); } }

        /// <summary>
        /// Right coordinate for the gradient. This value must be between 0 and 1, and specifies in percent where the last color of the gradient will be placed.
        /// </summary>
        public double Right { get { return FRight; } set { Set01Value(value, ref FRight, "Right"); } }

        /// <summary>
        /// Returns true if both classes contain the same gradient.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            if (!base.Equals(obj)) return false;
            TExcelRectangularGradient o2 = obj as TExcelRectangularGradient;
            if (o2 == null) return false;
            if (FTop != o2.FTop) return false;
            if (FLeft != o2.FLeft) return false;
            if (FBottom != o2.FBottom) return false;
            if (FRight != o2.FRight) return false;
            return true;
        }

        /// <summary>
        /// Returns the hashcode for this object.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHash(base.GetHashCode(), FTop.GetHashCode(), FLeft.GetHashCode(), FBottom.GetHashCode(), FRight.GetHashCode());
        }

    }

    #endregion
}
