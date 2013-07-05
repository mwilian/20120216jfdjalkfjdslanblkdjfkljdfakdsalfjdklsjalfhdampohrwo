using System;
using System.Collections.Generic;

using System.Text;

#if (MONOTOUCH)
  using Color = MonoTouch.UIKit.UIColor;
  using System.Drawing;
#else
	#if (WPF)
	using System.Windows.Media;
	using System.Windows;
	#else
	using System.Drawing;
	using Colors = System.Drawing.Color;
	#endif
#endif

namespace FlexCel.Core
{
    /// <summary>
    /// Defines the kind of colors that might be stored inside a color definition in a drawing or a theme.
    /// </summary>
    public enum TDrawingColorType
    {
        /// <summary>
        /// Hue, Saturation, Luminance.
        /// </summary>
        HSL,

        /// <summary>
        /// Color is from the list in <see cref="TPresetColor"/>.
        /// </summary>
        Preset,

        /// <summary>
        /// RGB expressed as components. Components are in the range 0-255
        /// </summary>
        RGB,

        /// <summary>
        /// scRGB color mode. Components are in the range 0-1.
        /// </summary>
        scRGB,

        /// <summary>
        /// System color. This is defined by windows, for example the color of the active caption.
        /// </summary>
        System,

        /// <summary>
        /// Links to a theme color. You can't use this value when defining the theme colors themselves.
        /// </summary>
        Theme
    }

    #region Long color enumerations
    /// <summary>
    /// Preset colors.
    /// </summary>
    public enum TPresetColor
    {
        /// <summary>
        /// Color is not specified.
        /// </summary>
        None,

        /// <summary>
        /// Specifies a color with RGB value (240,248,255)
        /// </summary>
        AliceBlue,

        /// <summary>
        /// Specifies a color with RGB value (250,235,215)
        /// </summary>
        AntiqueWhite,

        /// <summary>
        /// Specifies a color with RGB value (0,255,255)
        /// </summary>
        Aqua,

        /// <summary>
        /// Specifies a color with RGB value (127,255,212)
        /// </summary>
        Aquamarine,

        /// <summary>
        /// Specifies a color with RGB value (240,255,255)
        /// </summary>
        Azure,

        /// <summary>
        /// Specifies a color with RGB value (245,245,220)
        /// </summary>
        Beige,

        /// <summary>
        /// Specifies a color with RGB value (255,228,196)
        /// </summary>
        Bisque,

        /// <summary>
        /// Specifies a color with RGB value (0,0,0)
        /// </summary>
        Black,

        /// <summary>
        /// Specifies a color with RGB value (255,235,205)
        /// </summary>
        BlanchedAlmond,

        /// <summary>
        /// Specifies a color with RGB value (0,0,255)
        /// </summary>
        Blue,

        /// <summary>
        /// Specifies a color with RGB value (138,43,226)
        /// </summary>
        BlueViolet,

        /// <summary>
        /// Specifies a color with RGB value (165,42,42)
        /// </summary>
        Brown,

        /// <summary>
        /// Specifies a color with RGB value (222,184,135)
        /// </summary>
        BurlyWood,

        /// <summary>
        /// Specifies a color with RGB value (95,158,160)
        /// </summary>
        CadetBlue,

        /// <summary>
        /// Specifies a color with RGB value (127,255,0)
        /// </summary>
        Chartreuse,

        /// <summary>
        /// Specifies a color with RGB value (210,105,30)
        /// </summary>
        Chocolate,

        /// <summary>
        /// Specifies a color with RGB value (255,127,80)
        /// </summary>
        Coral,

        /// <summary>
        /// Specifies a color with RGB value (100,149,237)
        /// </summary>
        CornflowerBlue,

        /// <summary>
        /// Specifies a color with RGB value (255,248,220)
        /// </summary>
        Cornsilk,

        /// <summary>
        /// Specifies a color with RGB value (220,20,60)
        /// </summary>
        Crimson,

        /// <summary>
        /// Specifies a color with RGB value (0,255,255)
        /// </summary>
        Cyan,

        /// <summary>
        /// Specifies a color with RGB value (255,20,147)
        /// </summary>
        DeepPink,

        /// <summary>
        /// Specifies a color with RGB value (0,191,255)
        /// </summary>
        DeepSkyBlue,

        /// <summary>
        /// Specifies a color with RGB value (105,105,105)
        /// </summary>
        DimGray,

        /// <summary>
        /// Specifies a color with RGB value (0,0,139)
        /// </summary>
        DkBlue,

        /// <summary>
        /// Specifies a color with RGB value (0,139,139)
        /// </summary>
        DkCyan,

        /// <summary>
        /// Specifies a color with RGB value (184,134,11)
        /// </summary>
        DkGoldenrod,

        /// <summary>
        /// Specifies a color with RGB value (169,169,169)
        /// </summary>
        DkGray,

        /// <summary>
        /// Specifies a color with RGB value (0,100,0)
        /// </summary>
        DkGreen,

        /// <summary>
        /// Specifies a color with RGB value (189,183,107)
        /// </summary>
        DkKhaki,

        /// <summary>
        /// Specifies a color with RGB value (139,0,139)
        /// </summary>
        DkMagenta,

        /// <summary>
        /// Specifies a color with RGB value (85,107,47)
        /// </summary>
        DkOliveGreen,

        /// <summary>
        /// Specifies a color with RGB value (255,140,0)
        /// </summary>
        DkOrange,

        /// <summary>
        /// Specifies a color with RGB value (153,50,204)
        /// </summary>
        DkOrchid,

        /// <summary>
        /// Specifies a color with RGB value (139,0,0)
        /// </summary>
        DkRed,

        /// <summary>
        /// Specifies a color with RGB value (233,150,122)
        /// </summary>
        DkSalmon,

        /// <summary>
        /// Specifies a color with RGB value (143,188,139)
        /// </summary>
        DkSeaGreen,

        /// <summary>
        /// Specifies a color with RGB value (72,61,139)
        /// </summary>
        DkSlateBlue,

        /// <summary>
        /// Specifies a color with RGB value (47,79,79)
        /// </summary>
        DkSlateGray,

        /// <summary>
        /// Specifies a color with RGB value (0,206,209)
        /// </summary>
        DkTurquoise,

        /// <summary>
        /// Specifies a color with RGB value (148,0,211)
        /// </summary>
        DkViolet,

        /// <summary>
        /// Specifies a color with RGB value (30,144,255)
        /// </summary>
        DodgerBlue,

        /// <summary>
        /// Specifies a color with RGB value (178,34,34)
        /// </summary>
        Firebrick,

        /// <summary>
        /// Specifies a color with RGB value (255,250,240)
        /// </summary>
        FloralWhite,

        /// <summary>
        /// Specifies a color with RGB value (34,139,34)
        /// </summary>
        ForestGreen,

        /// <summary>
        /// Specifies a color with RGB value (255,0,255)
        /// </summary>
        Fuchsia,

        /// <summary>
        /// Specifies a color with RGB value (220,220,220)
        /// </summary>
        Gainsboro,

        /// <summary>
        /// Specifies a color with RGB value (248,248,255)
        /// </summary>
        GhostWhite,

        /// <summary>
        /// Specifies a color with RGB value (255,215,0)
        /// </summary>
        Gold,

        /// <summary>
        /// Specifies a color with RGB value (218,165,32)
        /// </summary>
        Goldenrod,

        /// <summary>
        /// Specifies a color with RGB value (128,128,128)
        /// </summary>
        Gray,

        /// <summary>
        /// Specifies a color with RGB value (0,128,0)
        /// </summary>
        Green,

        /// <summary>
        /// Specifies a color with RGB value (173,255,47)
        /// </summary>
        GreenYellow,

        /// <summary>
        /// Specifies a color with RGB value (240,255,240)
        /// </summary>
        Honeydew,

        /// <summary>
        /// Specifies a color with RGB value (255,105,180)
        /// </summary>
        HotPink,

        /// <summary>
        /// Specifies a color with RGB value (205,92,92)
        /// </summary>
        IndianRed,

        /// <summary>
        /// Specifies a color with RGB value (75,0,130)
        /// </summary>
        Indigo,

        /// <summary>
        /// Specifies a color with RGB value (255,255,240)
        /// </summary>
        Ivory,

        /// <summary>
        /// Specifies a color with RGB value (240,230,140)
        /// </summary>
        Khaki,

        /// <summary>
        /// Specifies a color with RGB value (230,230,250)
        /// </summary>
        Lavender,

        /// <summary>
        /// Specifies a color with RGB value (255,240,245)
        /// </summary>
        LavenderBlush,

        /// <summary>
        /// Specifies a color with RGB value (124,252,0)
        /// </summary>
        LawnGreen,

        /// <summary>
        /// Specifies a color with RGB value (255,250,205)
        /// </summary>
        LemonChiffon,

        /// <summary>
        /// Specifies a color with RGB value (0,255,0)
        /// </summary>
        Lime,

        /// <summary>
        /// Specifies a color with RGB value (50,205,50)
        /// </summary>
        LimeGreen,

        /// <summary>
        /// Specifies a color with RGB value (250,240,230)
        /// </summary>
        Linen,

        /// <summary>
        /// Specifies a color with RGB value (173,216,230)
        /// </summary>
        LtBlue,

        /// <summary>
        /// Specifies a color with RGB value (240,128,128)
        /// </summary>
        LtCoral,

        /// <summary>
        /// Specifies a color with RGB value (224,255,255)
        /// </summary>
        LtCyan,

        /// <summary>
        /// Specifies a color with RGB value (250,250,120)
        /// </summary>
        LtGoldenrodYellow,

        /// <summary>
        /// Specifies a color with RGB value (211,211,211)
        /// </summary>
        LtGray,

        /// <summary>
        /// Specifies a color with RGB value (144,238,144)
        /// </summary>
        LtGreen,

        /// <summary>
        /// Specifies a color with RGB value (255,182,193)
        /// </summary>
        LtPink,

        /// <summary>
        /// Specifies a color with RGB value (255,160,122)
        /// </summary>
        LtSalmon,

        /// <summary>
        /// Specifies a color with RGB value (32,178,170)
        /// </summary>
        LtSeaGreen,

        /// <summary>
        /// Specifies a color with RGB value (135,206,250)
        /// </summary>
        LtSkyBlue,

        /// <summary>
        /// Specifies a color with RGB value (119,136,153)
        /// </summary>
        LtSlateGray,

        /// <summary>
        /// Specifies a color with RGB value (176,196,222)
        /// </summary>
        LtSteelBlue,

        /// <summary>
        /// Specifies a color with RGB value (255,255,224)
        /// </summary>
        LtYellow,

        /// <summary>
        /// Specifies a color with RGB value (255,0,255)
        /// </summary>
        Magenta,

        /// <summary>
        /// Specifies a color with RGB value (128,0,0)
        /// </summary>
        Maroon,

        /// <summary>
        /// Specifies a color with RGB value (102,205,170)
        /// </summary>
        MedAquamarine,

        /// <summary>
        /// Specifies a color with RGB value (0,0,205)
        /// </summary>
        MedBlue,

        /// <summary>
        /// Specifies a color with RGB value (186,85,211)
        /// </summary>
        MedOrchid,

        /// <summary>
        /// Specifies a color with RGB value (147,112,219)
        /// </summary>
        MedPurple,

        /// <summary>
        /// Specifies a color with RGB value (60,179,113)
        /// </summary>
        MedSeaGreen,

        /// <summary>
        /// Specifies a color with RGB value (123,104,238)
        /// </summary>
        MedSlateBlue,

        /// <summary>
        /// Specifies a color with RGB value (0,250,154)
        /// </summary>
        MedSpringGreen,

        /// <summary>
        /// Specifies a color with RGB value (72,209,204)
        /// </summary>
        MedTurquoise,

        /// <summary>
        /// Specifies a color with RGB value (199,21,133)
        /// </summary>
        MedVioletRed,

        /// <summary>
        /// Specifies a color with RGB value (25,25,112)
        /// </summary>
        MidnightBlue,

        /// <summary>
        /// Specifies a color with RGB value (245,255,250)
        /// </summary>
        MintCream,

        /// <summary>
        /// Specifies a color with RGB value (255,228,225)
        /// </summary>
        MistyRose,

        /// <summary>
        /// Specifies a color with RGB value (255,228,181)
        /// </summary>
        Moccasin,

        /// <summary>
        /// Specifies a color with RGB value (255,222,173)
        /// </summary>
        NavajoWhite,

        /// <summary>
        /// Specifies a color with RGB value (0,0,128)
        /// </summary>
        Navy,

        /// <summary>
        /// Specifies a color with RGB value (253,245,230)
        /// </summary>
        OldLace,

        /// <summary>
        /// Specifies a color with RGB value (128,128,0)
        /// </summary>
        Olive,

        /// <summary>
        /// Specifies a color with RGB value (107,142,35)
        /// </summary>
        OliveDrab,

        /// <summary>
        /// Specifies a color with RGB value (255,165,0)
        /// </summary>
        Orange,

        /// <summary>
        /// Specifies a color with RGB value (255,69,0)
        /// </summary>
        OrangeRed,

        /// <summary>
        /// Specifies a color with RGB value (218,112,214)
        /// </summary>
        Orchid,

        /// <summary>
        /// Specifies a color with RGB value (238,232,170)
        /// </summary>
        PaleGoldenrod,

        /// <summary>
        /// Specifies a color with RGB value (152,251,152)
        /// </summary>
        PaleGreen,

        /// <summary>
        /// Specifies a color with RGB value (175,238,238)
        /// </summary>
        PaleTurquoise,

        /// <summary>
        /// Specifies a color with RGB value (219,112,147)
        /// </summary>
        PaleVioletRed,

        /// <summary>
        /// Specifies a color with RGB value (255,239,213)
        /// </summary>
        PapayaWhip,

        /// <summary>
        /// Specifies a color with RGB value (255,218,185)
        /// </summary>
        PeachPuff,

        /// <summary>
        /// Specifies a color with RGB value (205,133,63)
        /// </summary>
        Peru,

        /// <summary>
        /// Specifies a color with RGB value (255,192,203)
        /// </summary>
        Pink,

        /// <summary>
        /// Specifies a color with RGB value (221,160,221)
        /// </summary>
        Plum,

        /// <summary>
        /// Specifies a color with RGB value (176,224,230)
        /// </summary>
        PowderBlue,

        /// <summary>
        /// Specifies a color with RGB value (128,0,128)
        /// </summary>
        Purple,

        /// <summary>
        /// Specifies a color with RGB value (255,0,0)
        /// </summary>
        Red,

        /// <summary>
        /// Specifies a color with RGB value (188,143,143)
        /// </summary>
        RosyBrown,

        /// <summary>
        /// Specifies a color with RGB value (65,105,225)
        /// </summary>
        RoyalBlue,

        /// <summary>
        /// Specifies a color with RGB value (139,69,19)
        /// </summary>
        SaddleBrown,

        /// <summary>
        /// Specifies a color with RGB value (250,128,114)
        /// </summary>
        Salmon,

        /// <summary>
        /// Specifies a color with RGB value (244,164,96)
        /// </summary>
        SandyBrown,

        /// <summary>
        /// Specifies a color with RGB value (46,139,87)
        /// </summary>
        SeaGreen,

        /// <summary>
        /// Specifies a color with RGB value (255,245,238)
        /// </summary>
        SeaShell,

        /// <summary>
        /// Specifies a color with RGB value (160,82,45)
        /// </summary>
        Sienna,

        /// <summary>
        /// Specifies a color with RGB value (192,192,192)
        /// </summary>
        Silver,

        /// <summary>
        /// Specifies a color with RGB value (135,206,235)
        /// </summary>
        SkyBlue,

        /// <summary>
        /// Specifies a color with RGB value (106,90,205)
        /// </summary>
        SlateBlue,

        /// <summary>
        /// Specifies a color with RGB value (112,128,144)
        /// </summary>
        SlateGray,

        /// <summary>
        /// Specifies a color with RGB value (255,250,250)
        /// </summary>
        Snow,

        /// <summary>
        /// Specifies a color with RGB value (0,255,127)
        /// </summary>
        SpringGreen,

        /// <summary>
        /// Specifies a color with RGB value (70,130,180)
        /// </summary>
        SteelBlue,

        /// <summary>
        /// Specifies a color with RGB value (210,180,140)
        /// </summary>
        Tan,

        /// <summary>
        /// Specifies a color with RGB value (0,128,128)
        /// </summary>
        Teal,

        /// <summary>
        /// Specifies a color with RGB value (216,191,216)
        /// </summary>
        Thistle,

        /// <summary>
        /// Specifies a color with RGB value (255,99,71)
        /// </summary>
        Tomato,

        /// <summary>
        /// Specifies a color with RGB value (64,224,208)
        /// </summary>
        Turquoise,

        /// <summary>
        /// Specifies a color with RGB value (238,130,238)
        /// </summary>
        Violet,

        /// <summary>
        /// Specifies a color with RGB value (245,222,179)
        /// </summary>
        Wheat,

        /// <summary>
        /// Specifies a color with RGB value (255,255,255)
        /// </summary>
        White,

        /// <summary>
        /// Specifies a color with RGB value (245,245,245)
        /// </summary>
        WhiteSmoke,

        /// <summary>
        /// Specifies a color with RGB value (255,255,0)
        /// </summary>
        Yellow,

        /// <summary>
        /// Specifies a color with RGB value (154,205,50)
        /// </summary>
        YellowGreen
    }

    /// <summary>
    /// System colors.
    /// </summary>
    public enum TSystemColor
    {
        /// <summary>
        /// Color is not specified.
        /// </summary>
        None,

        /// <summary>
        /// Specifies a Dark shadow color for three-dimensional display elements.
        /// </summary>
        DkShadow3d,

        /// <summary>
        /// Specifies a Light color for three-dimensional display elements (for edges facing the light source).
        /// </summary>
        Light3d,

        /// <summary>
        /// Specifies an Active Window Border Color.
        /// </summary>
        ActiveBorder,

        /// <summary>
        /// Specifies the active window title bar color. In particular the left side color in the color gradient of an active window's title bar if the gradient effect is enabled.
        /// </summary>
        ActiveCaption,

        /// <summary>
        /// Specifies the Background color of multiple document interface (MDI) applications.
        /// </summary>
        AppWorkspace,

        /// <summary>
        /// Specifies the desktop background color.
        /// </summary>
        Background,

        /// <summary>
        /// Specifies the face color for three-dimensional display elements and for dialog box backgrounds.
        /// </summary>
        BtnFace,

        /// <summary>
        /// Specifies the highlight color for three-dimensional display elements (for edges facing the light source).
        /// </summary>
        BtnHighlight,

        /// <summary>
        /// Specifies the shadow color for three-dimensional display elements (for edges facing away from the light source).
        /// </summary>
        BtnShadow,

        /// <summary>
        /// Specifies the color of text on push buttons.
        /// </summary>
        BtnText,

        /// <summary>
        /// Specifies the color of text in the caption, size box, and scroll bar arrow box.
        /// </summary>
        CaptionText,

        /// <summary>
        /// Specifies the right side color in the color gradient of an active window's title bar.
        /// </summary>
        GradientActiveCaption,

        /// <summary>
        /// Specifies the right side color in the color gradient of an inactive window's title bar.
        /// </summary>
        GradientInactiveCaption,

        /// <summary>
        /// Specifies a grayed (disabled) text. This color is set to 0 if the current display driver does not support a solid gray color.
        /// </summary>
        GrayText,

        /// <summary>
        /// Specifies the color of Item(s) selected in a control.
        /// </summary>
        Highlight,

        /// <summary>
        /// Specifies the text color of item(s) selected in a control.
        /// </summary>
        HighlightText,

        /// <summary>
        /// Specifies the color for a hyperlink or hot-tracked item.
        /// </summary>
        HotLight,

        /// <summary>
        /// Specifies the color of the Inactive window border.
        /// </summary>
        InactiveBorder,

        /// <summary>
        /// Specifies the color of the Inactive window caption. Specifies the left side color in the color gradient of an inactive window's title bar if the gradient effect is enabled.
        /// </summary>
        InactiveCaption,

        /// <summary>
        /// Specifies the color of text in an inactive caption.
        /// </summary>
        InactiveCaptionText,

        /// <summary>
        /// Specifies the background color for tooltip controls.
        /// </summary>
        InfoBk,

        /// <summary>
        /// Specifies the text color for tooltip controls.
        /// </summary>
        InfoText,

        /// <summary>
        /// Specifies the menu background color.
        /// </summary>
        Menu,

        /// <summary>
        /// Specifies the background color for the menu bar when menus appear as flat menus.
        /// </summary>
        MenuBar,

        /// <summary>
        /// Specifies the color used to highlight menu items when the menu appears as a flat menu.
        /// </summary>
        MenuHighlight,

        /// <summary>
        /// Specifies the color of Text in menus.
        /// </summary>
        MenuText,

        /// <summary>
        /// Specifies the scroll bar gray area color.
        /// </summary>
        ScrollBar,

        /// <summary>
        /// Specifies window background color.
        /// </summary>
        Window,

        /// <summary>
        /// Specifies the window frame color.
        /// </summary>
        WindowFrame,

        /// <summary>
        /// Specifies the color of text in windows.
        /// </summary>
        WindowText
    }

    #endregion

    #region Colors
    /// <summary>
    /// Represents a Color for a drawing or a theme. Different from TExcelColor, this structure is defined in terms of DrawingML, not SpreadsheetML.
    /// </summary>
    public struct TDrawingColor : IComparable
    {
        #region Variables
        private readonly TDrawingColorType FColorType;
        private readonly THSLColor FHSL;
        private readonly long FRGB;
        private readonly TScRGBColor FScRGB;
        private readonly TPresetColor FPreset;
        private readonly TSystemColor FSystem;
        private readonly TThemeColor FTheme;
        private readonly TColorTransform[] FColorTransform; //Will be immutable.
        private static readonly Dictionary<Color, TPresetColor> ColorsFromPreset = GetColorsFromPreset();
        private static readonly Dictionary<Color, TSystemColor> ColorsFromSystem = GetColorsFromSystem();
        #endregion

        #region Properties
        /// <summary>
        /// Identifies which kind of color is the one to apply in this structure.
        /// </summary>
        public TDrawingColorType ColorType { get { return FColorType; } }

        /// <summary>
        /// Returns the color when this structure has an HSL color, as a 0xHHSSLL integer. 
        /// <br/><br/>If you try to read the value of this property and <see cref="ColorType"/> is not the right kind, an Exception will be raised.
        /// </summary>
        public THSLColor HSL
        {
            get
            {
                CheckColorType("HSL", TDrawingColorType.HSL);
                return FHSL;
            }
        }

        /// <summary>
        /// Returns the color when this structure has an RGB color, as a 0xRRGGBB integer. 
        /// <br/><br/>If you try to read the value of this property and <see cref="ColorType"/> is not the right kind, an Exception will be raised.
        /// </summary>
        public long RGB
        {
            get
            {
                CheckColorType("RGB", TDrawingColorType.RGB);
                return FRGB;
            }
        }

        /// <summary>
        /// Returns the color when this structure has an scRGB color. 
        /// <br/><br/>If you try to read the value of this property and <see cref="ColorType"/> is not the right kind, an Exception will be raised.
        /// </summary>
        public TScRGBColor ScRGB
        {
            get
            {
                CheckColorType("ScRGB", TDrawingColorType.scRGB);
                return FScRGB;
            }
        }

        /// <summary>
        /// Returns the color when this structure has a Preset color. 
        /// <br/><br/>If you try to read the value of this property and <see cref="ColorType"/> is not the right kind, an Exception will be raised.
        /// </summary>
        public TPresetColor Preset
        {
            get
            {
                CheckColorType("Preset", TDrawingColorType.Preset);
                return FPreset;
            }
        }

        /// <summary>
        /// Returns the color when this structure has a System color. 
        /// <br/><br/>If you try to read the value of this property and <see cref="ColorType"/> is not the right kind, an Exception will be raised.
        /// </summary>
        public TSystemColor System
        {
            get
            {
                CheckColorType("System", TDrawingColorType.System);
                return FSystem;
            }
        }

        /// <summary>
        /// Returns the color when this structure has a Themed color. 
        /// <br/><br/>If you try to read the value of this property and <see cref="ColorType"/> is not the right kind, an Exception will be raised.
        /// </summary>
        public TThemeColor Theme
        {
            get
            {
                CheckColorType("Theme", TDrawingColorType.Theme);
                return FTheme;
            }
        }

        /// <summary>
        /// Returns an array with all the color transforms in this object.
        /// </summary>
        public TColorTransform[] GetColorTransform()
        {
            if (FColorTransform == null) return null;
            TColorTransform[] Result = new TColorTransform[FColorTransform.Length];
            FColorTransform.CopyTo(Result, 0);
            return Result;
        }

        #endregion

        #region Setters

        private void CheckColorType(string PropName, TDrawingColorType aColorType)
        {
            if (aColorType != ColorType) FlxMessages.ThrowException(FlxErr.ErrInvalidColorType, "TDrawingColor." + PropName, TCompactFramework.EnumGetName(typeof(TDrawingColorType), ColorType), TCompactFramework.EnumGetName(typeof(TDrawingColorType), aColorType));
        }

        private TDrawingColor(TDrawingColorType aColorType, THSLColor aHSL, long aRGB, TScRGBColor aScRGB, TPresetColor aPreset, TSystemColor aSystem, TThemeColor aTheme, TColorTransform[] aColorTransform)
        {
            FColorType = aColorType;
            FHSL = aHSL;
            FRGB = aRGB & 0xFFFFFFFF;
            FScRGB = aScRGB;
            if (aColorType == TDrawingColorType.Preset && (!Enum.IsDefined(typeof(TPresetColor), aPreset) || aPreset == TPresetColor.None))
                FlxMessages.ThrowException(FlxErr.ErrInvalidColorEnum, "Preset");
            FPreset = aPreset;

            if (aColorType == TDrawingColorType.System && (!Enum.IsDefined(typeof(TSystemColor), aSystem) || aSystem == TSystemColor.None))
                FlxMessages.ThrowException(FlxErr.ErrInvalidColorEnum, "System");
            FSystem = aSystem;

            FTheme = aTheme;

            FColorTransform = aColorTransform;
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
            if (!(obj is TDrawingColor)) return -1;
            TDrawingColor Color2 = (TDrawingColor)obj;

#if (COMPACTFRAMEWORK && !FRAMEWORK20)
			int Result = ((int)ColorType).CompareTo((int)Color2.ColorType);
#else
            int Result = ColorType.CompareTo(Color2.ColorType);
#endif
            if (Result != 0) return Result;

            switch (ColorType)
            {
                case TDrawingColorType.HSL:
                    return HSL.CompareTo(Color2.HSL);

                case TDrawingColorType.Preset:
#if (COMPACTFRAMEWORK && !FRAMEWORK20)
					return ((int)Preset).CompareTo((int)Preset);
#else
                    return Preset.CompareTo(Preset);
#endif

                case TDrawingColorType.RGB:
                    return RGB.CompareTo(Color2.RGB);

                case TDrawingColorType.scRGB:
                    return ScRGB.CompareTo(Color2.ScRGB);

                case TDrawingColorType.System:

#if (COMPACTFRAMEWORK && !FRAMEWORK20)
					return ((int)System).CompareTo((int)Color2.System);
#else
                    return System.CompareTo(Color2.System);
#endif
                case TDrawingColorType.Theme:
#if (COMPACTFRAMEWORK && !FRAMEWORK20)
					return ((int)Theme).CompareTo((int)Color2.Theme);
#else
                    return Theme.CompareTo(Color2.Theme);
#endif
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
                case TDrawingColorType.HSL:
                    return HashCoder.GetHash(ColorType.GetHashCode(), HSL.GetHashCode());

                case TDrawingColorType.Preset:
                    return HashCoder.GetHash(ColorType.GetHashCode(), Preset.GetHashCode());

                case TDrawingColorType.RGB:
                    return HashCoder.GetHash(ColorType.GetHashCode(), RGB.GetHashCode());

                case TDrawingColorType.scRGB:
                    return HashCoder.GetHash(ColorType.GetHashCode(), ScRGB.GetHashCode());

                case TDrawingColorType.System:
                    return HashCoder.GetHash(ColorType.GetHashCode(), System.GetHashCode());

                case TDrawingColorType.Theme:
                    return HashCoder.GetHash(ColorType.GetHashCode(), Theme.GetHashCode());
            }
            return ColorType.GetHashCode();
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
        public static bool operator ==(TDrawingColor o1, TDrawingColor o2)
        {
            return o1.Equals(o2);
        }

        /// <summary>
        /// Returns true if both colors do not have the same value.
        /// </summary>
        /// <param name="o1">First color to compare.</param>
        /// <param name="o2">Second color to compare.</param>
        /// <returns></returns>
        public static bool operator !=(TDrawingColor o1, TDrawingColor o2)
        {
            return !(o1.Equals(o2));
        }

        /// <summary>
        /// Returns true if o1 is bigger than o2.
        /// </summary>
        /// <param name="o1"></param>
        /// <param name="o2"></param>
        /// <returns></returns>
        public static bool operator >(TDrawingColor o1, TDrawingColor o2)
        {
            return o1.CompareTo(o2) > 0;
        }

        /// <summary>
        /// Returns true if o1 is less than o2.
        /// </summary>
        /// <param name="o1"></param>
        /// <param name="o2"></param>
        /// <returns></returns>
        public static bool operator <(TDrawingColor o1, TDrawingColor o2)
        {
            return o1.CompareTo(o2) < 0;
        }

        #endregion

        #region Conversion
        /// <summary>
        /// Returns a color class with an specified preset color.
        /// </summary>
        /// <param name="aColor">Color that we want to set.</param>
        /// <returns>The corresponding preset color.</returns>
        public static TDrawingColor FromPreset(TPresetColor aColor)
        {
            return new TDrawingColor(TDrawingColorType.Preset, ColorUtil.Empty, -1, ColorUtil.Empty, aColor, TSystemColor.None, TThemeColor.None, null);
        }

        /// <summary>
        /// Returns a color class with an specified system color.
        /// </summary>
        /// <param name="aColor">Color that we want to set.</param>
        /// <returns>The corresponding system color.</returns>
        public static TDrawingColor FromSystem(TSystemColor aColor)
        {
            return new TDrawingColor(TDrawingColorType.System, ColorUtil.Empty, -1, ColorUtil.Empty, TPresetColor.None, aColor, TThemeColor.None, null);
        }

        /// <summary>
        /// Returns a color class with an specified color. There is no real need to call this method, since conversion between 
        /// <see cref="TDrawingColor"/> and <see cref="Color"/> is implicit. You can just assign Color to this class and viceversa.
        /// </summary>
        /// <param name="aColor">Color that we want to set.</param>
        /// <returns>The corresponding system color.</returns>
        public static TDrawingColor FromColor(Color aColor)
        {
#if (COMPACTFRAMEWORK || (!FRAMEWORK30 && !MONOTOUCH))
            if (aColor.IsSystemColor)
#else
            if (aColor.IsSystemColor())
#endif
            {
                TSystemColor SysColor = GetSysColor(aColor);
                if (SysColor != TSystemColor.None)
                {
                    return new TDrawingColor(TDrawingColorType.System, ColorUtil.Empty, -1, ColorUtil.Empty, TPresetColor.None, SysColor, TThemeColor.None, null);
                }
            }

#if (!COMPACTFRAMEWORK && (FRAMEWORK30 || MONOTOUCH) )
            if (aColor.IsNamedColor())
            {
                TPresetColor PresetColor = GetPresetColor(aColor);
                if (PresetColor != TPresetColor.None)
                {
                    return new TDrawingColor(TDrawingColorType.Preset, ColorUtil.Empty, -1, ColorUtil.Empty, PresetColor, TSystemColor.None, TThemeColor.None, null);
                }
            }
#endif
            unchecked
            {
                //Alpha here is always 1. We might set alpha later with a transform.
                return new TDrawingColor(TDrawingColorType.RGB, ColorUtil.Empty, ((UInt32)aColor.ToArgb()) | 0xFF000000, ColorUtil.Empty, TPresetColor.None, TSystemColor.None, TThemeColor.None, null);
            }
        }

        /// <summary>
        /// Returns a color from its byte components.
        /// </summary>
        /// <param name="r">Red component.</param>
        /// <param name="g">Green component.</param>
        /// <param name="b">Blue component.</param>
        /// <returns></returns>
        public static TDrawingColor FromRgb(byte r, byte g, byte b)
        {
            unchecked
            {
                //Alpha here is always 1. We might set alpha later with a transform.
                UInt32 ARGB = (UInt32)0xFF000000 | (UInt32)(r << 16) | (UInt32)(g << 8) | (b);
                return new TDrawingColor(TDrawingColorType.RGB, ColorUtil.Empty, ARGB,
                    ColorUtil.Empty, TPresetColor.None, TSystemColor.None, TThemeColor.None, null);
            }
        }

        /// <summary>
        /// Returns a color class with an specified scRGB color.
        /// </summary>
        /// <param name="aColor">Color that we want to set.</param>
        /// <returns>The corresponding system color.</returns>
        public static TDrawingColor FromScRgb(TScRGBColor aColor)
        {
            return new TDrawingColor(TDrawingColorType.scRGB, ColorUtil.Empty, -1, aColor, TPresetColor.None, TSystemColor.None, TThemeColor.None, null);
        }

        /// <summary>
        /// Returns a color class with an specified themed color.
        /// </summary>
        /// <param name="aColor">Color that we want to set.</param>
        /// <returns>The corresponding system color.</returns>
        public static TDrawingColor FromTheme(TThemeColor aColor)
        {
            return new TDrawingColor(TDrawingColorType.Theme, ColorUtil.Empty, -1, ColorUtil.Empty, TPresetColor.None, TSystemColor.None, aColor, null);
        }

        /// <summary>
        /// Returns a color class with an specified HSL color.
        /// </summary>
        /// <param name="aColor">Color that we want to set.</param>
        /// <returns>The corresponding system color.</returns>
        public static TDrawingColor FromHSL(THSLColor aColor)
        {
            return new TDrawingColor(TDrawingColorType.HSL, aColor, -1, ColorUtil.Empty, TPresetColor.None, TSystemColor.None, TThemeColor.None, null);
        }

        /// <summary>
        /// Returns the DrawingColor that results of applying the transform to the existing DrawingColor.
        /// </summary>
        /// <param name="oldColor">Original color where we will apply the transform.</param>
        /// <param name="aTransform">Transform to be applied to oldColor.</param>
        /// <returns></returns>
        public static TDrawingColor AddTransform(TDrawingColor oldColor, TColorTransform[] aTransform)
        {
            if ((aTransform == null || aTransform.Length == 0) && (oldColor.FColorTransform == null || oldColor.FColorTransform.Length == 0)) return oldColor;
            TColorTransform[] NewTransform = null; //clone it as it will be immutable later. This struct could be copied with memberwiseclone, and there should be no way to change this array (as it might be shared with others)
            if (aTransform != null)
            {
                NewTransform = (TColorTransform[])aTransform.Clone();
            }
            return new TDrawingColor(oldColor.FColorType, oldColor.FHSL, oldColor.FRGB, oldColor.FScRGB, oldColor.FPreset, oldColor.FSystem, oldColor.FTheme, NewTransform);
        }

        /// <summary>
        /// Returns the .NET Color specified by this structure. 
        /// </summary>
        /// <param name="xls">Excel file containing the themes and palettes for the color indexes.</param>
        /// <returns>The corresponding .NET color.</returns>
        public Color ToColor(IFlexCelPalette xls)
        {
            switch (ColorType)
            {
                case TDrawingColorType.HSL:
                    return Transform(ColorUtil.FromArgb(FHSL.R, FHSL.G, FHSL.B));
                case TDrawingColorType.Preset:
                    return Transform(GetPresetColor(FPreset));
                case TDrawingColorType.RGB:
                    unchecked
                    {
                        return Transform(ColorUtil.FromArgb((int)((uint)(0xFF000000 | FRGB))));
                    }
                case TDrawingColorType.scRGB:
                    return Transform(ScRGB);

                case TDrawingColorType.System:
                    return Transform(GetSystemColor(FSystem));

                case TDrawingColorType.Theme:
                    TDrawingColor ColTheme = xls.GetColorTheme(FTheme);
                    if (ColTheme.ColorType == TDrawingColorType.Theme) FlxMessages.ThrowException(FlxErr.ErrInternal); //This could create an infinite recursion. But it can't happen, beacuse themes can't contain themed colors as definition.
                    return Transform(ColTheme.ToColor(xls));
            }

            return ColorUtil.Empty;
        }

        private Color Transform(Color aColor)
        {
            if (FColorTransform == null || FColorTransform.Length == 0) return aColor;
            foreach (TColorTransform ct in FColorTransform)
            {
                aColor = ct.Transform(aColor);
            }
            return aColor;
        }

        #region Boring conversions
        /// <summary>
        /// Returns the color associated with a simple color.
        /// </summary>
        /// <param name="aSystem"></param>
        /// <returns></returns>
        public static Color GetSystemColor(TSystemColor aSystem)
        {
            switch (aSystem) //We could use Color.FromName here to make this code simpler, but it wouldn't guarantee that all Color.SystemColors map into TSystemColors.
            {
                case TSystemColor.None:
                    return ColorUtil.Empty;

                case TSystemColor.DkShadow3d:
                    return SystemColors.ControlDarkDark;

                case TSystemColor.Light3d:
                    return SystemColors.ControlLight;

                case TSystemColor.ActiveBorder:
                    return SystemColors.ActiveBorder;

                case TSystemColor.ActiveCaption:
                    return SystemColors.ActiveCaption;

                case TSystemColor.AppWorkspace:
                    return SystemColors.AppWorkspace;

                case TSystemColor.Background:
                    return SystemColors.Control;

                case TSystemColor.BtnFace:
#if (FRAMEWORK20 && !COMPACTFRAMEWORK)
                    return SystemColors.ButtonFace;
#else
                    return SystemColors.Control;
#endif

                case TSystemColor.BtnHighlight:
#if (FRAMEWORK20 && !COMPACTFRAMEWORK)
                    return SystemColors.ButtonHighlight;
#else
                    return SystemColors.ControlLightLight;
#endif

                case TSystemColor.BtnShadow:
#if (FRAMEWORK20 && !COMPACTFRAMEWORK)
                    return SystemColors.ButtonShadow;
#else
                    return SystemColors.ControlDarkDark;
#endif

                case TSystemColor.BtnText:
                    return SystemColors.ControlText;

                case TSystemColor.CaptionText:
                    return SystemColors.ActiveCaptionText;

                case TSystemColor.GradientActiveCaption:
#if (FRAMEWORK20 && !COMPACTFRAMEWORK)
                    return SystemColors.GradientActiveCaption;
#else
                    return SystemColors.ActiveCaption;
#endif

                case TSystemColor.GradientInactiveCaption:
#if (FRAMEWORK20 && !COMPACTFRAMEWORK)
                    return SystemColors.GradientInactiveCaption;
#else
                    return SystemColors.InactiveCaption;
#endif

                case TSystemColor.GrayText:
                    return SystemColors.GrayText;

                case TSystemColor.Highlight:
                    return SystemColors.Highlight;

                case TSystemColor.HighlightText:
                    return SystemColors.HighlightText;

                case TSystemColor.HotLight:
                    return SystemColors.ControlLightLight;

                case TSystemColor.InactiveBorder:
                    return SystemColors.InactiveBorder;

                case TSystemColor.InactiveCaption:
                    return SystemColors.InactiveCaption;

                case TSystemColor.InactiveCaptionText:
                    return SystemColors.InactiveCaptionText;

                case TSystemColor.InfoBk:
                    return SystemColors.Info;

                case TSystemColor.InfoText:
                    return SystemColors.InfoText;

                case TSystemColor.Menu:
                    return SystemColors.Menu;

                case TSystemColor.MenuBar:
#if (FRAMEWORK20 && !COMPACTFRAMEWORK)
                    return SystemColors.MenuBar;
#else
                    return SystemColors.Menu;
#endif

                case TSystemColor.MenuHighlight:
#if (FRAMEWORK20 && !COMPACTFRAMEWORK)
                    return SystemColors.MenuHighlight;
#else
                    return SystemColors.ControlLightLight;
#endif

                case TSystemColor.MenuText:
                    return SystemColors.MenuText;

                case TSystemColor.ScrollBar:
                    return SystemColors.ScrollBar;

                case TSystemColor.Window:
                    return SystemColors.Window;

                case TSystemColor.WindowFrame:
                    return SystemColors.WindowFrame;

                case TSystemColor.WindowText:
                    return SystemColors.WindowText;

            }
            return ColorUtil.Empty;
        }

        private static Color GetPresetColor(TPresetColor aPreset)
        {
            switch (aPreset)
            {
                case TPresetColor.None:
                    return ColorUtil.Empty;

                case TPresetColor.AliceBlue:
                    return Colors.AliceBlue;

                case TPresetColor.AntiqueWhite:
                    return Colors.AntiqueWhite;

                case TPresetColor.Aqua:
                    return Colors.Aqua;

                case TPresetColor.Aquamarine:
                    return Colors.Aquamarine;

                case TPresetColor.Azure:
                    return Colors.Azure;

                case TPresetColor.Beige:
                    return Colors.Beige;

                case TPresetColor.Bisque:
                    return Colors.Bisque;

                case TPresetColor.Black:
                    return Colors.Black;

                case TPresetColor.BlanchedAlmond:
                    return Colors.BlanchedAlmond;

                case TPresetColor.Blue:
                    return Colors.Blue;

                case TPresetColor.BlueViolet:
                    return Colors.BlueViolet;

                case TPresetColor.Brown:
                    return Colors.Brown;

                case TPresetColor.BurlyWood:
                    return Colors.BurlyWood;

                case TPresetColor.CadetBlue:
                    return Colors.CadetBlue;

                case TPresetColor.Chartreuse:
                    return Colors.Chartreuse;

                case TPresetColor.Chocolate:
                    return Colors.Chocolate;

                case TPresetColor.Coral:
                    return Colors.Coral;

                case TPresetColor.CornflowerBlue:
                    return Colors.CornflowerBlue;

                case TPresetColor.Cornsilk:
                    return Colors.Cornsilk;

                case TPresetColor.Crimson:
                    return Colors.Crimson;

                case TPresetColor.Cyan:
                    return Colors.Cyan;

                case TPresetColor.DeepPink:
                    return Colors.DeepPink;

                case TPresetColor.DeepSkyBlue:
                    return Colors.DeepSkyBlue;

                case TPresetColor.DimGray:
                    return Colors.DimGray;

                case TPresetColor.DkBlue:
                    return Colors.DarkBlue;

                case TPresetColor.DkCyan:
                    return Colors.DarkCyan;

                case TPresetColor.DkGoldenrod:
                    return Colors.DarkGoldenrod;

                case TPresetColor.DkGray:
                    return Colors.DarkGray;

                case TPresetColor.DkGreen:
                    return Colors.DarkGreen;

                case TPresetColor.DkKhaki:
                    return Colors.DarkKhaki;

                case TPresetColor.DkMagenta:
                    return Colors.DarkMagenta;

                case TPresetColor.DkOliveGreen:
                    return Colors.DarkOliveGreen;

                case TPresetColor.DkOrange:
                    return Colors.DarkOrange;

                case TPresetColor.DkOrchid:
                    return Colors.DarkOrchid;

                case TPresetColor.DkRed:
                    return Colors.DarkRed;

                case TPresetColor.DkSalmon:
                    return Colors.DarkSalmon;

                case TPresetColor.DkSeaGreen:
                    return Colors.DarkSeaGreen;

                case TPresetColor.DkSlateBlue:
                    return Colors.DarkSlateBlue;

                case TPresetColor.DkSlateGray:
                    return Colors.DarkSlateGray;

                case TPresetColor.DkTurquoise:
                    return Colors.DarkTurquoise;

                case TPresetColor.DkViolet:
                    return Colors.DarkViolet;

                case TPresetColor.DodgerBlue:
                    return Colors.DodgerBlue;

                case TPresetColor.Firebrick:
                    return Colors.Firebrick;

                case TPresetColor.FloralWhite:
                    return Colors.FloralWhite;

                case TPresetColor.ForestGreen:
                    return Colors.ForestGreen;

                case TPresetColor.Fuchsia:
                    return Colors.Fuchsia;

                case TPresetColor.Gainsboro:
                    return Colors.Gainsboro;

                case TPresetColor.GhostWhite:
                    return Colors.GhostWhite;

                case TPresetColor.Gold:
                    return Colors.Gold;

                case TPresetColor.Goldenrod:
                    return Colors.Goldenrod;

                case TPresetColor.Gray:
                    return Colors.Gray;

                case TPresetColor.Green:
                    return Colors.Green;

                case TPresetColor.GreenYellow:
                    return Colors.GreenYellow;

                case TPresetColor.Honeydew:
                    return Colors.Honeydew;

                case TPresetColor.HotPink:
                    return Colors.HotPink;

                case TPresetColor.IndianRed:
                    return Colors.IndianRed;

                case TPresetColor.Indigo:
                    return Colors.Indigo;

                case TPresetColor.Ivory:
                    return Colors.Ivory;

                case TPresetColor.Khaki:
                    return Colors.Khaki;

                case TPresetColor.Lavender:
                    return Colors.Lavender;

                case TPresetColor.LavenderBlush:
                    return Colors.LavenderBlush;

                case TPresetColor.LawnGreen:
                    return Colors.LawnGreen;

                case TPresetColor.LemonChiffon:
                    return Colors.LemonChiffon;

                case TPresetColor.Lime:
                    return Colors.Lime;

                case TPresetColor.LimeGreen:
                    return Colors.LimeGreen;

                case TPresetColor.Linen:
                    return Colors.Linen;

                case TPresetColor.LtBlue:
                    return Colors.LightBlue;

                case TPresetColor.LtCoral:
                    return Colors.LightCoral;

                case TPresetColor.LtCyan:
                    return Colors.LightCyan;

                case TPresetColor.LtGoldenrodYellow:
                    return Colors.LightGoldenrodYellow;

                case TPresetColor.LtGray:
                    return Colors.LightGray;

                case TPresetColor.LtGreen:
                    return Colors.LightGreen;

                case TPresetColor.LtPink:
                    return Colors.LightPink;

                case TPresetColor.LtSalmon:
                    return Colors.LightSalmon;

                case TPresetColor.LtSeaGreen:
                    return Colors.LightSeaGreen;

                case TPresetColor.LtSkyBlue:
                    return Colors.LightSkyBlue;

                case TPresetColor.LtSlateGray:
                    return Colors.LightSlateGray;

                case TPresetColor.LtSteelBlue:
                    return Colors.LightSteelBlue;

                case TPresetColor.LtYellow:
                    return Colors.LightYellow;

                case TPresetColor.Magenta:
                    return Colors.Magenta;

                case TPresetColor.Maroon:
                    return Colors.Maroon;

                case TPresetColor.MedAquamarine:
                    return Colors.MediumAquamarine;

                case TPresetColor.MedBlue:
                    return Colors.MediumBlue;

                case TPresetColor.MedOrchid:
                    return Colors.MediumOrchid;

                case TPresetColor.MedPurple:
                    return Colors.MediumPurple;

                case TPresetColor.MedSeaGreen:
                    return Colors.MediumSeaGreen;

                case TPresetColor.MedSlateBlue:
                    return Colors.MediumSlateBlue;

                case TPresetColor.MedSpringGreen:
                    return Colors.MediumSpringGreen;

                case TPresetColor.MedTurquoise:
                    return Colors.MediumTurquoise;

                case TPresetColor.MedVioletRed:
                    return Colors.MediumVioletRed;

                case TPresetColor.MidnightBlue:
                    return Colors.MidnightBlue;

                case TPresetColor.MintCream:
                    return Colors.MintCream;

                case TPresetColor.MistyRose:
                    return Colors.MistyRose;

                case TPresetColor.Moccasin:
                    return Colors.Moccasin;

                case TPresetColor.NavajoWhite:
                    return Colors.NavajoWhite;

                case TPresetColor.Navy:
                    return Colors.Navy;

                case TPresetColor.OldLace:
                    return Colors.OldLace;

                case TPresetColor.Olive:
                    return Colors.Olive;

                case TPresetColor.OliveDrab:
                    return Colors.OliveDrab;

                case TPresetColor.Orange:
                    return Colors.Orange;

                case TPresetColor.OrangeRed:
                    return Colors.OrangeRed;

                case TPresetColor.Orchid:
                    return Colors.Orchid;

                case TPresetColor.PaleGoldenrod:
                    return Colors.PaleGoldenrod;

                case TPresetColor.PaleGreen:
                    return Colors.PaleGreen;

                case TPresetColor.PaleTurquoise:
                    return Colors.PaleTurquoise;

                case TPresetColor.PaleVioletRed:
                    return Colors.PaleVioletRed;

                case TPresetColor.PapayaWhip:
                    return Colors.PapayaWhip;

                case TPresetColor.PeachPuff:
                    return Colors.PeachPuff;

                case TPresetColor.Peru:
                    return Colors.Peru;

                case TPresetColor.Pink:
                    return Colors.Pink;

                case TPresetColor.Plum:
                    return Colors.Plum;

                case TPresetColor.PowderBlue:
                    return Colors.PowderBlue;

                case TPresetColor.Purple:
                    return Colors.Purple;

                case TPresetColor.Red:
                    return Colors.Red;

                case TPresetColor.RosyBrown:
                    return Colors.RosyBrown;

                case TPresetColor.RoyalBlue:
                    return Colors.RoyalBlue;

                case TPresetColor.SaddleBrown:
                    return Colors.SaddleBrown;

                case TPresetColor.Salmon:
                    return Colors.Salmon;

                case TPresetColor.SandyBrown:
                    return Colors.SandyBrown;

                case TPresetColor.SeaGreen:
                    return Colors.SeaGreen;

                case TPresetColor.SeaShell:
                    return Colors.SeaShell;

                case TPresetColor.Sienna:
                    return Colors.Sienna;

                case TPresetColor.Silver:
                    return Colors.Silver;

                case TPresetColor.SkyBlue:
                    return Colors.SkyBlue;

                case TPresetColor.SlateBlue:
                    return Colors.SlateBlue;

                case TPresetColor.SlateGray:
                    return Colors.SlateGray;

                case TPresetColor.Snow:
                    return Colors.Snow;

                case TPresetColor.SpringGreen:
                    return Colors.SpringGreen;

                case TPresetColor.SteelBlue:
                    return Colors.SteelBlue;

                case TPresetColor.Tan:
                    return Colors.Tan;

                case TPresetColor.Teal:
                    return Colors.Teal;

                case TPresetColor.Thistle:
                    return Colors.Thistle;

                case TPresetColor.Tomato:
                    return Colors.Tomato;

                case TPresetColor.Turquoise:
                    return Colors.Turquoise;

                case TPresetColor.Violet:
                    return Colors.Violet;

                case TPresetColor.Wheat:
                    return Colors.Wheat;

                case TPresetColor.White:
                    return Colors.White;

                case TPresetColor.WhiteSmoke:
                    return Colors.WhiteSmoke;

                case TPresetColor.Yellow:
                    return Colors.Yellow;

                case TPresetColor.YellowGreen:
                    return Colors.YellowGreen;

            }

            return ColorUtil.Empty;
        }


#if(FRAMEWORK20)
        private static Dictionary<Color, TPresetColor> GetColorsFromPreset()
        {
            Dictionary<Color, TPresetColor> Result = new Dictionary<Color, TPresetColor>();
            foreach (TPresetColor pr in TCompactFramework.EnumGetValues(typeof(TPresetColor)))
            {
                Result[GetPresetColor(pr)] = pr;
            }
            return Result;
        }

        private static Dictionary<Color, TSystemColor> GetColorsFromSystem()
        {
            Dictionary<Color, TSystemColor> Result = new Dictionary<Color, TSystemColor>();
            foreach (TSystemColor sc in TCompactFramework.EnumGetValues(typeof(TSystemColor)))
            {
                Result[GetSystemColor(sc)] = sc;
            }
            return Result;
        }

        private static TPresetColor GetPresetColor(Color aColor)
        {
            TPresetColor Result;
            if (ColorsFromPreset.TryGetValue(aColor, out Result)) return Result;
            return TPresetColor.None;
        }

        private static TSystemColor GetSysColor(Color aColor)
        {
            TSystemColor Result;
            if (ColorsFromSystem.TryGetValue(aColor, out Result)) return Result;
            return TSystemColor.None;
        }
#else
        private static Hashtable GetColorsFromPreset()
        {
            Hashtable Result = new Hashtable();
            foreach (TPresetColor pr in TCompactFramework.EnumGetValues(typeof(TPresetColor)))
            {
                Result[GetPresetColor(pr)] = pr;
            }
            return Result;
        }

        private static Hashtable GetColorsFromSystem()
        {
            Hashtable Result = new Hashtable();
            foreach (TSystemColor sc in TCompactFramework.EnumGetValues(typeof(TSystemColor)))
            {
                Result[GetSystemColor(sc)] = sc;
            }
            return Result;
        }

        private static TPresetColor GetPresetColor(Color aColor)
        {
            object obj = ColorsFromPreset[aColor];
            if (obj == null) return TPresetColor.None;
            return (TPresetColor)obj;
        }

        private static TSystemColor GetSysColor(Color aColor)
        {
            object obj = ColorsFromSystem[aColor];
            if (obj == null) return TSystemColor.None;
            return (TSystemColor)obj;
        }
#endif

        #endregion

        /// <summary>
        /// Assigns a .NET color to this instance.
        /// </summary>
        /// <param name="aColor"></param>
        /// <returns></returns>
        public static implicit operator TDrawingColor(Color aColor)
        {
            return TDrawingColor.FromColor(aColor);
        }
        #endregion

        internal TDrawingColor ReplacePhClr(TDrawingColor basicColor)
        {
            if (ColorType == TDrawingColorType.Theme && Theme == TThemeColor.None)
            {
                return basicColor;
            }

            return this;
        }

        /// <summary>
        /// Returns the transparent color.
        /// </summary>
        public static TDrawingColor Transparent
        {
            get
            {
                TDrawingColor transp = TDrawingColor.FromRgb(0, 0, 0);
                transp = TDrawingColor.AddTransform(transp, new TColorTransform[] { new TColorTransform(TColorTransformType.Alpha, 0) });
                return transp;
            }
        }

        /// <summary>
        /// Returns R, G and B components of the color. If this is a theme or indexed color, it will
        /// be converted to RGB before getting the components.
        /// </summary>
        /// <param name="xls">ExcelFile that will be used to know the palette and themes of the file for indexed/themed colors.</param>
        /// <param name="R">Returns the red component of the color.</param>
        /// <param name="G">Returns the green component of the color.</param>
        /// <param name="B">Returns the blue component of the color.</param>
        public void GetComponents(IFlexCelPalette xls, out byte R, out byte G, out byte B)
        {
            Color c = ToColor(xls);
#if (COMPACTFRAMEWORK || (!FRAMEWORK30 && !MONOTOUCH))
            R = c.R;
            G = c.G;
            B = c.B;
#else
            R = c.R();
            G = c.G();
            B = c.B();
#endif
        }
    }

    /// <summary>
    /// A color scheme for a theme.
    /// </summary>
    public class TThemeColorScheme
    {
        private static readonly TDrawingColor[] StandardColors =
        {
            TDrawingColor.FromSystem(TSystemColor.Window),
            TDrawingColor.FromSystem(TSystemColor.WindowText),
            ColorUtil.FromArgb(0xEE, 0xEC, 0xE1),
            ColorUtil.FromArgb(0x1F, 0x49, 0x7D),
            ColorUtil.FromArgb(0x4F, 0x81, 0xBD),
            ColorUtil.FromArgb(0xC0, 0x50, 0x4D),
            ColorUtil.FromArgb(0x9B, 0xBB, 0x59),
            ColorUtil.FromArgb(0x80, 0x64, 0xA2),
            ColorUtil.FromArgb(0x4B, 0xAC, 0xC6),
            ColorUtil.FromArgb(0xF7, 0x96, 0x46),
            ColorUtil.FromArgb(0x00, 0x00, 0xFF),
            ColorUtil.FromArgb(0x80, 0x00, 0x80)
        };

        internal TDrawingColor[] Colors;
        private string FName;

        /// <summary>
        /// Creates a new ColorScheme with standard properties.
        /// </summary>
        public TThemeColorScheme()
        {
            Name = "Office";
            Colors = (TDrawingColor[])StandardColors.Clone();
        }

        /// <summary>
        /// Name of the color definition. This will be shown in Excel UI.
        /// </summary>
        public string Name { get { return FName; } set { FName = value; } }

        /// <summary>
        /// Resets the color scheme to be the Excel 2007 standard.
        /// </summary>
        public void Reset()
        {
            Name = "Office";
            Colors = (TDrawingColor[])StandardColors.Clone();
        }

        /// <summary>
        /// Returns a color definition for a themed color.
        /// </summary>
        /// <param name="themeColor"></param>
        /// <returns></returns>
        public TDrawingColor this[TThemeColor themeColor]
        {
            get
            {
                if (!Enum.IsDefined(typeof(TThemeColor), themeColor) || themeColor == TThemeColor.None
                    || (int)themeColor < 0 || (int)themeColor > Colors.Length)
                    FlxMessages.ThrowException(FlxErr.ErrInvalidColorEnum, "Theme");
                return Colors[(int)themeColor];
            }
            set
            {
                if (!Enum.IsDefined(typeof(TThemeColor), themeColor) || themeColor == TThemeColor.None)
                    FlxMessages.ThrowException(FlxErr.ErrInvalidColorEnum, "Theme");

                if (value.ColorType == TDrawingColorType.Theme) FlxMessages.ThrowException(FlxErr.ErrCantUseThemeColorsInsideATheme);
                Colors[(int)themeColor] = value;
            }
        }

        /// <summary>
        /// True if this is the standard Excel 2007 color palette.
        /// </summary>
        internal bool IsStandard
        {
            get
            {
                if (Name != "Office") return false;
                if (Colors.Length != StandardColors.Length) return false;

                for (int i = 0; i < StandardColors.Length; i++)
                {
                    if (!StandardColors[i].Equals(Colors[i])) return false;
                }

                return true;
            }
        }

        /// <summary>
        /// Returns a deep copy of this object.
        /// </summary>
        /// <returns></returns>        
        internal TThemeColorScheme Clone()
        {
            TThemeColorScheme Result = new TThemeColorScheme();
            Result.Name = Name;
            Result.Colors = (TDrawingColor[])Colors.Clone();

            return Result;
        }
    }
    #endregion

    #region Color Transforms
    /// <summary>
    /// List of transformations that can be done to a color.
    /// </summary>
    public enum TColorTransformType
    {
        /// <summary>
        /// Yields a lighter version of its input color. A 10% tint is 10% of the input color combined with 90% white. 
        /// </summary>
        Tint,

        /// <summary>
        /// Yields a darker version of its input color. A 10% shade is 10% of the input color combined with 90% black. 
        /// </summary>
        Shade,

        /// <summary>
        /// Yields the complement of its input color. For example, the complement of red is green. 
        /// </summary>
        Complement,

        /// <summary>
        /// Yields the inverse of its input color. For example, the inverse of red (1,0,0) is cyan (0,1,1). 
        /// </summary>
        Inverse,

        /// <summary>
        /// Yields a grayscale of its input color, taking into relative intensities of the red, green, and blue primaries. 
        /// </summary>
        Gray,

        /// <summary>
        /// Yields its input color with the specified opacity, but with its color unchanged. 
        /// </summary>
        Alpha,

        /// <summary>
        /// Yields a more or less opaque version of its input color. An alpha offset never increases the alpha beyond 100% or decreases below 0%; i.e., the result of the transform pins the alpha to the range of [0%,100%]. A 10% alpha offset increases a 50% opacity to 60%. A -10% alpha offset decreases a 50% opacity to 40%. 
        /// </summary>
        AlphaOff,

        /// <summary>
        /// Yields a more or less opaque version of its input color. An alpha modulate never increases the alpha beyond 100%. A 200% alpha modulate makes a input color twice as opaque as before. A 50% alpha modulate makes a input color half as opaque as before. 
        /// </summary>
        AlphaMod,

        /// <summary>
        /// Yields the input color with the specified hue, but with its saturation and luminance unchanged. 
        /// </summary>
        Hue,

        /// <summary>
        /// Yields the input color with its hue shifted, but with its saturation and luminance unchanged. 
        /// </summary>
        HueOff,

        /// <summary>
        /// Yields the input color with its hue modulated by the given percentage. 
        /// </summary>
        HueMod,

        /// <summary>
        /// Yields the input color with the specified saturation, but with its hue and luminance unchanged. Typically saturation values fall in the range [0%, 100%]. 
        /// </summary>
        Sat,

        /// <summary>
        /// Yields the input color with its saturation shifted, but with its hue and luminance unchanged. 
        /// </summary>
        SatOff,

        /// <summary>
        /// Yields the input color with its saturation modulated by the given percentage. A 50% saturation modulate reduces the saturation by half. A 200% saturation modulate doubles the saturation. 
        /// </summary>
        SatMod,

        /// <summary>
        /// Yields the input color with the specified luminance, but with its hue and saturation unchanged. Typically, luminance values fall in the range [0%,100%]. 
        /// </summary>
        Lum,

        /// <summary>
        /// Yields the input color with its luminance shifted, but with its hue and saturation unchanged. 
        /// </summary>
        LumOff,

        /// <summary>
        /// Yields the input color with its luminance modulated by the given percentage. A 50% luminance modulate reduces the luminance by half. A 200% luminance modulate doubles the luminance. 
        /// </summary>
        LumMod,

        /// <summary>
        /// Yields the input color with the specified red component, but with its green and blue components unchanged. 
        /// </summary>
        Red,

        /// <summary>
        /// Yields the input color with its red component shifted, but with its green and blue components unchanged.
        /// </summary>
        RedOff,

        /// <summary>
        /// Yields the input color with its red component modulated by the given percentage. A 50% red modulate reduces the red component by half. A 200% red modulate doubles the red component. 
        /// </summary>
        RedMod,

        /// <summary>
        /// Yields the input color with the specified green component, but with its red and blue components unchanged. 
        /// </summary>
        Green,

        /// <summary>
        /// Yields the input color with its green component shifted, but with its red and blue components unchanged. 
        /// </summary>
        GreenOff,

        /// <summary>
        /// Yields the input color with its green component modulated by the given percentage. A 50% green modulate reduces the green component by half. A 200% green modulate doubles the green component. 
        /// </summary>
        GreenMod,

        /// <summary>
        /// Yields the input color with the specified blue component, but with its red and green components unchanged. 
        /// </summary>
        Blue,

        /// <summary>
        /// Yields the input color with its blue component shifted, but with its red and green components unchanged. 
        /// </summary>
        BlueOff,

        /// <summary>
        /// Yields the input color with its blue component modulated by the given percentage. A 50% blue modulate reduces the blue component by half. A 200% blue modulate doubles the blue component. 
        /// </summary>
        BlueMod,

        /// <summary>
        /// Yields the sRGB gamma shift of its input color. 
        /// </summary>
        Gamma,

        /// <summary>
        /// Yields the inverse sRGB gamma shift of its input color. 
        /// </summary>
        InvGamma

    }

    /// <summary>
    /// Specifies a color transformation to be applied to a color.
    /// </summary>
    public struct TColorTransform
    {
        #region Variables
        private readonly TColorTransformType FColorTransformType;
        private readonly double FValue;
        #endregion

        #region Constructor
        /// <summary>
        /// Creates a new TColorTransform with the corresponding parameters.
        /// </summary>
        /// <param name="aColorTransformType">Type of transformation to be applied.</param>
        /// <param name="aValue">Value of the transform. The meaning of this field depends in the <see cref="ColorTransformType"/> value.</param>
        public TColorTransform(TColorTransformType aColorTransformType, double aValue)
        {
            FColorTransformType = aColorTransformType;
            FValue = aValue;
        }
        #endregion

        #region Properties
        /// <summary>
        /// Type of transformation to be applied.
        /// </summary>
        public TColorTransformType ColorTransformType { get { return FColorTransformType; } }

        /// <summary>
        /// Value of the transform. The meaning of this field depends in the <see cref="ColorTransformType"/> value.
        /// </summary>
        public double Value { get { return FValue; } }
        #endregion

        #region Compare
        /// <summary>
        /// Returns true if both color transforms are the same.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            if (!(obj is TColorTransform)) return false;
            TColorTransform o2 = (TColorTransform)obj;
            return ColorTransformType == o2.ColorTransformType && Value == o2.Value;
        }

        /// <summary>
        /// Hashcode for the color transform.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHash(FColorTransformType.GetHashCode(), FValue.GetHashCode());
        }

        /// <summary>
        /// Returns true if both color transforms have the same value.
        /// </summary>
        /// <param name="o1">First color transform to compare.</param>
        /// <param name="o2">Second color transform to compare.</param>
        /// <returns></returns>
        public static bool operator ==(TColorTransform o1, TColorTransform o2)
        {
            return o1.Equals(o2);
        }

        /// <summary>
        /// Returns true if both color transforms do not have the same value.
        /// </summary>
        /// <param name="o1">First color transform to compare.</param>
        /// <param name="o2">Second color transform to compare.</param>
        /// <returns></returns>
        public static bool operator !=(TColorTransform o1, TColorTransform o2)
        {
            return !(o1.Equals(o2));
        }
        #endregion

        #region Transform
        /// <summary>
        /// Applies the transform for a given color.
        /// </summary>
        /// <param name="c">Color to transform.</param>
        /// <returns>Transformed color.</returns>
        public Color Transform(Color c)
        {
#if (!COMPACTFRAMEWORK)
            TScRGBColor sc = c;
            THSLColor hsl;

#if (COMPACTFRAMEWORK || (!FRAMEWORK30 && !MONOTOUCH))
            byte A = c.A;
#else
            byte A = c.A();
#endif

            switch (ColorTransformType)
            {
                case TColorTransformType.Tint:
                    return ColorUtil.FromArgb(A, new TScRGBColor(sc.ScR * Value + 1 - Value, sc.ScG * Value + 1 - Value, sc.ScB * Value + 1 - Value));

                case TColorTransformType.Shade:
                    return ColorUtil.FromArgb(A, new TScRGBColor(sc.ScR * Value, sc.ScG * Value, sc.ScB * Value));

                case TColorTransformType.Complement:
                    hsl = c;
                    return ColorUtil.FromArgb(A, new THSLColor(hsl.Hue + 180, hsl.Sat, hsl.Lum));

                case TColorTransformType.Inverse:
                    return ColorUtil.FromArgb(A, new TScRGBColor(1 - sc.ScR, 1 - sc.ScG, 1 - sc.ScB));

                case TColorTransformType.Gray:
                    hsl = c;
                    return ColorUtil.FromArgb(A, new THSLColor(hsl.Hue, 0, hsl.Lum)); //not the exact same algorithm used by Excel

                case TColorTransformType.Alpha:
                    return ColorUtil.FromArgb((int)Math.Round(Value * 255), c);

                case TColorTransformType.AlphaOff:
                    return ColorUtil.FromArgb(Math.Min(255, (int)Math.Round((A / 255.0 + Value) * 255)), c);

                case TColorTransformType.AlphaMod:
                    return ColorUtil.FromArgb(Math.Min(255, (int)Math.Round(A * Value)), c);

                case TColorTransformType.Hue:
                    hsl = c;
                    return ColorUtil.FromArgb(A, new THSLColor(Value, hsl.Sat, hsl.Lum));

                case TColorTransformType.HueOff:
                    hsl = c;
                    double NewHue = hsl.Hue + Value;
                    if (NewHue < 0) NewHue = 0;
                    if (NewHue >= 360) NewHue = 360;
                    return ColorUtil.FromArgb(A, new THSLColor(NewHue, hsl.Sat, hsl.Lum));

                case TColorTransformType.HueMod:
                    hsl = c;
                    return ColorUtil.FromArgb(A, new THSLColor(hsl.Hue * Value, hsl.Sat, hsl.Lum));

                case TColorTransformType.Sat:
                    hsl = c;
                    return ColorUtil.FromArgb(A, new THSLColor(hsl.Hue, Value, hsl.Lum));

                case TColorTransformType.SatOff:
                    hsl = c;
                    return ColorUtil.FromArgb(A, new THSLColor(hsl.Hue, hsl.Sat + Value, hsl.Lum));

                case TColorTransformType.SatMod:
                    hsl = c;
                    return ColorUtil.FromArgb(A, new THSLColor(hsl.Hue, hsl.Sat * Value, hsl.Lum));

                case TColorTransformType.Lum:
                    hsl = c;
                    return ColorUtil.FromArgb(A, new THSLColor(hsl.Hue, hsl.Sat, Value));

                case TColorTransformType.LumOff:
                    hsl = c;
                    return ColorUtil.FromArgb(A, new THSLColor(hsl.Hue, hsl.Sat, hsl.Lum + Value));

                case TColorTransformType.LumMod:
                    hsl = c;
                    return ColorUtil.FromArgb(A, new THSLColor(hsl.Hue, hsl.Sat, hsl.Lum * Value));

                case TColorTransformType.Red:
                    return ColorUtil.FromArgb(A, new TScRGBColor(Value, sc.ScG, sc.ScB));

                case TColorTransformType.RedOff:
                    return ColorUtil.FromArgb(A, new TScRGBColor(sc.ScR + Value, sc.ScG, sc.ScB));

                case TColorTransformType.RedMod:
                    return ColorUtil.FromArgb(A, new TScRGBColor(sc.ScR * Value, sc.ScG, sc.ScB));

                case TColorTransformType.Green:
                    return ColorUtil.FromArgb(A, new TScRGBColor(sc.ScR, Value, sc.ScB));

                case TColorTransformType.GreenOff:
                    return ColorUtil.FromArgb(A, new TScRGBColor(sc.ScR, sc.ScG + Value, sc.ScB));

                case TColorTransformType.GreenMod:
                    return ColorUtil.FromArgb(A, new TScRGBColor(sc.ScR, sc.ScG * Value, sc.ScB));

                case TColorTransformType.Blue:
                    return ColorUtil.FromArgb(A, new TScRGBColor(sc.ScR, sc.ScG, Value));

                case TColorTransformType.BlueOff:
                    return ColorUtil.FromArgb(A, new TScRGBColor(sc.ScR, sc.ScG, sc.ScB + Value));

                case TColorTransformType.BlueMod:
                    return ColorUtil.FromArgb(A, new TScRGBColor(sc.ScR, sc.ScG, sc.ScB * Value));

                case TColorTransformType.Gamma:
                    return ColorUtil.FromArgb(A,
#if (COMPACTFRAMEWORK || (!FRAMEWORK30 && !MONOTOUCH))
                        TScRGBColor.To255(TScRGBColor.SRGBtoRGB(c.R / 255f)),
                        TScRGBColor.To255(TScRGBColor.SRGBtoRGB(c.G / 255f)),
                        TScRGBColor.To255(TScRGBColor.SRGBtoRGB(c.B / 255f)));
#else
                        TScRGBColor.To255(TScRGBColor.SRGBtoRGB(c.Rd())),
                        TScRGBColor.To255(TScRGBColor.SRGBtoRGB(c.Gd())),
                        TScRGBColor.To255(TScRGBColor.SRGBtoRGB(c.Bd())));
#endif

                case TColorTransformType.InvGamma:
                    return ColorUtil.FromArgb(A, TScRGBColor.To255(sc.ScR), TScRGBColor.To255(sc.ScG), TScRGBColor.To255(sc.ScB));

            }
#endif
            return c;
        }
        #endregion

    }
    #endregion
}
