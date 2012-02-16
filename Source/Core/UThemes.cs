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

    #region Fonts

    #region Enums
    /// <summary>
    /// Character set for a font as defined in the DrawingML standard.
    /// </summary>
    public enum TFontCharSet
    {
        /// <summary>
        /// Specifies the ANSI character set. (IANA name iso-8859-1) 
        /// </summary>
        Ansi = 0x00,

        /// <summary>
        /// Specifies the default character set. 
        /// </summary>
        Default = 0x01,

        /// <summary>
        /// Specifies the Symbol character set. This value specifies that the characters in the Unicode private use area (U+FF00 to U+FFFF) of the font should be used to display characters in the range U+0000 to U+00FF. 
        /// </summary>
        Symbol = 0x02,

        /// <summary>
        /// Specifies a Macintosh (Standard Roman) character set. (IANA name macintosh) 
        /// </summary>
        Macintosh = 0x4D,

        /// <summary>
        /// Specifies the JIS character set. (IANA name shift_jis) 
        /// </summary>
        Jis = 0x80,

        /// <summary>
        /// Specifies the Hangul character set. (IANA name ks_c_5601-1987) 
        /// </summary>
        Hangul = 0x81,

        /// <summary>
        /// Specifies a Johab character set. (IANA name KS C-5601-1992) 
        /// </summary>
        Johab = 0x82,

        /// <summary>
        /// Specifies the GB-2312 character set. (IANA name GBK) 
        /// </summary>
        GB2312 = 0x86,

        /// <summary>
        /// Specifies the Chinese Big Five character set. (IANA name Big5) ECMA-376 Part 1 
        /// </summary>
        Big5 = 0x88,

        /// <summary>
        /// Specifies a Greek character set. (IANA name windows-1253) 
        /// </summary>
        Greek = 0xA1,

        /// <summary>
        /// Specifies a Turkish character set. (IANA name iso-8859-9) 
        /// </summary>
        Turkish = 0xA2,

        /// <summary>
        /// Specifies a Vietnamese character set. (IANA name windows-1258) 
        /// </summary>
        Vietnamese = 0xA3,

        /// <summary>
        /// Specifies a Hebrew character set. (IANA name windows-1255) 
        /// </summary>
        Hebrew = 0xB1,

        /// <summary>
        /// Specifies an Arabic character set. (IANA name windows-1256) 
        /// </summary>
        Arabic = 0xB2,

        /// <summary>
        /// Specifies a Baltic character set. (IANA name windows-1257) 
        /// </summary>
        Baltic = 0xBA,

        /// <summary>
        /// Specifies a Russian character set. (IANA name windows-1251) 
        /// </summary>
        Russian = 0xCC,

        /// <summary>
        /// Specifies a Thai character set. (IANA name windows-874) 
        /// </summary>
        Thai = 0xDE,

        /// <summary>
        /// Specifies an Eastern European character set. (IANA name windows-1250) 
        /// </summary>
        EasternEuropean = 0xEE,

        /// <summary>
        /// Specifies an OEM character set not defined by ECMA-376. 
        /// </summary>
        OEM = 0xFF
    }

    /// <summary>
    /// Pitch family for a font.
    /// </summary>
    public enum TPitchFamily
    {
        /// <summary>
        /// DEFAULT PITCH + UNKNOWN FONT FAMILY 
        /// </summary>
        DEFAULT_PITCH__UNKNOWN_FONT_FAMILY = 0x00,

        /// <summary>
        /// FIXED PITCH + UNKNOWN FONT FAMILY 
        /// </summary>
        FIXED_PITCH__UNKNOWN_FONT_FAMILY = 0x01,

        /// <summary>
        /// VARIABLE PITCH + UNKNOWN FONT FAMILY 
        /// </summary>
        VARIABLE_PITCH__UNKNOWN_FONT_FAMILY = 0x02,

        /// <summary>
        /// DEFAULT PITCH + ROMAN FONT FAMILY 
        /// </summary>
        DEFAULT_PITCH__ROMAN_FONT_FAMILY = 0x10,

        /// <summary>
        /// FIXED PITCH + ROMAN FONT FAMILY 
        /// </summary>
        FIXED_PITCH__ROMAN_FONT_FAMILY = 0x11,

        /// <summary>
        /// VARIABLE PITCH + ROMAN FONT FAMILY 
        /// </summary>
        VARIABLE_PITCH__ROMAN_FONT_FAMILY = 0x12,

        /// <summary>
        /// DEFAULT PITCH + SWISS FONT FAMILY 
        /// </summary>
        DEFAULT_PITCH__SWISS_FONT_FAMILY = 0x20,

        /// <summary>
        /// FIXED PITCH + SWISS FONT FAMILY 
        /// </summary>
        FIXED_PITCH__SWISS_FONT_FAMILY = 0x21,

        /// <summary>
        /// VARIABLE PITCH + SWISS FONT FAMILY 
        /// </summary>
        VARIABLE_PITCH__SWISS_FONT_FAMILY = 0x22,

        /// <summary>
        /// DEFAULT PITCH + MODERN FONT FAMILY 
        /// </summary>
        DEFAULT_PITCH__MODERN_FONT_FAMILY = 0x30,

        /// <summary>
        /// FIXED PITCH + MODERN FONT FAMILY 
        /// </summary>
        FIXED_PITCH__MODERN_FONT_FAMILY = 0x31,

        /// <summary>
        /// VARIABLE PITCH + MODERN FONT FAMILY 
        /// </summary>
        VARIABLE_PITCH__MODERN_FONT_FAMILY = 0x32,

        /// <summary>
        /// DEFAULT PITCH + SCRIPT FONT FAMILY 
        /// </summary>
        DEFAULT_PITCH__SCRIPT_FONT_FAMILY = 0x40,

        /// <summary>
        /// FIXED PITCH + SCRIPT FONT FAMILY 
        /// </summary>
        FIXED_PITCH__SCRIPT_FONT_FAMILY = 0x41,

        /// <summary>
        /// VARIABLE PITCH + SCRIPT FONT FAMILY 
        /// </summary>
        VARIABLE_PITCH__SCRIPT_FONT_FAMILY = 0x42,

        /// <summary>
        /// DEFAULT PITCH + DECORATIVE FONT FAMILY 
        /// </summary>
        DEFAULT_PITCH__DECORATIVE_FONT_FAMILY = 0x50,

        /// <summary>
        /// FIXED PITCH + DECORATIVE FONT FAMILY 
        /// </summary>
        FIXED_PITCH__DECORATIVE_FONT_FAMILY = 0x51,

        /// <summary>
        /// VARIABLE PITCH + DECORATIVE FONT FAMILY 
        /// </summary>
        VARIABLE_PITCH__DECORATIVE_FONT_FAMILY = 0x52
    }
    #endregion

    internal struct TFontScriptDef
    {
        internal string Script;
        internal string TypeFace;

        internal TFontScriptDef(string aScript, string aTypeFace)
        {
            Script = aScript;
            TypeFace = aTypeFace;
        }
    }

    /// <summary>
    /// Represents the fonts for a theme.
    /// </summary>
    public class TThemeFontScheme
    {
        private TThemeFont FMajorFont;
        private TThemeFont FMinorFont;
        private string FName;

        private static readonly TFontScriptDef[] MajorScripts = CreateMajorScripts();
        private static readonly TFontScriptDef[] MinorScripts = CreateMinorScripts();

        private static readonly StringStringHashtable MajorDef = CreateMajorDef();
        private static readonly StringStringHashtable MinorDef = CreateMinorDef();

        private static StringStringHashtable CreateMinorDef()
        {
            StringStringHashtable MinorDef = new StringStringHashtable();
            foreach (TFontScriptDef fdef in MinorScripts)
            {
                MinorDef[fdef.Script] = fdef.TypeFace;
            }

            return MinorDef;
        }

        private static TFontScriptDef[] CreateMinorScripts()
        {
            return new TFontScriptDef[]
            {
            new TFontScriptDef("Jpan", "ＭＳ Ｐゴシック"),
            new TFontScriptDef("Hang", "맑은 고딕"),
            new TFontScriptDef("Hans", "宋体"),
            new TFontScriptDef("Hant", "新細明體"),
            new TFontScriptDef("Arab", "Arial"),
            new TFontScriptDef("Hebr", "Arial"),
            new TFontScriptDef("Thai", "Tahoma"),
            new TFontScriptDef("Ethi", "Nyala"),
            new TFontScriptDef("Beng", "Vrinda"),
            new TFontScriptDef("Gujr", "Shruti"),
            new TFontScriptDef("Khmr", "DaunPenh"),
            new TFontScriptDef("Knda", "Tunga"),
            new TFontScriptDef("Guru", "Raavi"),
            new TFontScriptDef("Cans", "Euphemia"),
            new TFontScriptDef("Cher", "Plantagenet Cherokee"),
            new TFontScriptDef("Yiii", "Microsoft Yi Baiti"),
            new TFontScriptDef("Tibt", "Microsoft Himalaya"),
            new TFontScriptDef("Thaa", "MV Boli"),
            new TFontScriptDef("Deva", "Mangal"),
            new TFontScriptDef("Telu", "Gautami"),
            new TFontScriptDef("Taml", "Latha"),
            new TFontScriptDef("Syrc", "Estrangelo Edessa"),
            new TFontScriptDef("Orya", "Kalinga"),
            new TFontScriptDef("Mlym", "Kartika"),
            new TFontScriptDef("Laoo", "DokChampa"),
            new TFontScriptDef("Sinh", "Iskoola Pota"),
            new TFontScriptDef("Mong", "Mongolian Baiti"),
            new TFontScriptDef("Viet", "Arial"),
            new TFontScriptDef("Uigh", "Microsoft Uighur"),
            new TFontScriptDef("Geor", "Sylfaen")
            };
        }

        private static StringStringHashtable CreateMajorDef()
        {
            StringStringHashtable MajorDef = new StringStringHashtable();
            foreach (TFontScriptDef fdef in MajorScripts)
            {
                MajorDef[fdef.Script] = fdef.TypeFace;
            }

            return MajorDef;
        }

        private static TFontScriptDef[] CreateMajorScripts()
        {
            return new TFontScriptDef[]
            {
            new TFontScriptDef("Jpan", "ＭＳ Ｐゴシック"),
            new TFontScriptDef("Hang", "맑은 고딕"),
            new TFontScriptDef("Hans", "宋体"),
            new TFontScriptDef("Hant", "新細明體"),
            new TFontScriptDef("Arab", "Times New Roman"),
            new TFontScriptDef("Hebr", "Times New Roman"),
            new TFontScriptDef("Thai", "Tahoma"),
            new TFontScriptDef("Ethi", "Nyala"),
            new TFontScriptDef("Beng", "Vrinda"),
            new TFontScriptDef("Gujr", "Shruti"),
            new TFontScriptDef("Khmr", "MoolBoran"),
            new TFontScriptDef("Knda", "Tunga"),
            new TFontScriptDef("Guru", "Raavi"),
            new TFontScriptDef("Cans", "Euphemia"),
            new TFontScriptDef("Cher", "Plantagenet Cherokee"),
            new TFontScriptDef("Yiii", "Microsoft Yi Baiti"),
            new TFontScriptDef("Tibt", "Microsoft Himalaya"),
            new TFontScriptDef("Thaa", "MV Boli"),
            new TFontScriptDef("Deva", "Mangal"),
            new TFontScriptDef("Telu", "Gautami"),
            new TFontScriptDef("Taml", "Latha"),
            new TFontScriptDef("Syrc", "Estrangelo Edessa"),
            new TFontScriptDef("Orya", "Kalinga"),
            new TFontScriptDef("Mlym", "Kartika"),
            new TFontScriptDef("Laoo", "DokChampa"),
            new TFontScriptDef("Sinh", "Iskoola Pota"),
            new TFontScriptDef("Mong", "Mongolian Baiti"),
            new TFontScriptDef("Viet", "Times New Roman"),
            new TFontScriptDef("Uigh", "Microsoft Uighur"),
            new TFontScriptDef("Geor", "Sylfaen")
            };

        }

        /// <summary>
        /// Creates a new font scheme with standard properties.
        /// </summary>
        public TThemeFontScheme()
        {
            FName = "Office";
            MajorFont = null; //don't remove, it will initialize the font.
            MinorFont = null;
        }

        private TThemeFont CreateDefaultMajorFont()
        {
            TThemeFont Result = new TThemeFont();
            Result.Latin = new TThemeTextFont("Cambria", null, TPitchFamily.DEFAULT_PITCH__UNKNOWN_FONT_FAMILY, TFontCharSet.Default);
            Result.EastAsian = new TThemeTextFont(string.Empty, null, TPitchFamily.DEFAULT_PITCH__UNKNOWN_FONT_FAMILY, TFontCharSet.Default);
            Result.ComplexScript = new TThemeTextFont(string.Empty, null, TPitchFamily.DEFAULT_PITCH__UNKNOWN_FONT_FAMILY, TFontCharSet.Default);
            foreach (TFontScriptDef fdef in MajorScripts)
            {
                Result.AddFont(fdef.Script, fdef.TypeFace);
            }

            return Result;
        }

        private TThemeFont CreateDefaultMinorFont()
        {
            TThemeFont Result = new TThemeFont();
            Result.Latin = new TThemeTextFont("Calibri", null, TPitchFamily.DEFAULT_PITCH__UNKNOWN_FONT_FAMILY, TFontCharSet.Default);
            Result.EastAsian = new TThemeTextFont(string.Empty, null, TPitchFamily.DEFAULT_PITCH__UNKNOWN_FONT_FAMILY, TFontCharSet.Default);
            Result.ComplexScript = new TThemeTextFont(string.Empty, null, TPitchFamily.DEFAULT_PITCH__UNKNOWN_FONT_FAMILY, TFontCharSet.Default);
            foreach (TFontScriptDef fdef in MinorScripts)
            {
                Result.AddFont(fdef.Script, fdef.TypeFace);
            }

            return Result;
        }


        /// <summary>
        /// Name of the font definition. This will be shown in Excel UI.
        /// </summary>
        public string Name { get { return FName; } set { FName = value; } }

        /// <summary>
        /// This element defines the set of major fonts which are to be used under different languages or locals. 
        /// </summary>
        public TThemeFont MajorFont { get { return FMajorFont; } set { if (value == null) FMajorFont = CreateDefaultMajorFont(); else FMajorFont = value; } }

        /// <summary>
        /// This element defines the set of minor fonts which are to be used under different languages or locals. 
        /// </summary>
        public TThemeFont MinorFont { get { return FMinorFont; } set { if (value == null) FMinorFont = CreateDefaultMinorFont(); else FMinorFont = value; } }

        /// <summary>
        /// Returns a deep copy of this object.
        /// </summary>
        /// <returns></returns>
        internal TThemeFontScheme Clone()
        {
            TThemeFontScheme Result = new TThemeFontScheme();
            Result.Name = Name;
            Result.FMajorFont = FMajorFont == null ? null : FMajorFont.Clone();
            Result.FMinorFont = FMinorFont == null ? null : FMinorFont.Clone();

            return Result;
        }

        /// <summary>
        /// Returns true is this is a standard theme.
        /// </summary>
        internal bool IsStandard
        {
            get
            {
                return Name == "Office" &&
                    (MajorFont == null || IsStandardFont(MajorFont, "Cambria")) &&
                    (MinorFont == null || IsStandardFont(MinorFont, "Calibri"));
            }
        }

        private bool IsStandardFont(TThemeFont ThemeFont, string DefTypeFace)
        {
            return ThemeFont.Latin.IsStandard(DefTypeFace)
                && ThemeFont.EastAsian.IsStandard(null)
                && ThemeFont.ComplexScript.IsStandard(null)
                && IsScriptStandard(MajorFont, MajorDef)
                && IsScriptStandard(MinorFont, MinorDef);
        }

        private static bool IsScriptStandard(TThemeFont MFont, StringStringHashtable MDef)
        {
            foreach (string script in MFont.GetFontScripts())
            {
                string mdefscript;
                if (!MDef.TryGetValue(script, out mdefscript)) return false;
                if (mdefscript != MFont.GetFont(script)) return false;
            }

            return true;
        }

    }

    /// <summary>
    /// Represents either a major or a minor font for the theme.
    /// </summary>
    public class TThemeFont
    {
        private TThemeTextFont FLatin;
        private TThemeTextFont FEastAsian;
        private TThemeTextFont FComplexScript;
        private StringStringHashtable FFontCollection;

        /// <summary>
        /// Creates a new TThemeFont instance.
        /// </summary>
        public TThemeFont()
        {
            FFontCollection = new StringStringHashtable();
        }

        /// <summary>
        /// Creates a new TThemeFont.
        /// </summary>
        /// <param name="aLatin">See <see cref="Latin" /></param>
        /// <param name="aEastAsian">See <see cref="EastAsian" /></param>
        /// <param name="aComplexScript">See <see cref="ComplexScript" /></param>
        public TThemeFont(TThemeTextFont aLatin, TThemeTextFont aEastAsian, TThemeTextFont aComplexScript)
            : this()
        {
            Latin = aLatin;
            EastAsian = aEastAsian;
            ComplexScript = aComplexScript;
        }

        /// <summary>
        /// <b>NORMALLY THIS IS ALL YOU HAVE TO CHANGE TO CHANGE A TYPEFACE.</b> Check with APIMate if unsure.
        /// This element specifies that a Latin font be used for a specific run of text. This font is specified with a typeface 
        /// attribute much like the others but is specifically classified as a Latin font. 
        /// </summary>
        public TThemeTextFont Latin { get { return FLatin; } set { FLatin = value; } }

        /// <summary>
        /// This element specifies that an East Asian font be used for a specific run of text. This font is specified with a 
        /// typeface attribute much like the others but is specifically classified as an East Asian font. 
        /// </summary>
        public TThemeTextFont EastAsian { get { return FEastAsian; } set { FEastAsian = value; } }

        /// <summary>
        /// This element specifies that a complex script font be used for a specific run of text. This font is specified with a 
        /// typeface attribute much like the others but is specifically classified as a complex script font. 
        /// </summary>
        public TThemeTextFont ComplexScript { get { return FComplexScript; } set { FComplexScript = value; } }

        /// <summary>
        /// Adds a new font to the font collection.
        /// </summary>
        /// <param name="script">Script to be associated with the font.</param>
        /// <param name="typeface">Typeface to associate. Set it to null to delete the element</param>
        public void AddFont(string script, string typeface)
        {
            if (typeface == null && FFontCollection.ContainsKey(script)) FFontCollection.Remove(script);
            FFontCollection[script] = typeface;
        }

        /// <summary>
        /// Returns the typeface associated with a script.
        /// </summary>
        /// <param name="script"></param>
        public string GetFont(string script)
        {
            string Result;
            if (FFontCollection.TryGetValue(script, out Result)) return Result;
            return null;
        }

        /// <summary>
        /// Clears all font associations.
        /// </summary>
        public void ClearFonts()
        {
            FFontCollection.Clear();
        }

        /// <summary>
        /// Returns all scripts that have a current association. You can use the values in this array as keys for <see cref="GetFont(string)"/>
        /// </summary>
        /// <returns></returns>
        public string[] GetFontScripts()
        {
            string[] Result = new string[FFontCollection.Count];

            int i = 0;
            foreach (string s in FFontCollection.Keys)
            {
                Result[i] = s;
                i++;
            }

            return Result;
        }

        /// <summary>
        /// Returns a deep copy of this object.
        /// </summary>
        /// <returns></returns>
        public TThemeFont Clone()
        {
            TThemeFont Result = new TThemeFont();
            Result.Latin = Latin;
            Result.EastAsian = EastAsian;
            Result.ComplexScript = ComplexScript;

            foreach (string s in FFontCollection.Keys)
            {
                Result.AddFont(s, (string)FFontCollection[s]);
            }

            return Result;

        }

        #region Compare
        /// <summary>
        /// Returns true if both objects are the same.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            TThemeFont fnt = obj as TThemeFont;
            if (fnt == null) return this == null;

            return Latin == fnt.Latin && EastAsian == fnt.EastAsian && ComplexScript == fnt.ComplexScript
                && SameFontCollection(FFontCollection, fnt.FFontCollection);
        }

        private bool SameFontCollection(StringStringHashtable f1, StringStringHashtable f2)
        {
            if (f1 == null) return f2 == null;
            if (f2 == null) return false;
            if (f1.Count != f2.Count) return false;

            foreach (string key in f1.Keys)
            {
                if (!f2.ContainsKey(key)) return false;
                if (f2[key] != f1[key]) return false;
            }

            return true;
        }

        /// <summary>
        /// Returns the hashcode for the object.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHashObj(Latin, EastAsian, ComplexScript, FFontCollection);
        }
        #endregion
    }

    /// <summary>
    /// The characteristics that define a font.
    /// </summary>
    public struct TThemeTextFont: IComparable, IComparable<TThemeTextFont>
    {
        private readonly string FTypeface;
        private readonly string FPanose;
        private readonly TPitchFamily FPitch;
        private readonly TFontCharSet FCharSet;

        /// <summary>
        /// Creates a new TThemeTextFont.
        /// </summary>
        /// <param name="aTypeface">See <see cref="Typeface" /></param>
        /// <param name="aPanose">See <see cref="Panose" /></param>
        /// <param name="aPitch">See <see cref="Pitch" /></param>
        /// <param name="aCharSet">See <see cref="CharSet" /></param>
        public TThemeTextFont(string aTypeface, string aPanose, TPitchFamily aPitch, TFontCharSet aCharSet)
        {
            FTypeface = aTypeface;
            FPanose = aPanose;
            FPitch = aPitch;
            FCharSet = aCharSet;
        }


        /// <summary>
        /// Specifies the typeface, or name of the font that is to be used. The typeface is a string 
        /// name of the specific font that should be used in rendering the presentation. If this font is 
        /// not available within the font list of the generating application than font substitution logic 
        /// should be utilized in order to select an alternate font. 
        /// </summary>
        public string Typeface { get { return FTypeface; } }

        /// <summary>
        /// Specifies the Panose-1 classification number for the current font. 
        /// This is a string consisting of 20 hexadecimal digits which defines the Panose-1 font classification
        /// </summary>
        public string Panose { get { return FPanose; } }

        /// <summary>
        /// Specifies the font pitch as well as the font family for the corresponding font.
        /// </summary>
        public TPitchFamily Pitch { get { return FPitch; } }

        /// <summary>
        /// Specifies the character set which is supported by the parent font. This information can be 
        /// used in font substitution logic to locate an appropriate substitute font when this font is 
        /// not available. This information is determined by querying the font when present and shall 
        /// not be modified when the font is not available.
        /// </summary>
        public TFontCharSet CharSet { get { return FCharSet; } }

        internal bool IsStandard(string DefTypeFace)
        {
            return EqualTypeFace(Typeface, DefTypeFace)
            && Pitch == TPitchFamily.DEFAULT_PITCH__UNKNOWN_FONT_FAMILY
            && CharSet == TFontCharSet.Default
            && (Panose == null || Panose.Length == 0);

        }

        private bool EqualTypeFace(string Typeface, string DefTypeFace)
        {
            if (Typeface == null) Typeface = String.Empty;
            if (DefTypeFace == null) DefTypeFace = String.Empty;
            return Typeface == DefTypeFace;
        }

        #region IComparable Members
        /// <summary>
        /// Returns -1, 0 or 1 depending if the objects is smaller, equal or bigger than the other.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public int CompareTo(TThemeTextFont obj)
        {
            int r;
            r = FlxUtils.CompareObjects(Typeface, obj.Typeface); if (r != 0) return r;
            r = FlxUtils.CompareObjects(Panose, obj.Panose);if (r != 0) return r; 
            r = Pitch.CompareTo(obj.Pitch);if (r != 0) return r;
            r = CharSet.CompareTo(obj.CharSet); if (r != 0) return r;

            return 0;
        }

        /// <summary></summary>
        public int CompareTo(object obj)
        {
            if (!(obj is TThemeTextFont)) return -1;
            return CompareTo((TThemeTextFont)obj);
        }


        /// <summary>
        /// Returns a hashcode for the object.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHash(Typeface.GetHashCode(), Panose.GetHashCode(), Pitch.GetHashCode(), CharSet.GetHashCode());
        }

        /// <summary></summary>
        public override bool Equals(object obj)
        {
            if (!(obj is TThemeTextFont)) return false;
            return CompareTo((TThemeTextFont)obj) == 0;
        }

        /// <summary></summary>
        public static bool operator ==(TThemeTextFont f1, TThemeTextFont f2)
        {
            return f1.CompareTo(f2) == 0;
        }

        /// <summary></summary>
        public static bool operator !=(TThemeTextFont f1, TThemeTextFont f2)
        {
            return f1.CompareTo(f2) != 0;
        }

        /// <summary></summary>
        public static bool operator >(TThemeTextFont o1, TThemeTextFont o2)
        {
            return o1.CompareTo(o2) > 0;
        }

        /// <summary></summary>
        public static bool operator <(TThemeTextFont o1, TThemeTextFont o2)
        {
            return o1.CompareTo(o2) < 0;
        }



        #endregion
    }

    #endregion

    #region Formats
    /// <summary>
    /// The elements here are different types of fill/effects/etc that can be applied to an object.
    /// Currently Excel only defines 3 types, but more can be added in future specifications. 
    /// Newer types can be accessed by casting an int to TFormattingType. (for example "(TFormattingType)5"  would refer to a now non existing formatting type)
    /// </summary>
    public enum TFormattingType
    {
        /// <summary>
        /// Subtle formatting type.
        /// </summary>
        Subtle = 0,

        /// <summary>
        /// Moderate formatting type.
        /// </summary>
        Moderate = 1,

        /// <summary>
        /// Intense formatting type.
        /// </summary>
        Intense = 2
    }

    /// <summary>
    /// Represents the drawing formats (fill styles, line styles, effects) for a theme.
    /// </summary>
    public class TThemeFormatScheme
    {
        private string FName;
        private TFillStyleList FFillStyleList;
        private TLineStyleList FLineStyleList;
        private TEffectStyleList FEffectStyleList;
        private TFillStyleList FBkFillStyleList;

        /// <summary>
        /// Creates a new Format scheme instance.
        /// </summary>
        public TThemeFormatScheme()
        {
            FFillStyleList = new TFillStyleList();
            FLineStyleList = new TLineStyleList();
            FEffectStyleList = new TEffectStyleList();
            FBkFillStyleList = new TFillStyleList();
            FName = "Office";
        }

        /// <summary>
        /// Name of the format definition. This will be shown in Excel UI.
        /// </summary>
        public string Name { get { return FName; } set { FName = value; } }

        /// <summary>
        /// This element defines a set of three fill styles that are used within a theme.
        /// </summary>
        public TFillStyleList FillStyleList { get { return FFillStyleList; } }

        /// <summary>
        /// This element defines a list of three line styles for use within a theme.
        /// </summary>
        public TLineStyleList LineStyleList { get { return FLineStyleList; } }

        /// <summary>
        /// This element defines a set of three effect styles that create the effect style list for a theme.
        /// </summary>
        public TEffectStyleList EffectStyleList { get { return FEffectStyleList; } }

        /// <summary>
        /// This element defines a list of background fills that are used within a theme.
        /// </summary>
        public TFillStyleList BkFillStyleList { get { return FBkFillStyleList; } }

        /// <summary>
        /// Returns a deep copy of this object.
        /// </summary>
        /// <returns></returns>
        internal TThemeFormatScheme Clone()
        {
            TThemeFormatScheme Result = new TThemeFormatScheme();
            Result.Name = Name;
            Result.FFillStyleList = FFillStyleList.Clone();
            Result.FLineStyleList = FLineStyleList.Clone();
            Result.FEffectStyleList = FEffectStyleList.Clone();
            Result.FBkFillStyleList = FBkFillStyleList.Clone();
            return Result;
        }

        internal bool IsStandard
        {
            get
            {
                return (Name == "Office") &&
                       FillStyleList.IsStandard(false) &&
                       BkFillStyleList.IsStandard(true);
            }
        }
    }

    /// <summary>
    /// Represents the fill style characteristics for an autoshape.
    /// </summary>
    public class TFillStyleList
    {
        private List<TFillStyle> FFillCollection;

        /// <summary>
        /// Creates a new TFillStyleList instance.
        /// </summary>
        public TFillStyleList()
        {
            FFillCollection = new List<TFillStyle>();
        }

        /// <summary>
        /// Creates a deep copy of this object.
        /// </summary>
        /// <returns></returns>
        public TFillStyleList Clone()
        {
            TFillStyleList Result = new TFillStyleList();
            foreach (TFillStyle fs in FFillCollection)
            {
                if (fs == null) Result.FFillCollection.Add(null); else Result.FFillCollection.Add(fs.Clone());
            }

            return Result;
        }


        /// <summary>
        /// Adds a new FillStyle to the collection. Fill styles must be added in order, first is "Subtle", second is "Moderate", third is "Intense"
        /// and there could be new definitions in newer versions of Excel.
        /// </summary>
        /// <param name="aFill">Fill style to add.</param>
        public void Add(TFillStyle aFill)
        {
            FFillCollection.Add(aFill);
        }

        /// <summary>
        /// Returns the fill style for a given formatting type. Currently Excel defines only 3 formatting types, but more could be added in the future.
        /// If you need to access a formatting type that is not defined in the <see cref="TFormattingType"/> enumeration, just cast an integer to TFormattingType.
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public TFillStyle this[TFormattingType index]
        {
            get
            {
                return FFillCollection[(int)index];
            }
            set
            {
                FFillCollection[(int)index] = value;
            }
        }

        /// <summary>
        /// Returns the fill style that results from applying the formatting type to a color.
        /// </summary>
        /// <param name="index">Index to the collection.</param>
        /// <param name="basicColor">Color that will be used as a base to calculate the fill style.</param>
        /// <returns></returns>
        public TFillStyle GetRealFillStyle(TFormattingType index, TDrawingColor basicColor)
        {
            return FFillCollection[(int)index].ReplacePhClr(basicColor);
        }

        /// <summary>
        /// Returns the number of elements stored in this collection.
        /// </summary>
        public int Count
        {
            get
            {
                return FFillCollection.Count;
            }
        }

        /// <summary>
        /// Clears all the formatting definitions.
        /// </summary>
        public void Clear()
        {
            FFillCollection.Clear();
        }

        internal bool IsStandard(bool IsBkFill)
        {
#if(FRAMEWORK30 && !COMPACTFRAMEWORK)
            if (FFillCollection.Count > 3) return false;

            for (int i = 0; i < 3; i++)
            {
                if (i >= FFillCollection.Count) continue;
                TFillStyle fs = FFillCollection[i];
                if (fs == null) continue;

                if (IsBkFill)
                {
                    if (!Object.Equals(fs, GetDefaultBkFillStyle(i))) return false;
                }
                else
                {
                    if (!Object.Equals(fs, GetDefaultFillStyle(i))) return false;
                }
            }
#endif
            return true;
        }

#if(FRAMEWORK30 && !COMPACTFRAMEWORK)
        internal static TFillStyle GetDefaultFillStyle(int i)
        {
            switch (i)
            {
                case 0:
                    return new TSolidFill(TDrawingColor.FromTheme(TThemeColor.None)); //phClr

                case 1:
                    return new TGradientFill(null, true, TFlipMode.None,
                        new TDrawingGradientStop[]
                        {
                            new TDrawingGradientStop(0, TDrawingColor.AddTransform(TDrawingColor.FromTheme(TThemeColor.None), 
                                new TColorTransform[]{new TColorTransform(TColorTransformType.Tint, 0.5), new TColorTransform(TColorTransformType.SatMod,3)})),

                            new TDrawingGradientStop(0.35, TDrawingColor.AddTransform(TDrawingColor.FromTheme(TThemeColor.None), 
                                new TColorTransform[]{new TColorTransform(TColorTransformType.Tint, 0.37), new TColorTransform(TColorTransformType.SatMod,3)})),

                            new TDrawingGradientStop(1, TDrawingColor.AddTransform(TDrawingColor.FromTheme(TThemeColor.None), 
                                new TColorTransform[]{new TColorTransform(TColorTransformType.Tint, 0.15), new TColorTransform(TColorTransformType.SatMod,3.5)})),

                        },
                        new TDrawingLinearGradient(270, true));

                case 2:
                    return new TGradientFill(null, true, TFlipMode.None,
                        new TDrawingGradientStop[]
                        {
                            new TDrawingGradientStop(0, TDrawingColor.AddTransform(TDrawingColor.FromTheme(TThemeColor.None), 
                                new TColorTransform[]{new TColorTransform(TColorTransformType.Shade, 0.51), new TColorTransform(TColorTransformType.SatMod,1.3)})),

                            new TDrawingGradientStop(0.8, TDrawingColor.AddTransform(TDrawingColor.FromTheme(TThemeColor.None), 
                                new TColorTransform[]{new TColorTransform(TColorTransformType.Shade, 0.93), new TColorTransform(TColorTransformType.SatMod,1.3)})),

                            new TDrawingGradientStop(1, TDrawingColor.AddTransform(TDrawingColor.FromTheme(TThemeColor.None), 
                                new TColorTransform[]{new TColorTransform(TColorTransformType.Shade, 0.94), new TColorTransform(TColorTransformType.SatMod,1.35)})),

                        },
                        new TDrawingLinearGradient(270, false));
                default:
                    return new TNoFill();
            }
        }

        internal static TFillStyle GetDefaultBkFillStyle(int i)
        {
            switch (i)
            {
                case 0:
                    return new TSolidFill(TDrawingColor.FromTheme(TThemeColor.None)); //phClr

                case 1:
                    return new TGradientFill(null, true, TFlipMode.None,
                        new TDrawingGradientStop[]
                        {
                            new TDrawingGradientStop(0, TDrawingColor.AddTransform(TDrawingColor.FromTheme(TThemeColor.None), 
                                new TColorTransform[]{new TColorTransform(TColorTransformType.Tint, 0.4), new TColorTransform(TColorTransformType.SatMod,3.5)})),

                            new TDrawingGradientStop(0.4, TDrawingColor.AddTransform(TDrawingColor.FromTheme(TThemeColor.None), 
                                new TColorTransform[]{new TColorTransform(TColorTransformType.Tint, 0.45), new TColorTransform(TColorTransformType.Shade, 0.99), new TColorTransform(TColorTransformType.SatMod,3.5)})),

                            new TDrawingGradientStop(1, TDrawingColor.AddTransform(TDrawingColor.FromTheme(TThemeColor.None), 
                                new TColorTransform[]{new TColorTransform(TColorTransformType.Shade, 0.2), new TColorTransform(TColorTransformType.SatMod,2.55)})),

                        },
                        new TDrawingPathGradient(new TDrawingRelativeRect(0.5, -0.8, 0.5, 1.8), TPathShadeType.Circle));

                case 2:
                    return new TGradientFill(null, true, TFlipMode.None,
                        new TDrawingGradientStop[]
                        {
                            new TDrawingGradientStop(0, TDrawingColor.AddTransform(TDrawingColor.FromTheme(TThemeColor.None), 
                                new TColorTransform[]{new TColorTransform(TColorTransformType.Tint, 0.8), new TColorTransform(TColorTransformType.SatMod,3)})),

                            new TDrawingGradientStop(1, TDrawingColor.AddTransform(TDrawingColor.FromTheme(TThemeColor.None), 
                                new TColorTransform[]{new TColorTransform(TColorTransformType.Shade, 0.3), new TColorTransform(TColorTransformType.SatMod,2)})),

                        },
                        new TDrawingPathGradient(new TDrawingRelativeRect(0.5, 0.5, 0.5, 0.5), TPathShadeType.Circle));
                default:
                    return new TNoFill();
            }
        }
#endif
    }

    /// <summary>
    /// Stores the different kind of fill styles for an autoshape or drawing.
    /// </summary>
    public enum TFillStyleType
    {
        /// <summary>
        /// No fill associated with the shape.
        /// </summary>
        NoFill,

        /// <summary>
        /// Shape is filled with a solid color.
        /// </summary>
        Solid,

        /// <summary>
        /// Shape is filled with a gradient.
        /// </summary>
        Gradient,

        /// <summary>
        /// Shape is filled with an image.
        /// </summary>
        Blip,

        /// <summary>
        /// Shape is filled with a pattern.
        /// </summary>
        Pattern,

        /// <summary>
        /// The shape is part of a group, and should inherit its parent fill style.
        /// </summary>
        Group
    }

    /// <summary>
    /// Specifies the alignment to be used for the underline stroke. 
    /// </summary>
    public enum TPenAlignment
    {
        /// <summary>
        /// Center pen (line drawn at center of path stroke).
        /// </summary>
        Center,

        /// <summary>
        /// Inset pen (the pen is aligned on the inside of the edge of the path). 
        /// </summary>
        Inset
    }

    /// <summary>
    /// How the line ends.
    /// </summary>
    public enum TLineCap
    {
        /// <summary>
        /// Line ends at end point.
        /// </summary>
        Flat,

        /// <summary>
        /// Rounded ends. Semi-circle protrudes by half line width. 
        /// </summary>
        Round,

        /// <summary>
        /// Square protrudes by half line width. 
        /// </summary>
        Square
    }

    /// <summary>
    /// Type of line.
    /// </summary>
    public enum TCompoundLineType
    {
        /// <summary>
        /// A single line with normal width.
        /// </summary>
        Single,

        /// <summary>
        /// A double line, both with normal width.
        /// </summary>
        Double,

        /// <summary>
        /// A double line, first think and the other thin.
        /// </summary>
        ThickThin,

        /// <summary>
        /// A double line, first think and the other thick.
        /// </summary>
        ThinThick,

        /// <summary>
        /// A triple line, Thin, thick and thin.
        /// </summary>
        Triple
    }

    /// <summary>
    /// Base definition for a Drawing fill style. This class is abstract, and you should use its descendants like <see cref="TSolidFill"/> or <see cref="TGradientFill"/>
    /// </summary>
    public abstract class TFillStyle: IComparable
    {
        private readonly TFillStyleType FFillStyleType;

        /// <summary>
        /// Initializes the fill style.
        /// </summary>
        /// <param name="aFillStyleType"></param>
        protected TFillStyle(TFillStyleType aFillStyleType)
        {
            FFillStyleType = aFillStyleType;
        }

        /// <summary>
        /// Stores which kind of fill style is used.
        /// </summary>
        public TFillStyleType FillStyleType { get { return FFillStyleType; } }

        /// <summary>
        /// Returns a deep copy of the fill style.
        /// </summary>
        /// <returns></returns>
        public virtual TFillStyle Clone()
        {
            return (TFillStyle)MemberwiseClone();
        }

        /// <summary>
        /// Returns true if this instance has the same data as the object obj.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            TFillStyle o2 = obj as TFillStyle;
            if (o2 == null) return false;
            return FFillStyleType == o2.FFillStyleType;
        }

        /// <summary>
        /// Returns -1 if obj is bigger than this, 0 if both objects are the same, and 1 if obj is smaller than this.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public virtual int CompareTo(object obj)
        {
            TFillStyle s2 = obj as TFillStyle;
            if (s2 == null) return -1;

            int r = FFillStyleType.CompareTo(s2.FFillStyleType);
            if (r != 0) return r;

            return 0;
        }

        /// <summary>
        /// Returns the hashcode for this object
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHash((int)FFillStyleType);
        }

        internal virtual TFillStyle ReplacePhClr(TDrawingColor basicColor)
        {
            return Clone();
        }
    }

    #region Specialized fill styles
    #region NoFill
    /// <summary>
    /// There is no fill associated with the shapes.
    /// </summary>
    public class TNoFill : TFillStyle
    {
        /// <summary>
        /// Creates a new TNoFill instance.
        /// </summary>
        public TNoFill()
            : base(TFillStyleType.NoFill)
        {
        }

        /// <summary>
        /// Returns true if this instance has the same data as the object obj.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            TNoFill o2 = obj as TNoFill;
            if (o2 == null) return false;
            return base.Equals(obj);
        }

        /// <summary>
        /// Returns the hashcode for this object
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return base.GetHashCode();
        }
    }
    #endregion

    #region Solid Fill
    /// <summary>
    /// Shapes are filled with a solid color.
    /// </summary>
    public class TSolidFill : TFillStyle
    {
        readonly TDrawingColor FColor;

        /// <summary>
        /// Creates a new Solid fill.
        /// </summary>
        /// <param name="aColor">Color for the fill.</param>
        public TSolidFill(TDrawingColor aColor)
            : base(TFillStyleType.Solid)
        {
            FColor = aColor;
        }

        /// <summary>
        /// Color used to fill the shape.
        /// </summary>
        public TDrawingColor Color { get { return FColor; } }


        internal override TFillStyle ReplacePhClr(TDrawingColor basicColor)
        {
            return new TSolidFill(FColor.ReplacePhClr(basicColor));
        }

        /// <summary>
        /// Returns true if this instance has the same data as the object obj.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            TSolidFill o2 = obj as TSolidFill;
            if (o2 == null) return false;
            return base.Equals(obj) &&
                FColor == o2.FColor;
        }

        /// <summary>
        /// Returns -1 if obj is bigger than this, 0 if both objects are the same, and 1 if obj is smaller than this.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override int CompareTo(object obj)
        {
            TSolidFill s2 = obj as TSolidFill;
            if (s2 == null) return -1;

            int r = FColor.CompareTo(s2.FColor);
            if (r != 0) return r;

            return base.CompareTo(obj);
        }

        /// <summary>
        /// Returns the hashcode for this object
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHash(base.GetHashCode(), FColor.GetHashCode());
        }

    }
    #endregion

    #region GradientFill
    /// <summary>
    /// Shapes are filled with a gradient.
    /// </summary>
    public class TGradientFill : TFillStyle
    {
        readonly TDrawingRelativeRect? FTileRect;
        readonly bool FRotateWithShape;
        readonly TFlipMode FFlip;
        readonly TDrawingGradientStop[] FGradientStops;
        readonly TDrawingGradientDef FGradientDef;

        /// <summary>
        /// Creates a new Gradient fill.
        /// </summary>
        public TGradientFill(TDrawingRelativeRect? aTileRect, bool aRotateWithShape, TFlipMode aFlip, TDrawingGradientStop[] aGradientStops, TDrawingGradientDef aGradientDef)
            : base(TFillStyleType.Gradient)
        {
            FTileRect = aTileRect;
            FRotateWithShape = aRotateWithShape;
            FFlip = aFlip;
            if (aGradientStops != null)
            {
                FGradientStops = new TDrawingGradientStop[aGradientStops.Length];
                aGradientStops.CopyTo(FGradientStops, 0);
            }

            FGradientDef = aGradientDef; //no need to clone, classes are immutable.
        }

        /// <summary>
        /// Specifies that the fill should rotate with the shape.
        /// </summary>
        public bool RotateWithShape { get { return FRotateWithShape; } }

        /// <summary>
        /// This element specifies a rectangular region of the shape to which the gradient is applied.  This region is then 
        /// tiled across the remaining area of the shape to complete the fill.  The tile rectangle is defined by percentage 
        /// offsets from the sides of the shape's bounding box.
        /// </summary>
        public TDrawingRelativeRect? TileRect { get { return FTileRect; }}

        /// <summary>
        /// Specifies the direction(s) in which to flip the gradient while tiling.   <br></br>
        /// Normally a gradient fill encompasses the entire bounding box of the shape which 
        /// contains the fill. However, with the tileRect element, it is possible to define a "tile" 
        /// rectangle which is smaller than the bounding box. In this situation, the gradient fill is 
        /// encompassed within the tile rectangle, and the tile rectangle is tiled across the bounding box to fill the entire area
        /// </summary>
        public TFlipMode Flip { get { return FFlip; } }

        /// <summary>
        /// The list of gradient stops that specifies the gradient colors and their relative positions in the color band.
        /// </summary>
        public TDrawingGradientStop[] GradientStops { get { return FGradientStops; } }

        /// <summary>
        /// Definition of the gradient. This can be a TDrawingLinearGradient class or a TDrawingPathGradient class.
        /// </summary>
        public TDrawingGradientDef GradientDef { get { return FGradientDef; } }

        internal override TFillStyle ReplacePhClr(TDrawingColor basicColor)
        {
            TDrawingGradientStop[] NewGradientStops = null;
            if (GradientStops != null)
            {
                NewGradientStops = new TDrawingGradientStop[GradientStops.Length];
                for (int i = 0; i < GradientStops.Length; i++)
                {
                    NewGradientStops[i] = new TDrawingGradientStop(GradientStops[i].Position, 
                        GradientStops[i].Color.ReplacePhClr(basicColor));
                }
            }

            TDrawingGradientDef NewGradientDef = null;
            if (GradientDef != null)
            {
                NewGradientDef = GradientDef.Clone();
            }
            return new TGradientFill(FTileRect, FRotateWithShape, FFlip, NewGradientStops, NewGradientDef);
        }

        /// <summary>
        /// Returns true if this instance has the same data as the object obj.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            TGradientFill o2 = obj as TGradientFill;
            if (o2 == null) return false;
            return base.Equals(obj) &&
                        FTileRect == o2.FTileRect &&
                        FRotateWithShape == o2.FRotateWithShape &&
                        FFlip == o2.FFlip &&
                        SameStops(FGradientStops, o2.FGradientStops) &&
                        object.Equals(FGradientDef, o2.FGradientDef);
        }

        /// <summary>
        /// Returns -1 if obj is bigger than this, 0 if both objects are the same, and 1 if obj is smaller than this.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override int CompareTo(object obj)
        {
            TGradientFill s2 = obj as TGradientFill;
            if (s2 == null) return -1;

            int r;
            if (!FTileRect.HasValue)
            {
                if (s2.FTileRect.HasValue) return -1;
            }
            else
            {
                if (!s2.FTileRect.HasValue) return 1;
                r = FTileRect.Value.CompareTo(s2.FTileRect.Value);
                if (r != 0) return r;
            }

            r = FRotateWithShape.CompareTo(s2.FRotateWithShape);
            if (r != 0) return r;

            r = FFlip.CompareTo(s2.FFlip);
            if (r != 0) return r;

            r = FlxUtils.CompareArray(FGradientStops, s2.FGradientStops);
            if (r != 0) return r;

            if (FGradientDef == null)
            {
                if (s2.FGradientDef != null) return -1;
            }
            else
            {
                r = FGradientDef.CompareTo(s2.FGradientDef);
                if (r != 0) return r;
            }

            return base.CompareTo(obj);
        }


        private static bool SameStops(TDrawingGradientStop[] stop1, TDrawingGradientStop[] stop2)
        {
            if (stop1 == null) return stop2 == null;
            if (stop2 == null) return false;

            if (stop1.Length != stop2.Length) return false;
            for (int i = 0; i < stop1.Length; i++)
            {
                if (stop1[i] != stop2[i]) return false;
            }

            return true;
        }

        /// <summary>
        /// Returns the hashcode for this object
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            int GrDefHashCode = GradientDef == null ? 0 : FGradientDef.GetHashCode();
            return HashCoder.GetHash(base.GetHashCode(), FTileRect.GetHashCode(), FRotateWithShape.GetHashCode(),FFlip.GetHashCode(), StopHashCode(FGradientStops), GrDefHashCode);
        }

        private static int StopHashCode(TDrawingGradientStop[] stops)
        {
            int Result = 0;
            if (stops == null) return 0;

            for (int i = 0; i < stops.Length; i++)
            {
                Result ^= stops[i].GetHashCode();
            }

            return Result;
        }

    }

    /// <summary>
    /// A base class for storing gradient definitions, be them Linear or Path gradients.
    /// </summary>
    public abstract class TDrawingGradientDef
    {
        /// <summary>
        /// Rerturns a deep copy of this object.
        /// </summary>
        /// <returns></returns>
        public virtual TDrawingGradientDef Clone()
        {
            return MemberwiseClone() as TDrawingGradientDef;
        }

        /// <summary>
        /// Returns if an objects is bigger than the other.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public virtual int CompareTo(object obj)
        {
            return 0;
        }
    }

    /// <summary>
    /// This class holds a linear gradient definition.
    /// </summary>
    public class TDrawingLinearGradient : TDrawingGradientDef
    {
        readonly double FAngle;
        readonly bool FScaled;

        /// <summary>
        /// Creates a new Linear gradient definition.
        /// </summary>
        /// <param name="aAngle"></param>
        /// <param name="aScaled"></param>
        public TDrawingLinearGradient(double aAngle, bool aScaled)
        {
            FAngle = aAngle;
            FScaled = aScaled;
        }

        /// <summary>
        /// Specifies the direction of color change for the gradient. To define this angle, let its value 
        /// be x measured clockwise. Then ( -sin x, cos x ) is a vector parallel to the line of constant 
        /// color in the gradient fill. 
        /// </summary>
        public double Angle { get { return FAngle; } }

        /// <summary>
        /// Whether the gradient angle scales with the fill region. Mathematically, if this flag is true, 
        /// then the gradient vector ( cos x , sin x ) is scaled by the width (w) and height (h) of the fill 
        /// region, so that the vector becomes ( w cos x, h sin x ) (before normalization).
        /// </summary>
        public bool Scaled { get { return FScaled; } }

        /// <summary>
        /// Returns true if this instance has the same data as the object obj.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            TDrawingLinearGradient o2 = obj as TDrawingLinearGradient;
            if (o2 == null) return false;
            return
                FAngle == o2.FAngle &&
                FScaled == o2.FScaled;
        }

        /// <summary>
        /// Returns -1 if obj is bigger than this, 0 if both objects are the same, and 1 if obj is smaller than this.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override int CompareTo(object obj)
        {
            TDrawingLinearGradient s2 = obj as TDrawingLinearGradient;
            if (s2 == null) return -1;

            int r = FAngle.CompareTo(s2.FAngle);
            if (r != 0) return r;

            r = FScaled.CompareTo(s2.FScaled);
            if (r != 0) return r;

            return base.CompareTo(obj);
        }

        /// <summary>
        /// Returns the hashcode for this object
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHash(FAngle.GetHashCode(), FScaled.GetHashCode());
        }

    }

    /// <summary>
    /// Enumerates the different gradient modes for a Path gradient.
    /// </summary>
    public enum TPathShadeType
    {
        /// <summary>
        /// Gradient follows the shape.
        /// </summary>
        Shape,

        /// <summary>
        /// Gradient follows a circular path.
        /// </summary>
        Circle,

        /// <summary>
        /// Gradient follows a rectangular path.
        /// </summary>
        Rect
    }


    /// <summary>
    /// Holds a Path gradient definition.
    /// </summary>
    public class TDrawingPathGradient : TDrawingGradientDef
    {
        readonly TDrawingRelativeRect? FFillToRect;
        readonly TPathShadeType FPath;

        /// <summary>
        /// Creates a new TDrawingPath object.
        /// </summary>
        /// <param name="aFillToRect"></param>
        /// <param name="aPath"></param>
        public TDrawingPathGradient(TDrawingRelativeRect? aFillToRect, TPathShadeType aPath)
        {
            FFillToRect = aFillToRect;
            FPath = aPath;
        }

        /// <summary>
        /// This element defines the "focus" rectangle for the center shade, specified relative to the fill tile rectangle.   <br></br>
        /// The center shade fills the entire tile except the margins specified by each attribute. <br></br>
        /// Each edge of the center shade rectangle is defined by a percentage offset from the corresponding edge of the 
        /// tile rectangle.  A positive percentage specifies an inset, while a negative percentage specifies an outset. 
        /// </summary>
        public TDrawingRelativeRect? FillToRect { get { return FFillToRect; }}

        /// <summary>
        /// Specifies the shape of the path to follow.
        /// </summary>
        public TPathShadeType Path { get { return FPath; } }

        /// <summary>
        /// Returns true if this instance has the same data as the object obj.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            TDrawingPathGradient o2 = obj as TDrawingPathGradient;
            if (o2 == null) return false;
            return 
                FFillToRect == o2.FFillToRect &&
                FPath == o2.FPath;
        }

        /// <summary>
        /// Returns -1 if obj is bigger than this, 0 if both objects are the same, and 1 if obj is smaller than this.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override int CompareTo(object obj)
        {
            TDrawingPathGradient s2 = obj as TDrawingPathGradient;
            if (s2 == null) return -1;

            int r;
            if (!FFillToRect.HasValue)
            {
                if (s2.FFillToRect.HasValue) return -1;
            }
            else
            {
                if (!s2.FFillToRect.HasValue) return 1;

                r = FFillToRect.Value.CompareTo(s2.FFillToRect.Value);
                if (r != 0) return r;
            }

            r = FPath.CompareTo(s2.FPath);
            if (r != 0) return r;

            return base.CompareTo(obj);
        }


        /// <summary>
        /// Returns the hashcode for this object
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHash(FFillToRect.GetHashCode(), FPath.GetHashCode());
        }
    }

        
    #endregion

    #region Blip Fill

    /// <summary>
    /// This type specifies the amount of compression that has been used for a particular picture.
    /// </summary>
    public enum TBlipCompression
    {
        /// <summary>
        /// No compression used.
        /// </summary>
        None,

        /// <summary>
        /// Compressed for email.
        /// </summary>
        Email,

        /// <summary>
        /// Compressed for screen viewing.
        /// </summary>
        Screen,

        /// <summary>
        /// Compressed for printing.
        /// </summary>
        Print,

        /// <summary>
        /// Compressed for high quality printing.
        /// </summary>
        HQPrint
    }

    /// <summary>
    /// Picture and properties used in a Blip fill.
    /// </summary>
    public class TBlip: IComparable
    {
        private readonly TBlipCompression FCompressionState;
        private readonly byte[] FPictureData;
        private readonly string FImageFileName;
        private readonly string FContentType;
        private readonly string[] FTransforms;

        /// <summary>
        /// Creates a new Blip.
        /// </summary>
        /// <param name="aCompressionState">Compression state.</param>
        /// <param name="aPictureData">Picture Data. This data will be copied in this object, so after using it, you can modify the original and this won't change.</param>
        /// <param name="aImageFileName">File name which will be used when saving the file inside the xlsx container.</param>
        /// <param name="aContentType">Content type for the image, like "image/jpeg".</param>
        public TBlip(TBlipCompression aCompressionState, byte[] aPictureData, string aImageFileName, string aContentType)
            : this(aCompressionState, aPictureData, aImageFileName, aContentType, null)
        {
        }


        /// <summary>
        /// This should really have a parsed Transforms, this is why it isn't public.
        /// </summary>
        internal TBlip(TBlipCompression aCompressionState, byte[] aPictureData, string aImageFileName, string aContentType, string[] aTransforms)
        {
            FCompressionState = aCompressionState;
            if (aPictureData != null)
            {
                FPictureData = (byte[])aPictureData.Clone();
            }
            else FPictureData = null;

            FTransforms = null;
            if (aTransforms != null)
            {
                FTransforms = new string[aTransforms.Length];
                aTransforms.CopyTo(FTransforms, 0);
            }

            FImageFileName = aImageFileName;
            FContentType = aContentType;
        }

        /// <summary>
        /// Specifies the compression state with which the picture is stored. This allows the 
        /// application to specify the amount of compression that has been applied to a picture. 
        /// </summary>
        public TBlipCompression CompressionState { get { return FCompressionState; } }

        /// <summary>
        /// Image data.
        /// </summary>
        public byte[] PictureData { get { return FPictureData; } }

        internal string[] Transforms { get { return FTransforms; } }

        /// <summary>
        /// File name which will be used when saving the file inside the xlsx container.
        /// </summary>
        public string ImageFileName { get { return FImageFileName; } }

        /// <summary>
        /// Content type for the image, like "image/jpeg".
        /// </summary>
        public string ContentType { get { return FContentType; } }

        internal TBlip Clone()
        {
            return new TBlip(CompressionState, PictureData, ImageFileName, ContentType, Transforms);
        }

        /// <summary>
        /// Returns true if this instance has the same data as the object obj.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            TBlip o2 = obj as TBlip;
            if (o2 == null) return false;
            return
                FCompressionState == o2.FCompressionState &&
                FlxUtils.CompareMem(FPictureData, o2.FPictureData) &&
                FImageFileName == o2.FImageFileName &&
                FContentType == o2.FContentType &&
                CompareTransforms(FTransforms, o2.FTransforms);
        }

        private static bool CompareTransforms(string[] a1, string[] a2)
        {
            if (a1 == null) return a2 == null;
            if (a2 == null) return false;
            if (a1.Length != a2.Length) return false;

            for (int i = 0; i < a1.Length; i++)
            {
                if (a1[i] != a2[i]) return false;
            }

            return true;
        }

        /// <summary>
        /// Returns -1 if obj is bigger than this, 0 if both objects are the same, and 1 if obj is smaller than this.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public int CompareTo(object obj)
        {
            TBlip s2 = obj as TBlip;
            if (s2 == null) return -1;

            int r = FCompressionState.CompareTo(s2.FCompressionState);
            if (r != 0) return r;

            r = FlxUtils.CompareArray(FPictureData, s2.FPictureData);
            if (r != 0) return r;

            r = string.Compare(FImageFileName, s2.FImageFileName);
            if (r != 0) return r;

            r = string.Compare(FContentType, s2.FContentType);
            if (r != 0) return r;

            r = FlxUtils.CompareArray(FTransforms, s2.FTransforms);
            if (r != 0) return r;

            return 0;
        }

        /// <summary>
        /// Returns the hashcode for this object
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHashObj(CompressionState, PictureData, ImageFileName, ContentType, Transforms);
        }
    }

    /// <summary>
    /// This class is sed to store tiling or stretching fill mode information.
    /// You need to use any of its descendants, <see cref="TBlipFillStretch"/> or <see cref="TBlipFillTile"/>
    /// </summary>
    public abstract class TBlipFillMode: IComparable
    {
        /// <summary>
        /// Returns -1 if obj is bigger than this, 0 if both objects are the same, and 1 if obj is smaller than this.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public virtual int CompareTo(object obj)
        {
            return 0;
        }

    }

    /// <summary>
    /// This element specifies that a BLIP should be stretched to fill the target rectangle. The other option is a tile where 
    /// a BLIP is tiled to fill the available area.
    /// </summary>
    public class TBlipFillStretch : TBlipFillMode
    {
        private readonly TDrawingRelativeRect FFillRect;

        /// <summary>
        /// Creates a new TBlipFillStretch object.
        /// </summary>
        /// <param name="aFillRect"></param>
        public TBlipFillStretch(TDrawingRelativeRect aFillRect)
        {
            FFillRect = aFillRect;
        }

        /// <summary>
        /// Rectangle where the picture will be streched.
        /// </summary>
        public TDrawingRelativeRect FillRect { get { return FFillRect; } }

        /// <summary>
        /// Returns true if this instance has the same data as the object obj.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            TBlipFillStretch o2 = obj as TBlipFillStretch;
            if (o2 == null) return false;
            return base.Equals(obj) &&
                FFillRect == o2.FFillRect;
        }

        /// <summary>
        /// Returns -1 if obj is bigger than this, 0 if both objects are the same, and 1 if obj is smaller than this.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override int CompareTo(object obj)
        {
            TBlipFillStretch s2 = obj as TBlipFillStretch;
            if (s2 == null) return -1;

            int r = FFillRect.CompareTo(s2.FillRect);
            if (r != 0) return r;

            return base.CompareTo(obj);
        }


        /// <summary>
        /// Returns the hashcode for this object
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHash(base.GetHashCode(), FFillRect.GetHashCode());
        }
    }


    /// <summary>
    /// This element specifies that a BLIP should be tiled to fill the available space.  This element defines a "tile" 
    /// rectangle within the bounding box.  The image is encompassed within the tile rectangle, and the tile rectangle is 
    /// tiled across the bounding box to fill the entire area. 
    /// </summary>
    public class TBlipFillTile : TBlipFillMode
    {
        private readonly TDrawingRectAlign FAlign;
        private readonly TFlipMode FFlip;

        private readonly TDrawingCoordinate FTx;
        private readonly TDrawingCoordinate FTy;
        private readonly double FScaleX;
        private readonly double FScaleY;

        /// <summary>
        /// Creates a new Blip fill tile instance.
        /// </summary>
        /// <param name="aAlign"></param>
        /// <param name="aFlip"></param>
        /// <param name="aTx"></param>
        /// <param name="aTy"></param>
        /// <param name="aScaleX"></param>
        /// <param name="aScaleY"></param>
        public TBlipFillTile(TDrawingRectAlign aAlign, TFlipMode aFlip, TDrawingCoordinate aTx, TDrawingCoordinate aTy, double aScaleX, double aScaleY)
        {
            FAlign = aAlign;
            FFlip = aFlip;
            FTx = aTx;
            FTy = aTy;
            FScaleX = aScaleX;
            FScaleY = aScaleY;
        }

        /// <summary>
        /// Specifies where to align the first tile with respect to the shape.  Alignment happens after the scaling, but before the additional offset.
        /// </summary>
        public TDrawingRectAlign Align { get { return FAlign; } }

        /// <summary>
        /// Specifies the direction(s) in which to flip the source image while tiling.  Images can be
        /// flipped horizontally, vertically, or in both directions to fill the entire region. 
        /// </summary>
        public TFlipMode Flip { get { return FFlip; } }

        /// <summary>
        /// Specifies an extra horizontal offest after alignment.
        /// </summary>
        public TDrawingCoordinate Tx { get { return FTx; } }

        /// <summary>
        /// Specifies an extra vertical offest after alignment.
        /// </summary>
        public TDrawingCoordinate Ty { get { return FTy; } }

        /// <summary>
        /// Indicates the amount ot horizontally scale the source rectangle.
        /// </summary>
        public double ScaleX { get { return FScaleX; } }

        /// <summary>
        /// Indicates the amount ot vertically scale the source rectangle.
        /// </summary>
        public double ScaleY { get { return FScaleY; } }

    #region Compare
        /// <summary>
        /// Returns true if this instance has the same data as the object obj.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            TBlipFillTile o2 = obj as TBlipFillTile;
            if (o2 == null) return false;
            return base.Equals(obj) &&
                        FAlign == o2.FAlign &&
                        FFlip == o2.FFlip &&

                        FTx == o2.FTx &&
                        FTy == o2.FTy &&
                        FScaleX == o2.FScaleX &&
                        FScaleY == o2.FScaleY;
        }

        /// <summary>
        /// Returns -1 if obj is bigger than this, 0 if both objects are the same, and 1 if obj is smaller than this.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override int CompareTo(object obj)
        {
            TBlipFillTile s2 = obj as TBlipFillTile;
            if (s2 == null) return -1;

            int r = FAlign.CompareTo(s2.FAlign); if (r != 0) return r;
            r = FFlip.CompareTo(s2.FFlip); if (r != 0) return r;
            r = FTx.CompareTo(s2.FTx); if (r != 0) return r;
            r = FTy.CompareTo(s2.FTy); if (r != 0) return r;
            r = FScaleX.CompareTo(s2.FScaleX); if (r != 0) return r;
            r = FScaleY.CompareTo(s2.FScaleY); if (r != 0) return r;

            return base.CompareTo(obj);
        }

        /// <summary>
        /// Returns the hashcode for this object
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHash(base.GetHashCode(), FAlign.GetHashCode(), FFlip.GetHashCode(), FTx.GetHashCode(), FTy.GetHashCode(),
                                    FScaleX.GetHashCode(), FScaleY.GetHashCode());
        }
        #endregion
    }


    /// <summary>
    /// Shapes are filled with an image.
    /// </summary>
    public class TBlipFill : TFillStyle
    {
        private readonly int FDpi;
        private readonly bool FRotateWithShape;
        private readonly TBlip FBlip;
        private readonly TDrawingRelativeRect? FSourceRect;
        private readonly TBlipFillMode FFillMode;

        /// <summary>
        /// Creates a new Image fill.
        /// </summary>
        internal TBlipFill(int aDpi, bool aRotateWithShape, TBlip aBlip, TDrawingRelativeRect? aSourceRect, TBlipFillMode aFillMode)
            : base(TFillStyleType.Blip)
        {
            FDpi = aDpi;
            FRotateWithShape = aRotateWithShape;

            if (aBlip != null) FBlip = aBlip.Clone(); else FBlip = null;
            FSourceRect = aSourceRect;
            FFillMode = aFillMode;
        }

        /// <summary>
        /// Specifies the DPI (dots per inch) used to calculate the size of the blip. If not present or zero, the DPI in the blip is used. 
        /// </summary>
        public int Dpi { get { return FDpi; } }

        /// <summary>
        /// Specifies that the fill should rotate with the shape.
        /// </summary>
        public bool RotateWithShape { get { return FRotateWithShape; } }

        /// <summary>
        /// Picture and properties used in the Blip fill.
        /// </summary>
        public TBlip Blip { get { return FBlip; } }

        /// <summary>
        /// This element specifies the portion of the blip used for the fill. 
        /// Each edge of the source rectangle is defined by a percentage offset from the corresponding edge of the 
        /// bounding box.  A positive percentage specifies an inset, while a negative percentage specifies an outset. 
        /// For example, a left offset of 25% specifies that the left edge of the source rectangle is located to the right of the 
        /// bounding box's left edge by an amount equal to 25% of the bounding box's width.
        /// </summary>
        public TDrawingRelativeRect? SourceRect { get { return FSourceRect; } }

        /// <summary>
        /// Specifies how the blip will be applied to the fill, either by stretching it to cover all the surface, of by tiling it.
        /// </summary>
        public TBlipFillMode FillMode { get { return FFillMode; } }

    #region Compare
        /// <summary>
        /// Returns true if this instance has the same data as the object obj.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            TBlipFill o2 = obj as TBlipFill;
            if (o2 == null) return false;
            return base.Equals(obj) &&
                    FDpi == o2.FDpi &&
                    FRotateWithShape == o2.FRotateWithShape &&
                    object.Equals(FBlip, o2.FBlip) &&
                    FSourceRect == o2.FSourceRect &&
                    object.Equals(FFillMode, o2.FFillMode);
        }

        /// <summary>
        /// Returns -1 if obj is bigger than this, 0 if both objects are the same, and 1 if obj is smaller than this.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override int CompareTo(object obj)
        {
            TBlipFill s2 = obj as TBlipFill;
            if (s2 == null) return -1;

            int r = FDpi.CompareTo(s2.FDpi);
            if (r != 0) return r;

            r = FRotateWithShape.CompareTo(s2.FRotateWithShape);
            if (r != 0) return r;

            r = FlxUtils.CompareObjects(FBlip, s2.FBlip);
            if (r != 0) return r;

            if (!FSourceRect.HasValue)
            {
                if (s2.FSourceRect.HasValue) return -1;
            }
            else
            {
                if (!s2.FSourceRect.HasValue) return 1;

                r = FSourceRect.Value.CompareTo(s2.FSourceRect.Value);
                if (r != 0) return r;
            }

            r = FlxUtils.CompareObjects(FFillMode, s2.FFillMode);
            if (r != 0) return r;

            return base.CompareTo(obj);
        }

        /// <summary>
        /// Returns the hashcode for this object
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHashObj(base.GetHashCode(), FDpi, FRotateWithShape, FBlip, FSourceRect, FFillMode);
        }        
        #endregion
    }
    #endregion

    #region Pattern Fill
    /// <summary>
    /// Shapes are filled with a texture.
    /// </summary>
    public class TPatternFill : TFillStyle
    {
        readonly TDrawingColor FFgColor;
        readonly TDrawingColor FBgColor;
        readonly TDrawingPattern FPattern;

        /// <summary>
        /// Creates a new Pattern fill.
        /// </summary>
        public TPatternFill(TDrawingColor aFgColor, TDrawingColor aBgColor, TDrawingPattern aPattern)
            : base(TFillStyleType.Pattern)
        {
            FFgColor = aFgColor;
            FBgColor = aBgColor;
            FPattern = aPattern;
        }

        /// <summary>
        ///  Foreground color of a pattern fill.
        /// </summary>
        public TDrawingColor FgColor { get { return FFgColor; } }

        /// <summary>
        /// Background color of a Pattern fill. 
        /// </summary>
        public TDrawingColor BgColor { get { return FBgColor; } }

        /// <summary>
        /// Type of hatching.
        /// </summary>
        public TDrawingPattern Pattern { get { return FPattern; } }

        internal override TFillStyle ReplacePhClr(TDrawingColor basicColor)
        {
            return new TPatternFill(FFgColor.ReplacePhClr(basicColor), FBgColor.ReplacePhClr(basicColor), FPattern);
        }

        /// <summary>
        /// Returns true if this instance has the same data as the object obj.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            TPatternFill o2 = obj as TPatternFill;
            if (o2 == null) return false;
            return base.Equals(obj) &&
                FFgColor == o2.FFgColor &&
                FBgColor == o2.FBgColor &&
                FPattern == o2.FPattern;
        }

        /// <summary>
        /// Returns -1 if obj is bigger than this, 0 if both objects are the same, and 1 if obj is smaller than this.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override int CompareTo(object obj)
        {
            TPatternFill s2 = obj as TPatternFill;
            if (s2 == null) return -1;

            int r = FFgColor.CompareTo(s2.FFgColor);
            if (r != 0) return r;

            r = FBgColor.CompareTo(s2.FBgColor);
            if (r != 0) return r;

            r = FPattern.CompareTo(s2.FPattern);
            if (r != 0) return r;

            return base.CompareTo(obj);
        }

        /// <summary>
        /// Returns the hashcode for this object
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHash(base.GetHashCode(), FFgColor.GetHashCode(), FBgColor.GetHashCode(), FPattern.GetHashCode());
        }

    }
    #endregion

    #region Group Fill
    /// <summary>
    /// Shape is part of a group and filled with its parent fill style.
    /// </summary>
    public class TGroupFill : TFillStyle
    {
        /// <summary>
        /// Creates a new TGroupFill instance.
        /// </summary>
        public TGroupFill()
            : base(TFillStyleType.Group)
        {
        }

        /// <summary>
        /// Returns true if this instance has the same data as the object obj.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            TGroupFill o2 = obj as TGroupFill;
            if (o2 == null) return false;
            return base.Equals(obj);
        }

        /// <summary>
        /// Returns -1 if obj is bigger than this, 0 if both objects are the same, and 1 if obj is smaller than this.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override int CompareTo(object obj)
        {
            TGroupFill s2 = obj as TGroupFill;
            if (s2 == null) return -1;

            return base.CompareTo(obj);
        }


        /// <summary>
        /// Returns the hashcode for this object
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return base.GetHashCode();
        }
    }
    #endregion
    #endregion

    /// <summary>
    /// Represents the line style characteristics.
    /// </summary>
    public class TLineStyleList
    {
#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
        private List<TLineStyle> FLineCollection;

        /// <summary>
        /// Creates a new instance.
        /// </summary>
        public TLineStyleList()
        {
            FLineCollection = new List<TLineStyle>();
        }

        /// <summary>
        /// Returns a deep copy of the object.
        /// </summary>
        /// <returns></returns>
        public TLineStyleList Clone()
        {
            TLineStyleList Result = new TLineStyleList();
            for (int i = 0; i < FLineCollection.Count; i++)
            {
                if (FLineCollection[i] == null) Result.FLineCollection.Add(null); else Result.FLineCollection.Add(FLineCollection[i].Clone());
            }
            return Result;
        }

        /// <summary>
        /// Returns the line color that results from applying the formatting type to a color.
        /// </summary>
        /// <param name="index">Index to the collection.</param>
        /// <param name="basicColor">Color that will be used as a base to calculate the fill style.</param>
        /// <returns></returns>
        public TFillStyle GetRealFillStyle(TFormattingType index, TDrawingColor basicColor)
        {
            return FLineCollection[(int)index].ReplacePhClr(basicColor);
        }

                /// <summary>
        /// Count of line styles.
        /// </summary>
        public int Count { get { return FLineCollection.Count; } }

        /// <summary>
        /// Adds a new LineStyle to the collection. Line styles must be added in order, first is "Subtle", second is "Moderate", third is "Intense"
        /// and there could be new definitions in newer versions of Excel.
        /// </summary>
        /// <param name="aLineStyle">Line style to add.</param>
        public void Add(TLineStyle aLineStyle)
        {
            FLineCollection.Add(aLineStyle);
        }

        /// <summary>
        /// Returns the line style for a given formatting type. Currently Excel defines only 3 formatting types, but more could be added in the future.
        /// If you need to access a formatting type that is not defined in the <see cref="TFormattingType"/> enumeration, just cast an integer to TFormattingType.
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public TLineStyle this[TFormattingType index]
        {
            get
            {
                return FLineCollection[(int)index];
            }
            set
            {
                FLineCollection[(int)index] = value;
            }
        }

        internal static TLineStyle GetDefaultLineStyle(int i)
        {
            switch (i)
            {
                case 0:
                    {
                        TFillStyle fs = new TSolidFill(
                            TDrawingColor.AddTransform(TDrawingColor.FromTheme(TThemeColor.None),
                            new TColorTransform[]{new TColorTransform(TColorTransformType.Shade, 0.95000),
                       new TColorTransform(TColorTransformType.SatMod, 1.05000)}));

                        return new
                             TLineStyle(fs, TLineStyle.DefaultWidth, TPenAlignment.Center, TLineCap.Flat, TCompoundLineType.Single, TLineDashing.Solid, null, null, null);
                    }
                case 1:
                    {
                        TFillStyle fs = new TSolidFill(TDrawingColor.FromTheme(TThemeColor.None));

                        return new
                             TLineStyle(fs, 25400, TPenAlignment.Center, TLineCap.Flat, TCompoundLineType.Single, TLineDashing.Solid, null, null, null);
                    }

                case 2:
                    {
                        TFillStyle fs = new TSolidFill(TDrawingColor.FromTheme(TThemeColor.None));

                        return new
                             TLineStyle(fs, 38100, TPenAlignment.Center, TLineCap.Flat, TCompoundLineType.Single, TLineDashing.Solid, null, null, null);
                    }
                default:
                    return new TLineStyle(new TNoFill(), 0);
            }
        }        
#else
        /// <summary>
        /// Creates a deep copy of this object.
        /// </summary>
        /// <returns></returns>
        public TLineStyleList Clone()
        {
            return this;
        }
#endif

    }

    /// <summary>
    /// Definition for a Drawing line style. 
    /// </summary>
    public class TLineStyle: IComparable
    {
        readonly TFillStyle FFill;
        readonly int? FWidth;
        readonly TPenAlignment? FPenAlign;
        readonly TLineCap? FLineCap;
        readonly TCompoundLineType? FCompoundLineType;
        readonly TLineDashing? FDashing;
        readonly TLineJoin? FJoin;
        readonly TLineArrow? FHeadArrow;
        readonly TLineArrow? FTailArrow;
        readonly internal string[] FExtra; //This shouldn't be public, as it contains xlsx specific data.

        /// <summary>
        /// Creates a simple line with most values by default, including default width.
        /// </summary>
        /// <param name="aFill"></param>
        public TLineStyle(TFillStyle aFill)
            : this(aFill, null, null, null, null, null, null, null, null, null)
        {
        }
        /// <summary>
        /// Creates a simple line with most values by default.
        /// </summary>
        /// <param name="aFill"></param>
        /// <param name="aWidth"></param>
        public TLineStyle(TFillStyle aFill, int? aWidth)
            : this(aFill, aWidth, null, null, null, null, null, null, null)
        {
        }

        /// <summary>
        /// Initializes the line style.
        /// </summary>
        public TLineStyle(TFillStyle aFill, int? aWidth, TPenAlignment? aPenAlign, 
            TLineCap? aLineCap, TCompoundLineType? aCompoundLineType, TLineDashing? aDashing, TLineJoin? aJoin, TLineArrow? aHeadArrow, TLineArrow? aTailArrow)
        {
            FFill = aFill; //immutable, no need to clone.
            FWidth = aWidth;
            FPenAlign = aPenAlign;
            FLineCap = aLineCap;
            FCompoundLineType = aCompoundLineType;
            FDashing = aDashing;
            FJoin = aJoin;
            FHeadArrow = aHeadArrow;
            FTailArrow = aTailArrow;
        }

        internal TLineStyle(TFillStyle aFill, int? aWidth, TPenAlignment? aPenAlign, TLineCap? aLineCap, TCompoundLineType? aCompoundLineType,
            TLineDashing? aDashing, TLineJoin? aJoin, TLineArrow? aHeadArrow, TLineArrow? aTailArrow, string[] aExtra)
        : this(aFill, aWidth, aPenAlign, aLineCap, aCompoundLineType, aDashing, aJoin, aHeadArrow, aTailArrow)
        {
            FExtra = aExtra;
        }

        /// <summary>
        /// Line color and/or fill.
        /// </summary>
        public TFillStyle Fill { get { return FFill; } }

        /// <summary>
        /// Width of the line in EMUs (1 pt = 12700 EMUs). If null, width of the theme will be used.
        /// </summary>
        public int? Width { get { return FWidth; } }

        /// <summary>
        /// Specifies the alignment to be used for the underline stroke. If null, default from the theme will be used. 
        /// </summary>
        public TPenAlignment? PenAlign { get { return FPenAlign; } }

        /// <summary>
        /// How the line ends. If null, default from the theme will be used. 
        /// </summary>
        public TLineCap? LineCap { get { return FLineCap; } }

        /// <summary>
        /// Compound line style. If null, default from the theme will be used. 
        /// </summary>
        public TCompoundLineType? CompoundLineType{ get { return FCompoundLineType; } }

        /// <summary>
        /// Line dashing. If null, default from the theme will be used. 
        /// </summary>
        public TLineDashing? Dashing { get { return FDashing; } }

        /// <summary>
        /// How the line joins with the next. If null, default from the theme will be used. 
        /// </summary>
        public TLineJoin? Join { get { return FJoin; } }

        /// <summary>
        /// Head arrow if it has one. If null, default from the theme will be used. 
        /// </summary>
        public TLineArrow? HeadArrow { get { return FHeadArrow; } }

        /// <summary>
        /// Tail arrow if it has one. If null, default from the theme will be used. 
        /// </summary>
        public TLineArrow? TailArrow { get { return FTailArrow; } }

        /// <summary>
        /// Default line width.
        /// </summary>
        public const int DefaultWidth = 9525;

        /// <summary>
        /// Returns a deep copy of the fill style.
        /// </summary>
        /// <returns></returns>
        public virtual TLineStyle Clone()
        {
            return (TLineStyle)MemberwiseClone();
        }

        /// <summary>
        /// Returns true if this instance has the same data as the object obj.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            TLineStyle o2 = obj as TLineStyle;
            if (o2 == null) return false;
            return
                FWidth == o2.Width &&
                FPenAlign == o2.PenAlign &&
                FLineCap == o2.LineCap &&
                FCompoundLineType == o2.CompoundLineType &&
                FDashing == o2.FDashing &&
                FJoin == o2.FJoin &&
                HeadArrow == o2.HeadArrow &&
                TailArrow == o2.TailArrow &&
                object.Equals(FFill, o2.FFill)
                && FlxUtils.CompareArray(FExtra, o2.FExtra) == 0;
        }

        private bool SameExtra(List<string> Extra1, List<string> Extra2)
        {
            if (Extra1 == null) return Extra2 == null;
            if (Extra1.Count != Extra2.Count) return false;
            for (int i = 0; i < Extra1.Count; i++)
            {
                if (Extra1[i] != Extra2[i]) return false;
            }

            return true;
        }

        /// <summary>
        /// Returns the hashcode for this object
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHashObj(FWidth, FPenAlign, FLineCap, FCompoundLineType, FDashing, FJoin, FHeadArrow, FTailArrow, FFill);
        }

        internal TFillStyle ReplacePhClr(TDrawingColor basicColor)
        {
            if (FFill == null) return null;
            return FFill.ReplacePhClr(basicColor);
        }

        #region IComparable Members
        /// <summary>
        /// Returns -1, 0 or 1 depending if the object is smaller, equal or bigger than the other.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public int CompareTo(object obj)
        {
            TLineStyle o2 = obj as TLineStyle;
            if (object.ReferenceEquals(o2, null)) return -1;
            int r;
            r = FlxUtils.CompareObjects(FWidth, o2.FWidth); if (r != 0) return r;
            r = FlxUtils.CompareObjects(FPenAlign, o2.FPenAlign); if (r != 0) return r;
            r = FlxUtils.CompareObjects(FLineCap, o2.FLineCap); if (r != 0) return r;
            r = FlxUtils.CompareObjects(FCompoundLineType, o2.FCompoundLineType); if (r != 0) return r;
            r = FlxUtils.CompareObjects(FDashing, o2.FDashing); if (r != 0) return r;
            r = FlxUtils.CompareObjects(FJoin, o2.FJoin); if (r != 0) return r;
            r = FlxUtils.CompareObjects(HeadArrow, o2.HeadArrow); if (r != 0) return r;
            r = FlxUtils.CompareObjects(TailArrow, o2.TailArrow); if (r != 0) return r;
            r = FlxUtils.CompareObjects(FFill, o2.FFill); if (r != 0) return r;
            r = FlxUtils.CompareArray(FExtra, o2.FExtra); if (r != 0) return r;

            return 0;
        }

        #endregion
    }

    /// <summary>
    /// This class holds the effects that are applied to a drawing. At this moment its members are not public.
    /// </summary>
    public class TEffectProperties: IComparable
    {
        internal readonly string xml;

        internal TEffectProperties(string aXml)
        {
            xml = aXml;
        }

        #region IComparable Members

        /// <summary>
        /// Returns -1, 0 or 1 depending if this object is smaller or bigger than the other.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public int CompareTo(object obj)
        {
            TEffectProperties o2 = obj as TEffectProperties;
            if (object.ReferenceEquals(o2, null)) return -1;

            return String.Compare(xml, o2.xml);
        }

        /// <summary>
        /// Returns the hashcode of the object.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHashObj(xml);
        }

        /// <summary>
        /// Returns true if both instances have the same data.
        /// </summary>
        /// <param name="obj">Object to compare.</param>
        /// <returns>True if both strings are the same.</returns>
        public override bool Equals(object obj)
        {
            return CompareTo(obj) == 0;
        }

        /// <summary>
        /// Returns true if both objects are equal.
        /// </summary>
        /// <param name="o1">First object to compare.</param>
        /// <param name="o2">Second object to compare.</param>
        /// <returns></returns>
        public static bool operator ==(TEffectProperties o1, TEffectProperties o2)
        {
            return Object.Equals(o1, o2);
        }

        /// <summary>
        /// Returns true if both objects do not have the same value.
        /// </summary>
        /// <param name="o1">First objects to compare.</param>
        /// <param name="o2">Second objects to compare.</param>
        /// <returns></returns>
        public static bool operator !=(TEffectProperties o1, TEffectProperties o2)
        {
            return !(Object.Equals(o1, o2));
        }

        #endregion

    }

    /// <summary>
    /// Represents the effect style characteristics.
    /// </summary>
    public class TEffectStyleList
    {
#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
        private List<String> FEffectCollection;
        internal string Xml;

        /// <summary>
        /// Creates a new instance.
        /// </summary>
        public TEffectStyleList()
        {
            FEffectCollection = new List<string>();
        }

        /// <summary>
        /// Returns a deep copy of the object.
        /// </summary>
        /// <returns></returns>
        public TEffectStyleList Clone()
        {
            TEffectStyleList Result = (TEffectStyleList)MemberwiseClone();
            return Result;
        }
#else
        /// <summary>
        /// Creates a deep copy of this object.
        /// </summary>
        /// <returns></returns>
        public TEffectStyleList Clone()
        {
            return this;
        }
#endif
    }


    #endregion

    #region Theme Elements
    /// <summary>
    /// Definitions of the elements in a theme (colors, fonts, formats). This is the main part of a theme.
    /// </summary>
    public class TThemeElements
    {
        private TThemeColorScheme FColorScheme;
        private TThemeFontScheme FFontScheme;
        private TThemeFormatScheme FFormatScheme;

        /// <summary>
        /// Creates a new theme elements.
        /// </summary>
        public TThemeElements()
        {
            FColorScheme = new TThemeColorScheme();
            FFontScheme = new TThemeFontScheme();
            FFormatScheme = new TThemeFormatScheme();
        }

        /// <summary>
        /// Returns true if the elements in this theme are the default ones in Office 2007.
        /// </summary>
        public bool IsStandard
        {
            get
            {
                return FColorScheme.IsStandard
                && FFontScheme.IsStandard
                && FFormatScheme.IsStandard;
            }
        }

        /// <summary>
        /// Color Scheme in the theme.
        /// </summary>
        public TThemeColorScheme ColorScheme { get { return FColorScheme; } }

        /// <summary>
        /// Font Scheme in the theme.
        /// </summary>
        public TThemeFontScheme FontScheme { get { return FFontScheme; } }

        /// <summary>
        /// Format Scheme (Effects). This won't affect cells in the spreadsheet, but can affect drawings.
        /// </summary>
        public TThemeFormatScheme FormatScheme { get { return FFormatScheme; } }

        /// <summary>
        /// Returns a deep copy of this object.
        /// </summary>
        /// <returns></returns>
        public TThemeElements Clone()
        {
            TThemeElements Result = new TThemeElements();
            Result.FColorScheme = FColorScheme.Clone();
            Result.FFontScheme = FFontScheme.Clone();
            Result.FFormatScheme = FFormatScheme.Clone();

            return Result;
        }
    }
    #endregion

    #region Theme
    /// <summary>
    /// Contains a complete definition for an Office Theme.
    /// </summary>
    public class TTheme
    {
        private string FName;
        private int FThemeVersion;
        private TThemeElements FElements;

        /// <summary>
        /// Creates a new standard "Office" theme.
        /// </summary>
        public TTheme()
        {
            FElements = new TThemeElements();
            FThemeVersion = 124226;
            Name = "Office Theme";
        }

        private TTheme(TThemeElements NewElements)
        {
            FElements = NewElements;
        }

        /// <summary>
        /// Name of the theme definition. This will be shown in Excel UI.
        /// </summary>
        public string Name { get { return FName; } set { FName = value; } }

        /// <summary>
        /// Elements of the theme.
        /// </summary>
        public TThemeElements Elements { get { return FElements; } }

        /// <summary>
        /// Excel version that saved this theme. For Excel 2007 this value is 124226. For themes new to Excel 2010, this value might be 144315
        /// </summary>
        public int ThemeVersion { get { return FThemeVersion; } set { FThemeVersion = value; } }

        /// <summary>
        /// Returns true if the theme is standard.
        /// </summary>
        public bool IsStandard
        {
            get
            {
                return
                    FThemeVersion == 124226
                    ||
                    (
                      Name == "Office Theme" &&
                      Elements.IsStandard
                    );
            }
        }

        /// <summary>
        /// Returns a deep copy of this object.
        /// </summary>
        /// <returns></returns>
        public TTheme Clone()
        {
            TTheme Result = new TTheme(Elements.Clone());
            return Result;
        }
    }
    #endregion
}
