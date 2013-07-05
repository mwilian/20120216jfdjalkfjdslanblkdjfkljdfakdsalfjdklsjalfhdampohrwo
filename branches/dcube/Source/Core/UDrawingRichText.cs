using System;
using System.Collections.Generic;
using System.Text;

namespace FlexCel.Core
{
    #region Enums
    /// <summary>
    /// Specifies the alignment types for text in a drawing.
    /// </summary>
    public enum TDrawingAlignment
    {
        /// <summary>
        /// Align text to the left margin. 
        /// </summary>
        Left,

        /// <summary>
        /// Align text in the center. 
        /// </summary>
        Center,

        /// <summary>
        /// Align text to the right margin. 
        /// </summary>
        Right,

        /// <summary>
        /// Align text so that it is justified across the whole line. It is smart in the sense that it does not justify sentences which are short. 
        /// </summary>
        Justified,

        /// <summary>
        /// Aligns the text with an adjusted kashida length for Arabic text.
        /// </summary>
        JustLow,

        /// <summary>
        /// Distributes the text words across an entire text line.
        /// </summary>
        Distributed,

        /// <summary>
        /// Distributes Thai text specially, because each character is treated as a word. 
        /// </summary>
        ThaiDist,
    }

    /// <summary>
    /// Different types of font alignment.
    /// </summary>
    public enum TDrawingFontAlign
    {
        /// <summary>
        /// When the text flow is horizontal or simple vertical same as fontBaseline but for other vertical modes same as fontCenter. 
        /// </summary>
        Automatic,

        /// <summary>
        /// The letters are anchored to the top baseline of a single line. 
        /// </summary>
        Top,

        /// <summary>
        /// The letters are anchored between the two baselines of a single line. 
        /// </summary>
        Center,

        /// <summary>
        /// The letters are anchored to the bottom baseline of a single line. 
        /// </summary>
        BaseLine,

        /// <summary>
        /// The letters are anchored to the very bottom of a single line. This is different than the bottom baseline because of letters such as "g," "q," "y," etc. 
        /// </summary>
        Bottom,
    }

    /// <summary>
    /// Possible underline types in a drawing.
    /// </summary>
    public enum TDrawingUnderlineStyle
    {
        /// <summary>
        /// No underline.
        /// </summary>
        None,

        /// <summary>
        /// Underline just the words and not the spaces between them. 
        /// </summary>
        Words,

        /// <summary>
        /// Underline the text with a single line of normal thickness. 
        /// </summary>
        Single,

        /// <summary>
        /// Underline the text with two lines of normal thickness.
        /// </summary>
        Double,

        /// <summary>
        /// Underline the text with a single, thick line.
        /// </summary>
        Heavy,

        /// <summary>
        /// Underline the text with a single, dotted line of normal thickness. 
        /// </summary>
        Dotted,

        /// <summary>
        /// Underline the text with a single, thick dotted line. 
        /// </summary>
        DottedHeavy,

        /// <summary>
        /// Underline the text with a single, dashed line of normal thickness. 
        /// </summary>
        Dash,

        /// <summary>
        /// Underline the text with a single, thick dashed line. 
        /// </summary>
        DashHeavy,

        /// <summary>
        /// Underline the text with a single line of normal thickness consisting of long dashes. 
        /// </summary>
        DashLong,

        /// <summary>
        /// Underline the text with a single thick line consisting of long dashes. 
        /// </summary>
        DashLongHeavy,

        /// <summary>
        /// Underline the text with a single line consisting of alternating dots and dashes of normal thcikness. 
        /// </summary>
        DotDash,

        /// <summary>
        /// Underline the text with a single thick line consisting of alternating dots and dashes. 
        /// </summary>
        DotDashHeavy,

        /// <summary>
        /// Underline the text with a single line of normal thickness consisting of repeating two dots and dashes.
        /// </summary>
        DotDotDash,

        /// <summary>
        /// Underline the text with a single thick line consisting of repeating two dots and dashes.
        /// </summary>
        DotDotDashHeavy,

        /// <summary>
        /// Underline the text with a single wavy line of normal thickness.
        /// </summary>
        Wavy,

        /// <summary>
        /// Underline the text with a single, thick wavy line.
        /// </summary>
        WavyHeavy,

        /// <summary>
        /// Underline the text with two wavy lines of normal thickness. 
        /// </summary>
        WavyDouble,
    }

    /// <summary>
    /// How the text is striked out.
    /// </summary>
    public enum TDrawingTextStrike
    {
        /// <summary>
        /// No strike is applied to the text. 
        /// </summary>
        None,

        /// <summary>
        /// A single strike is applied to the text. 
        /// </summary>
        Single,

        /// <summary>
        /// A double strike is applied to the text. 
        /// </summary>
        Double,
    }

    /// <summary>
    /// how text is capitalized when rendered.
    /// </summary>
    public enum TDrawingTextCapitalization
    {
        /// <summary>
        /// No capitalization.
        /// </summary>
        None,

        /// <summary>
        /// Apply small caps to the text. All letters are converted to lower case.
        /// </summary>
        Small,

        /// <summary>
        /// Apply all caps on the text. All lower case letters are 
        /// converted to upper case even though they are stored 
        /// differently in the backing store. 
        /// </summary>
        All
    }
    #endregion

    #region DrawingRichString
    /// <summary>
    /// A rich string used in drawings. It is similar to <see cref=" TRichString"/> but it has more 
    /// properties like for example wordart properties. Similar to a string, this class is immutable.
    /// </summary>
    public sealed class TDrawingRichString : IComparable
    {
        #region Variables
        private readonly TDrawingTextParagraph[] Paragraphs;
        private string CachedValue; //this class is immutable, this doesn't change.
        #endregion

        #region Constructors
        /// <summary>
        /// Creates a new TDrawingRichString with a null value.
        /// </summary>
        public TDrawingRichString()
            : this((TDrawingTextParagraph[])null)
        {
        }

        /// <summary>
        /// Creates a new TDrawingRichString with no formatting.
        /// </summary>
        /// <param name="s">String with the data.</param>
        public TDrawingRichString(string s)
            : this(new TDrawingTextParagraph[] { new TDrawingTextParagraph(s, TDrawingParagraphProperties.Empty, TDrawingTextProperties.Empty) })
        {
        }

        /// <summary>
        /// Creates a new TDrawingRichString from an array of paragraphs.
        /// </summary>
        /// <param name="aParagraphs">Array of RTF runs with the data and formatting for the string.</param>
        public TDrawingRichString(TDrawingTextParagraph[] aParagraphs)
        {
            if (aParagraphs == null) Paragraphs = new TDrawingTextParagraph[0];
            else
            {
                Paragraphs = new TDrawingTextParagraph[aParagraphs.Length];
                Array.Copy(aParagraphs, Paragraphs, aParagraphs.Length);
            }
        }
        #endregion

        #region Properties
        /// <summary>
        /// Text of the string without formatting. Might be null.
        /// </summary>
        public string Value
        {
            get
            {
                if (CachedValue != null) return CachedValue;
                if (Paragraphs.Length == 0) return null;
                StringBuilder sb = new StringBuilder();
                bool First = true;
                foreach (TDrawingTextParagraph paragraph in Paragraphs)
                {
                    if (!First) sb.Append((char)10);
                    First = false;
                    sb.Append(paragraph.Text);
                }

                CachedValue = sb.ToString();
                return CachedValue;
            }
        }

        /// <summary>
        /// A paragraph of the text.
        /// </summary>
        /// <param name="index">Index on the list. 0 based.</param>
        /// <returns>The paragraph.</returns>
        public TDrawingTextParagraph Paragraph(int index)
        {
            return Paragraphs[index];
        }

        /// <summary>
        /// The count of Paragraphs in this string.
        /// </summary>
        public int ParagraphCount
        {
            get
            {
                return Paragraphs.Length;
            }
        }

        /// <summary>
        /// Length of the DrawingRichString.
        /// </summary>
        public int Length
        {
            get
            {
                if (Value == null) return 0; else return Value.Length;
            }
        }
        #endregion

        #region String Utilities
        /// <summary>
        /// Retrieves a substring from this instance. The substring starts at a specified character position and has a specified length.
        /// </summary>
        /// <param name="index">Start of the substring (0 based)</param>
        /// <param name="count">Number of characters to copy.</param>
        public TDrawingRichString Substring(int index, int count)
        {

            if (index < 0 || index >= Length) FlxMessages.ThrowException(FlxErr.ErrInvalidValue, "index", index, 0, Length - 1);
            if (count < 0 || index + count > Length) FlxMessages.ThrowException(FlxErr.ErrInvalidValue, "count", count, 0, Length - index);
            List<TDrawingTextParagraph> NewParagraphs = new List<TDrawingTextParagraph>();
            int charPos = -1;
            foreach (TDrawingTextParagraph item in Paragraphs)
            {
                charPos++; //accounts for the #10 between breaks;

                if (string.IsNullOrEmpty(item.Text)) continue;

                int charEndPos = charPos + item.Text.Length;
                if (charEndPos < index)
                {
                    charPos = charEndPos;
                    continue;
                }

                int s1 = index - charPos;
                int s2 = index + count;
                bool FullParagraph = true;
                if (s1 < 0)
                {
                    s1 = 0;
                    FullParagraph = false;
                }

                if (s2 >= charEndPos)
                {
                    s2 = charEndPos;
                }
                else FullParagraph = false;
                if (FullParagraph) NewParagraphs.Add(item);
                else NewParagraphs.Add(item.Substring(s1, s2 - s1 + 1));
                charPos = charEndPos;
                if (charPos > index + count) break;

            }
            return new TDrawingRichString(NewParagraphs.ToArray());
        }

        /// <summary>
        /// Retrieves a substring from this instance. The substring starts at a specified character position and ends at the end of the string.
        /// </summary>
        /// <param name="index">Start of the substring (0 based)</param>
        public TDrawingRichString Substring(int index)
        {
            return Substring(index, Length - index);
        }

        /// <summary>
        /// Concatenates two TDrawingRichString objects.
        /// </summary>
        /// <param name="s1">First string to concatenate.</param>
        /// <param name="s2">Second string to concatenate.</param>
        /// <returns>The concatenated string.</returns>
        public static TDrawingRichString operator +(TDrawingRichString s1, TDrawingRichString s2)
        {
            if (s1 == null || s1.Value == null) return s2;
            if (s2 == null || s2.Value == null) return s1;

            TDrawingTextParagraph[] aParagraphs = new TDrawingTextParagraph[s1.ParagraphCount + s2.ParagraphCount];
            for (int i = 0; i < s1.ParagraphCount; i++)
            {
                aParagraphs[i] = s1.Paragraph(i);
            }

            for (int i = 0; i < s2.ParagraphCount; i++)
            {
                aParagraphs[i + s1.ParagraphCount] = s2.Paragraph(i);
            }

            return new TDrawingRichString(aParagraphs);
        }

        /// <summary>
        /// Adds two richstrings together. If using C#, you can just use the overloaded "+" operator to contactenate rich strings.
        /// </summary>
        /// <param name="s1"></param>
        /// <returns></returns>
        public TDrawingRichString Add(TDrawingRichString s1)
        {
            return this + s1;
        }

        /// <summary>
        /// Trims all the whitespace at the beginning and end of the string.
        /// </summary>
        /// <returns>The trimmed string.</returns>
        public TDrawingRichString Trim()
        {
            if (Value == null) return new TDrawingRichString();
            int i = 0;
            while (i < Value.Length && Value[i] == ' ') i++;
            int k = Value.Length - 1;
            while (k >= 0 && Value[k] == ' ') k--;
            if (i <= k) return Substring(i, k - i + 1); else return new TDrawingRichString();
        }

        /// <summary>
        /// Trims all the whitespace at the end of the string.
        /// </summary>
        /// <returns>The trimmed string.</returns>
        public TDrawingRichString RightTrim()
        {
            if (Value == null) return new TDrawingRichString();
            int i = 0;
            int k = Value.Length - 1;
            while (k >= 0 && Value[k] == ' ') k--;
            if (i <= k) return Substring(i, k - i + 1); else return new TDrawingRichString();
        }

        #endregion

        #region Convert
        /// <summary>
        /// Returns the string without Rich text info.
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            if (Value == null) return String.Empty;
            else
                return Value;
        }

        /// <summary>
        /// Converts this TDrawingRichString into a <see cref="TRichString"/>. Note that the conversion is not perfect 
        /// as a TDrawingRichString has different information from a TRichString.
        /// </summary>
        /// <param name="xls">Excel file with the fonts.</param>
        /// <param name="ShapeThemeFont">Theme font used by default when no formatting is specified. Set it to null to use the default.</param>
        /// <returns></returns>
        public TRichString ToRichString(ExcelFile xls, TShapeFont ShapeThemeFont)
        {
            List<TRTFRun> runs = new List<TRTFRun>();

            StringBuilder sb = new StringBuilder();
            bool First = true;
            foreach (TDrawingTextParagraph paragraph in Paragraphs)
            {
                if (!First) sb.Append((char)10);
                First = false;
                {
                    for (int i = 0; i < paragraph.TextRunCount; i++)
                    {
                        TDrawingTextRun r = paragraph.TextRun(i);
                        runs.Add(GetRTFRun(sb.Length, r.TextProperties, xls, ShapeThemeFont));
                        sb.Append(r.Text);
                    }
                }
            }
            return new TRichString(sb.ToString(), runs, xls);
        }

        private TRTFRun GetRTFRun(int p, TDrawingTextProperties rp, ExcelFile xls, TShapeFont ShapeThemeFont)
        {
            TRTFRun Result = new TRTFRun();
            Result.FirstChar = p;
            TFlxFont NewFont = new TFlxFont();
            
#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
            if (ShapeThemeFont != null)
            {
                TThemeFontScheme fs = xls.GetTheme().Elements.FontScheme;
                TThemeFont ThemeFont = null;
                switch (ShapeThemeFont.ThemeScheme)
                {
                    case TFontScheme.None:
                        ThemeFont = null;
                        break;

                    case TFontScheme.Minor:
                        ThemeFont = fs.MinorFont;
                        break;

                    case TFontScheme.Major:
                        ThemeFont = fs.MajorFont;
                        break;
                }

                if (ThemeFont != null)
                {
                    NewFont.Name = ThemeFont.Latin.Typeface;
                    NewFont.Color = ShapeThemeFont.ThemeColor.ToColor(xls);
                }
            }
#endif

            if (rp.Attributes.Size > 0) NewFont.Size20 = (int)Math.Round(rp.Attributes.Size * (20.0 / 100.0));
            if (rp.Attributes.Bold) NewFont.Style |= TFlxFontStyles.Bold;
            if (rp.Attributes.Italic) NewFont.Style |= TFlxFontStyles.Italic;
            if (rp.Attributes.Strike != TDrawingTextStrike.None) NewFont.Style |= TFlxFontStyles.StrikeOut;

            switch (rp.Attributes.Underline)
            {
                case TDrawingUnderlineStyle.Single: NewFont.Underline = TFlxUnderline.Single; break;
                case TDrawingUnderlineStyle.Double: NewFont.Underline = TFlxUnderline.Double; break;
                default: if (rp.Attributes.Underline != TDrawingUnderlineStyle.None) NewFont.Underline = TFlxUnderline.Single; break;
            }

            GetTypeFace(rp.Symbol, NewFont);
            GetTypeFace(rp.ComplexScript, NewFont);
            GetTypeFace(rp.EastAsian, NewFont);
            GetTypeFace(rp.Latin, NewFont);

            TSolidFill sFill = rp.Fill as TSolidFill;
            if (sFill != null) NewFont.Color = sFill.Color.ToColor(xls);

            Result.FontIndex = xls.AddFont(NewFont);
            return Result;
        }

        private static void GetTypeFace(TThemeTextFont? tf, TFlxFont NewFont)
        {
            if (tf.HasValue) NewFont.Name = tf.Value.Typeface;
        }

        /// <summary>
        /// Returns a new TDrawingRichString from a TRichString. Note that the conversion is not perfect since information in both kind of strings is different.
        /// </summary>
        /// <param name="aValue">String that we want to convert.</param>
        /// <param name="xls">Excel file with the fonts.</param>
        /// <returns></returns>
        public static TDrawingRichString FromRichString(TRichString aValue, ExcelFile xls)
        {
            if (aValue == null || aValue.Value == null) return null;
            string[] ps = aValue.Value.Split((char)10);
            TDrawingTextParagraph[] Paragraphs = new TDrawingTextParagraph[ps.Length];

            int AcumLen = 0;
            int UsedRuns = 1;
            TRTFList RTFRunArray = CalcRTFRunArray(aValue);
            for (int i = 0; i < Paragraphs.Length; i++)
            {
                UsedRuns--; //we can use the last one
                if (UsedRuns < 0) UsedRuns = 0;
                TDrawingTextRun[] runs = GetRuns(ref UsedRuns, AcumLen, ps[i], RTFRunArray, xls);
                AcumLen += ps[i].Length + 1; //add char 10
                Paragraphs[i] = new TDrawingTextParagraph(runs, TDrawingParagraphProperties.Empty, TDrawingTextProperties.Empty);
            }

            return new TDrawingRichString(Paragraphs);
        }

        private static TRTFList CalcRTFRunArray(TRichString aValue)
        {
            TRTFList Result = new TRTFList();
            //Ensure first and last runs
            for (int i = 0; i < aValue.RTFRunCount; i++)
            {
                Result.Add(aValue.RTFRun(i));
            }

            if (Result.Count == 0 || Result[0].FirstChar != 0)
            {
                TRTFRun Run0 = new TRTFRun();
                Run0.FirstChar = 0;
                Run0.FontIndex = 0;
                Result.Insert(0, Run0);
            }

            if (Result[Result.Count - 1].FirstChar < aValue.Length)
            {
                TRTFRun RunN = new TRTFRun();
                RunN.FirstChar = aValue.Length;
                RunN.FontIndex = 0;
                Result.Add(RunN);
            }
            return Result;
        }

        private static TDrawingTextRun[] GetRuns(ref int UsedRuns, int AcumLen, string p, TRTFList RTFRunArray, ExcelFile xls)
        {
            List<TDrawingTextRun> runs = new List<TDrawingTextRun>();
            while (UsedRuns < RTFRunArray.Count - 1)
            {
                TRTFRun r1 = RTFRunArray[UsedRuns];
                TRTFRun r2 = RTFRunArray[UsedRuns + 1];
                if (r1.FirstChar >= AcumLen + p.Length) return runs.ToArray();
                CalcRun(ref UsedRuns, AcumLen, p, xls, runs, ref r1, ref r2);
            }

            return runs.ToArray();
        }

        private static void CalcRun(ref int UsedRuns, int AcumLen, string p, ExcelFile xls, List<TDrawingTextRun> runs, ref TRTFRun r1, ref TRTFRun r2)
        {
            if (r2.FirstChar < AcumLen)
            {
                UsedRuns++;
                return;
            }

            int rs = r1.FirstChar - AcumLen; if (rs < 0) rs = 0;
            int re = r2.FirstChar - AcumLen - 1; if (re >= p.Length) re = p.Length - 1;
            if (re < rs)
            {
                UsedRuns++;
                return;
            }

            runs.Add(new TDrawingTextRun(p.Substring(rs, re - rs + 1), GetTextProps(xls.GetFont(r1.FontIndex), xls)));
            UsedRuns++;
        }

        private static TDrawingTextProperties GetTextProps(TFlxFont NewFont, ExcelFile xls)
        {
            int Size100 = NewFont.Size20 * 5;
            bool Bold = (NewFont.Style & TFlxFontStyles.Bold) != 0;
            bool Italic = (NewFont.Style & TFlxFontStyles.Italic) != 0;
            bool Strikeb = (NewFont.Style & TFlxFontStyles.StrikeOut) != 0;
            TDrawingTextStrike Strike = Strikeb ? TDrawingTextStrike.Single : TDrawingTextStrike.None;

            TDrawingUnderlineStyle Underline = TDrawingUnderlineStyle.None;
            switch (NewFont.Underline)
            {
                case TFlxUnderline.None:
                    break;

                case TFlxUnderline.Single:
                case TFlxUnderline.SingleAccounting:
                    Underline = TDrawingUnderlineStyle.Single;
                    break;

                case TFlxUnderline.Double:
                case TFlxUnderline.DoubleAccounting:
                    Underline = TDrawingUnderlineStyle.Double;
                    break;
            }

            TThemeTextFont Latin = new TThemeTextFont(NewFont.Name, null, TPitchFamily.DEFAULT_PITCH__UNKNOWN_FONT_FAMILY, TFontCharSet.Default);
            TSolidFill sFill = xls == null? null: new TSolidFill(NewFont.Color.ToColor(xls));

            TDrawingTextAttributes def = TDrawingTextAttributes.Empty;
            TDrawingTextAttributes atts = new TDrawingTextAttributes
            (
               def.Kumimoji,
               def.Lang,
               def.AltLang,
               Size100,
               Bold,
               Italic,
               Underline,
               Strike,
               def.Kern,
               def.Capitalization,
               def.Spacing,
               def.NormalizeH,
               def.Baseline,
               def.NoProof,
               def.Dirty,
               def.Err,
               def.SmartTagClean,
               def.SmartTagId,
               def.BookmarkLinkTarget
            );
            return new TDrawingTextProperties(sFill, null, null, null, null, Latin, null, null, null, null, null, false, atts);
        }
        #endregion

        #region Implicits
        /// <summary>
        /// Converts a string to a TDrawingRichString.
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static implicit operator TDrawingRichString(string s)
        {
            return new TDrawingRichString(s);
        }

        /// <summary>
        /// Converts a TDrawingRichString to a string.
        /// </summary>
        /// <param name="r"></param>
        /// <returns></returns>
        public static implicit operator String(TDrawingRichString r)
        {
            return r == null ? null : r.Value;
        }

        #endregion

        #region Compare
        /// <summary>
        /// Returns -1 if obj is bigger than this, 0 if both strings are the same, and 1 if obj is smaller than this.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public int CompareTo(object obj)
        {
            TDrawingRichString s2 = obj as TDrawingRichString;
            if (s2 == null) return -1;


            int r = Paragraphs.Length.CompareTo(s2.Paragraphs.Length);
            if (r != 0) return r;

            for (int i = 0; i < Paragraphs.Length; i++)
            {
                r = Paragraphs[i].CompareTo(s2.Paragraphs[i]);
                if (r != 0) return r;
            }
            return 0;
        }

        /// <summary>
        /// Returns the hashcode of the object.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHashObj(Paragraphs);
        }

        /// <summary>
        /// Returns true if both instances have the same string and formatting.
        /// </summary>
        /// <param name="obj">Object to compare.</param>
        /// <returns>True if both strings are the same.</returns>
        public override bool Equals(object obj)
        {
            return CompareTo(obj) == 0;
        }

        /// <summary>
        /// Returns true if both strings are equal.
        /// </summary>
        /// <param name="o1">First string to compare.</param>
        /// <param name="o2">Second string to compare.</param>
        /// <returns></returns>
        public static bool operator ==(TDrawingRichString o1, TDrawingRichString o2)
        {
            return Object.Equals(o1, o2);
        }

        /// <summary>
        /// Returns true if both strings do not have the same value.
        /// </summary>
        /// <param name="o1">First string to compare.</param>
        /// <param name="o2">Second string to compare.</param>
        /// <returns></returns>
        public static bool operator !=(TDrawingRichString o1, TDrawingRichString o2)
        {
            return !(Object.Equals(o1, o2));
        }

        /// <summary>
        /// Returns true is a string is less than the other.
        /// </summary>
        /// <param name="o1">First string to compare.</param>
        /// <param name="o2">Second string to compare.</param>
        /// <returns></returns>
        public static bool operator <(TDrawingRichString o1, TDrawingRichString o2)
        {
            if (o1 == null) return false;
            return o1.CompareTo(o2) < 0;
        }

        /// <summary>
        /// Returns true is a string is bigger than the other.
        /// </summary>
        /// <param name="o1">First string to compare.</param>
        /// <param name="o2">Second string to compare.</param>
        /// <returns></returns>
        public static bool operator >(TDrawingRichString o1, TDrawingRichString o2)
        {
            if (o1 == null)
            {
                if (o2 == null) return false;
                return true;
            }
            return o1.CompareTo(o2) > 0;
        }

        #endregion
    }
    #endregion

    #region DrawingTextParagraph
    /// <summary>
    /// A paragraph in the text inside a drawing. This struct is immutable.
    /// </summary>
    public struct TDrawingTextParagraph : IComparable
    {
        private readonly TDrawingTextRun[] FRuns;
        private readonly TDrawingParagraphProperties FProperties;
        private readonly TDrawingTextProperties FEndParagraphProperties;
        private string CachedText; //Struct won't change, so it never gets invalid.

        /// <summary>
        /// Creates a new TDrawingTextParagraph instance.
        /// </summary>
        public TDrawingTextParagraph(TDrawingTextRun[] aRuns, TDrawingParagraphProperties aProperties, TDrawingTextProperties aEndParagraphProperties)
        {
            if (aRuns == null) FRuns = new TDrawingTextRun[0];
            else
            {
                FRuns = new TDrawingTextRun[aRuns.Length];
                aRuns.CopyTo(FRuns, 0);
            }

            FProperties = aProperties;
            FEndParagraphProperties = aEndParagraphProperties;
            CachedText = null;
        }

        /// <summary>
        /// Creates a new TDrawingTextParagraph based on a simple string.
        /// </summary>
        public TDrawingTextParagraph(string s, TDrawingParagraphProperties aProperties, TDrawingTextProperties aEndParagraphProperties)
        {
            FRuns = new TDrawingTextRun[1];
            FRuns[0] = new TDrawingTextRun(s, new TDrawingTextProperties());
            FProperties = aProperties;
            FEndParagraphProperties = aEndParagraphProperties;
            CachedText = null;
        }


        /// <summary>
        /// The properties that apply to this paragraph.
        /// </summary>
        public TDrawingParagraphProperties Properties { get { return FProperties; } }

        /// <summary>
        /// Properties that apply to new paragraphs that are added after this one.
        /// </summary>
        public TDrawingTextProperties EndParagraphProperties { get { return FEndParagraphProperties; } }

        /// <summary>
        /// Returns a single text run for the paragraph.
        /// </summary>
        /// <param name="index">Index of the run (0 based)</param>
        /// <returns>Text run for position index.</returns>
        public TDrawingTextRun TextRun(int index)
        {
            return FRuns[index];
        }

        /// <summary>
        /// Returns the number of runs in the paragraph.
        /// </summary>
        public int TextRunCount
        {
            get { return FRuns.Length; }
        }

        /// <summary>
        /// Returns the contents of the paragraph as plain text.
        /// </summary>
        public string Text
        {
            get
            {
                if (CachedText != null) return CachedText;
                StringBuilder sb = new StringBuilder();
                foreach (TDrawingTextRun run in FRuns)
                {
                    sb.Append(run.Text);
                }

                CachedText = sb.ToString();
                return CachedText;
            }
        }

        #region String Utilities
        /// <summary>
        /// Retrieves a substring from this instance. The substring starts at a specified character position and has a specified length.
        /// </summary>
        /// <param name="index">Start of the substring (0 based)</param>
        /// <param name="count">Number of characters to copy.</param>
        public TDrawingTextParagraph Substring(int index, int count)
        {

            if (index < 0 || index >= Text.Length) FlxMessages.ThrowException(FlxErr.ErrInvalidValue, "index", index, 0, Text.Length - 1);
            if (count < 0 || index + count > Text.Length) FlxMessages.ThrowException(FlxErr.ErrInvalidValue, "count", count, 0, Text.Length - index);
            List<TDrawingTextRun> NewRuns = new List<TDrawingTextRun>();
            int charPos = -1;
            foreach (TDrawingTextRun item in FRuns)
            {
                if (string.IsNullOrEmpty(item.Text)) continue;

                int charEndPos = charPos + item.Text.Length;
                if (charEndPos < index)
                {
                    charPos = charEndPos;
                    continue;
                }

                int s1 = index - charPos;
                int s2 = index + count;
                bool FullRun = true;
                if (s1 < 0)
                {
                    s1 = 0;
                    FullRun = false;
                }

                if (s2 >= charEndPos)
                {
                    s2 = charEndPos;
                }
                else FullRun = false;
                if (FullRun) NewRuns.Add(item);
                else NewRuns.Add(new TDrawingTextRun(item.Text.Substring(s1, s2 - s1 + 1), item.TextProperties));
                charPos = charEndPos;
                if (charPos > index + count) break;

            }
            return new TDrawingTextParagraph(NewRuns.ToArray(), Properties, EndParagraphProperties);
        }
        #endregion

        #region Compare

        /// <summary>
        /// Returns -1 if obj is bigger than this, 0 if both strings are the same, and 1 if obj is smaller than this.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public int CompareTo(object obj)
        {
            if (!(obj is TDrawingTextParagraph)) return -1;
            TDrawingTextParagraph s2 = (TDrawingTextParagraph)obj;


            int r = FRuns.Length.CompareTo(s2.FRuns.Length);
            if (r != 0) return r;

            for (int i = 0; i < FRuns.Length; i++)
            {
                r = FRuns[i].CompareTo(s2.FRuns[i]);
                if (r != 0) return r;
            }

            r = FProperties.CompareTo(s2.FProperties); if (r != 0) return r;
            r = FEndParagraphProperties.CompareTo(s2.FEndParagraphProperties); if (r != 0) return r;

            return 0;
        }

        /// <summary>
        /// Returns the hashcode of the object.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHashObj(FRuns, FProperties, FEndParagraphProperties);
        }

        /// <summary>
        /// Returns true if both instances have the same string and formatting.
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
        public static bool operator ==(TDrawingTextParagraph o1, TDrawingTextParagraph o2)
        {
            return Object.Equals(o1, o2);
        }

        /// <summary>
        /// Returns true if both objects do not have the same value.
        /// </summary>
        /// <param name="o1">First objects to compare.</param>
        /// <param name="o2">Second objects to compare.</param>
        /// <returns></returns>
        public static bool operator !=(TDrawingTextParagraph o1, TDrawingTextParagraph o2)
        {
            return !(Object.Equals(o1, o2));
        }

        /// <summary>
        /// Returns true is an object is less than the other.
        /// </summary>
        /// <param name="o1">First object to compare.</param>
        /// <param name="o2">Second object to compare.</param>
        /// <returns></returns>
        public static bool operator <(TDrawingTextParagraph o1, TDrawingTextParagraph o2)
        {
            return o1.CompareTo(o2) < 0;
        }

        /// <summary>
        /// Returns true is an object is bigger than the other.
        /// </summary>
        /// <param name="o1">First object to compare.</param>
        /// <param name="o2">Second object to compare.</param>
        /// <returns></returns>
        public static bool operator >(TDrawingTextParagraph o1, TDrawingTextParagraph o2)
        {
            return o1.CompareTo(o2) > 0;
        }

        #endregion

    }


    /// <summary>
    /// Properties of a text paragraph inside a drawing.
    /// </summary>
    public struct TDrawingParagraphProperties : IComparable
    {
        #region Variables
        private readonly int FMarL;
        private readonly int FMarR;
        private readonly int FLvl;
        private readonly TDrawingCoordinate FIndent;
        private readonly TDrawingAlignment FAlgn;
        private readonly TDrawingCoordinate FDefTabSz;
        private readonly bool FRtl;
        private readonly bool FEaLnBrk;
        private readonly TDrawingFontAlign FFontAlgn;
        private readonly bool FLatinLnBrk;
        private readonly bool FHangingPunct;
        #endregion

        #region Empty
        static readonly TDrawingParagraphProperties FEmpty = new TDrawingParagraphProperties(
                                347663,
                                0,
                                0,
                                new TDrawingCoordinate(-342900),
                                TDrawingAlignment.Left,
                                new TDrawingCoordinate(0),
                                false,
                                true,
                                TDrawingFontAlign.BaseLine,
                                true,
                                false
                    );

        #endregion

        #region Constructors
        /// <summary>
        /// Returns a paragraph with the default values.
        /// </summary>
        public static TDrawingParagraphProperties Empty { get { return FEmpty; } }

        /// <summary>
        /// Creates a new instance by setting all properties.
        /// </summary>
        /// <param name="aFMarL"></param>
        /// <param name="aFMarR"></param>
        /// <param name="aFLvl"></param>
        /// <param name="aFIndent"></param>
        /// <param name="aFAlgn"></param>
        /// <param name="aFDefTabSz"></param>
        /// <param name="aFRtl"></param>
        /// <param name="aFEaLnBrk"></param>
        /// <param name="aFFontAlgn"></param>
        /// <param name="aFLatinLnBrk"></param>
        /// <param name="aFHangingPunct"></param>
        public TDrawingParagraphProperties(
           int aFMarL, int aFMarR, int aFLvl, TDrawingCoordinate aFIndent, TDrawingAlignment aFAlgn,
            TDrawingCoordinate aFDefTabSz, bool aFRtl, bool aFEaLnBrk,
            TDrawingFontAlign aFFontAlgn, bool aFLatinLnBrk, bool aFHangingPunct
            )
        {
            FMarL = aFMarL;
            FMarR = aFMarR;
            FLvl = aFLvl;
            FIndent = aFIndent;
            FAlgn = aFAlgn;
            FDefTabSz = aFDefTabSz;
            FRtl = aFRtl;
            FEaLnBrk = aFEaLnBrk;
            FFontAlgn = aFFontAlgn;
            FLatinLnBrk = aFLatinLnBrk;
            FHangingPunct = aFHangingPunct;
        }
        #endregion

        #region Properties
        /// <summary>
        ///  Specifies the left margin of the paragraph. 
        /// </summary>
        public int MarL { get { return FMarL; } }

        /// <summary>
        /// Specifies the right margin of the paragraph.
        /// </summary>
        public int MarR { get { return FMarR; } }

        /// <summary>
        /// Specifies the particular level text properties that this paragraph follows.
        /// </summary>
        public int Lvl { get { return FLvl; } }

        /// <summary>
        /// Specifies the indent size that is applied to the first line of text in the paragraph. An 
        /// indentation of 0 is considered to be at the same location as marL attribute. 
        /// </summary>
        public TDrawingCoordinate Indent { get { return FIndent; } } 

        /// <summary>
        /// Specifies the alignment that is to be applied to the paragraph.
        /// </summary>
        public TDrawingAlignment Algn { get { return FAlgn; } }

        /// <summary>
        /// Specifies the default size for a tab character within this paragraph. This attribute should 
        /// be used to describe the spacing of tabs within the paragraph instead of a leading 
        /// indentation tab. For indentation tabs there are the marL and indent attributes to assist with this.  
        /// </summary>
        public TDrawingCoordinate DefTabSz { get { return FDefTabSz; } }

        /// <summary>
        /// Specifies whether the text is right-to-left or left-to-right in its flow direction.
        /// </summary>
        public bool Rtl { get { return FRtl; } }

        /// <summary>
        /// Specifies whether an East Asian word can be broken in half and wrapped onto the next line without a hyphen being added.
        /// </summary>
        public bool EaLnBrk { get { return FEaLnBrk; } }

        /// <summary>
        /// Determines where vertically on a line of text the actual words are positioned. This deals 
        /// with vertical placement of the characters with respect to the baselines.
        /// </summary>
        public TDrawingFontAlign FontAlgn { get { return FFontAlgn; } }

        /// <summary>
        /// Specifies whether a Latin word can be broken in half and wrapped onto the next line without a hyphen being added. 
        /// </summary>
        public bool LatinLnBrk { get { return FLatinLnBrk; } }

        /// <summary>
        /// Specifies whether punctuation is to be forcefully laid out on a line of text or put on a different line of text.
        /// </summary>
        public bool HangingPunct { get { return FHangingPunct; } }

        #endregion

        #region Compare

        /// <summary>
        /// Returns -1 if obj is bigger than this, 0 if both strings are the same, and 1 if obj is smaller than this.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public int CompareTo(object obj)
        {
            if (!(obj is TDrawingParagraphProperties)) return -1;
            TDrawingParagraphProperties s2 = (TDrawingParagraphProperties)obj;

            int r;
            r = FMarL.CompareTo(s2.FMarL); if (r != 0) return r;
            r = FMarR.CompareTo(s2.FMarR); if (r != 0) return r;
            r = FLvl.CompareTo(s2.FLvl); if (r != 0) return r;
            r = FIndent.CompareTo(s2.FIndent); if (r != 0) return r;
            r = FAlgn.CompareTo(s2.FAlgn); if (r != 0) return r;
            r = FDefTabSz.CompareTo(s2.FDefTabSz); if (r != 0) return r;
            r = FRtl.CompareTo(s2.FRtl); if (r != 0) return r;
            r = FEaLnBrk.CompareTo(s2.FEaLnBrk); if (r != 0) return r;
            r = FFontAlgn.CompareTo(s2.FFontAlgn); if (r != 0) return r;
            r = FLatinLnBrk.CompareTo(s2.FLatinLnBrk); if (r != 0) return r;
            r = FHangingPunct.CompareTo(s2.FHangingPunct); if (r != 0) return r;
            return 0;
        }

        /// <summary>
        /// Returns the hashcode of the object.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHashObj(
                        FMarL,
                        FMarR,
                        FLvl,
                        FIndent,
                        FAlgn,
                        FDefTabSz,
                        FRtl,
                        FEaLnBrk,
                        FFontAlgn,
                        FLatinLnBrk,
                        FHangingPunct
                );
        }

        /// <summary>
        /// Returns true if both instances have the same string and formatting.
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
        public static bool operator ==(TDrawingParagraphProperties o1, TDrawingParagraphProperties o2)
        {
            return Object.Equals(o1, o2);
        }

        /// <summary>
        /// Returns true if both objects do not have the same value.
        /// </summary>
        /// <param name="o1">First objects to compare.</param>
        /// <param name="o2">Second objects to compare.</param>
        /// <returns></returns>
        public static bool operator !=(TDrawingParagraphProperties o1, TDrawingParagraphProperties o2)
        {
            return !(Object.Equals(o1, o2));
        }

        /// <summary>
        /// Returns true is an object is less than the other.
        /// </summary>
        /// <param name="o1">First object to compare.</param>
        /// <param name="o2">Second object to compare.</param>
        /// <returns></returns>
        public static bool operator <(TDrawingParagraphProperties o1, TDrawingParagraphProperties o2)
        {
            return o1.CompareTo(o2) < 0;
        }

        /// <summary>
        /// Returns true is an object is bigger than the other.
        /// </summary>
        /// <param name="o1">First object to compare.</param>
        /// <param name="o2">Second object to compare.</param>
        /// <returns></returns>
        public static bool operator >(TDrawingParagraphProperties o1, TDrawingParagraphProperties o2)
        {
            return o1.CompareTo(o2) > 0;
        }

        #endregion
    }
    #endregion

    #region DrawingTextRun
    /// <summary>
    /// A rich formatting run used in text inside of a drawing. This struct is immutable.
    /// </summary>
    public struct TDrawingTextRun : IComparable
    {
        private readonly string FText;
        private readonly TDrawingTextProperties FTextProperties;


        /// <summary>
        /// Creates a new Text Run.
        /// </summary>
        public TDrawingTextRun(string aText, TDrawingTextProperties aTextProperties)
        {
            FText = aText;
            FTextProperties = aTextProperties;
        }

        /// <summary>
        /// String that this text run holds.
        /// </summary>
        public string Text { get { return FText; } }

        /// <summary>
        /// Properties for this text run.
        /// </summary>
        public TDrawingTextProperties TextProperties { get { return FTextProperties; } }

        #region Compare

        /// <summary>
        /// Returns -1 if obj is bigger than this, 0 if both strings are the same, and 1 if obj is smaller than this.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public int CompareTo(object obj)
        {
            if (!(obj is TDrawingTextRun)) return -1;
            TDrawingTextRun s2 = (TDrawingTextRun)obj;

            if (FText == null)
            {
                if (s2.FText != null) return -1;
            }
            else
            {
                int r = FText.CompareTo(s2.FText);
                if (r != 0) return r;
            }


            int r1 = FTextProperties.CompareTo(s2.FTextProperties);
            if (r1 != 0) return r1;
            return 0;
        }

        /// <summary>
        /// Returns the hashcode of the object.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHashObj(FText, FTextProperties);
        }

        /// <summary>
        /// Returns true if both instances have the same string and formatting.
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
        public static bool operator ==(TDrawingTextRun o1, TDrawingTextRun o2)
        {
            return Object.Equals(o1, o2);
        }

        /// <summary>
        /// Returns true if both objects do not have the same value.
        /// </summary>
        /// <param name="o1">First objects to compare.</param>
        /// <param name="o2">Second objects to compare.</param>
        /// <returns></returns>
        public static bool operator !=(TDrawingTextRun o1, TDrawingTextRun o2)
        {
            return !(Object.Equals(o1, o2));
        }

        /// <summary>
        /// Returns true is an object is less than the other.
        /// </summary>
        /// <param name="o1">First object to compare.</param>
        /// <param name="o2">Second object to compare.</param>
        /// <returns></returns>
        public static bool operator <(TDrawingTextRun o1, TDrawingTextRun o2)
        {
            return o1.CompareTo(o2) < 0;
        }

        /// <summary>
        /// Returns true is an object is bigger than the other.
        /// </summary>
        /// <param name="o1">First object to compare.</param>
        /// <param name="o2">Second object to compare.</param>
        /// <returns></returns>
        public static bool operator >(TDrawingTextRun o1, TDrawingTextRun o2)
        {
            return o1.CompareTo(o2) > 0;
        }

        /// <summary>
        /// String in the text run.
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return FText;
        }
        #endregion


        /// <summary>
        /// Returns true if this run contains a single line break.
        /// </summary>
        public bool IsBreak { get { return FText == ((char)10).ToString(); } }
    }

    internal struct TMutableDrawingTextProperties
    {
        #region Variables
        public TFillStyle FFill;
        public TLineStyle FLine;
        public TEffectProperties FEffects;
        public TDrawingColor? FHighlight;
        public TDrawingUnderline FUnderline;
        public TThemeTextFont? FLatin;
        public TThemeTextFont? FEastAsian;
        public TThemeTextFont? FComplexScript;
        public TThemeTextFont? FSymbol;

        public TDrawingHyperlink FHyperlinkClick;
        public TDrawingHyperlink FHyperlinkMouseOver;

        public bool FRightToLeft;

        public TDrawingTextAttributes FAttributes;
        #endregion

        public TDrawingTextProperties GetProps()
        {
            return new TDrawingTextProperties(
                FFill,
                FLine,
                FEffects,
                FHighlight,
                FUnderline,
                FLatin,
                FEastAsian,
                FComplexScript,
                FSymbol,
                FHyperlinkClick,
                FHyperlinkMouseOver,
                FRightToLeft,
                FAttributes
                );
        }

        public static TMutableDrawingTextProperties Empty
        {
            get
            {
                TDrawingTextProperties def = TDrawingTextProperties.Empty;
                TMutableDrawingTextProperties Result = new TMutableDrawingTextProperties();
                Result.FFill = def.Fill;
                Result.FLine = def.Line;
                Result.FEffects = def.Effects;
                Result.FHighlight = def.Highlight;
                Result.FUnderline = def.Underline;
                Result.FLatin = def.Latin;
                Result.FEastAsian = def.EastAsian;
                Result.FComplexScript = def.ComplexScript;
                Result.FSymbol = def.Symbol;
                Result.FHyperlinkClick = def.HyperlinkClick;
                Result.FHyperlinkMouseOver = def.HyperlinkMouseOver;
                Result.FRightToLeft = def.RightToLeft;
                Result.FAttributes = def.Attributes;

                return Result;
            }
        }


    }

    /// <summary>
    /// Properties of a text run inside a drawing.
    /// </summary>
    public struct TDrawingTextProperties : IComparable
    {
        #region Variables
        readonly TFillStyle FFill;
        readonly TLineStyle FLine;
        readonly TEffectProperties FEffects;
        readonly TDrawingColor? FHighlight;
        readonly TDrawingUnderline FUnderline;
        readonly TThemeTextFont? FLatin;
        readonly TThemeTextFont? FEastAsian;
        readonly TThemeTextFont? FComplexScript;
        readonly TThemeTextFont? FSymbol;

        readonly TDrawingHyperlink FHyperlinkClick;
        readonly TDrawingHyperlink FHyperlinkMouseOver;

        readonly bool FRightToLeft;

        readonly TDrawingTextAttributes FAttributes;
        #endregion

        #region Empty
        static readonly TDrawingTextProperties FEmpty = new TDrawingTextProperties(
        );

        #endregion

        #region Constructors

        /// <summary>
        /// Returns the text attributes with the default values.
        /// </summary>
        public static TDrawingTextProperties Empty { get { return FEmpty; } }

        /// <summary>
        /// Creates a new instance.
        /// </summary>
        public TDrawingTextProperties(
                    TFillStyle aFill,
                    TLineStyle aLine,
                    TEffectProperties aEffects,
                    TDrawingColor? aHighlight,
                    TDrawingUnderline aUnderline,
                    TThemeTextFont? aLatin,
                    TThemeTextFont? aEastAssian,
                    TThemeTextFont? aComplexScript,
                    TThemeTextFont? aSymbol,
                    TDrawingHyperlink aHyperlinkClick,
                    TDrawingHyperlink aHyperlinkMouseOver,
                    bool aRightToLeft,
                    TDrawingTextAttributes aAttributes
            )
        {
            FFill = aFill;
            FLine = aLine;
            FEffects = aEffects;
            FHighlight = aHighlight;
            FUnderline = aUnderline;
            FLatin = aLatin;
            FEastAsian = aEastAssian;
            FComplexScript = aComplexScript;
            FSymbol = aSymbol;

            FHyperlinkClick = aHyperlinkClick;
            FHyperlinkMouseOver = aHyperlinkMouseOver;

            FRightToLeft = aRightToLeft;
            FAttributes = aAttributes;
        }
        #endregion


        #region Properties
        /// <summary>
        /// Fill style for the text.
        /// </summary>
        public TFillStyle Fill { get { return FFill; } }

        /// <summary>
        /// Line style for the text.
        /// </summary>
        public TLineStyle Line { get { return FLine; } }

        /// <summary>
        /// Effects applied to the text.
        /// </summary>
        public TEffectProperties Effects { get { return FEffects; } }

        /// <summary>
        /// Highlight color that is present for a run of text.
        /// </summary>
        public TDrawingColor? Highlight { get { return FHighlight; } }

        /// <summary>
        /// Underline fill for the text.
        /// </summary>
        public TDrawingUnderline Underline { get { return FUnderline; } }

        /// <summary>
        /// This element specifies that a Latin font be used for a specific run of text. This font is specified with a typeface 
        /// attribute much like the others but is specifically classified as a Latin font. 
        /// </summary>
        public TThemeTextFont? Latin { get { return FLatin; } }

        /// <summary>
        /// This element specifies that an East Asian font be used for a specific run of text. This font is specified with a 
        /// typeface attribute much like the others but is specifically classified as an East Asian font. 
        /// </summary>
        public TThemeTextFont? EastAsian { get { return FEastAsian; } }

        /// <summary>
        /// This element specifies that a complex script font be used for a specific run of text. This font is specified with a 
        /// typeface attribute much like the others but is specifically classified as a complex script font.
        /// </summary>
        public TThemeTextFont? ComplexScript { get { return FComplexScript; } }

        /// <summary>
        /// This element specifies that a symbol script font be used for a specific run of text. This font is specified with a 
        /// typeface attribute much like the others but is specifically classified as a symbol script font.
        /// </summary>
        public TThemeTextFont? Symbol { get { return FSymbol; } }

        /// <summary>
        /// Specifies the on-click hyperlink information to be applied to a run of text. When the hyperlink text is clicked the 
        /// link is fetched.
        /// </summary>
        public TDrawingHyperlink HyperlinkClick { get { return FHyperlinkClick; } }

        /// <summary>
        /// Specifies the mouse-over hyperlink information to be applied to a run of text. When the mouse is hovered over 
        /// this hyperlink text the link is fetched. 
        /// </summary>
        public TDrawingHyperlink HyperlinkMouseOver { get { return FHyperlinkMouseOver; } }


        /// <summary>
        /// This element specifies whether the contents of this run shall have right-to-left characteristics. 
        /// </summary>
        public bool RightToLeft { get { return FRightToLeft; } }

        /// <summary>
        /// Group of simple attributes applied to the text run.
        /// </summary>
        public TDrawingTextAttributes Attributes { get { return FAttributes; } }
        #endregion

        #region Compare

        /// <summary>
        /// Returns -1 if obj is bigger than this, 0 if both strings are the same, and 1 if obj is smaller than this.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public int CompareTo(object obj)
        {
            if (!(obj is TDrawingTextProperties)) return -1;
            TDrawingTextProperties s2 = (TDrawingTextProperties)obj;

            int r;
            r = FlxUtils.CompareObjects(FFill, s2.FFill); if (r != 0) return r;
            r = FlxUtils.CompareObjects(FLine, s2.FLine); if (r != 0) return r;
            r = FlxUtils.CompareObjects(FEffects, s2.FEffects); if (r != 0) return r;
            r = FlxUtils.CompareObjects(FHighlight, s2.FHighlight); if (r != 0) return r;
            r = FlxUtils.CompareObjects(FUnderline, s2.FUnderline); if (r != 0) return r;
            r = FlxUtils.CompareObjects(FLatin, s2.FLatin); if (r != 0) return r;
            r = FlxUtils.CompareObjects(FEastAsian, s2.FEastAsian); if (r != 0) return r;
            r = FlxUtils.CompareObjects(FComplexScript, s2.FComplexScript); if (r != 0) return r;
            r = FlxUtils.CompareObjects(FSymbol, s2.FSymbol); if (r != 0) return r;

            r = FlxUtils.CompareObjects(FHyperlinkClick, s2.FHyperlinkClick); if (r != 0) return r;
            r = FlxUtils.CompareObjects(FHyperlinkMouseOver, s2.FHyperlinkMouseOver); if (r != 0) return r;

            r = FlxUtils.CompareObjects(FRightToLeft, s2.FRightToLeft); if (r != 0) return r;

            r = FlxUtils.CompareObjects(FAttributes, s2.FAttributes); if (r != 0) return r;
            return 0;
        }

        /// <summary>
        /// Returns the hashcode of the object.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHashObj(FFill);
        }

        /// <summary>
        /// Returns true if both instances have the same string and formatting.
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
        public static bool operator ==(TDrawingTextProperties o1, TDrawingTextProperties o2)
        {
            return Object.Equals(o1, o2);
        }

        /// <summary>
        /// Returns true if both objects do not have the same value.
        /// </summary>
        /// <param name="o1">First objects to compare.</param>
        /// <param name="o2">Second objects to compare.</param>
        /// <returns></returns>
        public static bool operator !=(TDrawingTextProperties o1, TDrawingTextProperties o2)
        {
            return !(Object.Equals(o1, o2));
        }

        /// <summary>
        /// Returns true is an object is less than the other.
        /// </summary>
        /// <param name="o1">First object to compare.</param>
        /// <param name="o2">Second object to compare.</param>
        /// <returns></returns>
        public static bool operator <(TDrawingTextProperties o1, TDrawingTextProperties o2)
        {
            return o1.CompareTo(o2) < 0;
        }

        /// <summary>
        /// Returns true is an object is bigger than the other.
        /// </summary>
        /// <param name="o1">First object to compare.</param>
        /// <param name="o2">Second object to compare.</param>
        /// <returns></returns>
        public static bool operator >(TDrawingTextProperties o1, TDrawingTextProperties o2)
        {
            return o1.CompareTo(o2) > 0;
        }

        #endregion

    }

    /// <summary>
    /// Group of simple attributes for text properties.
    /// </summary>
    public struct TDrawingTextAttributes : IComparable
    {
        #region Variables
        readonly bool FKumimoji;
        readonly string FLang;
        readonly string FAltLang;
        readonly int FSize;
        readonly bool FBold;
        readonly bool FItalic;
        readonly TDrawingUnderlineStyle FUnderline;
        readonly TDrawingTextStrike FStrike;
        readonly int FKern;
        readonly TDrawingTextCapitalization FCapitalization;
        readonly TDrawingCoordinate FSpacing;
        readonly bool FNormalizeH;
        readonly double FBaseline;
        readonly bool FNoProof;
        readonly bool FDirty;
        readonly bool FErr;
        readonly bool FSmartTagClean;
        readonly int FSmartTagId;
        readonly string FBookmarkLinkTarget;
        #endregion

        #region Empty
        static readonly TDrawingTextAttributes FEmpty = new TDrawingTextAttributes(
                false,
                null,
                null,
                0,
                false,
                false,
                TDrawingUnderlineStyle.None,
                TDrawingTextStrike.None,
                0,
                TDrawingTextCapitalization.None,
                new TDrawingCoordinate(0),
                false,
                0,
                false,
                false,
                false,
                false,
                0,
                null
             );

        #endregion

        #region Constructors

        /// <summary>
        /// Returns the text attributes with the default values.
        /// </summary>
        public static TDrawingTextAttributes Empty { get { return FEmpty; } }

        /// <summary>
        /// Creates a new instance.
        /// </summary>
        public TDrawingTextAttributes(
                    bool aKumimoji,
                    string aLang,
                    string aAltLang,
                    int aSize,
                    bool aBold,
                    bool aItalic,
                    TDrawingUnderlineStyle aUnderline,
                    TDrawingTextStrike aStrike,
                    int aKern,
                    TDrawingTextCapitalization aCapitalization,
                    TDrawingCoordinate aSpacing,
                    bool aNormalizeH,
                    double aBaseline,
                    bool aNoProof,
                    bool aDirty,
                    bool aErr,
                    bool aSmartTagClean,
                    int aSmartTagId,
                    string aBookmarkLinkTarget
            )
        {
            FKumimoji = aKumimoji;
            FLang = aLang;
            FAltLang = aAltLang;
            FSize = aSize;
            FBold = aBold;
            FItalic = aItalic;
            FUnderline = aUnderline;
            FStrike = aStrike;
            FKern = aKern;
            FCapitalization = aCapitalization;
            FSpacing = aSpacing;
            FNormalizeH = aNormalizeH;
            FBaseline = aBaseline;
            FNoProof = aNoProof;
            FDirty = aDirty;
            FErr = aErr;
            FSmartTagClean = aSmartTagClean;
            FSmartTagId = aSmartTagId;
            FBookmarkLinkTarget = aBookmarkLinkTarget;
        }
        #endregion

        #region Properties
        /// <summary>
        /// Specifies whether the numbers contained within vertical text continue vertically with the 
        /// text or whether they are to be displayed horizontally while the surrounding characters 
        /// continue in a vertical fashion.
        /// </summary>
        public bool Kumimoji { get { return FKumimoji; } }

        /// <summary>
        /// Specifies the language to be used when the generating application is displaying the user 
        /// interface controls.
        /// </summary>
        public string Lang { get { return FLang; } }

        /// <summary>
        /// Specifies the alternate language to use when the generating application is displaying the 
        /// user interface controls.
        /// </summary>
        public string AltLang { get { return FAltLang; } }

        /// <summary>
        /// Specifies the size of text within a text run. Whole points are specified in increments of 
        /// 100 starting with 100 being a point size of 1. For instance a font point size of 12 would be 
        /// 1200 and a font point size of 12.5 would be 1250. 
        /// </summary>
        public int Size { get { return FSize; } }

        /// <summary>
        /// Specifies whether a run of text is formatted as bold text.
        /// </summary>
        public bool Bold { get { return FBold; } }

        /// <summary>
        /// Specifies whether a run of text is formatted as italic text.
        /// </summary>
        public bool Italic { get { return FItalic; } }

        /// <summary>
        /// Specifies whether a run of text is formatted as underlined text. 
        /// </summary>
        public TDrawingUnderlineStyle Underline { get { return FUnderline; } }

        /// <summary>
        /// Specifies whether a run of text is formatted as strikethrough text.
        /// </summary>
        public TDrawingTextStrike Strike { get { return FStrike; } }

        /// <summary>
        /// Specifies the minimum font size at which character kerning occurs for this text run. 
        /// Whole points are specified in increments of 100 starting with 100 being a point size of 1. 
        ///For instance a font point size of 12 would be 1200 and a font point size of 12.5 would be 
        ///1250.
        /// </summary>
        public int Kern { get { return FKern; } }

        /// <summary>
        /// Specifies the capitalization that is to be applied to the text run. This is a render-only 
        /// modification and does not affect the actual characters stored in the text run. This 
        /// attribute is also distinct from the toggle function where the actual characters stored in 
        /// the text run are changed. 
        /// </summary>
        public TDrawingTextCapitalization Capitalization { get { return FCapitalization; } }

        /// <summary>
        /// Specifies the spacing between characters within a text run. This spacing is specified 
        /// numerically and should be consistently applied across the entire run of text by the 
        /// generating application. Whole points are specified in increments of 100 starting with 100 
        /// being a point size of 1. For instance a font point size of 12 would be 1200 and a font point 
        /// size of 12.5 would be 1250.
        /// </summary>
        public TDrawingCoordinate Spacing { get { return FSpacing; } }

        /// <summary>
        /// Specifies the normalization of height that is to be applied to the text run. This is a render-
        /// only modification and does not affect the actual characters stored in the text run. This 
        /// attribute is also distinct from the toggle function where the actual characters stored in 
        /// the text run are changed. 
        /// </summary>
        public bool NormalizeH { get { return FNormalizeH; } }

        /// <summary>
        ///   Specifies the baseline for both the superscript and subscript fonts.
        /// </summary>
        public double Baseline { get { return FBaseline; } }

        /// <summary>
        /// Specifies that a run of text has been selected by the user to not be checked for mistakes.
        /// </summary>
        public bool NoProof { get { return FNoProof; } }

        /// <summary>
        /// Specifies that the content of a text run has changed since the proofing tools have last been run. 
        /// </summary>
        public bool Dirty { get { return FDirty; } }

        /// <summary>
        /// Specifies that when this run of text was checked for spelling, grammar, etc. that a mistake was indeed found. 
        /// </summary>
        public bool Err { get { return FErr; } }

        /// <summary>
        /// Specifies whether or not a text run has been checked for smart tags.
        /// </summary>
        public bool SmartTagClean { get { return FSmartTagClean; } }

        /// <summary>
        ///  Specifies a smart tag identifier for a run of text. This ID is unique throughout the 
        /// presentation and is used to reference corresponding auxiliary information about the 
        /// smart tag.
        /// </summary>
        public int SmartTagId { get { return FSmartTagId; } }

        /// <summary>
        /// Specifies the link target name that is used to reference to the proper link properties in a 
        /// custom XML part within the document. 
        /// </summary>
        public string BookmarkLinkTarget { get { return FBookmarkLinkTarget; } }

        #endregion

        #region Compare

        /// <summary>
        /// Returns -1 if obj is bigger than this, 0 if both strings are the same, and 1 if obj is smaller than this.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public int CompareTo(object obj)
        {
            if (!(obj is TDrawingTextAttributes)) return -1;
            TDrawingTextAttributes s2 = (TDrawingTextAttributes)obj;

            int r;
            r = FKumimoji.CompareTo(s2.FKumimoji); if (r != 0) return r;
            r = String.Compare(FLang, s2.FLang); if (r != 0) return r;
            r = String.Compare(FAltLang, s2.FAltLang); if (r != 0) return r;
            r = FSize.CompareTo(s2.FSize); if (r != 0) return r;
            r = FBold.CompareTo(s2.FBold); if (r != 0) return r;
            r = FItalic.CompareTo(s2.FItalic); if (r != 0) return r;
            r = FUnderline.CompareTo(s2.FUnderline); if (r != 0) return r;
            r = FStrike.CompareTo(s2.FStrike); if (r != 0) return r;
            r = FKern.CompareTo(s2.FKern); if (r != 0) return r;
            r = FCapitalization.CompareTo(s2.FCapitalization); if (r != 0) return r;
            r = FSpacing.CompareTo(s2.FSpacing); if (r != 0) return r;
            r = FNormalizeH.CompareTo(s2.FNormalizeH); if (r != 0) return r;
            r = FBaseline.CompareTo(s2.FBaseline); if (r != 0) return r;
            r = FNoProof.CompareTo(s2.FNoProof); if (r != 0) return r;
            r = FDirty.CompareTo(s2.FDirty); if (r != 0) return r;
            r = FErr.CompareTo(s2.FErr); if (r != 0) return r;
            r = FSmartTagClean.CompareTo(s2.FSmartTagClean); if (r != 0) return r;
            r = FSmartTagId.CompareTo(s2.FSmartTagId); if (r != 0) return r;
            r = String.Compare(FBookmarkLinkTarget, s2.FBookmarkLinkTarget); if (r != 0) return r;

            return 0;
        }

        /// <summary>
        /// Returns the hashcode of the object.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHashObj(
                        FKumimoji,
                        FLang,
                        FAltLang,
                        FSize,
                        FBold,
                        FItalic,
                        FUnderline,
                        FStrike,
                        FKern,
                        FCapitalization,
                        FSpacing,
                        FNormalizeH,
                        FBaseline,
                        FNoProof,
                        FDirty,
                        FErr,
                        FSmartTagClean,
                        FSmartTagId,
                        FBookmarkLinkTarget
                        );
        }

        /// <summary>
        /// Returns true if both instances have the same string and formatting.
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
        public static bool operator ==(TDrawingTextAttributes o1, TDrawingTextAttributes o2)
        {
            return Object.Equals(o1, o2);
        }

        /// <summary>
        /// Returns true if both objects do not have the same value.
        /// </summary>
        /// <param name="o1">First objects to compare.</param>
        /// <param name="o2">Second objects to compare.</param>
        /// <returns></returns>
        public static bool operator !=(TDrawingTextAttributes o1, TDrawingTextAttributes o2)
        {
            return !(Object.Equals(o1, o2));
        }

        /// <summary>
        /// Returns true is an object is less than the other.
        /// </summary>
        /// <param name="o1">First object to compare.</param>
        /// <param name="o2">Second object to compare.</param>
        /// <returns></returns>
        public static bool operator <(TDrawingTextAttributes o1, TDrawingTextAttributes o2)
        {
            return o1.CompareTo(o2) < 0;
        }

        /// <summary>
        /// Returns true is an object is bigger than the other.
        /// </summary>
        /// <param name="o1">First object to compare.</param>
        /// <param name="o2">Second object to compare.</param>
        /// <returns></returns>
        public static bool operator >(TDrawingTextAttributes o1, TDrawingTextAttributes o2)
        {
            return o1.CompareTo(o2) > 0;
        }

        #endregion
    }
    #endregion

    #region DrawingUnderline
    /// <summary>
    /// Specifies the Fill style and line style of underlined text, when it is underlined.
    /// </summary>
    public sealed class TDrawingUnderline: IComparable
    {
        internal readonly string xmlLine;
        internal readonly string xmlFill;

        internal TDrawingUnderline(string aXmlLine, string aXmlFill)
        {
            xmlLine = aXmlLine;
            xmlFill = aXmlFill;
        }

        #region IComparable Members

        /// <summary>
        /// Returns -1, 0 or 1 depending if this object is smaller or bigger than the other.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public int CompareTo(object obj)
        {
            TDrawingUnderline o2 = obj as TDrawingUnderline;
            if (object.ReferenceEquals(o2, null)) return -1;

            int r;
            r = String.Compare(xmlLine, o2.xmlFill); if (r != 0) return r;
            r = String.Compare(xmlFill, o2.xmlFill); if (r != 0) return r;

            return 0;
        }

        /// <summary>
        /// Returns the hashcode of the object.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHashObj(xmlLine, xmlFill);
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
        public static bool operator ==(TDrawingUnderline o1, TDrawingUnderline o2)
        {
            return Object.Equals(o1, o2);
        }

        /// <summary>
        /// Returns true if both objects do not have the same value.
        /// </summary>
        /// <param name="o1">First objects to compare.</param>
        /// <param name="o2">Second objects to compare.</param>
        /// <returns></returns>
        public static bool operator !=(TDrawingUnderline o1, TDrawingUnderline o2)
        {
            return !(Object.Equals(o1, o2));
        }

        #endregion
    }
    #endregion

    #region Drawing Hyperlinks
    /// <summary>
    /// Specifies an hyerlink in a drawing. This class has no public members yet.
    /// </summary>
    public sealed class TDrawingHyperlink : IComparable
    {
        internal readonly string xml;

        internal TDrawingHyperlink(string aXml)
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
            TDrawingHyperlink o2 = obj as TDrawingHyperlink;
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
        public static bool operator ==(TDrawingHyperlink o1, TDrawingHyperlink o2)
        {
            return Object.Equals(o1, o2);
        }

        /// <summary>
        /// Returns true if both objects do not have the same value.
        /// </summary>
        /// <param name="o1">First objects to compare.</param>
        /// <param name="o2">Second objects to compare.</param>
        /// <returns></returns>
        public static bool operator !=(TDrawingHyperlink o1, TDrawingHyperlink o2)
        {
            return !(Object.Equals(o1, o2));
        }

        #endregion
    }
    #endregion
}
