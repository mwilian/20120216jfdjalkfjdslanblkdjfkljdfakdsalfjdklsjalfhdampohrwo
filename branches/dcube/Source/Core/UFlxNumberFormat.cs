using System;
using System.Text;
using System.Globalization;
using System.Diagnostics;

#if (MONOTOUCH)
  using Color = MonoTouch.UIKit.UIColor;
  using System.Drawing;
#else
	#if (WPF)
	using System.Windows.Media;
	#else
	using System.Drawing;
	using Colors = System.Drawing.Color;
	#endif
#endif

namespace FlexCel.Core
{
	internal enum TLocalDateTime
	{
		None,
		LongDate,
		LongTime
	}

    /// <summary>
    /// A simple structure containing a position and a character.
    /// </summary>
    public struct TCharAndPos
    {
        private int FPos;
        private string FChar;

        /// <summary>
        /// Position of the character in the string (0 based).
        /// </summary>
        public int Pos { get { return FPos; } set { FPos = value; } }

        /// <summary>
        /// Character that should go at position. Note that if this is a surrogate pair (UTF32) the string might have 2 UTF16 characters.
        /// </summary>
        public string Char { get { return FChar; } set { FChar = value; } }

        /// <summary>
        /// True if both structs are the same.
        /// </summary>
        /// <param name="obj">Object to check.</param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            if (!(obj is TCharAndPos)) return false;
            TCharAndPos o2 = (TCharAndPos)obj;
            return this == o2;
        }

        /// <summary>
        /// Compares two structures.
        /// </summary>
        /// <param name="o1"></param>
        /// <param name="o2"></param>
        /// <returns></returns>
        public static bool operator ==(TCharAndPos o1, TCharAndPos o2)
        {
            return o1.FPos == o2.FPos && o1.FChar == o2.FChar;
        }

        /// <summary>
        /// Compares two structures.
        /// </summary>
        /// <param name="o1"></param>
        /// <param name="o2"></param>
        /// <returns></returns>
        public static bool operator !=(TCharAndPos o1, TCharAndPos o2)
        {
            return !(o1 == o2);
        }

        /// <summary>
        /// Returns the hashcode for the object.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHashObj(FPos, FChar);
        }
    }

    /// <summary>
    /// Information about characters in a numeric format that need to be adpted when rendering. For example,
    /// if the cell A1 has value 1 and format "*_0" it will print as "______1" when the cell is wide, and as "_1" when the cell is shorter.
    /// </summary>
    public class TAdaptativeFormats
    {
        #region Privates
        private int FWildcardPos = -1;
        private TCharAndPos[] FSeparators;
        #endregion

        /// <summary>
        /// Position of the last wildcard ("*") character in the format (0 based). If a wildcard is present in the format,
        /// the string has to be expanded with the character at position until it fit the width of the cell.
        /// A negative value means there is no wildcard in the format.
        /// </summary>
        public int WildcardPos { get { return FWildcardPos; } set { FWildcardPos = value; } }

        /// <summary>
        /// An array of positions and characters that must be used to pad the string. In this field you have the the "_" and "?" numeric format delimeters from Excel.
        /// The spaces at the positions in the position array should have the width of the character specified in this field.
        /// if null, there are no separators in this class.
        /// </summary>
        public TCharAndPos[] Separators { get { return FSeparators; } set { FSeparators = value; } }

        /// <summary>
        /// Returns true if this class has no adaptative formats.
        /// </summary>
        public bool IsEmpty
        {
            get
            {
                return FWildcardPos == 0 && (Separators == null || Separators.Length > 0);
            }
        }

        /// <summary>
        /// Returns the string with the characters at the positions in Separator changed by the characters specified in Separators.
        /// </summary>
        public string ApplySeparators(string s)
        {
            if (Separators == null || Separators.Length == 0) return s;

            StringBuilder sb = new StringBuilder(s);
            foreach (TCharAndPos cp in Separators)
            {
                if (cp.Char == null) continue;

                for (int i = 0; i < cp.Char.Length; i++)
                {
                    sb[cp.Pos + i] = cp.Char[i];
                }
            }
            return sb.ToString();

        }

        /// <summary>
        /// Adds a separator at a given position. Separator array will be kept in order when you add a value. This routine considers utf32 characters.
        /// </summary>
        /// <param name="Format">String with the characters.</param>
        /// <param name="fp">Position in the format string. (0 based)</param>
        /// <param name="sp">Position in the final string. (0 based)</param>
        public void AddSeparator(string Format, int fp, int sp)
        {
            if (Format == null || Format.Length == 0) return;
            if (fp >= Format.Length) return;

            string Sep = String.Empty + Format[fp];
            if (CharUtils.IsSurrogatePair(Format, fp)) Sep += Format[fp + 1];
            AddSeparator(Sep, sp);
        }

        /// <summary>
        /// Adds a char separator at a given position. Separator array will be kept in order when you add a value. 
        /// </summary>
        /// <param name="sep">Character that will be used to calculat the width of the string.</param>
        /// <param name="sp">Position in the final string. (0 based)</param>
        public void AddSeparator(string sep, int sp)
        {
            int SepPos =  -1;
            int SepOfs = 0;
            if (Separators == null) 
            {
                Separators = new TCharAndPos[1];
                SepPos = 0;
            }
            else
            {
                TCharAndPos[] NewSeparators = new TCharAndPos[Separators.Length + 1];
                for (int i = 0; i < Separators.Length; i++)
                {
                    if (Separators[i].Pos == sp) //We don't need to insert in this case.
                    {
                        Separators[i].Char = sep;
                        return;
                    }

                    if (Separators[i].Pos > sp && SepPos < 0)
                    {
                        SepPos = i;
                        SepOfs = 1;
                    }

                    NewSeparators[i + SepOfs] = Separators[i];
                }
                if (SepOfs == 0) //goes at the end.
                {
                    SepPos = Separators.Length;
                }

                Separators = NewSeparators;
            }

            Separators[SepPos] = new TCharAndPos();
            Separators[SepPos].Pos = sp;
            Separators[SepPos].Char = sep;
        }

        /// <summary>
        /// This assumes separators in New are all sorted and after separators in old.
        /// </summary>
        internal void Mix(TAdaptativeFormats NewAdaptativeFormats, int offset)
        {
            if (WildcardPos < 0 && NewAdaptativeFormats.WildcardPos >= 0) WildcardPos = NewAdaptativeFormats.WildcardPos + offset;
            TCharAndPos[] NewSeparators;
            int Divide = 0;
            if (Separators == null || Separators.Length == 0)
            {
                if (NewAdaptativeFormats.Separators == null || NewAdaptativeFormats.Separators.Length == 0) return;
                NewSeparators = (TCharAndPos[])NewAdaptativeFormats.Separators.Clone();
            }
            else
            {
                Divide = Separators.Length;
                if (NewAdaptativeFormats.Separators == null || NewAdaptativeFormats.Separators.Length == 0) return;
                NewSeparators = new TCharAndPos[Separators.Length + NewAdaptativeFormats.Separators.Length];
            }

            for (int i = 0; i < Divide; i++)
            {
                NewSeparators[i] = Separators[i];
            }

            for (int i = Divide; i < NewSeparators.Length; i++)
            {
                NewSeparators[i].Char = NewAdaptativeFormats.Separators[i - Divide].Char;
                NewSeparators[i].Pos = NewAdaptativeFormats.Separators[i - Divide].Pos + offset;
            }

            Separators = NewSeparators;

        }

        internal void RemovedPosition(int pos, int len)
        {
            if (Separators == null) return;

            for (int i = Separators.Length - 1; i >= 0; i--)
            {
                if (Separators[i].Pos < pos) return;
                if (Separators[i].Pos == pos)
                {
                    if (Separators.Length == 1) Separators = null;
                    else
                    {
                        TCharAndPos[] NewSeparators = new TCharAndPos[Separators.Length - 1];
                        for (int z = 0; z < i; z++) NewSeparators[z] = Separators[z];
                        for (int z = i+1; z < Separators.Length; z++) NewSeparators[z - 1] = Separators[z];
                        Separators = NewSeparators;
                    }
                    return;
                }
                Separators[i].Pos-=len;
            }
        }

        internal void InsertedPosition(int p)
        {
            InsertedPosition(p, 1);
        }

        internal void InsertedPosition(int p, int ofs)
        {
            if (Separators == null) return;

            for (int i = Separators.Length - 1; i >= 0; i--)
            {
                if (Separators[i].Pos < p) return;
                Separators[i].Pos += ofs;
            }
        }

        internal TAdaptativeFormats CopyTo(int start, int TextLen)
        {
            TAdaptativeFormats Result = new TAdaptativeFormats();
            if (WildcardPos >= start && WildcardPos < start + TextLen)
            {
                Result.WildcardPos = WildcardPos - start;
            }

            if (Separators != null)
            {
                int ItemCount = 0;
                foreach (TCharAndPos ch in Separators)
                {
                    if (ch.Pos >= start)
                    {
                        if (ch.Pos < start + TextLen) ItemCount++; else break;
                    }
                }

                if (ItemCount > 0)
                {
                    Result.Separators = new TCharAndPos[ItemCount];
                    ItemCount = 0;
                    foreach (TCharAndPos ch in Separators)
                    {
                        if (ch.Pos >= start)
                        {
                            if (ch.Pos < start + TextLen)
                            {
                                Result.Separators[ItemCount].Pos = ch.Pos - start;
                                Result.Separators[ItemCount].Char = ch.Char;
                                ItemCount++;
                            }
                            else break;
                        }
                    }
                }
            }
            return Result;
        }

        internal static TAdaptativeFormats CopyTo(TAdaptativeFormats AdaptFormat, int idx, int TextLen)
        {
            if (AdaptFormat == null) return null;
            return AdaptFormat.CopyTo(idx, TextLen);
        }
    }

	/// <summary>
	/// Static class to convert cells to formatted strings. It uses format strings from Excel, that
	/// are different to those on .net, so we have to try to reconcile the diffs.
	/// </summary>
	public sealed class TFlxNumberFormat
	{
		private const string NegativeDate = "###################";//STATIC*
		internal const string EmptySection = "###################";//STATIC*

		private TFlxNumberFormat(){}

		private const string RegionalFormatStr = ""; //"*";
		
		/// <summary>
		/// Returns the string used on a standard date on the current locale
		/// </summary>
		public static string RegionalDateString
		{
			get
			{
				string s = CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern;
                string Sep = "/";  //CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator;  We need always to use "/" here, even if our sep is for example ".", since this si a *format* string. When converting this to a real string we will change it to wht the date sep is.
				//It looks Excel will only show 3 formats:
				// m/d/y
				// d/m/y
				// y/m/d
				//depending on which thing you begin. yy and yyyy are both allowed.

				foreach (char c in s)
				{
					if (c=='y' || c == 'Y') return RegionalFormatStr + "YYYY" + Sep + "mm" + Sep + "dd";
					if (c=='m' || c == 'M') return RegionalFormatStr + "mm" + Sep + "dd" + Sep + "YYYY";
					if (c=='d' || c == 'D') return RegionalFormatStr + "dd" + Sep + "mm" + Sep + "YYYY";
				}

				return "mm/dd/YYYY"; //should not come here. Just to ensure something valid is returned.
			}
		}

		/// <summary>
		/// Returns the string used on a standard date and time on the current locale
		/// </summary>
		public static string RegionalDateTimeString
		{
			get
			{
				string s = CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern;
				string Sep = CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator;
				//It looks Excel will only show 3 formats:
				// m/d/y
				// d/m/y
				// y/m/d
				//depending on which thing you begin. yy and yyyy are both allowed.

				string Hour = " hh" + CultureInfo.CurrentCulture.DateTimeFormat.TimeSeparator + "mm";
				foreach (char c in s)
				{
					if (c=='y' || c == 'Y') return RegionalFormatStr + "YYYY" + Sep + "mm" + Sep + "dd" + Hour;
					if (c=='m' || c == 'M') return RegionalFormatStr + "mm" + Sep + "dd" + Sep + "YYYY" + Hour;
					if (c=='d' || c == 'D') return RegionalFormatStr + "dd" + Sep + "mm" + Sep + "YYYY" + Hour;
				}

				return "mm/dd/YYYY hh:mm"; //should not come here. Just to ensure something valid is returned.

			}
		}

         /// <summary>
		/// Formats a value as it would be shown by Excel.
		/// </summary>
		/// <param name="Value">Value to format.</param>
		/// <param name="Format">Cell Format. (For example, "yyyy-mm-dd" for a date format, or "#.00" for a numeric 2 decimal format)
        /// <br/>This format string is the same you use in Excel under "Custom" format when formatting a cell, and it is documented
        /// in Excel documentation. Under <b>"Finding out what format string to use in TFlxFormat.Format"</b> section in <b>UsingFlexCelAPI.pdf</b>
        /// you can find more detailed information on how to create this string.
        /// </param>
		/// <param name="aColor">Final color of the text. (Depending on the format, color might change. I.e. Red for negatives) </param>
		/// <param name="Workbook">Workbook with the cell. If null, no color information will be returned and the base date fill be assumed to be 1900 (windows) and not 1904 (macs).</param>
		/// <param name="HasDate">Returns if the format contains a date.</param>
		/// <param name="HasTime">Returns if the format contains a time.</param>
		/// <returns>Formatted string.</returns>
        public static TRichString FormatValue(object Value, string Format, ref Color aColor, ExcelFile Workbook, out bool HasDate, out bool HasTime)
        {
            TAdaptativeFormats AdaptativeFormats;
            return FormatValue(Value, Format, ref aColor, Workbook, out HasDate, out HasTime, out AdaptativeFormats);
        }

		/// <summary>
		/// Formats a value as it would be shown by Excel.
		/// </summary>
		/// <param name="Value">Value to format.</param>
        /// <param name="Format">Cell Format. (For example, "yyyy-mm-dd" for a date format, or "#.00" for a numeric 2 decimal format)
        /// <br/>This format string is the same you use in Excel under "Custom" format when formatting a cell, and it is documented
        /// in Excel documentation. Under <b>"Finding out what format string to use in TFlxFormat.Format"</b> section in <b>UsingFlexCelAPI.pdf</b>
        /// you can find more detailed information on how to create this string.
        /// </param>
        /// <param name="aColor">Final color of the text. (Depending on the format, color might change. I.e. Red for negatives) </param>
		/// <param name="Workbook">Workbook with the cell. If null, no color information will be returned and the base date fill be assumed to be 1900 (windows) and not 1904 (macs).</param>
		/// <param name="HasDate">Returns if the format contains a date.</param>
		/// <param name="HasTime">Returns if the format contains a time.</param>
        /// <param name="AdaptativeFormats">Returns micro-justification information needed to adapt the text better to a cell. Null if there are no adaptative formats.</param>
		/// <returns>Formatted string.</returns>
		public static TRichString FormatValue(object Value, string Format, ref Color aColor, ExcelFile Workbook, out bool HasDate, out bool HasTime, out TAdaptativeFormats AdaptativeFormats)
		{
            AdaptativeFormats = null;
			TFormula fm = (Value as TFormula);
			if (fm != null) Value = fm.Result;

			HasTime=false; HasDate=false;
			if (Value!=null && Value.Equals(0.0) && Workbook!=null && Workbook.HideZeroValues) return new TRichString(String.Empty);

			TRichString rs = Value as TRichString;
			if (rs!=null && (Format.IndexOf('@')<0)) return rs; //If we have a format like "@@", rich text info is lost.

			bool Dates1904 = Workbook == null? false: Workbook.OptionsDates1904;
			string s=XlsFormatValueEx2(Workbook, Value, Format, Dates1904, ref aColor, ref HasDate, ref HasTime, out AdaptativeFormats);
			if (Workbook==null)
				return new TRichString(s);

			return new TRichString(s, new TRTFRun[0], Workbook);
		}

		/// <summary>
		/// Formats a value as it would be shown by Excel.
		/// </summary>
		/// <param name="Value">Value to format.</param>
        /// <param name="Format">Cell Format. (For example, "yyyy-mm-dd" for a date format, or "#.00" for a numeric 2 decimal format)
        /// <br/>This format string is the same you use in Excel under "Custom" format when formatting a cell, and it is documented
        /// in Excel documentation. Under <b>"Finding out what format string to use in TFlxFormat.Format"</b> section in <b>UsingFlexCelAPI.pdf</b>
        /// you can find more detailed information on how to create this string.
        /// </param>
        /// <param name="aColor">Final color of the text. (Depending on the format, color might change. I.e. Red for negatives) </param>
		/// <param name="Workbook">Workbook with the cell. If null, no color information will be returned and the base date fill be assumed to be 1900 (windows) and not 1904 (macs).</param>
		/// <returns>Formatted string.</returns>
		public static TRichString FormatValue(object Value, string Format, ref Color aColor, ExcelFile Workbook)
		{
			bool HasTime=false; bool HasDate=false;
			return FormatValue(Value, Format, ref aColor, Workbook, out HasDate, out HasTime);

		}

		private static int FindFrom(string c, string s, int pos)
		{
			return s.IndexOf(c, pos)-pos;
		}

		private static double GetconditionNumber(string Format, int p, out bool HasErrors)
		{
			HasErrors = true;
			int p2=FindFrom("]", Format, p);
			if (p2 < 0) 
				return 0;

			string number = Format.Substring(p, p2);

			double Result = 0;
            HasErrors = !TCompactFramework.ConvertToNumber(number, CultureInfo.InvariantCulture, out Result);
			return Result;
		}

		private static bool EvalCondition(string Format, int position, double V, out bool ResultValue, out bool SuppressNegativeSign, out bool SuppressNegativeSignComp)
		{
			SuppressNegativeSign = false;
			SuppressNegativeSignComp = false;
			ResultValue = false;
			if (position + 2 >= Format.Length) return false; //We need at least a sign and a bracket.
			switch (Format[position])
			{
				case '=':
				{
					bool HasErrors;
					double c = GetconditionNumber(Format, position + 1, out HasErrors);
					if (HasErrors) return false;
					ResultValue = V == c;
					SuppressNegativeSign = true;
					SuppressNegativeSignComp = false;
					return true;
				}
				case '<':
				{
					if (Format[position + 1] == '=')
					{
						bool HasErrors;
						double c = GetconditionNumber(Format, position + 2, out HasErrors);
						if (HasErrors) return false;
						ResultValue = V <= c;
						if (c <= 0) SuppressNegativeSign = true; else SuppressNegativeSign = false;
						SuppressNegativeSignComp = true;
						return true;
					}
					if (Format[position + 1] == '>')
					{
						bool HasErrors;
						double c = GetconditionNumber(Format, position + 2, out HasErrors);
						if (HasErrors) return false;
						ResultValue = V != c;
						SuppressNegativeSign = false;
						SuppressNegativeSignComp = true;
						return true;
					}
					
				
				{
					bool HasErrors;
					double c = GetconditionNumber(Format, position + 1, out HasErrors);
					if (HasErrors) return false;
					ResultValue = V < c;
					if (c <= 0) SuppressNegativeSign = true; else SuppressNegativeSign = false;
					SuppressNegativeSignComp = true;
					return true;
				}
				
				}

				case '>':
				{
					if (Format[position + 1] == '=')
					{
						bool HasErrors;
						double c = GetconditionNumber(Format, position + 2, out HasErrors);
						if (HasErrors) return false;
						ResultValue = V >= c;
						if (c <= 0) SuppressNegativeSignComp = true; else SuppressNegativeSignComp = false;
						SuppressNegativeSign = false;
						return true;
					}
				
				{
					bool HasErrors;
					double c = GetconditionNumber(Format, position + 1, out HasErrors);
					if (HasErrors) return false;
					ResultValue = V > c;
					if (c <= 0) SuppressNegativeSignComp = true; else SuppressNegativeSignComp = false;
					SuppressNegativeSign = false;
					return true;
				}
				
				}
			}

			return false;
		}

		private static bool GetNegativeSign(TResultCondition[]Conditions, int SectionCount, ref int TargetedSection, double V)
		{
			if (TargetedSection < 0) 
			{
				if (Conditions[0] == null && (V>0 || SectionCount <= 1 || (V == 0 && SectionCount <=2))) 
				{
					TargetedSection = 0;
					return false; //doesn't matter.
				}
				if (Conditions[1] == null && (V<0 || SectionCount <= 2)) 
				{
					TargetedSection = 1;
					if (SectionCount == 2 && Conditions[0] != null) return Conditions[0].SuppressNegComp;
					return true; 
				}

				if (Conditions[2] == null) TargetedSection = 2; else TargetedSection = 3;
				return false;
			}

			if (Conditions[TargetedSection] != null)
			{
				Debug.Assert(!Conditions[TargetedSection].Complement, "Selected section can not be a complement");
				return Conditions[TargetedSection].SuppressNeg;
			}

			//Find Complement, if any
			int NullCount = 0;
			int CompCount = 0;
			TResultCondition Comp = null;
			for (int i = 0; i < SectionCount; i++)
			{
				if (Conditions[i] != null)
				{
					Debug.Assert(Conditions[i].Complement);
					CompCount++;
					if (CompCount > 1) return false;
					Comp = Conditions[i];
				}
				else
				{
					NullCount ++;
					if (NullCount > 1) return false;
				}
			}

			if (Comp != null) return Comp.SuppressNegComp;
			return false;
		}

		private static string[] GetSections(string Format, double V, out int TargetedSection, out int SectionCount, out bool SuppressNegativeSign)
		{
			bool InQuote = false;

			string[] Result = new string[4];
			TResultCondition[] Conditions = new TResultCondition[4];

			int CurrentSection = 0;
			
			int StartSection = 0;
			TargetedSection = -1;

			int i = 0; 
			while (i < Format.Length)
			{
				if (Format[i] == '\"')
				{
					InQuote = !InQuote;
				}

				if (InQuote) 
				{
					i++;
					continue; //escaped characters inside a quote like \" are not valid.
				}

				if (Format[i] == '\\')
				{
					i+=2;
					continue;
				}

				if (Format[i] == '[') 
				{
					if (i + 2 < Format.Length)
					{
						bool TargetsThis;
						bool SuppressNegs; bool SuppressNegsComp;
						if (EvalCondition(Format, i + 1, V, out TargetsThis, out SuppressNegs, out SuppressNegsComp)) 
						{
							Conditions[CurrentSection] = new TResultCondition(SuppressNegs, SuppressNegsComp, !TargetsThis);

							if (TargetedSection < 0)
							{
								if (TargetsThis) 
								{
									TargetedSection = CurrentSection;
								}
							}
						}
					}

					//Quotes inside brackets are not quotes. So we need to process the full bracket.
					while (i < Format.Length && Format[i] != ']') i++;
					i++;
					continue;

				}

				if (Format[i] == ';') 
				{
					if (i > StartSection) Result[CurrentSection] = Format.Substring(StartSection, i - StartSection);
					CurrentSection++;
					SectionCount = CurrentSection;
					if (CurrentSection >= Result.Length) 
					{
						SuppressNegativeSign = GetNegativeSign(Conditions, SectionCount, ref TargetedSection, V);
						return Result;
					}
					StartSection = i + 1;
				}

				i++;
			}

			if (i > StartSection) Result[CurrentSection] = Format.Substring(StartSection, i - StartSection);
			CurrentSection++;
			SectionCount = CurrentSection;
			SuppressNegativeSign = GetNegativeSign(Conditions, SectionCount, ref TargetedSection, V);
			return Result;
		}

		internal static string GetSection(string Format, double V, out bool SectionMatches, out bool SuppressNegativeSign)
		{
			SectionMatches = true;

			int TargetedSection; int SectionCount;
			string[] Sections = GetSections(Format, V, out TargetedSection, out SectionCount, out SuppressNegativeSign);

			if (TargetedSection >= SectionCount) 
			{
				SectionMatches = false; //No section matches condition. This has changed in Excel 2007, older versions would show an empty cell here, and Excel 2007 displays "####". We will use Excel2007 formatting.
				return String.Empty;
			}
			if (Sections[TargetedSection] == null) return String.Empty; else return Sections[TargetedSection];
		}

		private static bool GetColor(string s, ref Color aColor, ExcelFile Workbook)
		{
			if (Workbook == null) return false;
			string ColorId = s.Substring(5); //Skip "COLOR" in string.

			int ColorIndex = 0;
			for (int i = 0; i < ColorId.Length; i++)
			{
				if (i > 2) return false;
				if (ColorId[i] < '0' || ColorId[i] > '9') return false;
                
				ColorIndex = ColorIndex * 10 + ColorId[i] - '0';
			}

			if (ColorIndex < 1 || ColorIndex > Workbook.ColorPaletteCount) return false; 
			aColor = Workbook.GetColorPalette(ColorIndex);
			return true;
		}

		/// <summary>
		/// Checks for a [Color] tag. This should be always the first on the format.
		/// </summary>
		/// <param name="Workbook"></param>
		/// <param name="Format"></param>
		/// <param name="aColor"></param>
		/// <param name="p"></param>
		private static void CheckColor(ExcelFile Workbook, string Format, ref Color aColor, ref int p)
		{
			p=0;
			if ((Format.Length>0) && (Format[0]=='[') && (Format.IndexOf("]")>0)) 
			{
				bool IgnoreIt=false;
				string s=Format.Substring(1, Format.IndexOf("]")-1);
				if (String.Equals(s, "BLACK", StringComparison.InvariantCultureIgnoreCase)) aColor= Colors.Black; else
					if (String.Equals(s, "CYAN", StringComparison.InvariantCultureIgnoreCase)) aColor=Colors.Cyan; else
					if (String.Equals(s, "BLUE", StringComparison.InvariantCultureIgnoreCase)) aColor=Colors.Blue; else
					if (String.Equals(s, "GREEN", StringComparison.InvariantCultureIgnoreCase)) aColor=Colors.Green; else
					if (String.Equals(s, "MAGENTA", StringComparison.InvariantCultureIgnoreCase)) aColor=Colors.Magenta; else
					if (String.Equals(s, "RED", StringComparison.InvariantCultureIgnoreCase)) aColor=Colors.Red; else
					if (String.Equals(s, "WHITE", StringComparison.InvariantCultureIgnoreCase)) aColor=Colors.White; else
					if (String.Equals(s, "YELLOW", StringComparison.InvariantCultureIgnoreCase)) aColor=Colors.Yellow;else  
					if (String.Compare(s,0,"COLOR",0, 5, StringComparison.InvariantCultureIgnoreCase) == 0 && GetColor(s, ref aColor, Workbook)) {} 
				
				else IgnoreIt=true;

				if (!IgnoreIt) p= Format.IndexOf("]")+1;
			}
		}

		/// <summary>
		/// Checks for an optional tag. (conditions like "[>=100"]) should be checked at the beginning, to know which section to use.
		/// </summary>
		/// <param name="Format"></param>
		/// <param name="p"></param>
		/// <param name="TextOut"></param>
		private static void CheckOptional(string Format, ref int p, StringBuilder TextOut)
		{
			if (p>=Format.Length) return;
			if (Format[p]=='[')
			{
				int p2=FindFrom("]", Format, p);
				if ((p<Format.Length-1)&&(Format[p+1]=='$'))  //currency
				{
					int p3=FindFrom("-", Format+"-", p);
					TextOut.Append(Format.Substring(p+2, Math.Min(p2,p3)-2));
				}
				p+= p2+1;
			}
		}

		private static void CheckLiteral(string Format, ref int p, StringBuilder TextOut, TAdaptativeFormats AdaptativeFormats)
		{
			if (p>=Format.Length) return;
			if (
				(Format[p]<'\u00FF') && 
				((Format[p]=='\u0020')||(Format[p]=='$')||(Format[p]=='(')||(Format[p]==')')||(Format[p]=='!')||
				(Format[p]=='^')||(Format[p]=='&')||(Format[p]=='\'')||(Format[p]=='\u00B4')||(Format[p]=='~')||
				(Format[p]=='{')||(Format[p]=='}')||(Format[p]=='=')||(Format[p]=='<')||(Format[p]=='>')) 
				)
			{
				TextOut.Append(Format[p]);
				p++;
				return;
			}

            if (Format[p] == '\\')
            {
                if (p < Format.Length - 1) TextOut.Append(Format[p + 1]);
                if (CharUtils.IsSurrogatePair(Format, p + 1))
                {
                    p++;
                    TextOut.Append(Format[p + 1]);
                }
                p += 2;
                return;
            }

			if (Format[p]=='*')
			{
                if (p < Format.Length - 1)
                {
                    TextOut.Append(Format[p + 1]);
                    AdaptativeFormats.WildcardPos = TextOut.Length - 1;
                    if (CharUtils.IsSurrogatePair(Format, p+1)) p++;
                }
				p+=2;
				return;
			}

			if (Format[p]=='_')
			{
                if (p < Format.Length - 1)
                {
                    TextOut.Append(' ');
                    AdaptativeFormats.AddSeparator(Format, p + 1, TextOut.Length - 1);
                }
                if (CharUtils.IsSurrogatePair(Format, p + 1)) p++;
                p += 2;
				return;
			}

			if (Format[p]=='"')                    
			{
                p = AdvanceQuote(Format, p, TextOut);
                if (p < Format.Length)
                {
                    p++;
                }
            }
		}

        private static int AdvanceQuote(string Format, int p, StringBuilder TextOut)
        {
			p++;
			while ((p<Format.Length) && (Format[p]!='"'))
			{
				TextOut.Append(Format[p]);
				p++;
			}
            return p;
        }

		private static bool IsDateChar(char c)
		{
			return (c=='D')||(c=='Y');
		}

		private static bool IsTimeChar(char c)
		{
			return (c=='H')||(c=='S');
		}

		private static readonly char[] HArray={'H','h'};

		private static void CheckEllapsedTime(double value, string UpFormat, ref int p, StringBuilder TextOut, ref bool HasDate, ref bool HasTime, out int HourPos)
		{
			HourPos = -1;
			if (p >= UpFormat.Length || UpFormat[p] != '[') return;

			int endP = p + 1;
			int HCount = 0; 
			int MCount = 0;
			int SCount = 0;
			while (endP < UpFormat.Length && UpFormat[endP] != ']') 
			{
				if (UpFormat[endP] == 'H') HCount++;
				else
					if (UpFormat[endP] == 'M') MCount++;
				else
					if (UpFormat[endP] == 'S') SCount++;
				else return; //only h and m formats here.

				endP++;
			}
			
			if (endP >= UpFormat.Length) return;
			if (HCount <= 0 && MCount <= 0 && SCount <=0) return;
			if (HCount * MCount != 0 || HCount * SCount != 0 || MCount * SCount != 0) return;

			HasTime = true;
            double d = value;

			int Count = 0;
			if (HCount > 0) 
			{
				d = d * 24;
				Count = HCount;
			} 
			else
				if (MCount > 0) 
			{
				d = d * 24 * 60;
				Count = MCount;
			}
			else
				if (SCount > 0) 
			{
				d = d * 24 * 3600;
				Count = SCount;
			}
			
			d = Math.Floor(Math.Abs(d)) * Math.Sign(d);
			TextOut.AppendFormat(CultureInfo.CurrentCulture, "{0:" + new string('0', Count) + "}", d);
			p = endP + 1;

			if (HCount > 0) HourPos = p;
		}

		private static bool NextIsSecond(String UpFormat, int p, int MCount)
		{
			const string DateChars = "[]DMYHM";

			if (MCount < 1 || MCount > 2) return false; //mmm is always month.
			int i = p;
			while (i < UpFormat.Length)
			{
				if (UpFormat[i] == 'S') return true;
				if (DateChars.IndexOf(UpFormat[i]) >= 0) return false;
				
				if (UpFormat[i] == '\"')
				{
					i++;
					while (i < UpFormat.Length && UpFormat[i] != '\"')
					{
						i++;
					}
				}
				if (i < UpFormat.Length && UpFormat[i] == '\\') i++;
				i++;
			}
			return false;
		}

		private static void CheckDate(CultureInfo RegionalCulture, double value, string Format, string UpFormat, bool Dates1904, ref int p, StringBuilder TextOut, bool LastHour, ref bool HasDate, ref bool HasTime)
		{
			const string AmPm="AM/PM"; 
			const string AP="A/P";
			StringBuilder Fmt=new StringBuilder("");  //We cant add the space at the beginning, because " -1" is converted into "- 1"
			bool Ok=false;
			int StartP=p;

            DateTime TheDate = DateTime.MinValue;
            bool UsedTheDate = false;

			while (p<UpFormat.Length)
			{
				int q = p;
				CheckRegionalSettings(Format, ref RegionalCulture, ref p, Fmt, true);
				if (p != q) continue;

				if (IsDateChar(UpFormat[p])|| (!LastHour && (UpFormat[p]=='M'))) HasDate=true;
				if (IsTimeChar(UpFormat[p])|| (LastHour && (UpFormat[p]=='M'))) HasTime=true;

				if (UpFormat[p]=='H') LastHour=true;
				if (UpFormat[p]=='M')
				{
					int pIni = p;
					while ((p<UpFormat.Length) && (UpFormat[p]=='M')) p++;
					int MCount = p - pIni;

					if (LastHour || NextIsSecond(UpFormat, p, MCount))
					{
						Fmt.Append('m', MCount);  //Lower m is minute.
						LastHour=false;
					} 
					else
					{
						if (MCount == 5) //5 m is a one letter month.
						{
                            if (!UsedTheDate)
                            {
                                if (!GetDateValue(value, Dates1904, TextOut, out TheDate)) return;
                                UsedTheDate = true;
                            }

							string Month = RegionalCulture.DateTimeFormat.MonthNames[TheDate.Month - 1];
							if (Month != null && Month.Length > 0)
							{
								Fmt.Append("\\" + Month[0].ToString(CultureInfo.CurrentCulture)); //this is not uppercase. i.e. Finland has lowercase names, and there we have a lowercase letter.
							}
						}
						else
						{
							Fmt.Append('M', MCount);  //Lower m is minute.
						}
					}

					Ok=true;
				}	
				else       
					if ((UpFormat.Length>=p+AmPm.Length)&&(UpFormat.Substring(p, AmPm.Length)==AmPm))
				{
					int i= Fmt.ToString().LastIndexOfAny(HArray); //Look always for the last. If we have "h h am/pm am/pm", the final string should be "H h tt tt"
					if (i>=0) Fmt[i]='h';  //lower h is 12 hour clock.
					if ((i-1>=0) && (Fmt[i-1]=='H')) Fmt[i-1]='h';  //This is for hh format.
					Fmt.Append("tt");
					p+=AmPm.Length;
					Ok=true;
				}
				else
					if ((UpFormat.Length>=p+AP.Length)&&(UpFormat.Substring(p, AP.Length)==AP))
				{
					int i= Fmt.ToString().LastIndexOfAny(HArray); //Look always for the last. If we have "h h am/pm am/pm", the final string should be "H h tt tt"
					if (i>=0) Fmt[i]='h';  //lower h is 12 hour clock.
					if ((i-1>=0) && (Fmt[i-1]=='H')) Fmt[i-1]='h';  //This is for hh format.
					Fmt.Append("t");
					p+=AP.Length;
					Ok=true;
				}
				else
					if ((UpFormat[p]=='H'))
				{
					Fmt.Append(UpFormat[p]);
					p++;
					Ok=true;
				}
				else
					if ((UpFormat[p]=='F')||(UpFormat[p]=='Z')||(UpFormat[p]=='T')||(UpFormat[p]=='G'))  //Those are not used by excel, but used by .net
				{
					Fmt.Append("\\"+Format[p]);
					p++;
				}
				else
					if ((UpFormat[p]=='\\'))
				{
					if (p + 1 < Format.Length) Fmt.Append("\\"+Format[p+1]);
					p+=2;
				}
				else
					if ((UpFormat[p]=='D')||(UpFormat[p]=='Y')||(UpFormat[p]=='S'))
				{
					Fmt.Append(Char.ToLower(Format[p], CultureInfo.InvariantCulture));
					p++;
					Ok=true;
				}
				else
					if ((UpFormat[p]=='"'))
				{
					int QuoteLen=-1;
					if(p<UpFormat.Length-1) QuoteLen=UpFormat.IndexOf('"',p+1);
					if (QuoteLen<0)QuoteLen=UpFormat.Length-(p); else QuoteLen-=p-1;
					Fmt.Append(Format.Substring(p, QuoteLen).Replace("\\", "\\\\"));
					p+=QuoteLen;
				}
				else
					if ((UpFormat[p]=='['))
				{
					p--;
					break;
				}
				else
					if (UpFormat[p]=='.' && (HasTime || HasDate)) //This one is complex. "." might be used as a date separator (for example dd.mm) or as a time separator (for example "hh.ss") In the second case, we need to change "." to DecimalSeparator, in the first not.
				{
					Fmt.Append("\\" + CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator); //.NET will not convert a "." here to the corresponding "," etc in other cultures. Excel also does not use RegionalCulture here, it uses the current culture.
					p++;
					while (p< UpFormat.Length && UpFormat[p] == '0')
					{
						Fmt.Append('f');
						p++;
					}
					Ok = true;
					break;
				}
                    
                else if (UpFormat[p] == '_')
                {
                   Fmt.Append(' '); //We won't use adaptative formats here. That would imply a whole rewrite of this method, to be like CheckNumber
  			       p+=2;
                }
                else if (UpFormat[p] == '*')
                {
                    if (p + 1 < Format.Length)Fmt.Append(Format[p+1]); //We won't use adaptative formats here. That would imply a whole rewrite of this method, to be like CheckNumber
                    p += 2;
                }

				else
				{
					Fmt.Append(Format[p]);
					p++;
				}
			}

			if (!Ok)
			{
				p=StartP;
				return;
			}

            if (!UsedTheDate)
            {
                if (!GetDateValue(value, Dates1904, TextOut, out TheDate)) return;
            }

			string s="";
			if (Fmt.Length>0)
			{
				try
				{
					CultureInfo FormatProvider = null;
					//.NET uses empty AM/PM designators as empty. Excel uses AM/PM. This happens for example on German locale.
					if (RegionalCulture.DateTimeFormat.AMDesignator == null || RegionalCulture.DateTimeFormat.AMDesignator.Length == 0)
					{
						FormatProvider = (CultureInfo)RegionalCulture.Clone();
						FormatProvider.DateTimeFormat.AMDesignator = FlxMessages.GetString(FlxMessage.TxtDefaultTimeAMString);
					}
					if (RegionalCulture.DateTimeFormat.PMDesignator == null || RegionalCulture.DateTimeFormat.PMDesignator.Length == 0)
					{
						if (FormatProvider == null) FormatProvider =  (CultureInfo)RegionalCulture.Clone();
						FormatProvider.DateTimeFormat.PMDesignator = FlxMessages.GetString(FlxMessage.TxtDefaultTimePMString);
					}
					
					if (FormatProvider == null) FormatProvider = RegionalCulture;
					s=TheDate.ToString(Fmt.ToString()+" ", FormatProvider); //We add this space at the beginning so .net won't use global ids. For ex: "m" is MonthDayPattern, but " m" will be correctely interpreted as month
				}
				catch (FormatException)
				{
					p=StartP;
					return;
				}

				TextOut.Append(s.Substring(0,s.Length-1)); //Remove the leading space.
				p++;
			}
		}

        private static bool GetDateValue(double value, bool Dates1904, StringBuilder TextOut, out DateTime TheDate)
        {
            TheDate = DateTime.MinValue;

            if (value < 0 || !FlxDateTime.IsValidDate(value, Dates1904))
            {
                TextOut.Length = 0;
                TextOut.Append(NegativeDate); //Negative dates are shown this way
                return false;
            }

            TheDate = FlxDateTime.FromOADate(value, Dates1904);
            return true;

        }

        private static char GetDigit(string StrValue, int DecPos, int LastNot0Pos, int Digit, int ExpValue, out bool HasDigit)
        {
            int PosInString = DecPos + ExpValue - (Digit + 1);
            if (PosInString >= DecPos) PosInString++;
            HasDigit = (Digit >= 0 || PosInString <= LastNot0Pos) && (Digit < 0 || PosInString >= 0);
            if (PosInString < 0 || PosInString > LastNot0Pos) return '0';
            return StrValue[PosInString];
        }

        private static void CheckNumber(double value, bool SuppressNegativeSign, string Format, string UpFormat, ref CultureInfo RegionalCulture, ref int p, StringBuilder TextOut, TAdaptativeFormats AdaptativeFormats)
        {
            if (p >= Format.Length) return;

            int q = p;
            StringBuilder FormattedNumber = new StringBuilder();
            if (!SuppressNegativeSign && value < 0) FormattedNumber.Append(RegionalCulture.NumberFormat.NegativeSign);

            TAdaptativeFormats NewAdaptativeFormats = new TAdaptativeFormats();
            double NewValue = Math.Abs(value);
            TDigitCollection MantissaDigits = new TDigitCollection();
            TDigitCollection ExpDigits = new TDigitCollection();
            UInt32List ThousandSeps = new UInt32List();
            int ExpPosInResult; bool ExpSkipPlus;
            if (!ScanFormat(Format, RegionalCulture, ref q, FormattedNumber, NewAdaptativeFormats, MantissaDigits, ExpDigits, ThousandSeps, ref NewValue, out ExpSkipPlus, out ExpPosInResult)) return;

            bool HasExp = ExpPosInResult >=0;
           
            int LastDigitPos = MantissaDigits.Count > 0? MantissaDigits[MantissaDigits.Count - 1]: -1;
            bool HasThousands = GetThousands(ThousandSeps, LastDigitPos, ref NewValue);
            if (HasExp) NewValue = Math.Abs(value);

            NumberFormatInfo Invariant = CultureInfo.InvariantCulture.NumberFormat;
            double dExp10;
            if (NewValue == 0) dExp10 = 0; else dExp10 = Math.Log10(NewValue);
            int Exp10 = (int)Math.Floor(dExp10);

            double Pow10 = Math.Pow(10, Exp10);
            int ExpValue = HasExp ? 0 : Exp10;

            int DecDigits = MantissaDigits.Count - MantissaDigits.DecimalSep + ExpValue; //So it gets rounded properly.
            if (HasExp)
            {
                if (DecDigits < 0) DecDigits = 0;
            }
            else
            {
                while (DecDigits < 0) { ExpValue++; DecDigits++; Pow10 *= 10; };
            }
            if (DecDigits > 15) DecDigits = 15;

            string StrValue = (NewValue / Pow10).ToString("F" + DecDigits.ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);

            //This will only happen when value = 0
            while (StrValue.Length > 0 && StrValue[0] == '0') StrValue = StrValue.Remove(0, 1); //So "?" shows " " when the value is 0.

            int DecPos = StrValue.IndexOf(Invariant.NumberDecimalSeparator);
            int EndPos = StrValue.Length;

            if (StrValue.Length > 0 && !Char.IsDigit(StrValue[0]) && StrValue[0] != '.' && StrValue[0] != ',') //nan, infinity, etc
            {
                TextOut.Remove(0, TextOut.Length);
                TextOut.Append(TFormulaMessages.ErrString(TFlxFormulaErrorValue.ErrNum)); 
                AdaptativeFormats.Separators = new TCharAndPos[0];
                AdaptativeFormats.WildcardPos = -1;

                p = Format.Length + 1;
                return;

            }

            if (DecPos < 0) 
            {
                DecPos = EndPos;
                StrValue = StrValue.Insert(EndPos, ".");
                EndPos++;
            }

            //Exponent. Must go before mantissa so we don't mess with the MantissaDigits positions.
            if (HasExp)
            {
                String StrExp = Math.Abs(Exp10).ToString(CultureInfo.InvariantCulture);
                ReplaceDigits(RegionalCulture, NewAdaptativeFormats, FormattedNumber, ExpDigits, false, 0, StrExp, StrExp.Length, StrExp.Length, false);
                if (Exp10 < 0)
                {
                    FormattedNumber.Insert(ExpPosInResult + 1, RegionalCulture.NumberFormat.NegativeSign);
                    NewAdaptativeFormats.InsertedPosition(ExpPosInResult);
                }
                else
                {
                    if (!ExpSkipPlus)
                    {
                        FormattedNumber.Insert(ExpPosInResult + 1, RegionalCulture.NumberFormat.PositiveSign);
                        NewAdaptativeFormats.InsertedPosition(ExpPosInResult);
                    }
                }
            }
            //Mantissa. 
            ReplaceDigits(RegionalCulture, NewAdaptativeFormats, FormattedNumber, MantissaDigits, HasThousands, ExpValue, StrValue, DecPos, EndPos, HasExp && value == 0);

            AdaptativeFormats.Mix(NewAdaptativeFormats, TextOut.Length);
            TextOut.Append(FormattedNumber);
            p = Format.Length + 1;
        }

        private static void ReplaceDigits(CultureInfo RegionalCulture, TAdaptativeFormats AdaptativeFormats, StringBuilder FormattedNumber, TDigitCollection Digits, bool HasThousands, int ExpValue, string StrValue, int DecPos, int EndPos, bool EverythingBehavesAs0)
        {
            for (int i = Digits.Count - 1; i >= 0; i--) //has to be done in reverse, so when we delete/insert characters we are ok.
            {
                int pos = Digits[i];
                bool HasDigit;
                int Digit = Digits.DecimalSep - i - 1;
                int LastNot0Pos = EndPos - 1;
                while (LastNot0Pos > DecPos && StrValue[LastNot0Pos] == '0') LastNot0Pos--;
                char d = GetDigit(StrValue, DecPos, LastNot0Pos, Digit, ExpValue, out HasDigit);

                if (HasDigit || EverythingBehavesAs0)
                {
                    FormattedNumber[pos] = d;
                    if (HasThousands) AddThousands(RegionalCulture, AdaptativeFormats, FormattedNumber, pos, Digit);
                }
                else
                {
                    switch (FormattedNumber[pos])
                    {
                        case '0':
                            FormattedNumber[pos] = d;
                            if (HasThousands) AddThousands(RegionalCulture, AdaptativeFormats, FormattedNumber, pos, Digit);
                            break;
                        case '?':
                            FormattedNumber[pos] = ' ';
                            AdaptativeFormats.AddSeparator("0", pos);
                            break;
                        case '#':
                            FormattedNumber.Remove(pos, 1);
                            AdaptativeFormats.RemovedPosition(pos, 1);
                            break;
                    }
                }
            }

            //Add missing digits to the left of the last positive digit.
            AddCharatersAtLeft(RegionalCulture, AdaptativeFormats, FormattedNumber, HasThousands, StrValue, DecPos, EndPos, ExpValue, Digits);
        }

        private static void AddCharatersAtLeft(CultureInfo RegionalCulture, TAdaptativeFormats AdaptativeFormats, StringBuilder FormattedNumber, bool HasThousands, string StrValue, int DecPos, int ExpPos, int ExpValue, TDigitCollection Digits)
        {
            if (Digits.Count < 1) return;
            int pos = Math.Min(Digits[0], Digits.DecimalPos);
            int Digit = Digits.DecimalSep;
            bool HasDigit = false;
            do
            {
                char c = GetDigit(StrValue, DecPos, ExpPos - 1, Digit, ExpValue, out HasDigit);
                if (HasDigit)
                {
#if (COMPACTFRAMEWORK)
                    FormattedNumber.Insert(pos, c.ToString());
#else
                    FormattedNumber.Insert(pos, c);
#endif
                    AdaptativeFormats.InsertedPosition(pos); //This could optimized to avoid inserting too much, but it isn't worth. We will keep it simpler.
                    if (HasThousands) AddThousands(RegionalCulture, AdaptativeFormats, FormattedNumber, pos, Digit);
                    Digit++;
                }
            } while (HasDigit);
        }

        private static bool GetThousands(UInt32List ThousandSeps, int LastDigitPos, ref double NewValue)
        {
            for (int i = ThousandSeps.Count - 1; i >= 0; i--)
            {
                if (ThousandSeps[i] < LastDigitPos) return true;
                NewValue /= 1000;
            }

            return false;
        }

        private static void AddThousands(CultureInfo RegionalCulture, TAdaptativeFormats AdaptativeFormats, StringBuilder FormattedNumber, int pos, int Digit)
        {
            if (RegionalCulture.NumberFormat.NumberGroupSeparator != null && IsGroupSeparator(Digit, RegionalCulture.NumberFormat))
            {
                FormattedNumber.Insert(pos + 1, RegionalCulture.NumberFormat.NumberGroupSeparator);
                AdaptativeFormats.InsertedPosition(pos + 1);
            }
        }

        private static bool IsGroupSeparator(int Digit, NumberFormatInfo FormatInfo)
        {
            if (Digit <= 0 || FormatInfo.NumberGroupSizes == null || FormatInfo.NumberGroupSizes.Length == 0) return false;
            
            int tp = 0;
            int i = 0;
            while (tp < Digit)
            {
                tp += FormatInfo.NumberGroupSizes[i];
                if (i < FormatInfo.NumberGroupSizes.Length - 1) i++;
                else
                {
                    if (FormatInfo.NumberGroupSizes[FormatInfo.NumberGroupSizes.Length - 1] == 0) return false; //No more thousand seps.
                } 
            }

            return tp == Digit;
        }

        private static bool ScanFormat(string Format, CultureInfo RegionalCulture, ref int q, StringBuilder FormattedNumber, 
            TAdaptativeFormats AdaptativeFormats, TDigitCollection MantissaDigits, TDigitCollection ExpDigits, 
            UInt32List ThousandSeps, ref double value, out bool ExpSkipPlus, out int ExpPosInResult)
        {
            ExpSkipPlus = false;
            ExpPosInResult = -1;
            MantissaDigits.DecimalSep = -1;
            ExpDigits.DecimalSep = -1;
            TDigitCollection Digits = MantissaDigits;
            
            while (q < Format.Length)
            {
                int oldq = q;
                CheckRegionalSettings(Format, ref RegionalCulture, ref q, FormattedNumber, false);
                if (q != oldq) continue;

                switch (Format[q])
                {
                    case '0':
                    case '#': 
                    case '?': // ? behaves like "#", but adds a space if not used.
                        {
                            Digits.Add(FormattedNumber.Length);
                            FormattedNumber.Append(Format[q]);
                            break;
                        }

                    case '.':
                        if (Digits.DecimalSep < 0)
                        {
                            Digits.DecimalSep = Digits.Count;
                            Digits.DecimalPos = FormattedNumber.Length;
                        }
                        FormattedNumber.Append(RegionalCulture.NumberFormat.NumberDecimalSeparator);
                        break;

                    case ',':
                        ThousandSeps.Add((UInt32) FormattedNumber.Length); //Thousandseps might be seps or divide the value, depending if there is any digit after them. We don't know here yet if we are going to find digits after this, so we save the value.
                        break;

                    case '_': //Means spaces the width of the next character.
                        {
                            if (q < Format.Length - 1)
                            {
                                FormattedNumber.Append(' ');
                                AdaptativeFormats.AddSeparator(Format, q + 1, FormattedNumber.Length - 1);
                            }
                            q++;
                            if (CharUtils.IsSurrogatePair(Format, q)) q++;
                            break;
                        }

                    case '*': //Means repeat next character.
                        {
                            if (q < Format.Length - 1)
                            {
                                FormattedNumber.Append(Format[q + 1]);
                                AdaptativeFormats.WildcardPos = FormattedNumber.Length - 1;
                                q++;
                                if (CharUtils.IsSurrogatePair(Format, q)) q++;
                            }
                            break;
                        }

                    case '\\':
                        {
                            if (q + 1 < Format.Length)
                            {
                                FormattedNumber.Append(Format[q + 1]);
                                if (CharUtils.IsSurrogatePair(Format, q))
                                {
                                    q++;
                                    FormattedNumber.Append(Format[q + 1]);
                                }
                                q++;
                            }
                            break;
                        }

                    case '\"':
                        {
                            q = AdvanceQuote(Format, q, FormattedNumber);
                            break;
                        }

                    case 'e':
                    case 'E':
                        if (q + 1 < Format.Length)
                        {
                            if (Format[q + 1] == '-') ExpSkipPlus = true;
                            Digits = ExpDigits;
                            ExpPosInResult = FormattedNumber.Length;
                            FormattedNumber.Append(Format[q]);
                            q++; //skip +/-
                        }
                        break;

                    case '%':
                        {
                            value *= 100;
                            FormattedNumber.Append(Format[q]);
                            break;
                        }

                    case '$':
                    case '-':
                    case '+':
                    case '/':
                    case '(':
                    case ')':
                    case ':':
                    case ' ':
                    default:
                        FormattedNumber.Append(Format[q]);
                        break;
                }
                q++;

            }

            if (ExpDigits.DecimalSep < 0) { ExpDigits.DecimalSep = ExpDigits.Count; ExpDigits.DecimalPos = FormattedNumber.Length; }
            int ExpPos = (ExpPosInResult < 0) ? FormattedNumber.Length + 1 : ExpPosInResult;
            if (MantissaDigits.DecimalSep < 0) { MantissaDigits.DecimalSep = MantissaDigits.Count; MantissaDigits.DecimalPos = ExpPos - 1; }
            return MantissaDigits.Count > 0 || ExpDigits.Count > 0;
        }

		private static void CheckText(object value, string Format, ref int p, StringBuilder TextOut)
		{
			if (p>=Format.Length) return;
			if (Format[p]=='@')
			{
				TextOut.Append(value.ToString());
				p++;
			}
		}

		private static string FormatNumber(ExcelFile Workbook, double value, bool SuppressNegativeSign, string Format, bool Dates1904, ref Color aColor, ref bool HasDate, ref bool HasTime, TAdaptativeFormats AdaptativeFormats)
		{
			CultureInfo RegionalCulture = CultureInfo.CurrentCulture;

			//Numbers/dates and text formats can't be on the same format string. It is a number XOR a date XOR a text
			int p=0;
			String UpFormat=Format.ToUpper(CultureInfo.InvariantCulture);

			CheckColor(Workbook, Format, ref aColor, ref p);
			StringBuilder Result= new StringBuilder("");  
			while (p<Format.Length)
			{
				int p1=p;
				CheckRegionalSettings(Format, ref RegionalCulture, ref p, Result, false);
				int HourPos;
				CheckEllapsedTime (value, UpFormat, ref p, Result, ref HasDate, ref HasTime, out HourPos); //Shold be before CheckOptional
				CheckOptional(Format, ref p, Result);
                CheckLiteral(Format, ref p, Result, AdaptativeFormats);
				CheckDate    (RegionalCulture, value , Format, UpFormat, Dates1904, ref p, Result, HourPos == p, ref HasDate, ref HasTime);
				CheckNumber  (value, SuppressNegativeSign, Format, UpFormat, ref RegionalCulture, ref p, Result, AdaptativeFormats);
				CheckText    (value, Format, ref p, Result);
				if (p1==p) //not found
				{
					if (SuppressNegativeSign && value < 0) return (-value).ToString();
					return value.ToString();
				}
        
			}
			return Result.ToString();
		}

		private static string FormatText(ExcelFile Workbook, object value, string Format, ref Color aColor, TAdaptativeFormats AdaptativeFormats)
		{
			//Numbers/dates and text formats can't be on the same format string. It is a number XOR a date XOR a text

			int SectionCount; int ts; bool SuppressNegativeSign;
			string[] Sections = GetSections(Format, 0, out ts, out SectionCount, out SuppressNegativeSign);
			if (SectionCount < 4)
			{
				Format = Sections[0];
				if (Format == null || Format.IndexOf('@')<0) 
				{
					int p1 = 0;
					Color newColor = aColor;
					if (Format != null) 
					{
						CheckColor(Workbook, Format, ref newColor, ref p1);
						if (p1 >= Format.Length || String.Equals(Format.Substring(p1), "general", StringComparison.InvariantCultureIgnoreCase))  aColor = newColor; //Excel only uses the color if the format is empty or has an "@".
					}
					return Convert.ToString(value); //everything is ignored here.
				}
			}
			else
			{
				Format = Sections[3];
				if (Format == null) Format = String.Empty;
			}

			int p=0;
			CheckColor(Workbook, Format, ref aColor, ref p);
			StringBuilder Result= new StringBuilder("");  
			while (p<Format.Length)
			{
				int p1=p;
				CheckOptional(Format, ref p, Result);
				CheckLiteral (Format, ref p, Result, AdaptativeFormats);
				CheckText  (value, Format, ref p, Result);
				if (p1==p) //not found
					return value.ToString();
			}
			return Result.ToString();

		}

		private static TLocalDateTime GetLocalDateTime(string Format, ref bool HasDate, ref bool HasTime)
		{
			string FmtUp = Format.ToUpper(CultureInfo.InvariantCulture);
			if (FmtUp.IndexOf("[$-F800]") >= 0) //This means format with long date from regional settings. This is new on Excel 2002
			{
				HasDate = true;
				HasTime = false;
				return TLocalDateTime.LongDate;
			}
			if (FmtUp.IndexOf("[$-F400]") >= 0) //This means format with long hour from regional settings. This is new on Excel 2002
			{
				HasDate = false;
				HasTime = true;
				return TLocalDateTime.LongTime;
			}

			return TLocalDateTime.None;
		}

		private static string GetLocalDateTime(DateTime dt, TLocalDateTime Ldt)
		{
			switch (Ldt)
			{
				case TLocalDateTime.LongDate: return dt.ToLongDateString();
				case TLocalDateTime.LongTime: return dt.ToLongTimeString();
			}
			return NegativeDate; //should not come here.

		}

		private static void CheckRegionalSettings(String Format, ref CultureInfo RegionalSettings, ref int p, StringBuilder TextOut, bool Quote)
		{
			if(p >= Format.Length - 3) return;
			if (String.CompareOrdinal(Format, p, "[$", 0, 2) == 0)  //format is [$Currency-Locale]
			{
				p += 2;
                
				//Currency
				int StartCurr = p;
				while (Format[p] != '-' && Format[p] != ']') 
				{
					p++;
					if (p >= Format.Length) return; //no tag found.
				}
	
				if (p - StartCurr > 0) 
				{
					if (Quote) TextOut.Append("\"");
					string v = Format.Substring(StartCurr, p - StartCurr );
					if (Quote) v = v.Replace("\"","\"\\\"\"");
					TextOut.Append(v);
					if (Quote) TextOut.Append("\"");
				}

				if (Format[p] != '-') 
				{
					p++;
					return; //no culture info.
				}


				p++;
				int StartStr = p;
				while (p < Format.Length && Format[p] != ']') 
				{
					p++;
				}
				if (p < Format.Length) //We actually found a close tag
				{
					int EndStr = p;
					p++; //add the ']' char.
					int Len = Math.Min(4, EndStr - StartStr);
					
					//to avoid issues with non existing tryparse we will convert from hexa directly.
					int Result = 0;
					int Offset = 0;
					for (int i = EndStr - 1; i >= EndStr - Len; i--)
					{
						char digit = Char.ToUpper(Format[i], CultureInfo.InvariantCulture);
						if (digit >= '0' && digit <= '9')
						{
							Result += ((int)digit - (int) '0') << Offset;
							Offset += 4;
							continue;
						}
						if (digit >= 'A' && digit <= 'F')
						{
							Result += (10 + (int)digit - (int) 'A') << Offset;
							Offset += 4;
							continue;
						}
						return; //Cannot parse.						
					}

					if (Result < 0) return;
					try
					{
						CultureInfo aRegionalSettings = new CultureInfo(Result & 0xFFFF);
                        if (aRegionalSettings.IsNeutralCulture) //neutral cultures will raise an exception when trying to retrieve their information.
                        {
                            aRegionalSettings = CultureInfo.CreateSpecificCulture(aRegionalSettings.Name);
                        }

                        aRegionalSettings.NumberFormat = CultureInfo.CurrentCulture.NumberFormat; //Excel won't change this.
                        aRegionalSettings.DateTimeFormat.DateSeparator = CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator; //again, Excel won't
                        aRegionalSettings.DateTimeFormat.TimeSeparator = CultureInfo.CurrentCulture.DateTimeFormat.TimeSeparator; //again, Excel won't

                        RegionalSettings = aRegionalSettings;
                    }
					catch (System.ArgumentException)
					{
						//We could not create the culture, so we will continue with the existing one.					
					}
				}
			}
		}


		private static string XlsFormatValueEx2(ExcelFile Workbook, object value, string Format, bool Dates1904, ref Color aColor, ref bool HasDate, ref bool HasTime, out TAdaptativeFormats AdaptativeFormats)
		{
            AdaptativeFormats = new TAdaptativeFormats();
			HasDate=false;
			HasTime=false;
			if ((Format==null)||(Format.Length==0))  //General 
			{  
				if (value is TFlxFormulaErrorValue)
					return TFormulaMessages.ErrString((TFlxFormulaErrorValue)value);
                if (value is double)
                {
                    double d = (double)value;
                    double ad = Math.Abs(d);
                     //general format never has more than 11 characters, including the decimal separator,  but not the "-" sign.
                    int fmt = 10;
                    if (ad >= 1e11) fmt = 6;
                    else if (ad >= 1e10) fmt = 11;

                    return d.ToString("G" + fmt.ToString(CultureInfo.InvariantCulture), CultureInfo.CurrentCulture);
                }
				return Convert.ToString(value);
			}
      
			if (value is char[]) value=new string((char[])value);  //Delphi 8 bdp returns char[] for memo fields.
			switch (TExcelTypes.ObjectToCellType(value))
			{
				case TCellType.Empty:
					return String.Empty;
				case TCellType.Number:
				{
					double V=Convert.ToDouble(value);
					bool SectionMatches; bool SuppressNegativeSign;
					string FormatSection =  GetSection(Format,V, out SectionMatches, out SuppressNegativeSign);
					if (!SectionMatches) return EmptySection; //This is Excel2007 way. Older version would show an empty cell.

					TLocalDateTime Res = GetLocalDateTime(FormatSection, ref HasDate, ref HasTime);
					if (Res != TLocalDateTime.None) 
					{
						DateTime Dt;
						if (FlxDateTime.TryFromOADate(V, Dates1904, out Dt))
						{
							return GetLocalDateTime(Dt, Res);
						}
						return NegativeDate;
					}
                    
					return FormatNumber(Workbook, V, SuppressNegativeSign, FormatSection, Dates1904, ref aColor, ref HasDate, ref HasTime, AdaptativeFormats);
				}

				case TCellType.DateTime:
				{
					double V;
					if (!FlxDateTime.TryToOADate((DateTime)value, Dates1904, out V)) return NegativeDate;

					bool SectionMatches; bool SuppressNegativeSign;
					string FormatSection =  GetSection(Format,V, out SectionMatches, out SuppressNegativeSign);
					if (!SectionMatches) return NegativeDate;
					
					TLocalDateTime Res2 = GetLocalDateTime(FormatSection, ref HasDate, ref HasTime);
					if (Res2 != TLocalDateTime.None) 
					{
						return GetLocalDateTime((DateTime)value, Res2);
					}

					if (V<0) return NegativeDate; //Negative dates are shown this way

					return FormatNumber(Workbook, V, SuppressNegativeSign, FormatSection, Dates1904, ref aColor, ref HasDate, ref HasTime, AdaptativeFormats);
				}

				case TCellType.String:
					return FormatText(Workbook, value, Format, ref aColor, AdaptativeFormats);
                
				case TCellType.Bool:
					return value.ToString();
				case TCellType.Error:
					if (value is TFlxFormulaErrorValue)
						return TFormulaMessages.ErrString((TFlxFormulaErrorValue)value);
					return value.ToString();
				default:
					return value.ToString();
			} //case
		}
	}

	class TResultCondition 
	{
		internal bool SuppressNeg;
		internal bool SuppressNegComp;
		internal bool Complement;

		internal TResultCondition(bool aSuppressNeg, bool aSuppressNegComp, bool aComplement)
		{
			SuppressNeg = aSuppressNeg;
			SuppressNegComp = aSuppressNegComp;
			Complement = aComplement;
		}
	}

	class TDigitCollection
	{
		internal int DecimalSep;
		internal int DecimalPos;
		private UInt32List Decimals = new UInt32List();

		internal int Count { get { return Decimals.Count; } }

		internal int this[int index] 
		{ 
			get 
			{ 
				unchecked{return (int)Decimals[index];}
			} 
			set  
			{
                unchecked { Decimals[index] = (uint)value; } 
			}
		}
	
        internal void Add(int p)
        {
			unchecked
			{
                Decimals.Add((uint)p);
			}
        }
    }

}
