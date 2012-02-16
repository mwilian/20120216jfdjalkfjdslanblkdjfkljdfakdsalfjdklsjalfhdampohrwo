using System;
using System.Text;
using System.Globalization;
using System.IO;
using System.Collections.Generic;


namespace FlexCel.Core
{
	/// <summary>
	/// Excel formula p(arse)t(hin)gs.
	/// </summary>
	internal enum ptg
	{
		Exp = 0x01,
		Tbl = 0x02,
		Add = 0x03,
		Sub = 0x04,
		Mul = 0x05,
		Div = 0x06,
		Power = 0x07,
		Concat = 0x08,
		LT = 0x09,
		LE = 0x0A,
		EQ = 0x0B,
		GE = 0x0C,
		GT = 0x0D,
		NE = 0x0E,
		Isect = 0x0F,
		Union = 0x10,
		Range = 0x11,
		Uplus = 0x12,
		Uminus = 0x13,
		Percent = 0x14,
		Paren = 0x15,
		MissArg = 0x16,
		Str = 0x17,
		Attr = 0x19,
		Sheet = 0x1A,
		EndSheet = 0x1B,
		Err = 0x1C,
		Bool = 0x1D,
		Int = 0x1E,
		Num = 0x1F,
		Array = 0x20,
		Func = 0x21,
		FuncVar = 0x22,
		Name = 0x23,
		Ref = 0x24,
		Area = 0x25,
		MemArea = 0x26,
		MemErr = 0x27,
		MemNoMem = 0x28,
		MemFunc = 0x29,
		RefErr = 0x2A,
		AreaErr = 0x2B,
		RefN = 0x2C,
		AreaN = 0x2D,
		MemAreaN = 0x2E,
		MemNoMemN = 0x2F,
		NameX = 0x39,
		Ref3d = 0x3A,
		Area3d = 0x3B,
		Ref3dErr = 0x3C,
		Area3dErr = 0x3D
	}

	/// <summary>
	/// Operators that can be found inside a formula
	/// </summary>
	public enum TOperator
	{
		/// <summary>
		/// No operator
		/// </summary>
		Nop=0,
		/// <summary>
		/// -
		/// </summary>
		Neg = (byte)ptg.Uminus,
		/// <summary>
		/// %
		/// </summary>
		Percent= (byte)ptg.Percent,
		/// <summary>
		/// ^
		/// </summary>
		Power = (byte)ptg.Power,
		/// <summary>
		/// *
		/// </summary>
		Mul = (byte)ptg.Mul,
		/// <summary>
		/// /
		/// </summary>
		Div = (byte)ptg.Div,
		/// <summary>
		/// +
		/// </summary>
		Add = (byte)ptg.Add,
		/// <summary>
		/// -
		/// </summary>
		Sub = (byte)ptg.Sub,
		/// <summary>
		/// &amp;
		/// </summary>
		Concat = ptg.Concat,
		/// <summary>
		/// greater or equal
		/// </summary>
		GE = (byte)ptg.GE,
		/// <summary>
		/// less or equal
		/// </summary>
		LE = (byte)ptg.LE,
		/// <summary>
		/// Not equal
		/// </summary>
		NE = (byte)ptg.NE,
		/// <summary>
		/// =
		/// </summary>
		EQ = (byte)ptg.EQ,
		/// <summary>
		/// less than
		/// </summary>
		LT = (byte)ptg.LT,
		/// <summary>
		/// greater than
		/// </summary>
		GT = (byte)ptg.GT,
		/// <summary>
		/// Unary plus.
		/// </summary>
		UPlus = (byte)ptg.Uplus,
	}

	/// <summary>
	/// Represents a Missing Argument on a Formula.
	/// </summary>
	internal class TMissingArg: IConvertible
	{
		internal static readonly TMissingArg Instance = new TMissingArg();

		private TMissingArg(){}

		#region IConvertible Members

		public ulong ToUInt64(IFormatProvider provider)
		{
			return 0;
		}

		public sbyte ToSByte(IFormatProvider provider)
		{
			return 0;
		}

		public double ToDouble(IFormatProvider provider)
		{
			return 0;
		}

		public DateTime ToDateTime(IFormatProvider provider)
		{
			return new DateTime ();
		}

		public float ToSingle(IFormatProvider provider)
		{
			return 0;
		}

		public bool ToBoolean(IFormatProvider provider)
		{
			return false;
		}

		public int ToInt32(IFormatProvider provider)
		{
			return 0;
		}

		public ushort ToUInt16(IFormatProvider provider)
		{
			return 0;
		}

		public short ToInt16(IFormatProvider provider)
		{
			return 0;
		}

		public string ToString(IFormatProvider provider)
		{
			return String.Empty;
		}

		public byte ToByte(IFormatProvider provider)
		{
			return 0;
		}

		public char ToChar(IFormatProvider provider)
		{
			return '\0';
		}

		public long ToInt64(IFormatProvider provider)
		{
			return 0;
		}

		public System.TypeCode GetTypeCode()
		{
			return new System.TypeCode ();
		}

		public decimal ToDecimal(IFormatProvider provider)
		{
			return 0;
		}

		public object ToType(Type conversionType, IFormatProvider provider)
		{
			return null;
		}

		public uint ToUInt32(IFormatProvider provider)
		{
			return 0;
		}

		#endregion
	}

	internal interface INameRecordList
	{
		int GetCount();
		string GetName(int pos);
		int NameSheet(int pos);
	}

    internal struct TParseState
    {
        internal bool ForcedArrayClass;
        internal int Level;
        internal bool DirectlyInFunction;

        internal TParseState(bool aForcedArrayClass, int aLevel, bool aDirectlyInFunction)
        {
            ForcedArrayClass = aForcedArrayClass;
            Level = aLevel;
            DirectlyInFunction = aDirectlyInFunction;
        }

        internal TParseState WithDirectlyInFunction(bool Value)
        {
            return new TParseState(ForcedArrayClass, Level, Value);
        }

        internal TParseState WithOneMoreLevel()
        {
            return new TParseState(ForcedArrayClass, Level + 1, DirectlyInFunction);
        }
    }

	/// <summary>
	/// A class for parsing a string into an RPN byte stream. 
	/// This base class does not define the format of the RPN stream, this is done on the children.
	/// </summary>
	internal abstract class TBaseFormulaParser
	{
		protected int MaxFormulaLen; //1024 for excel formulas, longer for flexcelreport formulas
		protected int ParsePos;
		private string Fw;

		private TWhiteSpaceStack StackWs;
		protected TFmReturnType InitialRefMode;

		protected INameRecordList FNameTable;
        internal bool IsArrayFormula;

		protected ExcelFile Xls;
        protected bool CanModifyXls;
		protected bool ForceAbsolute;

        protected bool ReadingXlsx;
        private int WorkingSheet;

        protected bool R1C1;
        protected int CurrentRow;
        protected int CurrentCol;

		protected TBaseFormulaParser (ExcelFile aXls, bool aCanModifyXls, string aw, 
            TFmReturnType ReturnType, bool aReadingXlsx, int aWorkingSheet)
		{
			MaxFormulaLen = FlxConsts.Max_FormulaLen;
			Fw = aw;
			ParsePos = 0;
			StackWs = new TWhiteSpaceStack();
			InitialRefMode=ReturnType;
			Xls = aXls;
            CanModifyXls = aCanModifyXls;
            ReadingXlsx = aReadingXlsx;
            WorkingSheet = aWorkingSheet;

			FNameTable = aXls ==  null? null: aXls.GetNameRecordList();

            if (aXls != null) R1C1 = aXls.FormulaReferenceStyle == TReferenceStyle.R1C1;
		}        

		internal static char ft(TFormulaToken t)
		{
			return TFormulaMessages.TokenChar(t);   
		}

		internal static string fts(TFormulaToken t)
		{
			return TFormulaMessages.TokenString(t);   
		}
        
		private static bool IsNumber(char c)
		{
			return Char.IsDigit(c);
		}

        private static bool IsNumber(char c, ref TParseNumState ParseNumState)
        {
            ParseNumState.LastNumValid = ParseNumState.NumValid;
            if (!ParseNumState.NumValid) return IsNumber(c);
            if (ParseNumState.FirstChar && c == '.') { ParseNumState.FirstChar = false; ParseNumState.DotCount++; ParseNumState.LastExp = false; return true; }
            if (IsNumber(c)) { ParseNumState.FirstChar = false; ParseNumState.LastExp = false; return true; }

            if (ParseNumState.FirstChar) { ParseNumState.NumValid = false; return false; } //first char must be a number.
            ParseNumState.FirstChar = false;

            if (c == 'e' || c == 'E') 
            {
                ParseNumState.ExpCount++; ParseNumState.LastExp = true;
                if (ParseNumState.ExpCount > 1) { ParseNumState.NumValid = false; return false; };
                return true;
            }

            if (c == '+' || c == '-')
            {
                if (ParseNumState.LastExp) { ParseNumState.LastExp = false; return true; }
            }

            if (c == '.')
            {
                ParseNumState.DotCount++;
                if (ParseNumState.DotCount < 2) { ParseNumState.LastExp = false; return true; }
            }

            ParseNumState.NumValid = false;
            return false;
            
        }

		private static bool IsAlpha(char c)
		{
            return Char.IsLetter(c) || (c == '_') || (c == '\\'); //IsLetter is not perfect here, but the best I could find. A character like "0x2031" will need quotes in Excel, so it is not just any character above 0xFF. But, a character like "0x256a" that doesn't need quotes in Excel, returns false for Char.IsLetter.
		}
        
		private static bool IsAZ(char c)
		{
			return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z');
		}

		private static int ATo1(char c)
		{
			return Char.ToUpper(c, CultureInfo.InvariantCulture) - 'A'+1;
		}


        protected bool NextChar()
        {
            bool Result = ParsePos < Fw.Length;
            if (Result)
            {
                ParsePos++;
                if (ParsePos >= MaxFormulaLen) FlxMessages.ThrowException(FlxErr.ErrFormulaTooLong, Fw);
            }
            return Result;
        }

		public string RemainingFormula
		{
			get
			{
				if (ParsePos<Fw.Length) return Fw.Substring(ParsePos); else return String.Empty;
			}
		}

		private bool PeekChar(ref char c)
		{
			bool Result=ParsePos<Fw.Length;
			if (Result)
			{
				c=Fw[ParsePos];
			}
			return Result;
		}

		private bool Peek2Char(ref char c)
		{
			bool Result=ParsePos+1<Fw.Length;
			if (Result)
			{
				c=Fw[ParsePos+1];
			}
			return Result;
		}
		private bool PeekCharWs(ref char c)
		{
			int aParsePos=ParsePos;
            while (aParsePos < Fw.Length && (Fw[aParsePos] == ' ' || Fw[aParsePos] == '\n' || Fw[aParsePos] == '\r')) aParsePos++;

			bool Result=aParsePos<Fw.Length;
			if (Result)
			{
				c=Fw[aParsePos];
			}
			return Result;
		}


		private void GetString()
		{
			StringBuilder s=new StringBuilder();
			SkipWhiteSpace();
			char c=' ';
			if ( !PeekChar(ref c) || (c!=ft(TFormulaToken.fmStr))) FlxMessages.ThrowException(FlxErr.ErrNotAString, Fw);
			NextChar();

			bool More=false;
			do 
			{
				More=false;
				if (PeekChar(ref c) && (c!=ft(TFormulaToken.fmStr)))
				{
					s.Append(c);
					NextChar();
					More=true;
				}
				else
				{
					char d=' '; char e=' ';
      
					if ( PeekChar(ref d) && (d==ft(TFormulaToken.fmStr)) && Peek2Char(ref e) && (e==ft(TFormulaToken.fmStr)))  //We found a double quote, this means a simple quote, and the string doesn't end here.
					{
						s.Append(ft(TFormulaToken.fmStr));
						NextChar();
						NextChar();
						More=true;
					}
				}
			} 
			while(More);

			if ( !PeekChar(ref c)) FlxMessages.ThrowException(FlxErr.ErrUnterminatedString,Fw);
			NextChar();
			AddParsed(s.ToString());
		}

		private void GetAlpha(TParseState ParseState)
		{
			// Possibilities:
			/* 1 -> Formula - We know by the "(" at the end
				2 -> Boolean - We just see if text is "true" or "false"
                2.1->Number - If it is a valid number and not a cell ref like "1:2" or a sheet name. Names can't start with a number or a dot.
				3 -> Error   - No, we already catched this
				4 -> Reference - Look if it is one of the strings between A1..IV65536 (and $A$1) IF it starts with $, we don't look at it here.
				5 -> 3d Ref    - Search for a '!' at the end. As it might start with "'", we need to look at it in other places too.
                6 -> 2007 table reference  Table1[..].. do not confuse with R[1]C[1]
				7 -> Named Range - if it isn't anything else...
				*/
        
			SkipWhiteSpace();
            int start = ParsePos;

			char c=' ';
            int FirstColon = -1; //we want to allow A1:B1 but not A1:B1:C1, which is valid, but it is 2 alpha objects joined by ":"
            TParseNumState ParseNumState = TParseNumState.Create(); //to allow for 1e+3, but not for name+3, which is an ADD operator.
            while (PeekChar(ref c) && ValidId(c, FirstColon, ref ParseNumState)) 
            {
                NextChar(); 
                if (c == ':') FirstColon = ParsePos; 
            }

			string sOrig = Fw.Substring(start, ParsePos - start);
            string s = sOrig.ToUpper(CultureInfo.InvariantCulture);
            string sNoColon = FirstColon < 0 ? sOrig : Fw.Substring(start, FirstColon - start - 1);

            bool HasNext = PeekChar(ref c);
            if (HasNext)
            {
                if (FirstColon < 0 && c == ft(TFormulaToken.fmOpenParen))
                {
                    GetFunction(null, s, ParseState);
                    return;
                }

                if (c == ft(TFormulaToken.fmExternalRef)) 
                {
                    GetRef3d(s, ParseState);
                    return;
                }
            }

            
            if (GetReference(start)) return;

            if (HasNext) //should be tested after references, to ensure it is not R[1]C[1]
            {
                if (c == ft(TFormulaToken.fmWorkbookOpen))
                {
                    GetTable2007Ref(s, ParseState);
                    return;
                }

            }

            if (ParseNumState.LastNumValid && GetNumber(s)) return; //Is number must go after GetReference, to handle things like =sum(1:2)
            if (GetBool(s)) return;

            if (FirstColon >= 0) { ParsePos = FirstColon - 1; s = sNoColon.ToUpper(CultureInfo.InvariantCulture); } //from here we don't want colons.
            bool IsInternal;
            if (GetNamedRange(TXlsNamedRange.GetInternal(sNoColon, out IsInternal), WorkingSheet, true)) return;

            bool ValidName = TXlsNamedRange.IsValidRangeName(s, out IsInternal);
            if (!ReadingXlsx || !ValidName || IsInternal)
            {
                FlxMessages.ThrowException(FlxErr.ErrUnexpectedId, s, Fw);
            }

            //Add a private named range so we can load the file
            CreateSupportingName(TXlsNamedRange.GetInternal(sNoColon, out IsInternal));

		}

        private void GetTable2007Ref(string s, TParseState ParseState)
        {
            char c = ' ';
            if (!PeekChar(ref c) || (c != ft(TFormulaToken.fmWorkbookOpen)))
                FlxMessages.ThrowException(FlxErr.ErrUnexpectedChar, c, ParsePos, Fw);

            int level = 1;
            NextChar();
            while (PeekChar(ref c) && level > 0)
            {
                if (c == ft(TFormulaToken.fmWorkbookOpen)) level++;
                if (c == ft(TFormulaToken.fmWorkbookClose)) level--;
                NextChar();
            }

            if (level > 0)
                FlxMessages.ThrowException(FlxErr.ErrUnexpectedChar, c, ParsePos, Fw);
            AddParsedRefErr(); //Fix this to add real support.
        }

        private void CreateSupportingName(string s)
        {
            int NamePos = Xls.AddEmptyName(s, 0);
            AddParsedName(NamePos);
        }

        private static bool ValidId(char c, int FirstColon, ref TParseNumState ParseNumState)
        {
            bool num = IsNumber(c, ref ParseNumState); //We need to run this always, so it shouldn't go in the conditional branch of the return below.
            return num || IsAlpha(c) || (c == '$') || (c == '.') || (c == ':' && FirstColon < 0);
        }

		private static string UnQuote(string s)
		{
			int index = s.IndexOf(ft(TFormulaToken.fmStr), 0);
			while (index > 0)
			{
				if (index+1 >= s.Length || s[index+1] != ft(TFormulaToken.fmStr))
					FlxMessages.ThrowException(FlxErr.ErrUnterminatedString, s);
				if (index+2 < s.Length)
				{
					index = s.IndexOf(ft(TFormulaToken.fmStr), index+2);
				}
				else
					break;
			}

			return s.Replace(fts(TFormulaToken.fmStr)+fts(TFormulaToken.fmStr), fts(TFormulaToken.fmStr));
		}

		private static object GetSimpleValue(string s)
		{
			if (s.Length <= 0) return null;
			if (s[0] == ft(TFormulaToken.fmStr))
			{
				if (s.Length <= 1 || s[s.Length - 1] != ft(TFormulaToken.fmStr))
					FlxMessages.ThrowException(FlxErr.ErrUnterminatedString, s);
				return UnQuote(s.Substring(1, s.Length-2));
			}

			if (String.Equals(s, fts(TFormulaToken.fmTrue), StringComparison.InvariantCultureIgnoreCase)) return true;
			if (String.Equals(s, fts(TFormulaToken.fmFalse), StringComparison.InvariantCultureIgnoreCase)) return false;

			if (s[0] == ft(TFormulaToken.fmErrStart))
				return (TFlxFormulaErrorValue) TFormulaMessages.StringToErrCode(s, true);

			return TFormulaMessages.StringToFloat(s);

		}

        private bool GetNumber(string s)
        {
            if (s == null || s.Length < 1 || !(IsNumber(s[0]) || s[0] == '.')) return false; //fast check specially for CF that doesn't has double.tryparse.

            double d = 0;
            if (!TCompactFramework.ConvertToNumber(s, CultureInfo.InvariantCulture, out d)) return false;

            if ((Math.Floor(d) == d) && (d <= 0xFFFF) && (d >= 0))
            {
                UInt16 w = Convert.ToUInt16(d);
                AddParsed(w);
            }
            else
            {
                AddParsed(d);
            }
            return true;
        }

        
        private void GetArray()
		{
			List<object> Objects = new List<object>();
			SkipWhiteSpace();
			int Rows = 1; int Cols = 1;
			int ExpectedCols = 0;
			char c = ' ';
			if (!PeekChar(ref c) || (c!=ft(TFormulaToken.fmOpenArray)))
				FlxMessages.ThrowException(FlxErr.ErrUnexpectedChar, c, ParsePos, Fw);

			StringBuilder Param = new StringBuilder();
			NextChar();
			while (PeekChar(ref c) && (c!=ft(TFormulaToken.fmCloseArray)))
			{
				NextChar();
				if (c==ft(TFormulaToken.fmArrayRowSep)) 
				{
					if (Rows == 1) 
						ExpectedCols = Cols;
					else
						if (Cols != ExpectedCols)
						FlxMessages.ThrowException(FlxErr.ErrArrayNotSquared, Fw);
					Rows++; 
					Cols = 1;
					Objects.Add(GetSimpleValue(Param.ToString()));
					Param.Length = 0;
				}
				else
					if (c==ft(TFormulaToken.fmArrayColSep)) 
				{
					Cols++;
					Objects.Add(GetSimpleValue(Param.ToString()));
					Param.Length = 0;
				}
				else 
					Param.Append(c);
			}
			Objects.Add(GetSimpleValue(Param.ToString()));

			if (Rows > 1 && ExpectedCols != Cols)
				FlxMessages.ThrowException(FlxErr.ErrArrayNotSquared, Fw);
			if (Objects.Count != Rows * Cols)
				FlxMessages.ThrowException(FlxErr.ErrArrayNotSquared, Fw);
			
			if (!PeekChar(ref c))
				FlxMessages.ThrowException(FlxErr.ErrMissingParen, Fw);
			NextChar();

			object[,] ArrayData = new object[Rows, Cols];
			int aPos = 0;
			for (int r = 0; r < Rows; r++)
				for (int col = 0; col < Cols; col++)
				{
					ArrayData[r, col] = Objects[aPos++];
				}

			AddParsedArray(ArrayData); 
		}

		private void GetFunctionArgs(TCellFunctionData Func, ref int ArgCount, TParseState ParseState)
		{
			NextChar(); //skip parenthesis
			char c=' ';
			bool MoreToCome = true;
			while (MoreToCome)
			{
				int ActualPos = ParsePos;
				if (!ParseState.ForcedArrayClass)
				{
					ParseState.ForcedArrayClass = Func.ParamType(ArgCount) == TFmReturnType.Array;
				}

				Expression(ParseState.WithOneMoreLevel());

				if (ParsePos == ActualPos) //Missing argument.
				{
					SkipWhiteSpace();
					if (ArgCount >0 || 
						(PeekChar(ref c)&& c==ft(TFormulaToken.fmFunctionSep)))
					{
						MakeLastWhitespaceNormal(); //No real need to call this here, but this way it will behave the same as Excel. (An space before the closing parenthesis on a missing arg is not a post whitespace but a normal space)
						AddParsedMissingArg();
					}
					else
					{
						PopWhiteSpace();
						ArgCount--;  //This is not a real argument, as in PI()
					}
				}
				else
				{
					ConvertLastRefValueType(Func.ParamType(ArgCount), ParseState, Func.DoesNotAlterArray);
					SkipWhiteSpace();
					DiscardNormalWhiteSpace();  //No space is allowed before a ",". We only keep the whitespace if it is for closing a parenthesis.
				}

				if (PeekChar(ref c))
				{
					//We should not call SkipWhitespace here, as it was already called.
					if (c==ft(TFormulaToken.fmFunctionSep)) 
					{
						NextChar(); 
						if (!PeekChar(ref c))
							FlxMessages.ThrowException(FlxErr.ErrUnexpectedEof, Fw);
					}
					else
						if (c==ft(TFormulaToken.fmCloseParen))	MoreToCome = false;
					else
						FlxMessages.ThrowException(FlxErr.ErrUnexpectedChar, c, ParsePos, Fw);
				}
				else
					FlxMessages.ThrowException(FlxErr.ErrUnexpectedEof, Fw);

				ArgCount++;
			}

			if (!PeekChar(ref c)) FlxMessages.ThrowException(FlxErr.ErrMissingParen, Fw);
			NextChar();

			if ((ArgCount < Func.MinArgCount) || (ArgCount > Func.MaxArgCount))
				FlxMessages.ThrowException(FlxErr.ErrInvalidNumberOfParams,Func.Name, Func.MinArgCount, Func.MaxArgCount, ArgCount);
		}

		private void GetFunction(string ExternalLocation, string FunctionName, TParseState ParseState)
		{
            int ArgCount = 0;
            int ExtraArgCount = 0;

            if (ReadingXlsx && FunctionName.StartsWith("_xlfn.", StringComparison.InvariantCultureIgnoreCase)) FunctionName = FunctionName.Substring("_xlfn.".Length); //This happens in 2010, where new functions are "future" to xlsx
			TCellFunctionData Func= FuncNameArray(FunctionName);

            if (Xls != null && Func == null && FunctionName != null && FunctionName.Length > 0)
            {
                TUserDefinedFunctionContainer FnContainer = Xls.GetUserDefinedFunctionFromDisplayName(FunctionName);
                if (FnContainer != null || ReadingXlsx)
                {
                    TUserDefinedFunctionLocation Location = TUserDefinedFunctionLocation.External;
                    string FnName = FunctionName;
                    if (FnContainer != null)
                    {
                        FnName = FnContainer.Function.InternalName;
                        Location = FnContainer.Location;
                    }

                    if (ReadingXlsx)
                    {
                        //if (ExternalLocation == null) Location = TUserDefinedFunctionLocation.Internal; else Location = TUserDefinedFunctionLocation.External;
                        //This is a complex one. External locations are used for functions that work in 2007 but not in 2003, and Internal locations in unknown functions.
                        //So we let the user decide, depending in what he defines in the function itself.
                    }
                    Func = TXlsFunction.GetData(255); //user function.	
                    int ExternSheet = 0; int ExternName = 0;

                    if (Location == TUserDefinedFunctionLocation.Internal)
                    {
                        if (CanModifyXls)
                        {
                            Xls.EnsureAddInInternalName(FnName, false, out ExternName);
                        }

                        AddParsedName(ExternName);
                    }
                    else
                    {
                        if (CanModifyXls)
                        {
                            if (ExternalLocation != null)
                            {
                                AddParsedExternName(ExternalLocation, FnName);
                            }
                            else
                            {
                                Xls.EnsureAddInExternalName(FnName, out ExternSheet, out ExternName);
                                AddParsedExternName(ExternSheet, ExternName);
                            }
                        }
                    }
                }

                ExtraArgCount = 1; //We can't increment ArgCount here, GetParams expects 0 in a formula like =PI() to detect missing arguments.
                SkipWhiteSpace();
            }


            if (Func == null)
                FlxMessages.ThrowException(FlxErr.ErrFunctionNotFound, FunctionName, Fw);

            if (CanModifyXls && Func.FutureInXls)
            {
                int NameIndex;
                Xls.EnsureAddInInternalName(Func.FutureName, true, out NameIndex);
            }

			GetFunctionArgs(Func, ref ArgCount, ParseState.WithDirectlyInFunction(true));

			AddParsedFunction(Func, (byte) (ArgCount + ExtraArgCount));
		}

		private bool GetBool(string s)
		{
			if (String.Equals(s,fts(TFormulaToken.fmTrue), StringComparison.InvariantCultureIgnoreCase))
			{
				AddParsed(true);
				return true;
			}
			if (String.Equals(s,fts(TFormulaToken.fmFalse), StringComparison.InvariantCultureIgnoreCase)) 
			{
				AddParsed(false);
				return true;
			}
			return false;
		}

        private void GetInternalNamedRange(string s, int sheet)
        {
            SkipWhiteSpace();
            int start = ParsePos;
            if (GetNamedRange(s, sheet, true)) { NextChar(); return; }

            UndoSkipWhiteSpace(start);
        }

		private bool GetNamedRange(string s, int sheet, bool trylocal)
		{
			int i = GetNamedRangeIndex(s, sheet, trylocal);
			if (i >= 0)
			{
				AddParsedName(i);
				return true;
			}

			return false;
		}

        protected int GetNamedRangeIndex(string s, int sheet, bool trylocal)
        {
            //We will try to find the name first in the active sheet, if not possible and trylocal is true, we will try to find it in the global list.
            if (FNameTable == null) return -1;
            int aCount = FNameTable.GetCount();
            int LocalName = -1;
            for (int i = 0; i < aCount; i++)
            {
                if (String.Equals(s, FNameTable.GetName(i), StringComparison.CurrentCultureIgnoreCase))
                {
                    int NameSheet = FNameTable.NameSheet(i) + 1;
                    if (NameSheet == sheet)
                    {
                        return i;
                    }
                    else if (NameSheet == 0) LocalName = i;
                }
            }

            if (trylocal && LocalName >= 0)
            {
                return LocalName;
            }

            return -1;
        }

		private static bool IsErrorCode(string s, ref TFlxFormulaErrorValue b, ref bool more)
		{
			s = s.ToUpper(CultureInfo.InvariantCulture);
			more=false;
			foreach (TFlxFormulaErrorValue err in TCompactFramework.EnumGetValues(typeof(TFlxFormulaErrorValue)))
			{
				string e=TFormulaMessages.ErrString(err);
				bool hasmore = s.Length<=e.Length && e.StartsWith(s);
				if (hasmore && s.Length==e.Length)
				{
					b=err;
					return true;
				}
				if (hasmore) more=true;
			}
			return false;
		}

		private void GetError()
		{
			SkipWhiteSpace();
			int start=ParsePos;

			char c=' ';
			string s=String.Empty;
			while (PeekChar(ref c)) 
			{
				NextChar();
				s=Fw.Substring(start, ParsePos-start).ToUpper(CultureInfo.InvariantCulture);
                TFlxFormulaErrorValue b = TFlxFormulaErrorValue.ErrNA;
				bool more = false;
                if (IsErrorCode(s, ref b, ref more))
                {
                    if (b == TFlxFormulaErrorValue.ErrRef)
                    {
                        AddParsedRefErr();
                    }
                    else AddParsed(b);
                    return;
                }

				if (!more) break;
			}

			FlxMessages.ThrowException(FlxErr.ErrUnexpectedId,s,Fw);
		}

        private bool GetRefErr()
        {
            char c = ' ';
            int index = 0;
            string ErrRef = TFormulaMessages.ErrString(TFlxFormulaErrorValue.ErrRef);
            while (PeekChar(ref c) && c== ErrRef[index])
            {
                NextChar();
                if (index >= ErrRef.Length - 1) return true;
                index++;
            }

            return false;
        }

        public void GetOneReference(ref bool RowAbs, ref bool ColAbs, ref int Row, ref int Col, out bool IsFullRowRange, out bool IsFullColRange, out bool IsRefErr)
        {
            RowAbs = ForceAbsolute; ColAbs = ForceAbsolute; IsRefErr = false;
            IsFullColRange = true;  //Something like 'B:B'
            IsFullRowRange = true; //something like '1:3'

            char c = ' ';
            if (PeekChar(ref c) && c == ft(TFormulaToken.fmErrStart))
            {
                IsRefErr = GetRefErr();
                Row = -1;
                Col = -1;
                IsFullRowRange = false;
                IsFullColRange = false;
                return;
            }

            if (R1C1)
            {
                if (!ReadR1C1Ref(ForceAbsolute, ref RowAbs, ref ColAbs, out Row, out Col, ref IsFullRowRange, ref IsFullColRange))
                {
                    Row = -1;
                    Col = -1;
                    return;
                }
            }
            else
            {
                ReadA1Ref(ref RowAbs, ref ColAbs, out Row, out Col, ref IsFullRowRange, ref IsFullColRange);
            }
         
            if (IsFullColRange) RowAbs = true; //Excel 2007 doesn't like it other way.
            if (IsFullRowRange) ColAbs = true;
        }

        public bool ReadR1C1Ref(bool ForceAbs, ref bool RowAbs, ref bool ColAbs, out int Row, out int Col, ref bool IsFullRowRange, ref bool IsFullColRange)
        {
            Row = -1; Col = -1;
            bool LocalRowAbs, LocalColAbs;
            bool HasRow = ReadRowOrCol(ft(TFormulaToken.fmR1C1_R), ref Row, FlxConsts.Max_Rows + 1, out LocalRowAbs, CurrentRow);
            bool HasCol = ReadRowOrCol(ft(TFormulaToken.fmR1C1_C), ref Col, FlxConsts.Max_Columns + 1, out LocalColAbs, CurrentCol);

            if (!HasRow && !HasCol) return false;

            if (!ForceAbs)
            {
                RowAbs = LocalRowAbs;
                ColAbs = LocalColAbs;
            }
            IsFullRowRange = !HasCol;
            IsFullColRange = !HasRow;
            return true;
        }

        private bool ReadRowOrCol(char RC, ref int RowOrCol, int MaxRowOrCol, out bool LocalRowColAbs, int CurrentRowOrCol)
        {
            int SavePos = ParsePos;
            bool ok = false;
            try
            {
                char c = ' ';
                LocalRowColAbs = true;
#if (COMPACTFRAMEWORK)
                if (!PeekChar(ref c) || Char.ToUpper(c, CultureInfo.InvariantCulture) != RC) return false;
#else
                if (!PeekChar(ref c) || Char.ToUpperInvariant(c) != RC) return false;
#endif
                NextChar();


                bool NegativeRowCol = false;
                if (PeekChar(ref c) && c == ft(TFormulaToken.fmR1C1RelativeRefStart))
                {
                    LocalRowColAbs = false;
                    NextChar();
                    if (PeekChar(ref c) && c == ft(TFormulaToken.fmMinus)) //No whitespace allowed here.
                    {
                        NegativeRowCol = true;
                        NextChar();
                    }
                }

                RowOrCol = 0;
                bool HasNumber = false;
                while (PeekChar(ref c) && IsNumber(c) && RowOrCol <= MaxRowOrCol)
                {
                    NextChar();
                    HasNumber = true;
                    RowOrCol = RowOrCol * 10 + ((int)c - (int)'0');
                }

                if (!LocalRowColAbs)
                {
                    if (c != ft(TFormulaToken.fmR1C1RelativeRefEnd)) return false;
                    NextChar();
                }

                if (!HasNumber) LocalRowColAbs = false;
                if (!LocalRowColAbs)
                {
                    if (NegativeRowCol) RowOrCol = -RowOrCol;
                    RowOrCol = CurrentRowOrCol + 1 + RowOrCol;
                    while (RowOrCol <= 0) RowOrCol += MaxRowOrCol;
                    while (RowOrCol > MaxRowOrCol) RowOrCol -= MaxRowOrCol;
                }

                ok = true;
                return true;
            }
            finally
            {
                if (!ok) ParsePos = SavePos;
            }
        }

        private void ReadA1Ref(ref bool RowAbs, ref bool ColAbs, out int Row, out int Col, ref bool IsFullRowRange, ref bool IsFullColRange)
        {
            char c = ' ';
            if (PeekChar(ref c) && c == ft(TFormulaToken.fmAbsoluteRef))
            {
                ColAbs = true;
                NextChar();
            }

            Col = 0;
            while (PeekChar(ref c) && IsAZ(c) && Col <= FlxConsts.Max_Columns + 1)
            {
                IsFullRowRange = false;
                NextChar();
                Col = Col * ATo1('Z') + ATo1(c);
            }

            if (ColAbs && IsFullRowRange)
            {
                ColAbs = false;
                RowAbs = true;
            }
            else
            {
                if (PeekChar(ref c) && c == ft(TFormulaToken.fmAbsoluteRef))
                {
                    RowAbs = true;
                    NextChar();
                }
            }

            Row = 0;
            while (PeekChar(ref c) && IsNumber(c) && Row <= FlxConsts.Max_Rows + 1)
            {
                IsFullColRange = false;
                NextChar();
                Row = Row * 10 + ((int)c - (int)'0');
            }
        }

		private void DoExternNamedRange(string ExternSheet, string s)
		{
			AddParsedExternName(ExternSheet, s);
		}

        private string ReadWord()
        {
            int start = ParsePos;
            char c = ' ';
            while (PeekChar(ref c) && (IsAlpha(c) || IsNumber(c) || (c == '.') || (c == ':'))) NextChar();
            return Fw.Substring(start, ParsePos - start);
        }

		private void GetGeneric3dRef(string ExternSheet, TParseState ParseState)
		{
			bool RowAbs1=false; bool ColAbs1=false;
			int Row1=0; int Col1=0;
            bool IsFullRowRange1; bool IsFullColRange1; bool IsRefErr;

			int SavedPos = ParsePos;
            string slo = ReadWord(); //we could reuse this s in getref, but needs some refactoring.
            string sup = slo.ToUpper(CultureInfo.InvariantCulture);
            int SParsedPos = ParsePos;
            
            char c = ' ';
            if (PeekChar(ref c) && c == ft(TFormulaToken.fmOpenParen))
            {
                GetFunction(ExternSheet, sup, ParseState);
                return;
            }

            
            ParsePos = SavedPos;

			char d = ' ';
			GetOneReference(ref RowAbs1, ref ColAbs1, ref Row1, ref Col1, out IsFullRowRange1, out IsFullColRange1, out IsRefErr);
            if (IsRefErr)
            {
                //We can't have a #ref:#ref range in excel. So we will use the first ref, and let the colon operator do its work.
                AddParsed3dRefErr(ExternSheet);
                return;
            }

			if ((Row1 <= 0 && !IsFullColRange1) || (Col1 <=0 && !IsFullRowRange1) || Row1 > FlxConsts.Max_Rows + 1 || Col1 > FlxConsts.Max_Columns + 1 ||
				(PeekChar(ref d) && IsAlpha(d)) //something like "a3a" or "IV"
				) //Wasn't a valid reference. It might be a name
			{
				ParsePos = SParsedPos;
				DoExternNamedRange(ExternSheet, slo);
				return;
			}
			
			if (!IsFullRowRange1 && !IsFullColRange1)
			{
				if (Row1>FlxConsts.Max_Rows+1 || Row1<=0 || Col1<=0 || Col1>FlxConsts.Max_Columns+1)
				{
					FlxMessages.ThrowException(FlxErr.ErrInvalidRef, Row1.ToString()+", "+Col1.ToString());                
				}
			}

		    c=' ';
			bool IsArea = false;
			if (PeekChar(ref c) && c==ft(TFormulaToken.fmRangeSep)) 
			{
				IsArea = GetSecondAreaPart(ExternSheet, Row1, Col1, RowAbs1, ColAbs1, IsFullRowRange1, IsFullColRange1, -1);		
			} 
			
			if (!IsArea)
			{
				if (IsFullColRange1 || IsFullRowRange1)
				{
                    if (R1C1 && (IsFullColRange1 ^ IsFullRowRange1))
                    {
                        int r1 = IsFullColRange1 ? 0 : Row1 - 1;
                        int r2 = IsFullColRange1 ? FlxConsts.Max_Rows : Row1 - 1;
                        int c1 = IsFullRowRange1 ? 0 : Col1 - 1;
                        int c2 = IsFullRowRange1 ? FlxConsts.Max_Columns : Col1 - 1;
                        AddParsed3dArea(ExternSheet, r1, r2, c1, c2, RowAbs1, RowAbs1, ColAbs1, ColAbs1);
                        return;
                    }

                    ParsePos = SParsedPos;
                    DoExternNamedRange(ExternSheet, slo);
                    return;
                }

				int rw1=Row1-1;
				int cl1=Col1-1;

				AddParsed3dRef(ExternSheet, rw1, cl1, RowAbs1, ColAbs1);
			}
		}

		private bool GetSecondAreaPart(string ExternSheet, int Row1, int Col1, bool RowAbs1, bool ColAbs1, bool IsFullRowRange1, bool IsFullColRange1, int EndPos)
		{
			bool RowAbs2=false; bool ColAbs2=false;
			int Row2=0; int Col2=0;
			int ActualPos = ParsePos;

			NextChar();
			bool IsFullRowRange2; bool IsFullColRange2;bool IsRefErr;
			GetOneReference(ref RowAbs2, ref ColAbs2, ref Row2, ref Col2, out IsFullRowRange2, out IsFullColRange2, out IsRefErr);
            if (IsRefErr)
            {
                ParsePos = ActualPos;
                return false;
            }
            
            if (IsFullRowRange1 && IsFullRowRange2)
			{
				Col1 = 1;
				Col2 = FlxConsts.Max_Columns + 1;
			}
			if (IsFullColRange1 && IsFullColRange2)
			{
				Row1 = 1;
				Row2 = FlxConsts.Max_Rows + 1;
			}

			if (Row2>FlxConsts.Max_Rows+1 || Row2<=0 || Col2<=0 || Col2>FlxConsts.Max_Columns+1)
			{
				ParsePos = ActualPos;
				return false;
			}

			int rw1=Row1-1;
			int cl1=(Col1-1);

			int rw2=Row2-1;
			int cl2=(Col2-1);

            if (IsValidEndPos(EndPos))
            {
                if (ExternSheet != null) AddParsed3dArea(ExternSheet, rw1, rw2, cl1, cl2, RowAbs1, RowAbs2, ColAbs1, ColAbs2); 
                else AddParsedArea(rw1, rw2, cl1, cl2, RowAbs1, RowAbs2, ColAbs1, ColAbs2);
            }
			return true;
		}

        private bool GetReference(int StartPos)
        {
            int EndPos = ParsePos;
            ParsePos = StartPos;
            if (!ReadReference(false, EndPos))
            {
                ParsePos = EndPos;
                return false;
            }

            return true;
        }

        private bool IsValidEndPos(int EndPos)
        {
            if (EndPos < 0) return true;
            if (ParsePos >= EndPos) return true;

            //parsepos might end in a different place than where it started. For example a1:sheet1!b1 would return just a1, since ":" here is the colon operator.
            if (ParsePos < EndPos && Fw[ParsePos] == ft(TFormulaToken.fmRangeSep)) return true;

            return false;
        }

        private void GetReference()
        {
            int start = ParsePos;
            if (!ReadReference(true, -1))
            {
    			string s=Fw.Substring(start, ParsePos - start).ToUpper(CultureInfo.InvariantCulture);
                FlxMessages.ThrowException(FlxErr.ErrUnexpectedId, s, Fw);
            }
        }

		private bool ReadReference(bool SkipWs, int EndPos)
		{
            int SaveParsePos = ParsePos;
            if (SkipWs)SkipWhiteSpace();

			bool RowAbs1=false; bool ColAbs1=false;
			int Row1=0; int Col1=0;
            bool IsFullRowRange1; bool IsFullColRange1; bool IsRefErr;
			GetOneReference(ref RowAbs1, ref ColAbs1, ref Row1, ref Col1, out IsFullRowRange1, out IsFullColRange1, out IsRefErr);
            if (IsRefErr) FlxMessages.ThrowException(FlxErr.ErrInternal); //this method should always be called from a non ref error.

			if (!IsFullRowRange1 && !IsFullColRange1)
			{
				if (Row1>FlxConsts.Max_Rows+1 || Row1<=0 || Col1<=0 || Col1>FlxConsts.Max_Columns+1)
				{
                    if (SkipWs) UndoSkipWhiteSpace(SaveParsePos);
					return false;
				}
			}

			char c=' ';
            if (R1C1 && (IsFullColRange1 ^ IsFullRowRange1))
            {
                //R1C1 can have areas with a single character like =R1. A1 refs need = 1:1
                if (IsValidEndPos(EndPos))
                {
                    int r1 = IsFullColRange1 ? 0 : Row1 - 1;
                    int r2 = IsFullColRange1 ? FlxConsts.Max_Rows : Row1 - 1;
                    int c1 = IsFullRowRange1 ? 0 : Col1 - 1;
                    int c2 = IsFullRowRange1 ? FlxConsts.Max_Columns : Col1 - 1;

                    AddParsedArea(r1, r2, c1, c2, RowAbs1, RowAbs1, ColAbs1, ColAbs1);
                }
            }
            else
            {
                bool IsArea = false;
                if (PeekChar(ref c) && c == ft(TFormulaToken.fmRangeSep))
                {
                    IsArea = GetSecondAreaPart(null, Row1, Col1, RowAbs1, ColAbs1, IsFullRowRange1, IsFullColRange1, EndPos);
                }

                if (!IsArea)
                {
                    if (IsFullColRange1 || IsFullRowRange1)
                    {
                        if (SkipWs) UndoSkipWhiteSpace(SaveParsePos);
                        return false;
                    }

                    int rw1 = Row1 - 1;
                    int cl1 = Col1 - 1;

                    if (IsValidEndPos(EndPos)) AddParsedRef(rw1, cl1, RowAbs1, ColAbs1);
                }
            }
			return IsValidEndPos(EndPos);
		}

		private void GetRef3d(string s, TParseState ParseState)
		{
			char c=' ';
			if (!PeekChar(ref c) || (c!=ft(TFormulaToken.fmExternalRef)))
				FlxMessages.ThrowException(FlxErr.ErrUnexpectedChar, c, ParsePos, Fw);
			NextChar();
			GetGeneric3dRef(s, ParseState);               
		}

		private void GetQuotedRef3d(TParseState ParseState)
		{
			SkipWhiteSpace();
			StringBuilder s= new StringBuilder();
			char c=' ';
			char sq=ft(TFormulaToken.fmSingleQuote);
			if (!PeekChar(ref c) || (c!=sq))
				FlxMessages.ThrowException(FlxErr.ErrUnexpectedChar, c, ParsePos, Fw);
			NextChar();
          
			bool More=false;
			do 
			{
				More=false;
				if (PeekChar(ref c) && (c!=sq))
				{
					s.Append(c);
					NextChar();
					More=true;
				}
				else
				{
					char d=' '; char e=' ';
      
					if ( PeekChar(ref d) && (d==sq) && Peek2Char(ref e) && (e==sq))  //We found a double quote, this means a simple quote, and the string doesn't end here.
					{
						s.Append(sq);
						NextChar();
						NextChar();
						More=true;
					}
				}
			} 
			while(More);

			if ( !PeekChar(ref c) || c!= sq) FlxMessages.ThrowException(FlxErr.ErrUnterminatedString,Fw);
			NextChar();
			GetRef3d(s.ToString(), ParseState);

		}

		//Gets a reference starting with "["
        private void GetExternRef3d(TParseState ParseState)
		{
			SkipWhiteSpace();
			StringBuilder s= new StringBuilder();
			char c=' ';
			char sq=ft(TFormulaToken.fmWorkbookOpen);
			if (!PeekChar(ref c) || (c!=sq))
				FlxMessages.ThrowException(FlxErr.ErrUnexpectedChar, c, ParsePos, Fw);
			NextChar();
			s.Append(sq);

			//First find the closing "]"  There could be "!" signs inside that.
			sq=ft(TFormulaToken.fmWorkbookClose);
			while (PeekChar(ref c) && c !=sq)
			{
				s.Append(c);
				NextChar();
			}

			if ( !PeekChar(ref c) || c!= sq) FlxMessages.ThrowException(FlxErr.ErrUnterminatedString,Fw);
			NextChar();
			s.Append(sq);

			sq=ft(TFormulaToken.fmExternalRef);
			while (PeekChar(ref c) && c !=sq)
			{
				s.Append(c);
				NextChar();
			}

			if ( !PeekChar(ref c) || c!= sq) FlxMessages.ThrowException(FlxErr.ErrUnterminatedString,Fw);

			GetRef3d(s.ToString(), ParseState);

		}


		/// <summary>
		/// [Whitespace]* Function | Number | String | Cell Reference | 3d Ref | (Expression) | NamedRange | Boolean | Err | Array
		/// </summary>
		private void Factor(TParseState ParseState)
		{
			char c=' ';
			if (PeekCharWs(ref c)) 
			{
				//if (c>'\u00FF') FlxMessages.ThrowException(FlxErr.ErrUnexpectedChar, c, ParsePos, Fw);
				
				if ( c == ft(TFormulaToken.fmOpenParen)) 
				{
					SkipWhiteSpace();
					NextChar();

					Expression(ParseState.WithDirectlyInFunction(false));

                    if (!(PeekCharWs(ref c)) || (c!=ft(TFormulaToken.fmCloseParen))) FlxMessages.ThrowException(FlxErr.ErrMissingParen, Fw);
					SkipWhiteSpace();
					NextChar();
					PopWhiteSpace();
					AddParsedParen();
				}
				else 
					if (c==ft(TFormulaToken.fmStr)) GetString();
				else 
					if (c==ft(TFormulaToken.fmOpenArray)) GetArray();
				else 
					if (c==ft(TFormulaToken.fmErrStart)) GetError();
				else 
					if (IsAlpha(c) || IsNumber(c) || c == '.') GetAlpha(ParseState);
                else
    				if (c==ft(TFormulaToken.fmAbsoluteRef)) GetReference();
                else
    				if (c==ft(TFormulaToken.fmSingleQuote)) GetQuotedRef3d(ParseState);
				else
					if (c==ft(TFormulaToken.fmWorkbookOpen)) GetExternRef3d(ParseState); //We need to fix 2007 tables here too, not in xlsx, but in the UI you can enter a table without a name before it.
				else
#if (COMPACTFRAMEWORK && !FRAMEWORK20)
					if (IsInternalName(c)) GetInternalNamedRange(Convert.ToString(c), WorkingSheet);
#else
                    if (IsInternalName(c)) GetInternalNamedRange(Convert.ToString(c, CultureInfo.InvariantCulture), WorkingSheet);
#endif
                else
					DoExtraToken(c);
			}
			else
				FlxMessages.ThrowException(FlxErr.ErrUnexpectedEof, Fw);

		}

        private static bool IsInternalName(char c)
        {
            return Enum.IsDefined(typeof(InternalNameRange), (InternalNameRange)c);
        }

		private bool NextIsReference()
		{
			//This is not perfect, as it does not fix something like Indirect("a1")
			//but it does cover the most usual cases.
            //return GetReference(true);

            //Problem with above is that it will fail with Sheet1!a1.  In fact, if we have an space, and the next thing is an alpha character
            //the only possibility is that it is an intersect, or an error.
            char c= ' ';
            if (!PeekCharWs(ref c)) return false;
            return IsAlpha(c) || IsNumber(c) || c == ft(TFormulaToken.fmErrStart) || c == ft(TFormulaToken.fmSingleQuote) || c == ft(TFormulaToken.fmOpenParen);
		}

		private bool IsIntersect()
		{
			char c = 'x';
			if (! PeekChar(ref c)) return false;
			if (c != ft(TFormulaToken.fmIntersect)) return false;
			if (!LastIsReference()) return false;
			if (!NextIsReference()) return false;
			return true;
		}

        /// <summary>
        /// RefTerm [' ' Factor]
        /// </summary>
        private void ISectTerm(TParseState ParseState)  //Intersect has more priority than union, even when not documented
        {
            Factor(ParseState);
            char c = ' ';
            bool First = true;

            //Intersect is tricky, as it uses spaces, and we can't know at first if it will be just whitespace or an intersect.
            //If both sides on the space are references, it is an intersect.
            while (PeekCharWs(ref c) && IsIntersect())
            {
                ConvertLastRefValueTypeOnce(TFmReturnType.Ref, ParseState, ref First);

                //before skipping ws, if this is an intersect, move on.
                NextChar();
                SkipWhiteSpace();

                Factor(ParseState);
                ConvertLastRefValueType(TFmReturnType.Ref, ParseState);
                AddParsedSep((byte)ptg.Isect);
            }
        }

		/// <summary>
		/// ISectTerm [ : | , ISectTerm]
		/// </summary>
		private void RefTerm(TParseState ParseState)    
		{
			ISectTerm(ParseState);
			char c=' ';
			bool First = true;
       
			//Union is only valid if we are not inside a function. For example A2:A3,B5 is ok. But HLookup(0,A2:A3,B5,1, true) is not ok.
            //=HLOOKUP(0;(A2:A3;B5);1; TRUE) is ok
			// =sum((A2:A3,B5 B6)) should be sum( (a2:a3) + (B5 isect B6)) and not (a2:a3,B5) isect b6
			while (PeekCharWs(ref c) && 
				(
				(c==ft(TFormulaToken.fmUnion) && !ParseState.DirectlyInFunction) 
				|| (c==ft(TFormulaToken.fmRangeSep))
				)
				)
			{
				ConvertLastRefValueTypeOnce(TFmReturnType.Ref, ParseState, ref First);
                
				SkipWhiteSpace();
			    NextChar();

				ISectTerm(ParseState);
				ConvertLastRefValueType(TFmReturnType.Ref, ParseState);
				byte b=0;
				if (c==ft(TFormulaToken.fmUnion)) b=(byte)ptg.Union; else
					b=(byte)ptg.Range;
				AddParsedSep(b);
			}
		}

		/// <summary>
		/// [- | +] * RefTerm
		/// </summary>
		private void NegTerm(TParseState ParseState)  
		{
			StringBuilder sb = null; //avoid creation except when needed.
			char c=' ';
			while (PeekCharWs(ref c) && (c==ft(TFormulaToken.fmMinus) || c==ft(TFormulaToken.fmPlus)))
			{
				SkipWhiteSpace();
				NextChar();
				if (sb == null) sb = new StringBuilder();
				sb.Append(c);
			}

			RefTerm(ParseState);
			
			if (sb != null)
			{
				ConvertLastRefValueType(TFmReturnType.Value, ParseState);

				for (int i=sb.Length-1; i>=0;i--) 
				{
					if (sb[i]==ft(TFormulaToken.fmMinus)) 
						AddParsedOp(TOperator.Neg);
					else
						AddParsedOp(TOperator.UPlus);
				}
			}
		}

		/// <summary>
		/// NegTerm [%]*
		/// </summary>
		private void PerTerm(TParseState ParseState)
		{
			NegTerm(ParseState);
			char c=' ';

			bool First = true;
			while (PeekCharWs(ref c) && (c==ft(TFormulaToken.fmPercent))) 
			{
				ConvertLastRefValueTypeOnce(TFmReturnType.Value, ParseState, ref First);
				SkipWhiteSpace();
				NextChar();
				AddParsedOp(TOperator.Percent);
			}
		}

		/// <summary>
		/// PerTerm [ ^ PerTerm]*
		/// </summary>
		private void ExpTerm(TParseState ParseState)   
		{
			PerTerm(ParseState);
			char c=' ';

			bool First = true;
			while (PeekCharWs(ref c) && (c==ft(TFormulaToken.fmPower)))
			{
				ConvertLastRefValueTypeOnce(TFmReturnType.Value, ParseState, ref First);
				SkipWhiteSpace();
				NextChar();
				PerTerm(ParseState);
				ConvertLastRefValueType(TFmReturnType.Value, ParseState);
				AddParsedOp(TOperator.Power);
			}
		}

		/// <summary>
		/// ExpTerm [ *|/ ExpTerm ]*
		/// </summary>
		private void MulTerm(TParseState ParseState)
		{
			char c=' ';
			ExpTerm(ParseState);

			bool First = true;
			while (PeekCharWs(ref c) && ((c==ft(TFormulaToken.fmMul)) || (c==ft(TFormulaToken.fmDiv))) )
			{
				ConvertLastRefValueTypeOnce(TFmReturnType.Value, ParseState, ref First);
				SkipWhiteSpace();
				NextChar();
				ExpTerm(ParseState);
				ConvertLastRefValueType(TFmReturnType.Value, ParseState);

				if (c==ft(TFormulaToken.fmMul)) AddParsedOp(TOperator.Mul); else AddParsedOp(TOperator.Div);
			}
		}

		/// <summary>
		/// MulTerm [ +|- MulTerm]*
		/// </summary>
		private void AddTerm(TParseState ParseState)
		{
			char c=' ';
			MulTerm(ParseState);

			bool First = true;
			while (PeekCharWs(ref c) && ((c==ft(TFormulaToken.fmPlus)) || (c==ft(TFormulaToken.fmMinus))) )
			{
				ConvertLastRefValueTypeOnce(TFmReturnType.Value, ParseState, ref First);
				SkipWhiteSpace();
				NextChar();
				MulTerm(ParseState);
				ConvertLastRefValueType(TFmReturnType.Value, ParseState);
				if (c==ft(TFormulaToken.fmPlus)) AddParsedOp(TOperator.Add); else AddParsedOp(TOperator.Sub);
			}
		}

		/// <summary>
		/// AddTerm [ &amp; AddTerm]*
		/// </summary>
		private void AndTerm(TParseState ParseState)   
		{
			char c=' ';
			AddTerm(ParseState);

			bool First = true;
			while (PeekCharWs(ref c) && c==ft(TFormulaToken.fmAnd)) 
			{
				ConvertLastRefValueTypeOnce(TFmReturnType.Value, ParseState, ref First);
				SkipWhiteSpace();
				NextChar();
				AddTerm(ParseState);
				ConvertLastRefValueType(TFmReturnType.Value, ParseState);
				AddParsedOp(TOperator.Concat);
			}
		}


		private bool FindComTerm(ref TOperator Ptg)
		{
			char c=' ';
			bool Result= PeekCharWs(ref c) && ((c==ft(TFormulaToken.fmEQ)) || (c==ft(TFormulaToken.fmLT)) || (c==ft(TFormulaToken.fmGT)));
			if (Result)
			{
				bool One=true;
				SkipWhiteSpace(); //Already granted we will add a ptg
				NextChar();
				char d=' ';
				if ( PeekChar(ref d) && ((d==ft(TFormulaToken.fmEQ)) || (d==ft(TFormulaToken.fmGT))) )
				{
					string s=c.ToString()+d; One=false;
					if (String.Equals(s,fts(TFormulaToken.fmGE),StringComparison.InvariantCulture)) {NextChar(); Ptg=TOperator.GE;}  
					else
						if (String.Equals(s,fts(TFormulaToken.fmLE), StringComparison.InvariantCulture)) {NextChar(); Ptg=TOperator.LE;}  
					else
						if (String.Equals(s,fts(TFormulaToken.fmNE), StringComparison.InvariantCulture)) {NextChar(); Ptg=TOperator.NE;}  
					else
						One=true;
				}
				if (One) 
					if (c == ft(TFormulaToken.fmEQ)) Ptg=TOperator.EQ; else
						if (c == ft(TFormulaToken.fmLT)) Ptg=TOperator.LT; else
						if (c == ft(TFormulaToken.fmGT)) Ptg=TOperator.GT; else
						FlxMessages.ThrowException(FlxErr.ErrInternal);
			}
			return Result;
		}

		/// <summary>
		/// AndTerm [ = | &lt; | &gt; | &lt;= | &gt;= | &lt;&gt;  AndTerm]*
		/// </summary>
		private void ComTerm(TParseState ParseState)    
		{
			char c=' ';
			TOperator Ptg= TOperator.Nop;
			AndTerm(ParseState);

			bool First = true;
			while (PeekCharWs(ref c) && FindComTerm(ref Ptg))
			{
				//no NextChar or SkipWhitespace here. It is added by FindComTerm
				ConvertLastRefValueTypeOnce(TFmReturnType.Value, ParseState, ref First);
				AndTerm(ParseState);
				ConvertLastRefValueType(TFmReturnType.Value, ParseState);
				AddParsedOp(Ptg);
			}
		}
       
		private void Expression(TParseState ParseState)
		{
			ComTerm(ParseState);
		}


        private void SkipSpace(TWhiteSpace Ws, out char c)
        {
            c = 'x';
            while (PeekChar(ref c) && (c == ' '))
            {
                NextChar();
                if (Ws.SpaceCount < 255) Ws.SpaceCount++;
            }
        }

        private void SkipEnter(TWhiteSpace Ws, out char c)
        {
            Ws.SpaceCount = 0; //spaces go after the last enter.

            c = 'x';
            char Lastc = c;
            while (PeekChar(ref c) && (c == '\n' || c == '\r'))
            {
                NextChar();
                if (Ws.EnterCount < 255 && (c == '\r' || Lastc != '\r')) Ws.EnterCount++;
                Lastc = c;
            }
        }


		protected void SkipWhiteSpace()
		{
            char c = 'x';
            TWhiteSpace Ws= new TWhiteSpace();
            do
            {
                SkipEnter(Ws, out c);
                SkipSpace(Ws, out c);
            }

            while (c == '\n' || c == '\r');

			if (ParsePos<Fw.Length)
			{
				c=Fw[ParsePos];
				
                if (c==ft(TFormulaToken.fmOpenParen)) {Ws.SpaceKind=FormulaAttr.bitFPreSpace; Ws.EnterKind = FormulaAttr.bitFPreEnter;}
                else if (c == ft(TFormulaToken.fmCloseParen)) { Ws.SpaceKind = FormulaAttr.bitFPostSpace; Ws.EnterKind = FormulaAttr.bitFPostEnter; }
                else {Ws.SpaceKind= FormulaAttr.bitFSpace; Ws.EnterKind = FormulaAttr.bitFEnter;}
 
                StackWs.Push(Ws);
			}
        }

		protected void UndoSkipWhiteSpace(int SaveParsePos)
		{
			StackWs.Pop();
			ParsePos=SaveParsePos;
		}

        protected void PopWhiteSpace()
        {
            TWhiteSpace Ws = StackWs.Pop();
            if (Ws.EnterCount > 0)
                AddParsedSpace(Ws.EnterCount, Ws.EnterKind);
            if (Ws.SpaceCount > 0)
                AddParsedSpace(Ws.SpaceCount, Ws.SpaceKind);
        }

        protected void DiscardNormalWhiteSpace()
        {
            TWhiteSpace Ws = StackWs.Pop();
            if (Ws.EnterCount > 0 && Ws.EnterKind != FormulaAttr.bitFEnter)
                AddParsedSpace(Ws.EnterCount, Ws.EnterKind);
            if (Ws.SpaceCount > 0 && Ws.SpaceKind != FormulaAttr.bitFSpace)
                AddParsedSpace(Ws.SpaceCount, Ws.SpaceKind);
        }

		protected void MakeLastWhitespaceNormal()
		{
			StackWs.NormalizeLastWhiteSpace();
		}

		#region AddParsed
		protected abstract void AddParsedUInt16(int w);
		protected abstract void AddParsed(double d);
		protected abstract void AddParsed(string s); 
		protected abstract void AddParsed(bool b);
		protected abstract void AddParsed(TFlxFormulaErrorValue err);
		protected abstract void AddParsedName(int NamePos);
		protected abstract void AddParsedExternName(int ExternSheet, int ExternName);
		protected abstract void AddParsedExternName(string ExternSheet, string ExternName);
        protected abstract void AddParsedRef(int Row, int Col, bool RowAbs, bool ColAbs);
        protected abstract void AddParsedRefErr();
		protected abstract void AddParsedArea(int Row1, int Row2, int Col1, int Col2, bool RowAbs1, bool RowAbs2, bool ColAbs1, bool ColAbs2);
        protected abstract void AddParsed3dRef(string ExternSheet, int Row, int Col, bool RowAbs, bool ColAbs);
        protected abstract void AddParsed3dRefErr(string ExternSheet);
        protected abstract void AddParsed3dArea(string ExternSheet, int Row1, int Row2, int Col1, int Col2, bool RowAbs1, bool RowAbs2, bool ColAbs1, bool ColAbs2);
		protected abstract void AddParsedSpace(byte Count, FormulaAttr Kind);
		protected abstract void AddParsedParen();
		protected abstract void AddParsedSep(byte b);
		protected abstract void AddParsedOp(TOperator op);
		protected abstract void AddParsedFunction(TCellFunctionData Func, byte ArgCount);
		protected abstract void AddParsedMissingArg();
		protected abstract void AddParsedArray(object[,] ArrayData);
		#endregion

		#region ConvertValueType	
		protected abstract void ConvertLastRefValueType(TFmReturnType RefMode, TParseState ParseState, bool IgnoreArray);
		protected void ConvertLastRefValueType(TFmReturnType RefMode, TParseState ParseState)
		{
            ConvertLastRefValueType(RefMode, ParseState, false);
		}
		protected void ConvertLastRefValueTypeOnce(TFmReturnType RefMode, TParseState ParseState, ref bool First)
		{
			if (First) ConvertLastRefValueType(RefMode, ParseState, false);
			First = false;
		}
		protected abstract bool LastIsReference();
		#endregion

		protected abstract TCellFunctionData FuncNameArray(string FuncName);

		protected virtual bool DoExtraToken(char c)
		{
			return false;
		}
		#region Public
		/// <summary>
		/// Does the encoding.
		/// </summary>
		protected void Go()
		{
            IsArrayFormula = false;

			char c= ' ';
            if (PeekChar(ref c) && (c==ft(TFormulaToken.fmOpenArray))) 
            {
                IsArrayFormula = true;
                NextChar();
            }

			if (!PeekChar(ref c) || (c!=ft(TFormulaToken.fmStartFormula))) FlxMessages.ThrowException(FlxErr.ErrFormulaStart,Fw);
			NextChar();
            Expression(new TParseState(IsArrayFormula, 0, false));
			ConvertLastRefValueType(InitialRefMode, new TParseState(false, 0, false));
            
            if (IsArrayFormula)
            {
                if (!PeekChar(ref c)) FlxMessages.ThrowException(FlxErr.ErrUnexpectedEof, Fw);
                if (c != ft(TFormulaToken.fmCloseArray)) FlxMessages.ThrowException(FlxErr.ErrUnexpectedEof, Fw);
                NextChar();
            }

            if (PeekChar(ref c)) FlxMessages.ThrowException(FlxErr.ErrUnexpectedChar,c, ParsePos, Fw);
			if (StackWs.Count!=0) FlxMessages.ThrowException(FlxErr.ErrInternal);
		}
		#endregion
	}


    internal struct TParseNumState
    {
        internal bool NumValid;
        internal bool LastNumValid;
        internal bool LastExp;
        internal bool FirstChar;
        internal int ExpCount;
        internal int DotCount;

        internal static TParseNumState Create()
        {
            TParseNumState Result = new TParseNumState();
            Result.NumValid = true;
            Result.LastNumValid = false;
            Result.LastExp = false;
            Result.FirstChar = true;
            Result.ExpCount = 0;
            Result.DotCount = 0;
            return Result;
        }
    }
}
