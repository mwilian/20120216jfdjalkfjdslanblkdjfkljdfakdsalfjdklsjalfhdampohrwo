using System;
using System.Resources;
using System.Globalization;
using System.Reflection;
using System.Diagnostics;

namespace FlexCel.Core
{
	/*
     *   Constants referred to formulas.
	 *
	 *   Resources on this unit are not localized, to avoid having different
	 *   interfaces to formulas for different languages.
	 * 
	 *    Of course, you are free to translate them to your language,
	 *    so a user can read or write the formula text the same way
	 *    he does it on his Excel version. But keep in mind that if you
	 *    write a formula in your code, you will have to write it in your language,
	 *    and if you later compile your app on other language, it will not work.
	 * 
	 */ 


	/// <summary>
	/// Known token on a formula.  They are not supposed to be localized, but they can be by editing formulamsg.resx
	/// </summary>
	public enum TFormulaToken
	{
		///<summary>'#';</summary>
        fmErrStart, 

		///<summary>'TRUE';</summary>
        fmTrue, 
		///<summary>'FALSE';</summary>
        fmFalse, 

		///<summary>'=' // as in =a1*2 or {=1+2}</summary>
        fmStartFormula, 

		///<summary>'.' as in '1.2'</summary>
        fmFormulaDecimal, 
		///<summary>',' as in '1,300'</summary>
        fmFormulaThousands, 

		///<summary>',' Argument separator on a function.
        ///For example, if fmFunctionSep=';' we should
        ///write "Max(a1;a2)"
        ///If you want to localize this, you could use fmFunctionSep=ListSeparator
        ///</summary>
        fmFunctionSep, 

		///<summary>',' as in "=a1, b2"</summary>
        fmUnion      , 
		///<summary>' ' as in a1 a2</summary>
        fmIntersect  , 

		///<summary>'{';</summary>
        fmOpenArray, 
		///<summary>'}';</summary>
        fmCloseArray, 
		///<summary>'(';</summary>
        fmOpenParen, 
		///<summary>')';</summary>
        fmCloseParen, 
		///<summary>'/' Separates 2 rows on an array. It is '\' in spanish</summary>
        fmArrayRowSep, 
		///<summary>',' Separates 2 columns on an array. Ex: {1,2}. It is ';' in spanish
		///If you want to localize this, you could use fmArrayColSep=ListSeparator
		///</summary>
        fmArrayColSep, 
		

		///<summary>'$' as in $A$3</summary>
        fmAbsoluteRef, 
		///<summary>':' as in A1:A3</summary>
        fmRangeSep,

        ///<summary>' (as in 'Sheet 1'!A1</summary>
        fmSingleQuote,
        ///<summary>'!' as in Sheet1!a1</summary>
        fmExternalRef, 
		///<summary>'[' as in c:\[book1.xls]Sheet1!a1</summary>
        fmWorkbookOpen, 
		///<summary>']';</summary>
        fmWorkbookClose, 

		///<summary>'TABLE';</summary>
        fmTableText, 

		//those here  shouldn't change
		///<summary>'+';</summary>
        fmPlus, 
		///<summary>'-';</summary>
        fmMinus, 
		///<summary>'*';</summary>
        fmMul, 
		///<summary>'/';</summary>
        fmDiv, 
		///<summary>'^';</summary>
        fmPower, 
		///<summary>'%';</summary>
        fmPercent, 
		///<summary>'"';</summary>
        fmStr, 

		///<summary>'&amp;';</summary>
        fmAnd, 

		///<summary>'&lt;';</summary>
        fmLT, 
		///<summary>'&lt;=';</summary>
        fmLE, 
		///<summary>'=';</summary>
        fmEQ, 
		///<summary>'&gt;=';</summary>
        fmGE, 
		///<summary>'&gt;';</summary>
        fmGT, 
		///<summary>'&lt;&gt;';</summary>
        fmNE,

        ///<summary>'R' in R1C1 notation as in R1C2</summary>
        fmR1C1_R,
        ///<summary>'C' in R1C1 notation as in R1C2</summary>
        fmR1C1_C,
        ///<summary>'[' in R1C1 notation as in R[1]C[-2]</summary>
        fmR1C1RelativeRefStart,
        ///<summary>']' in R1C1 notation as in R[1]C[-2]</summary>
        fmR1C1RelativeRefEnd
	}

	internal enum FormulaAttr
	{
		bitFSpace=0x00,
		bitFEnter=0x01,
		bitFPreSpace=0x02,
		bitFPreEnter=0x03,
		bitFPostSpace=0x04,
		bitFPostEnter=0x05,
		bitFPreFmlaSpace=0x06
	}

	/// <summary>
	/// Formula return types... A value, an array or a reference.
	/// </summary>
	public enum TFmReturnType{
        /// <summary>
        /// Formula returns a value.
        /// </summary>
        Value, 
        /// <summary>
        /// Formula returns a reference.
        /// </summary>
        Ref, 
        /// <summary>
        /// Formula returns an array.
        /// </summary>
        Array
        };

    /// <summary>
    /// The function description. Only use it if you have to override the formula parser.
    /// </summary>
	internal class TCellFunctionData
	{
		private int FIndex;
		private string FName;
		private int FMinArgCount;
		private int FMaxArgCount;
        private bool FVolatile;
		private TFmReturnType FReturnType;
        private string FParamType;
		private bool FDoesNotAlterArray;
  
        public int Index {get {return FIndex;}}
        public string Name {get {return FName;}}
		public int MinArgCount {get {return FMinArgCount;} }
        public bool FutureInXls;
        public bool FutureInXlsx;

        public int MaxArgCount
        {
            get
            {
                if (FlxConsts.ExcelVersion == TExcelVersion.v97_2003 && FMaxArgCount > FlxConsts.Max_FormulaArguments2003) return FlxConsts.Max_FormulaArguments2003;
                return FMaxArgCount;
            }
        }
		
        public bool Volatile {get {return FVolatile;}}
		public bool DoesNotAlterArray {get {return FDoesNotAlterArray;}}
        public TFmReturnType ReturnType {get {return FReturnType;}}
		public string ParamTypeStr {get {return FParamType;}}
        
		public TFmReturnType ParamType(int pos) 
		{
            int Par = FParamType.IndexOf("(");
            int ParLen = FParamType.Length - 1;
            if (Par >= 0 && pos >= Par)
            {
                pos++;
                ParLen--;
            }

            if (pos > ParLen)
            {
                if (Par < 0)
                {
                    pos = FParamType.Length - 1;
                }
                else
                {
                    int RepLen = FParamType.Length - Par - 2;
                    pos = Par + 1 + (pos - Par - 1) % RepLen;
                }
            }
			switch (FParamType[pos])
			{
				case 'A': return TFmReturnType.Array;
				case 'R': return TFmReturnType.Ref;
				case 'V': return TFmReturnType.Value;
				case '-': return TFmReturnType.Value; //Missing Arg.
			}
			FlxMessages.ThrowException(FlxErr.ErrInternal);
			return TFmReturnType.Value;  //just to please compiler.
		}

        public TCellFunctionData(int aIndex, string aName, int aMinArgCount, int aMaxArgCount, bool aNonVolatile, TFmReturnType aReturnType, string aParamType)
        {
            FIndex            =  aIndex;
            FName             =  aName;
			FMinArgCount      =  aMinArgCount;
			FMaxArgCount      =  aMaxArgCount;
            FVolatile         =  !aNonVolatile;
            FReturnType       =  aReturnType;
            FParamType        =  aParamType;
        }
		public TCellFunctionData(int aIndex, string aName, int aMinArgCount, int aMaxArgCount, bool aNonVolatile, TFmReturnType aReturnType, string aParamType, bool aDoesNotAlterArray):
			this(aIndex, aName, aMinArgCount, aMaxArgCount, aNonVolatile, aReturnType, aParamType)
		{
			FDoesNotAlterArray = aDoesNotAlterArray;
		}

        public TCellFunctionData(int aIndex, string aName, int aMinArgCount, int aMaxArgCount, bool aNonVolatile, TFmReturnType aReturnType, string aParamType, bool aFutureInXls, bool aFutureInXlsx) :
            this(aIndex, aName, aMinArgCount, aMaxArgCount, aNonVolatile, aReturnType, aParamType)
        {
            FutureInXls = aFutureInXls;
            FutureInXlsx = aFutureInXlsx;
        }

        public string FutureName
        {
            get { return FutureInXls ? "_xlfn." + Name : null; }
        }


	}

    /// <summary>
    /// Tokens that can be used on a formula.
    /// </summary>
	public sealed class TFormulaMessages
	{
		internal static readonly ResourceManager rm = new ResourceManager("FlexCel.Core.formulamsg", Assembly.GetExecutingAssembly()); //STATIC*
		private static readonly string[] FmlaToken = InitFmlaToken(); //STATIC*
        private static readonly string[] FmlaErr = InitFmlaErr();   //STATIC*

        private TFormulaMessages(){}

        /// <summary>
        /// Message for the ErrorCode.
        /// </summary>
        /// <param name="ErrCode">Error Code.</param>
        /// <returns>Message.</returns>
        public static string ErrString(TFlxFormulaErrorValue ErrCode)
        {
            return FmlaErr[(int)ErrCode];
        }

		private static string[] InitFmlaErr()
		{               
			Array ErrCodes = TCompactFramework.EnumGetValues(typeof(TFlxFormulaErrorValue));
			string[] Result = new string[64]; //It should be the real number but...
			foreach (TFlxFormulaErrorValue ErrCode in ErrCodes)
			{
				Result[(int)ErrCode] = rm.GetString("fm"+ErrCode.ToString());
			}

			return Result;
		}

        /// <summary>
        /// Formula tokens.
        /// </summary>
        /// <param name="Token">Token.</param>
        /// <returns>Message.</returns>
        public static string TokenString(TFormulaToken Token)
        {
            return FmlaToken[(int)Token];
        }

		private static string[] InitFmlaToken()
		{
			Array Tokens = TCompactFramework.EnumGetValues(typeof(TFormulaToken));
			string[] Result = new string[255];
			foreach (TFormulaToken Token in Tokens)
			{
				string s = rm.GetString(Token.ToString());
				if ((s.Length==0)) s = " "; //Xml resource does not save a single space!
				Result[(int)Token]=s;
			}
			return Result;
		}

        /// <summary>
        /// Returns the formula token as a character.
        /// </summary>
        /// <param name="Token">Token.</param>
        /// <returns>Message</returns>
        public static char TokenChar(TFormulaToken Token)
        {
            string s=TokenString(Token);
            if (s.Length>0) return s[0]; else return (char)0;
        }

		internal static byte StringToErrCode(string ErrString, bool RaiseException)
		{
			//We don't use Enum.GetValues here to enumerate all, as it isn't compatible with CF.
		{
			if (ErrString == TFormulaMessages.ErrString(TFlxFormulaErrorValue.ErrDiv0)) return (byte)TFlxFormulaErrorValue.ErrDiv0;
			if (ErrString == TFormulaMessages.ErrString(TFlxFormulaErrorValue.ErrNA)) return (byte)TFlxFormulaErrorValue.ErrNA;
			if (ErrString == TFormulaMessages.ErrString(TFlxFormulaErrorValue.ErrName)) return (byte)TFlxFormulaErrorValue.ErrName;
			if (ErrString == TFormulaMessages.ErrString(TFlxFormulaErrorValue.ErrNull)) return (byte)TFlxFormulaErrorValue.ErrNull;
			if (ErrString == TFormulaMessages.ErrString(TFlxFormulaErrorValue.ErrNum)) return (byte)TFlxFormulaErrorValue.ErrNum;
			if (ErrString == TFormulaMessages.ErrString(TFlxFormulaErrorValue.ErrRef)) return (byte)TFlxFormulaErrorValue.ErrRef;
			if (ErrString == TFormulaMessages.ErrString(TFlxFormulaErrorValue.ErrValue)) return (byte)TFlxFormulaErrorValue.ErrValue;
		}
			if (RaiseException) FlxMessages.ThrowException(FlxErr.ErrInvalidErrorCode, ErrString);
			return (byte)TFlxFormulaErrorValue.ErrNA;
		}


		/// <summary>
		///  This is a non-localized version of FloatToStr
		///  It will always use "." as decimal separator.
		///  If you are localizing this unit to your language, change this function
		///  to be:
		///  public string FloatToString(double Value)
		///  {
		///	    return Value.ToString();
		///  }
		///  
		/// And it will use your current locale to get the decimal separator.
		/// Just remember that if you for example use "," as decimal sep,
		/// you should also change fmArrayColSep, fmFunctionSep and all vars with value=","
		/// </summary>
		/// <param name="Value">Value to convert</param>
		/// <returns>String using ALWAYS "." as decimal separator, regardless of the regional settings</returns>
		public static string FloatToString(double Value)
		{
			return Value.ToString(NumberFormatInfo.InvariantInfo);
		}

		/// <summary>
		///  This is a non-localized version of StrToFloat
		///  It will always use "." as decimal separator.
		///  If you are localizing this unit to your language, change this function
		///  to be:
		///  public double FloatToString(string Value)
		///  {
		///	    return Value.ToDouble();
		///  }
		///  
		/// And it will use your current locale to get the decimal separator.
		/// Just remember that if you for example use "," as decimal sep,
		/// you should also change fmArrayColSep, fmFunctionSep and all vars with value=","
		/// </summary>
		/// <param name="Value">String to Convert. ALWAYS uses "." as decimal separator</param>
		/// <returns>value of the string</returns>
		public static double StringToFloat(string Value)
		{
			return Convert.ToDouble(Value, NumberFormatInfo.InvariantInfo);
		}
	}
}
