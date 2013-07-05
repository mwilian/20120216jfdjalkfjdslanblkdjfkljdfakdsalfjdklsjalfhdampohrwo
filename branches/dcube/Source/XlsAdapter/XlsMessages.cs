using System;
using System.Resources;
using System.Reflection;
using System.Threading;
using System.Diagnostics;
using System.Runtime.Serialization;
using System.Security.Permissions;
using FlexCel.Core;
using System.Security;


namespace FlexCel.XlsAdapter
{
	/// <summary>
	/// Error Codes. We use this and not actual strings to make sure all are correctly spelled.
	/// </summary>
	public enum XlsErr
	{
        ///<summary>Internal</summary>
        ErrInternal,
        ///<summary>TooManyEntries</summary>
		ErrTooManyEntries,
		///<summary>InvalidContinue</summary>
		ErrInvalidContinue,
		///<summary>WrongExcelRecord</summary>
		ErrWrongExcelRecord,
		///<summary>ExcelInvalid</summary>
		ErrExcelInvalid,
		///<summary>FileIsPasswordProtected</summary>
		ErrFileIsPasswordProtected,
		///<summary>NotSupportedOnCE</summary>
		ErrNotSupportedOnCE,
		///<summary>FileIsNotXLS</summary>
		ErrFileIsNotXLS,
		///<summary>BadFormula</summary>
		ErrBadFormula,
		///<summary>BadName</summary>
		ErrBadName,
		///<summary>BadToken</summary>
		ErrBadToken,
		///<summary>WrongType</summary>
		ErrWrongType,
		///<summary>InvalidStream</summary>
		ErrInvalidStream,
		///<summary>EofReached</summary>
		ErrEofReached,
		///<summary>ReadingRecord</summary>
		ErrReadingRecord,
		///<summary>InvalidStringRecord</summary>
		ErrInvalidStringRecord,
		///<summary>InvalidVersion</summary>
		ErrInvalidVersion,
		///<summary>DupRow</summary>
		ErrDupRow,
		///<summary>RowMissing</summary>
		ErrRowMissing,
		///<summary>EscherNotLoaded</summary>
		ErrEscherNotLoaded,
		///<summary>LoadingEscher</summary>
		ErrLoadingEscher,
		///<summary>BStoreDuplicated</summary>
		ErrBStoreDuplicated,
		///<summary>DgDuplicated</summary>
		ErrDgDuplicated,
        ///<summary>DggDuplicated</summary>
        ErrDggDuplicated,
		///<summary>SolverDuplicated</summary>
		ErrSolverDuplicated,
		///<summary>ChangingEscher</summary>
		ErrChangingEscher,
		///<summary>NotImplemented</summary>
		ErrNotImplemented,
		///<summary>CantCopyPictFmla</summary>
		ErrCantCopyPictFmla,
        ///<summary>BadChartFormula</summary>
        ErrBadChartFormula,
        ///<summary>SectionNotLoaded</summary>
        ErrSectionNotLoaded,
        ///<summary>InvalidDrawing</summary>
        ErrInvalidDrawing,
        ///<summary>XlsIndexOutBounds</summary>
        ErrXlsIndexOutBounds,
        ///<summary>BadCF</summary>
        ErrBadCF,
        ///<summary>InvalidCF</summary>
        ErrInvalidCF,
        ///<summary>InvalidSheetNo</summary>
        ErrInvalidSheetNo,
        ///<summary>DuplicatedSheetName</summary>
        ErrDuplicatedSheetName,
        ///<summary>TooManyPageBreaks</summary>
        ErrTooManyPageBreaks,
        ///<summary>InvalidRow</summary>
        ErrInvalidRow,
        ///<summary>InvalidCol</summary>
        ErrInvalidCol,
        ///<summary>BadRowCount</summary>
        ErrBadRowCount,
        ///<summary>ShrFmlaNotFound</summary>
        ErrShrFmlaNotFound,
        ///<summary>HiddenSheetSelected</summary>
        ErrHiddenSheetSelected,
        ///<summary>NoSheetVisible</summary>
        ErrNoSheetVisible,
		///<summary>BadCopyRows</summary>
		ErrBadCopyRows,
		///<summary>BadCopyCols</summary>
		ErrBadCopyCols,
		///<summary>BadMoveCall</summary>
		ErrBadMoveCall,
        ///<summary>eSheetName</summary>
        BaseSheetName,
        ///<summary>CantDeleteSheetWithMacros</summary>
        ErrCantDeleteSheetWithMacros,
        ///<summary>TooManyRows</summary>
        ErrTooManyRows,
        ///<summary>TooManyColumns</summary>
        ErrTooManyColumns,

        /// <summary>Trying to insert more than 65536 sheets.</summary>
        ErrTooManySheets,

		///<summary>Invalid password.</summary>
		ErrInvalidPassword,
		///<summary>Password too long.</summary>
		ErrPasswordTooLong,
		/// <summary>Encryption method is not supported. </summary>
        ErrNotSupportedEncryption,
        /// <summary>The name for a named range is invalid. It should have no more than 255 characters, must no start with a number, and must not contain some special characters. </summary>
        ErrInvalidNameForARange,
		/// <summary>Could not find the object path inside the shape.</summary>
		ErrObjectNotFound,
		/// <summary>The array element is not any of the supported types.</summary>
		ErrInvalidArrayElement,

        /// <summary>The Pocket Excel file (pxl) is not on a format FlexCel can understand.</summary>
        ErrPxlIsInvalid,

        /// <summary>The Pocket Excel file (pxl) is not on a format FlexCel can understand. Token invalid.</summary>
        ErrPxlIsInvalidToken,

        /// <summary>The string is longer than the maximum allowed by Excel.</summary>
        ErrStringTooLong,

        /// <summary>The string can't be empty.</summary>
        ErrStringEmpty,

        /// <summary>The string for a header or footer is longer than the maximum allowed by Excel.</summary>
        ErrHeaderFooterStringTooLong,
        
        /// <summary>The file is not on any on the formats supported by FlexCel (Excel 97 or newer, pxl). Note that to read or write xlsx files (Excel 2007 or newer)
        /// you need FlexCel for .NET Framework 3.5 or newer.</summary>
        ErrFileIsNotSupported,

        /// <summary>Pocket Excel does not have support for formulas that reference external files.</summary>
        ErrPxlDoesNotHaveExternalFormulas,

		/// <summary>Cannot read the properties of this file.</summary>
		ErrInvalidPropertySector,

		/// <summary>Chart is invalid.</summary>
		ErrInvalidChart,

		/// <summary>
		/// Ranges in move calls cannot intersect.
		/// </summary>
		ErrMoveRangesCanNotIntersect,

		/// <summary>
		/// Can't move a part of a table.
		/// </summary>
		ErrCantMovePartOfTable,

		/// <summary>
		/// Can't move a part of an Array formula.
		/// </summary>
		ErrCantMovePartOfArrayFormula,

		/// <summary>
		/// Trying to move a range outside bounds.
		/// </summary>
		ErrMoveRangeOutsideBounds,

		/// <summary>
		/// Too many Conditional format rules.
		/// </summary>
		ErrTooManyCFRules,

        /// <summary>
        /// Too many format definitions.
        /// </summary>
        ErrTooManyXFDefs,

		/// <summary>
		/// First formula of Data validation has more than 255 characters.
		/// </summary>
		ErrDataValidationFmla1TooLong,

		/// <summary>
		/// Second Formula of Data validation has more than 255 characters.
		/// </summary>
		ErrDataValidationFmla2TooLong,

		/// <summary>
		/// First formula of Data validation is null.
		/// </summary>
		ErrDataValidationFmla1Null,

		/// <summary>
		/// Second formula of Data validation is null.
		/// </summary>
		ErrDataValidationFmla2Null,

        /// <summary>
        /// Invalid type of Hyperlink.
        /// </summary>
        ErrInvalidHyperLinkType,

        /// <summary>
        /// Maximum length for a format string is 255 characters.
        /// </summary>
        ErrInvalidFormatStringLength,
        
        /// <summary>
        /// Too many custom numeric formats.
        /// </summary>
        ErrInvalidFormatId,

        /// <summary>
        /// In order to save as xlsx, the stream needs to have read access besides write.
        /// </summary>
        ErrStreamNeedsReadAccess,

        /// <summary>
        /// If you specify a gradient pattern, then gradient definition can't be null
        /// </summary>
        ErrNullGradient,

        /// <summary>
        /// Both sheets in the range must be the same
        /// </summary>
        ErrRangeMustHaveSameSheet,

        /// <summary>
        /// There is a part missing from the xlsx file.
        /// </summary>
        ErrMissingPart
}

	/// <summary>
	/// FlexCel Native XLS Constants. It reads the resources from the active locale, and
	/// returns the correct string.
	/// If your language is not supported and you feel like translating the messages,
	/// please send us a copy. We will include it on the next FlexCel version. 
	/// <para>To add a new language:
	/// <list type="number">
	/// <item>
	///    Copy the file xlsmsg.resx to your language (for example, xlsmsg.es.resx to translate to spanish)
	/// </item><item>
	///    Edit the new file and change the messages(you can do this visually with visual studio)
    /// </item><item>
    ///    Add the .resx file to the FlexCel project
    /// </item>
    /// </list>
	/// </para>
	/// </summary>
    internal sealed class XlsMessages
    {
        private XlsMessages(){}
        internal static readonly ResourceManager rm = new ResourceManager("FlexCel.XlsAdapter.xlsmsg", Assembly.GetExecutingAssembly()); //STATIC*
        public static string GetString( XlsErr ResName, params object[] args)
        {
			if (args.Length==0) return rm.GetString(ResName.ToString()); //To test without args
            return (String.Format(rm.GetString(ResName.ToString()), args));
        } 
		public static void ThrowException( XlsErr ResName, params object[] args)
		{
			throw new FlexCelXlsAdapterException(GetString(ResName, args), ResName); 
		}
    }


    /// <summary>
    /// Exception thrown when an exception on the XlsAdapter engine happens. 
    /// </summary>
    [Serializable]
    public class FlexCelXlsAdapterException : FlexCelException
    {
#if (FRAMEWORK40)
        [NonSerialized]
#endif
        private FlexCelXlsAdapterExceptionState FState;

        private void InitState(XlsErr aErrorCode)
        {
            FState.ErrorCode = aErrorCode;
#if(FRAMEWORK40)
            SerializeObjectState += delegate(object exception,
                SafeSerializationEventArgs eventArgs)
            {
                eventArgs.AddSerializedState(FState);
            };
#endif
        }

        /// <summary>
        /// Creates a new FlexCelXlsAdapterException
        /// </summary>
        public FlexCelXlsAdapterException() : base() { InitState(XlsErr.ErrInternal); }
        
        /// <summary>
        /// Creates a new FlexCelXlsAdapterException with an error message.
        /// </summary>
        /// <param name="message">Error Message.</param>
        public FlexCelXlsAdapterException(string message) : base(message) { InitState(XlsErr.ErrInternal); }

        /// <summary>
        /// Creates a new FlexCelXlsAdapterException with an error message and an exception code.
        /// </summary>
        /// <param name="message">Error Message</param>
        /// <param name="aErrorCode">Error code of the exception.</param>
        public FlexCelXlsAdapterException(string message, XlsErr aErrorCode) : base(message) { InitState(aErrorCode); }

        /// <summary>
        /// Creates a nested Exception.
        /// </summary>
        /// <param name="message">Error Message.</param>
        /// <param name="inner">Inner Exception.</param>
        public FlexCelXlsAdapterException(string message, Exception inner) : base(message, inner) { InitState(XlsErr.ErrInternal); }

#if(!COMPACTFRAMEWORK && !SILVERLIGHT && !FRAMEWORK40)
        /// <summary>
        /// Creates an exception from a serialization context.
        /// </summary>
        /// <param name="info">Serialization information.</param>
        /// <param name="context">Streaming Context.</param>
		[SecurityPermissionAttribute(SecurityAction.Demand,SerializationFormatter=true)]
        protected FlexCelXlsAdapterException(SerializationInfo info, StreamingContext context)
            : base(info, context)
		{
			if (info != null) FState.ErrorCode= (XlsErr)info.GetInt32("FErrorCode");
		}

		/// <summary>
		/// Implements standard GetObjectData.
		/// </summary>
		/// <param name="info"></param>
		/// <param name="context"></param>
		[SecurityPermissionAttribute(SecurityAction.Demand,SerializationFormatter=true)]
        public override void GetObjectData(SerializationInfo info, StreamingContext context)
		{
			if (info != null) info.AddValue("FErrorCode", (int)FState.ErrorCode);
			base.GetObjectData (info, context);
		}

#endif

        /// <summary>
        /// Error code on the Exception.
        /// </summary>
        public XlsErr ErrorCode
        {
            get
            {
                return FState.ErrorCode;
            }
        }

        [Serializable]
        private struct FlexCelXlsAdapterExceptionState : ISafeSerializationData
        {
            private XlsErr FErrorCode;

            public XlsErr ErrorCode
            {
                get { return FErrorCode; }
                set { FErrorCode = value; }
            }

            void ISafeSerializationData.CompleteDeserialization(object obj)
            {
                FlexCelXlsAdapterException ex = obj as FlexCelXlsAdapterException;
                ex.FState = this;
            }
        }
    }

    /// <summary>
    /// Exception thrown when an error parsing a formula happens.
    /// </summary>
  	[Serializable]
	public class ETokenException : FlexCelXlsAdapterException
	{
#if (FRAMEWORK40)
        [NonSerialized]
#endif
        private ETokenExceptionState FState;

        private void InitState(int aToken)
        {
            FState.Token = aToken;
#if(FRAMEWORK40)
            SerializeObjectState += delegate(object exception,
                SafeSerializationEventArgs eventArgs)
            {
                eventArgs.AddSerializedState(FState);
            };
#endif
        }
        /// <summary>
        /// Formula token with the error.
        /// </summary>
        public int Token{get {return FState.Token;}}

        /// <summary>
        /// Creates an empty ETokenException.
        /// </summary>
        public ETokenException() : base(XlsMessages.GetString(XlsErr.ErrBadToken, 0), XlsErr.ErrBadToken) { InitState(0); }

        /// <summary>
        /// Creates a new ETokenException for a specific Token.
        /// </summary>
        /// <param name="aToken"></param>
        public ETokenException(int aToken) : base(XlsMessages.GetString(XlsErr.ErrBadToken, aToken), XlsErr.ErrBadToken) { InitState(aToken); }

        /// <summary>
        /// Creates a new ETokenException with a message.
        /// </summary>
        /// <param name="message"></param>
        public ETokenException(string message) : base(message, XlsErr.ErrBadToken) { InitState(0); }

        /// <summary>
        /// Creates a nested Exception.
        /// </summary>
        /// <param name="message">Error Message.</param>
        /// <param name="inner">Inner Exception.</param>
        public ETokenException(string message, Exception inner):base(message, inner) { }

#if(!COMPACTFRAMEWORK && !SILVERLIGHT && !FRAMEWORK40)
        /// <summary>
        /// Creates an exception from a serialization context.
        /// </summary>
        /// <param name="info">Serialization information.</param>
        /// <param name="context">Streaming Context.</param>
		[SecurityPermissionAttribute(SecurityAction.Demand,SerializationFormatter=true)]
        protected ETokenException(SerializationInfo info, StreamingContext context): base(info, context)
		{
			if (info != null) FState.Token = info.GetInt32("FToken");
        }

		/// <summary>
		/// Implements standard GetObjectData.
		/// </summary>
		/// <param name="info"></param>
		/// <param name="context"></param>
		[SecurityPermissionAttribute(SecurityAction.Demand,SerializationFormatter=true)]
        public override void GetObjectData(SerializationInfo info, StreamingContext context)
		{
			if (info != null) info.AddValue("FToken", FState.Token);
			base.GetObjectData (info, context);
		}

#endif

        [Serializable]
        private struct ETokenExceptionState : ISafeSerializationData
        {
            private int FToken;

            public int Token
            {
                get { return FToken; }
                set { FToken = value; }
            }

            void ISafeSerializationData.CompleteDeserialization(object obj)
            {
                ETokenException ex = obj as ETokenException;
                ex.FState = this;
            }
        }
    }
}
