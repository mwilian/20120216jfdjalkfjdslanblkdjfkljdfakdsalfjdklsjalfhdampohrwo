using System;
using System.Resources;
using System.Reflection;
using System.Diagnostics;
using System.Runtime.Serialization;
using System.Security.Permissions;
using System.Security;

namespace FlexCel.Core
{
	/// <summary>
	/// Error Codes used in Exceptions.
	/// </summary>
	public enum FlxErr
	{
        /// <summary>
        /// Unexpected error
        /// </summary>
        ErrInternal, //keep it at pos 0.
        
        /// <summary>Can't convert the object into a valid value.</summary>
		ErrInvalidCellValue,

		/// <summary>
		/// Adapter's "Connect" Method has not been called.
		/// </summary>
		ErrNotConnected,

		/// <summary>
		/// Trying to access to an invalid row.
		/// </summary>
		ErrInvalidRow,

		/// <summary>
		/// Trying to access to an invalid column.
		/// </summary>
		ErrInvalidColumn,

		/// <summary>
		/// Generic out of range error. Params for this message should be taken from FlxParam.
		/// </summary>
		ErrInvalidValue,

        /// <summary>
        /// Generic out of range error when there are no objects to index.
        /// </summary>
        ErrInvalidValue2,

		/// <summary>
		/// Can't find a dataset for a named range.
		/// </summary>
		ErrDataSetNotFound,

		/// <summary>
		/// Sheet name does not exist.
		/// </summary>
		ErrInvalidSheet,

		/// <summary>
		/// Formula is longer than 0xFFFF
		/// </summary>
		ErrFormulaTooLong,

		/// <summary>
		/// When parsing a formula and expecting a string.
		/// </summary>
		ErrNotAString,

		/// <summary>
		/// A string missing the ending "
		/// </summary>
		ErrUnterminatedString,

		/// <summary>
		/// A formula with an invalid identifier.
		/// </summary>
		ErrUnexpectedId,

		/// <summary>
		/// A formula with an invalid character.
		/// </summary>
		ErrUnexpectedChar,

		/// <summary>
		/// A missing parenthesis.
		/// </summary>
		ErrMissingParen,

		/// <summary>
		/// Invalid number of parameters for the function.
		/// </summary>
		ErrInvalidNumberOfParams,

		/// <summary>
		/// Unexpected end of formula.
		/// </summary>
		ErrUnexpectedEof,

		/// <summary>
		/// Formula must start with "="
		/// </summary>
		ErrFormulaStart,

		/// <summary>
		/// Formula is invalid.
		/// </summary>
		ErrFormulaInvalid,

		/// <summary>
		/// Function on the formula does not exist.
		/// </summary>
		ErrFunctionNotFound,

		/// <summary>
		/// A formula refers to itself.
		/// </summary>
		ErrCircularReference,

		/// <summary>
		/// An invalid cell reference. A valid one is "A1"
		/// </summary>
		ErrInvalidRef,

		/// <summary>
		/// An invalid range reference. A valid one is "A1:A2"
		/// </summary>
		ErrInvalidRange,

        /// <summary>
        /// The name defining band "{0}" has a negative number of rows or columns. Verify it has absolute references ($A$1:$B$2) intead of relative ($A$1:B2)
        /// </summary>
        ErrInvalidRangeRef,

		/// <summary>
		/// An invalid End of tag (for example, &lt;#...&gt;&gt;)
		/// </summary>
		ErrMissingEOT,

		/// <summary>
		/// Missing arguments for a function.
		/// </summary>
		ErrMissingArgs,

		/// <summary>
		/// Too much arguments on a function.
		/// </summary>
		ErrTooMuchArgs,

		/// <summary>
		/// Reading after the end of a datatable.
		/// </summary>
		ErrReadAfterEOF,

		/// <summary>
		/// Reading before the beginning of a datatable.
		/// </summary>
		ErrReadBeforeBOF,

		/// <summary>
		/// 2 data ranges intersect on a report.
		/// </summary>
		ErrIntersectingRanges,

		///<summary>Function is not defined on file FunctionNames.resx</summary>
		ErrUndefinedFunction,

		///<summary>Id is not defined on file xxxx</summary>
		ErrUndefinedId,

		/// <summary>
		/// Invalid format on report.
		/// </summary>
		ErrInvalidFormat,

		/// <summary>
		/// Column does not exits.
		/// </summary>
		ErrColumNotFound,

		/// <summary>
		/// The property does not exist.
		/// </summary>
		ErrPropertyNotFound,

		/// <summary>
		/// Dataset does not have a field.
		/// </summary>
		ErrMemberNotFound,

		/// <summary>
		/// Dataset is not inside any named range.
		/// </summary>
		ErrDataSetNotInRange,

		/// <summary>
		/// Image is invalid.
		/// </summary>
		ErrInvalidImage,

        /// <summary>
        /// Error trying to create an empty image. Probably out of memory.
        /// </summary>
        ErrCreatingImage,

		/// <summary>
		/// Error in an included report.
		/// </summary>
		ErrOnIncludeReport,

		/// <summary>
		/// Too many nested includes, probably one file is recursively including itself.
		/// </summary>
		ErrTooManyNestedIncludes,

		/// <summary>
		/// Cannot find the included named range.
		/// </summary>
		ErrCantFindNamedRange,

		/// <summary>
		/// Range is not __, _, I_ or II_
		/// </summary>
		ErrUnknownRangeType,

		/// <summary>
		/// This feature is not supported on non commercial version of flexcel.
		/// </summary>
		ErrNotSupportedOnTrial,

		/// <summary>
		/// User defined function implementation is null.
		/// </summary>
		ErrUndefinedUserFunction,

		/// <summary>
		/// Can't find a dataset for a configsheet.
		/// </summary>
		ErrDataSetNotFoundInConfig,

		/// <summary>
		/// Can't find a dataset for an expression on the sheet.
		/// </summary>
		ErrDataSetNotFoundInExpression,

		/// <summary>
		/// RecalcMode can only be changed when no file is open.
		/// </summary>
		ErrCantChangeRecalcMode,

		/// <summary>
		/// Workbook can't be null on this version of the method.
		/// </summary>
		ErrWorkbookNull,

		/// <summary>
		/// Error with the compression engine.
		/// </summary>
		ErrCompression,

		/// <summary>
		/// A relationship is invalid.
		/// </summary>
		ErrInvalidRelationshipNullValues,

		/// <summary>
		/// A relationship is invalid.
		/// </summary>
		ErrInvalidRelationship2Tables,

		/// <summary>
		/// A relationship is invalid.
		/// </summary>
		ErrInvalidRelationshipFields,

		/// <summary>
		/// A relationship is invalid.
		/// </summary>
		ErrInvalidRelationshipDatasetNull,

        /// <summary>
        /// A relationship is invalid.
        /// </summary>
        ErrInvalidManualRelationshipDatasetNull,

        /// <summary>
        /// A relationship is invalid.
        /// </summary>
        ErrInvalidManualRelationshipFieldCount,

		/// <summary>
		/// A relationship is invalid.
		/// </summary>
		ErrInvalidRelationship2Fields,

		/// <summary>
		/// SQL function should have only 2 parameters.
		/// </summary>
		ErrInvalidSql2Params,

		/// <summary>
		/// SPLIT function should have only 2 parameters.
		/// </summary>
		ErrInvalidSplit2Params,

		/// <summary>
		/// Second parameter of a split function must be a positive integer.
		/// </summary>
		ErrInvalidSplitCount,

		/// <summary>
		/// TOP function should have only 2 parameters.
		/// </summary>
		ErrInvalidTop2Params,

		/// <summary>
		/// Second parameter of a TOP function must be a positive integer.
		/// </summary>
		ErrInvalidTopCount,

        /// <summary>
        /// ROWS function should have only 1 parameter.
        /// </summary>
        ErrInvalidNRows1Param,

        /// <summary>
        /// Parameter of a ROWS function must be a positive integer.
        /// </summary>
        ErrInvalidNRowsCount,

        /// <summary>
        /// COLUMNS function should have only 1 parameter.
        /// </summary>
        ErrInvalidColumns1Param,

		/// <summary>
		/// Missing closing parenthesis.
		/// </summary>
		ErrInvalidSqlParen,

        /// <summary>
        /// Missing closing parenthesis.
        /// </summary>
        ErrInvalidSplitParen,

        /// <summary>
        /// Missing closing parenthesis.
        /// </summary>
        ErrInvalidTopParen,

        /// <summary>
        /// Missing closing parenthesis.
        /// </summary>
        ErrInvalidNRowsParen,

        /// <summary>
        /// Missing closing parenthesis.
        /// </summary>
        ErrInvalidColumnsParen,

		/// <summary>
		/// Missing closing parenthesis.
		/// </summary>
		ErrInvalidUserTableParen,

		/// <summary>
		/// Adapter not found.
		/// </summary>
		ErrInvalidSqlAdapterNotFound,

		/// <summary>
		/// The adapter has not SQL select command defined.
		/// </summary>
		ErrInvalidSqlAdapterNoSelect,

		/// <summary>
		/// SQL string contains invalid characters.
		/// </summary>
		ErrInvalidSqlString,

		/// <summary>
		/// DataSet should not be null.
		/// </summary>
		ErrDataSetNull,

		/// <summary>
		/// Adapter should not be null.
		/// </summary>
		ErrAdapterNull,

		/// <summary>
		/// Missing parenthesis.
		/// </summary>
		ErrInvalidFilterParen,

		/// <summary>
		/// Error reading config tables.
		/// </summary>
		ErrOnConfigTables,

		/// <summary>
		/// Parameters for distinct filter should not be null.
		/// </summary>
		ErrInvalidDistinctParams,

		/// <summary>
		/// Parameters for SQL parameters should not be null.
		/// </summary>
		ErrInvalidSqlParams,

		/// <summary>
		/// The parameter for a direct sql query was not found.
		/// </summary>
		ErrSqlParameterNotFound,

		///<summary>EofReached</summary>
		ErrEofReached,

		/// <summary>All rows on an array must have the same number of columns.</summary>
		ErrArrayNotSquared,

		/// <summary>The columns on an array must be between 1 and 256</summary>
		ErrInvalidCols,

		/// <summary>The rows on an array must be between 1 and 65536</summary>
		ErrInvalidRows,

		/// <summary>Invalid error code</summary>
		ErrInvalidErrorCode,

		/// <summary>The folder does not contain any file of the needed type.</summary>
		ErrEmptyFolder,

		/// <summary>Cannot find a Tiff Encoder.</summary>
		ErrTiffEncoderNotFound,

		/// <summary>Invalid image format.</summary>
		ErrInvalidImageFormat,

		/// <summary>Invalid parameter for html tag.</summary>
		ErrInvalidHtmlParam,

		/// <summary>Invalid parameter for rowheight or columnwidth tag.</summary>
		ErrInvalidRowColParameters,

		/// <summary>Invalid second parameter for rowheight or columnwidth tag.</summary>
		ErrInvalidRowColParameters2,

		/// <summary>Invalid third parameter for rowheight or columnwidth tag.</summary>
		ErrInvalidRowColParameters3,
		
		/// <summary>Invalid fourth parameter for rowheight or columnwidth tag.</summary>
		ErrInvalidRowColParameters4,

		/// <summary>Invalid parameter for ref tag, 1 parameter version.</summary>
		ErrInvalidRefTag,

		/// <summary>Invalid parameter for ref tag, 2 parameter version.</summary>
		ErrInvalidRefTag2,

		/// <summary>Invalid parameter for ref tag, 3 parameter version.</summary>
		ErrInvalidRefTag3,

		/// <summary>
		/// Invalid parameter on Autofit settings tag.
		/// </summary>
		ErrInvalidGlobalAdjustment,

		/// <summary>
		/// Invalid parameter on Autofit settings tag.
		/// </summary>
		ErrInvalidGlobalAdjustmentFixed,

		/// <summary>
		/// Invalid adjustment on Autofit settings tag.
		/// </summary>
		ErrInvalidGlobalAutofit,

		/// <summary>
		/// Invalid Merge mode.
		/// </summary>
		ErrInvalidAutoFitMerged,

		/// <summary>
		/// When using Splitted tables, the detail must directly follow the master.
		/// </summary>
		ErrSplitNeedsOneAndOnlyOneDetail,

		/// <summary>
		/// A table cannot be assigned to 2 different splits masters.
		/// </summary>
		ErrSplitNeedsOnlyOneMaster,

		/// <summary>
		/// The template has User table tags, but there is no UserTable event assigned to the report.
		/// </summary>
		ErrUserTableEventNotAssigned,

		/// <summary>
		/// This is a virtual table, and it didn't define the methods needed to filter it.
		/// </summary>
		ErrTableDoesNotSupportFilter,

		/// <summary>
		/// This is a virtual table, and it didn't define the methods needed to support this tag.
		/// </summary>
		ErrTableDoesNotSupportTag,

		/// <summary>
		/// This is a virtual table, and it didn't define the methods needed to support master detail relationships.
		/// </summary>
		ErrTableDoesNotSupportMasterDetail,

		/// <summary>
		/// This is a virtual table, and it didn't define the methods needed to be used as a Lookpup source.
		/// </summary>
		ErrTableDoesNotSupportLookup,

		/// <summary>
		/// The encoding con not be used to create a MIME file.
		/// </summary>
		ErrInvalidEncodingForMIME,

		/// <summary>
		/// Before exporting a sheet you need to call BeginExport
		/// </summary>
		ErrBeginExportNotCalled,

		/// <summary>
		/// A preprocess tag must not be nested.
		/// </summary>
		ErrNoDuplicatedPreprocess,

		/// <summary>
		/// If the Defined tag has 2 parameters, the second must be GLOBAL or LOCAL.
		/// </summary>
		ErrInvalidDefinedGlobal,

		/// <summary>
		/// Invalide parameters for delete row/col
		/// </summary>
		ErrInvalidRowColDelete,

		/// <summary>
		/// The operation needs unamanaged permissions in order to complete.
		/// </summary>
		ErrNeedsUnmanaged,

		/// <summary>
		/// A format definition in the config sheet is not valid.
		/// </summary>
		ErrInvalidApplyTag,

		/// <summary>
		/// Error at cell n.
		/// </summary>
		ErrAtCell,

		/// <summary>
		/// Message that will be written in a cell when there is an error in a report and <see cref="FlexCel.Report.FlexCelReport.ErrorsInResultFile"/> is true.
		/// </summary>
		CellError,

		/// <summary>
		/// Invalid parameter for include.
		/// </summary>
		ErrInvalidIncludeStatic,

		/// <summary>
		/// Invalid parameter for include.
		/// </summary>
		ErrInvalidIncludeRowCol,

		/// <summary>
		/// Invalid value in the percent of page used in an automatic page break tag.
		/// </summary>
		ErrInvalidAutoPageBreaksPercent,

		/// <summary>
		/// Invalid PageScale in automatic page break tag.
		/// </summary>
		ErrInvalidAutoPageBreaksPageScale,

		/// <summary>
		/// Invalid parameter for a ImgPos tag.
		/// </summary>
		ErrInvalidImgPosParameter,

		/// <summary>
		/// Invalid parameter for a ImgFit tag.
		/// </summary>
		ErrInvalidImgFitParameter,

		/// <summary>
		/// Trying to add an xls file that already exists to a Workspace collection.
		/// </summary>
		ErrDuplicatedLinkedFile,

		/// <summary>
		/// Aggregate must be "SUM", "AVG", "MAX" or "MIN"
		/// </summary>
		ErrInvalidAggParameter,

		/// <summary>
		/// A filter must return a boolean value.
		/// </summary>
		ErrFilterMustReturnABooleanValue,

		/// <summary>
		/// Font was not found in the system.
		/// </summary>
		ErrFontNotFound,

		/// <summary>
		/// Font doesn't contain the character.
		/// </summary>
		ErrGlyphNotFound,

		/// <summary>
		/// Font is missing a character and a fallback font was used instead.
		/// </summary>
		ErrUsedFallbackFont,

		/// <summary>
		/// Font doesn't have italic or bold variant.
		/// </summary>
		ErrFauxBoldOrItalic,

        /// <summary>
        /// Font is not supported.
        /// </summary>
        ErrFontNotSupported,

		/// <summary>
		/// Name cannot be empty.
		/// </summary>
		ErrInvalidEmptyName,

		/// <summary>
		/// Name is too long.
		/// </summary>
		ErrNameTooLong,

		/// <summary>
		/// Named style doesn't exists.
		/// </summary>
		ErrStyleDoesntExists,

		/// <summary>
		/// Style already exists.
		/// </summary>
		ErrStyleAlreadyExists,

		/// <summary>
		/// Built-in styles can't be renamed.
		/// </summary>
		ErrCantRenameBuiltInStyle,

		/// <summary>
		/// Built-in styles can't be deleted.
		/// </summary>
		ErrCantDeleteBuiltInStyle,

		/// <summary>
		/// Style formats have to be added using SetStyle method, not with addformat.
		/// </summary>
		ErrCantAddStyleFormats,

		/// <summary>
		/// You can't replace a cell format with a style format or viceversa.
		/// </summary>
		ErrCantMixCellAndStyleFormats,

        /// <summary>
        /// ColorType is of the wrong type.
        /// </summary>
        ErrInvalidColorType,

        /// <summary>
        /// Value for the color is outside bounds.
        /// </summary>
        ErrInvalidColorValue,

        /// <summary>
        /// Value must be one of the enum, and can't be none.
        /// </summary>
        ErrInvalidColorEnum,

        /// <summary>
        /// String constant in formula is too long. (max 255 chars)
        /// </summary>
        ErrStringConstantInFormulaTooLong,

        /// <summary>
        /// When you set a border style different from none, then the border color must be set too.
        /// </summary>
        ErrNoBorderColorSet,

        /// <summary>
        /// You can use a themed color to define a color inside a theme.
        /// </summary>
        ErrCantUseThemeColorsInsideATheme,

        /// <summary>
        /// Parameter can't be null.
        /// </summary>
        ErrNullParameter,

        /// <summary>
        /// Parameter can't be null or empty.
        /// </summary>
        ErrNullOrEmptyParameter,

        /// <summary>
        /// The name for the range already exists.
        /// </summary>
        ErrRangeNameAlreadyExists,

        /// <summary>
        /// Name is invalid.
        /// </summary>
        ErrInvalidName,

        /// <summary>
        /// Enum must exist.
        /// </summary>
        ErrInvalidEnum,
    
        /// <summary>
        /// Detail table doesn't have a parent.
        /// </summary>
        ErrInvalidLinQDetail,

        /// <summary>
        /// The sort string is not supported.
        /// </summary>
        ErrInvalidSortString,

        /// <summary>
        /// The keys in a lookup can't be empty.
        /// </summary>
        ErrEmptyKeyNames,

        /// <summary>
        /// Keys and values in a lookup must have the same number of elements.
        /// </summary>
        ErrValuesAndKeysMismatch,

        /// <summary>
        /// When using LINQ, FlexCel implements filters by calling a "Where(string)" method in the data table. If the table doesn't
        /// have a "Where(string)" method, FlexCel can't filter the records.
        /// </summary>
        ErrDatasetDoesntSupportWhere,

        /// <summary>
        /// The count of records in a table is different from the actual number of records when reading the data. This means that 
        /// someone else has inserted or deleted records while this report was running. To avoid this issue, make sure
        /// you run the report with isolation level = SNAPSHOT.
        /// </summary>
        ErrInvalidReportRowCount,

        /// <summary>
        ///Dates in filters must be written as "yyy-dd-mm hh:mm:ss"
        /// </summary>
        ErrInvalidFilterDateTime


	}

	/// <summary>
	/// Some custom strings.
	/// </summary>
	public enum FlxMessage
	{
		/// <summary>
		/// Name of a custom page size.
		/// </summary>
		CustomPageSize,

		/// <summary>
		/// Text that will be shown on chart legends when no data is assigned.
		/// </summary>
		TxtSeries,

		/// <summary>
		/// Default text to show when AM time string is set to empty on the regional settings dialog from the control panel.
		/// When the default is empty, Excel defaults to AM, while .NET defaults to an empty string.
		/// As we want to behave the same as Excel, we provide this constant here.
		/// </summary>
		TxtDefaultTimeAMString,

		/// <summary>
		/// Default text to show when PM time string is set to empty on the regional settings dialog from the control panel.
		/// When the default is empty, Excel defaults to PM, while .NET defaults to an empty string.
		/// As we want to behave the same as Excel, we provide this constant here.
		/// </summary>
		TxtDefaultTimePMString

	}
		
	/// <summary>
	/// FlexCel Native XLS Constants. It reads the resources from the active locale, and
	/// returns the correct string.
	/// If your language is not supported and you feel like translating the messages,
	/// please send us a copy. We will include it on the next FlexCel version. 
	/// <para>To add a new language:
	/// <list type="number">
	/// <item>
	///    Copy the file flxmsg.resx to your language (for example, flxmsg.es.resx to translate to spanish)
	/// </item><item>
	///    Edit the new file and change the messages(you can do this visually with visual studio)
	/// </item><item>
	///    Add the .resx file to the FlexCel project
	/// </item>
	/// </list>
	/// </para>
	/// </summary>
	public sealed class FlxMessages
	{
		private FlxMessages(){}
		internal static readonly ResourceManager rm = new ResourceManager("FlexCel.Core.flxmsg", Assembly.GetExecutingAssembly()); //STATIC*

        /// <summary>
        /// Returns a string based on the FlxErr enumerator, formatted with args.
        /// This method is used to get an Exception error message. 
        /// </summary>
        /// <param name="ResName">Error Code.</param>
        /// <param name="args">Params for this error.</param>
        /// <returns></returns>
		public static string GetString( FlxErr ResName, params object[] args)
		{
			if (args.Length==0) return rm.GetString(ResName.ToString()); //To test without args
			return (String.Format(rm.GetString(ResName.ToString()), args));
		}

		/// <summary>
		/// Returns a string from the FlxMessage enumeration.
		/// </summary>
		/// <param name="ResName">Message code.</param>
		/// <returns>Associated string.</returns>
		public static string GetString(FlxMessage ResName)
		{
			return rm.GetString(ResName.ToString());
		}

        /// <summary>
        /// Throws a standard FlexCelException.
        /// </summary>
        /// <param name="ResName">Error Code.</param>
        /// <param name="args">Parameters for this error.</param>
		public static void ThrowException(FlxErr ResName, params object[] args)
		{
			throw new FlexCelCoreException(GetString(ResName, args), ResName);
		}
    
        /// <summary>
        /// Throws a standard FlexCelException with innerException.
        /// </summary>
        /// <param name="e">Inner exception.</param>
        /// <param name="ResName">Error Code.</param>
        /// <param name="args">Parameters for this error.</param>
        public static void ThrowException(Exception e, FlxErr ResName, params object[] args)
        {
            throw new FlexCelCoreException(GetString(ResName, args), ResName, e);
        }
    }

#if (!FRAMEWORK40)
    /// <summary>
    /// Internal use when not using .NET 4.0
    /// </summary>
    interface ISafeSerializationData
    {
        /// <summary>
        /// Dummy method to replace the exising in .NET 4.0
        /// </summary>
        /// <param name="obj"></param>
        void CompleteDeserialization(object obj);
    }
#endif

	/// <summary>
	/// Exception thrown when an specific FlexCel error happens. Base of all FlexCel hierarchy list.
	/// </summary>
    [Serializable]
    public class FlexCelException : Exception
    {
        /// <summary>
        /// Creates a new FlexCelException
        /// </summary>
        public FlexCelException(): base(){}
        
        /// <summary>
        /// Creates a new FlexCelException with an error message.
        /// </summary>
        /// <param name="message">Error Message</param>
        public FlexCelException(string message): base(message){}

        /// <summary>
        /// Creates a nested Exception.
        /// </summary>
        /// <param name="message">Error Message.</param>
        /// <param name="inner">Inner Exception.</param>
        public FlexCelException(string message, Exception inner):base(message, inner) {}

#if(!COMPACTFRAMEWORK && !SILVERLIGHT && !FRAMEWORK40)
        /// <summary>
        /// Creates an exception from a serialization context.
        /// </summary>
        /// <param name="info">Serialization information.</param>
        /// <param name="context">Streaming Context.</param>
		[SecurityPermissionAttribute(SecurityAction.Demand,SerializationFormatter=true)]
        protected FlexCelException(SerializationInfo info, StreamingContext context) : base(info, context) { }

        /// <summary>
        /// Implements standard GetObjectData.
        /// </summary>
        /// <param name="info"></param>
        /// <param name="context"></param>
		[SecurityPermissionAttribute(SecurityAction.Demand,SerializationFormatter=true)]
        public override void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            base.GetObjectData(info, context);
        }
#endif

    }

    /// <summary>
    /// Exception thrown when an exception on the core engine happens. 
    /// </summary>
    [Serializable]
    public class FlexCelCoreException : FlexCelException
    {
#if (FRAMEWORK40)
        [NonSerialized]
#endif
        private FlexCelCoreExceptionState FState;

        private void InitState(FlxErr aErrorCode)
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
        /// Creates a new FlexCelCoreException
        /// </summary>
        public FlexCelCoreException() : base() { InitState(FlxErr.ErrInternal); }
        
        /// <summary>
        /// Creates a new FlexCelCoreException with an error message.
        /// </summary>
        /// <param name="message">Error Message</param>
        public FlexCelCoreException(string message) : base(message) { InitState(FlxErr.ErrInternal); }

        /// <summary>
        /// Creates a new FlexCelCoreException with an error message and an exception code.
        /// </summary>
        /// <param name="message">Error Message</param>
        /// <param name="aErrorCode">Error code of the exception.</param>
        public FlexCelCoreException(string message, FlxErr aErrorCode) : base(message) { InitState(aErrorCode); }

        /// <summary>
        /// Creates a nested Exception.
        /// </summary>
        /// <param name="message">Error Message.</param>
        /// <param name="inner">Inner Exception.</param>
        public FlexCelCoreException(string message, Exception inner) : base(message, inner) { InitState(FlxErr.ErrInternal); }

        /// <summary>
        /// Creates a nested Exception.
        /// </summary>
        /// <param name="message">Error Message.</param>
        /// <param name="inner">Inner Exception.</param>
        /// <param name="aErrorCode">Error code of the exception.</param>
        public FlexCelCoreException(string message, FlxErr aErrorCode, Exception inner) : base(message, inner) { InitState(aErrorCode); }

#if(!COMPACTFRAMEWORK && !SILVERLIGHT && !FRAMEWORK40)
        /// <summary>
        /// Creates an exception from a serialization context.
        /// </summary>
        /// <param name="info">Serialization information.</param>
        /// <param name="context">Streaming Context.</param>
		[SecurityPermissionAttribute(SecurityAction.Demand,SerializationFormatter=true)]
        protected FlexCelCoreException(SerializationInfo info, StreamingContext context)
            : base(info, context)
		{
			if (info != null) FState.ErrorCode= (FlxErr)info.GetInt32("FErrorCode");
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
        public FlxErr ErrorCode
        {
            get
            {
                return FState.ErrorCode;
            }
        }

        [Serializable]
        private struct FlexCelCoreExceptionState : ISafeSerializationData
        {
            private FlxErr FErrorCode;

            public FlxErr ErrorCode
            {
                get { return FErrorCode; }
                set { FErrorCode = value; }
            }

            void ISafeSerializationData.CompleteDeserialization(object obj)
            {
                FlexCelCoreException ex = obj as FlexCelCoreException;
                ex.FState = this;
            }
        }
    }

}
