using System;
using System.Resources;
using System.Globalization;
using System.Reflection;
using FlexCel.Core;

using System.Diagnostics;
using System.Collections.Generic;

namespace FlexCel.Report
{
	internal enum TagEnum
	{
		FStrOpen=0,
		FStrClose=1,
		FParamDelim=2,
		FStrOpenParen=3,
		FStrCloseParen=4,
		FStrQuote=5,
		FDbSeparator=6,
		FRowRange1=7,
		FRowRange2=8,
		FColRange1=9,
		FColRange2=10,
		FRowFull1=11,
		FRowFull2=12,
		FColFull1=13,
		FColFull2=14,
		FCrossTabRange1=15,
		FCrossTabRange2=16,
		FCrossTabFull1=17,
		FCrossTabFull2=18,
		FStrFullDs=19,
		FStrFullDsCaptions=20,
		FStrOpenHLink=21,
		FStrDeleteLastRow=22,
		FStrExcludeSheet=23,
		FStrOpenHLink2=24,
		FStrCloseHLink=25,
		FRelationshipSeparator=26,
		FInternalDB=27,
		FStrRowCountColumn=28,
		FStrRowPosColumn=29,
		FStrShow = 30,
		FStrHide = 31,
		FStrAutofit = 32,
		FStrDontInsertRanges = 33,
		FStrAutofitOn = 34,
		FStrAutofitOff = 35,
		FStrKeepAutofit = 36,
		FStrDontKeepAutofit = 37,
		FStrDefinedLocal = 38,
		FStrDefinedGlobal = 39,
		FStrRelativeDelete = 40,
		FStrFullDelete = 41,
		FStrStaticInclude = 42,
		FStrDynamicInclude = 43,
		FStrCopyRows = 44,
		FStrCopyCols = 45,
	    FStrCopyRowsAndCols = 46,

        FKeepRowsTogether = 47,
        FKeepColsTogether = 48,

		FStrEndKeepTogether = 49,

		FStrAlignLeft = 50,
		FStrAlignCenter = 51,
		FStrAlignRight = 52,
		FStrAlignTop = 53,
		FStrAlignBottom = 54,

		FStrInRow = 55,
		FStrInCol = 56,

		FStrDontGrow = 57,
		FStrDontShrink = 58,

		FStrAggSum = 59,
		FStrAggAvg = 60,
		FStrAggMax = 61,
		FStrAggMin = 62,

		FStrAutofitModeFirst = 63,
		FStrAutofitModeLast = 64,
		FStrAutofitModeNone = 65,
		FStrAutofitModeBalanced = 66
	    
	}


	//This is public so names do not change when using an obfuscator.
	/// <summary>
	/// Configuration strings. Internal use.
	/// </summary>
    public enum ConfigTagEnum
    {
        /// <summary>
        /// SQL tag.
        /// </summary>
        SQL = 0,

        /// <summary>
        /// Relationship tag.
        /// </summary>
        Relationship = 1,

        /// <summary>
        /// Distinct filter.
        /// </summary>
        Distinct = 2,

        /// <summary>
        /// Identifier for sql parameters on the config sheet.
        /// </summary>
        SQLParam = 3,

        /// <summary>
        /// Split Tag on config sheet.
        /// </summary>
        Split = 4,

        /// <summary>
        /// User table tag on config sheet.
        /// </summary>
        UserTable = 5,

        /// <summary>
        /// Top n records in a dataset.
        /// </summary>
        Top = 6,

        /// <summary>
        /// An empty table with n rows.
        /// </summary>
        NRows = 7,

        /// <summary>
        /// A table with the columns of other table.
        /// </summary>
        Columns = 8
    }

	//This is public so names do not change when using an obfuscator.
	/// <summary>
	/// Configuration strings used in applying a part of a format. Internal use.
	/// </summary>
	public enum ApplyFormatTagEnum
	{
		/// <summary></summary>
		All,
        /// <summary></summary>
		Border,
        /// <summary></summary>
        BorderLeft,
        /// <summary></summary>
        BorderRight,
        /// <summary></summary>
        BorderTop,
		/// <summary></summary>        
		BorderBottom,
		/// <summary></summary>        
		BorderExterior,
		/// <summary></summary>
        Font,
        /// <summary></summary>
        FontFamily,
        /// <summary></summary>
        FontSize,
        /// <summary></summary>
        FontColor,
        /// <summary></summary>
        FontStyle,
        /// <summary></summary>
        FontUnderline,
        /// <summary></summary>
        NumericFormat,
        /// <summary></summary>
        Background,
        /// <summary></summary>
        BackgroundPattern,
        /// <summary></summary>
        BackgroundColor,
        /// <summary></summary>
        TextAlign,
        /// <summary></summary>
        TextAlignHoriz,
        /// <summary></summary>
        TextAlignVert,
        /// <summary></summary>
        Locked,
        /// <summary></summary>
        Hidden,
        /// <summary></summary>
        TextWrap,
        /// <summary></summary>
        ShrinkToFit,
        /// <summary></summary>
        Rotation,
        /// <summary></summary>
        TextIndent
	}

	/// <summary>
	/// Tags used in reports.
	/// </summary>
	public sealed class ReportTag
	{
		private ReportTag(){}

		private static readonly string[] Tags = CreateTags();  //STATIC* 
		private static readonly string[] ConfigTags = CreateConfigTags();  //STATIC* 
		private static readonly Dictionary<string, ApplyFormatTagEnum> ApplyFormatTags = CreateApplyFormatTags(); //STATIC* 
		private static readonly string[] ConfigTagsParams = CreateConfigTagsParams(); //STATIC* 
        private static readonly Dictionary<string, TValueType> FTagTable = CreateTagTable(); //STATIC*
        private static readonly Dictionary<string, int> FTagParams = CreateTagParams();  //STATIC*  


		#region Basic tags
		/// <summary>Open a Tag.</summary>
		public static string StrOpen {get {return Tags[(int)TagEnum.FStrOpen];}}

		/// <summary>**WARNING**Excel2003 does not let you write this either on hyperlinks, so use StrOpenHLink2.
		/// Open an Hyperlink Tag. As we can't use # on hyperlink texts, this gives an alternative.
		/// </summary>
		public static string StrOpenHLink {get {return Tags[(int)TagEnum.FStrOpenHLink];}}

		/// <summary>Open an Hyperlink Tag. As we can't use # on hyperlink texts, this gives an alternative.</summary>
		public static string StrOpenHLink2 {get {return Tags[(int)TagEnum.FStrOpenHLink2];}}

		/// <summary>Close a Tag.</summary>
		public static string StrClose {get {return Tags[(int)TagEnum.FStrClose];}}
		
		/// <summary>Close an Hyperlink Tag. As we can't use # on hyperlink texts, this gives an alternative.</summary>
		public static string StrCloseHLink {get {return Tags[(int)TagEnum.FStrCloseHLink];}}

		/// <summary>Function delimiter. (for example: "&lt;#if(xx ; yy ; zz)&gt;</summary>
		public static char ParamDelim {get {return Tags[(int)TagEnum.FParamDelim][0];}}

		/// <summary>Open Parenthesis.</summary>
		public static char StrOpenParen {get {return Tags[(int)TagEnum.FStrOpenParen][0];}}

		/// <summary>Close Parenthesis.</summary>
		public static char StrCloseParen {get {return Tags[(int)TagEnum.FStrCloseParen][0];}}

		/// <summary>Quote.</summary>
		public static char StrQuote {get {return Tags[(int)TagEnum.FStrQuote][0];}}

		/// <summary>"."</summary>
		public static string DbSeparator {get {return Tags[(int)TagEnum.FDbSeparator];}}

		/// <summary>"*"</summary>
		public static string StrFullDs {get {return Tags[(int)TagEnum.FStrFullDs];}}

		/// <summary>"**"</summary>
		public static string StrFullDsCaptions {get {return Tags[(int)TagEnum.FStrFullDsCaptions];}}

		/// <summary>"X"</summary>
		public static string StrDeleteLastRow {get {return Tags[(int)TagEnum.FStrDeleteLastRow];}}

		/// <summary>"FIXED"</summary>
		public static string StrDontInsertRanges {get {return Tags[(int)TagEnum.FStrDontInsertRanges];}}

		/// <summary>"."</summary>
		public static string StrExcludeSheet {get {return Tags[(int)TagEnum.FStrExcludeSheet];}}
		
		/// <summary>"->"</summary>
		public static string RelationshipSeparator {get {return Tags[(int)TagEnum.FRelationshipSeparator];}}

		/// <summary>"__##INT_RNAL_"</summary>
		public static string InternalDB {get {return Tags[(int)TagEnum.FInternalDB];}}
        
		/// <summary>"#RecordCount"</summary>
		public static string StrRowCountColumn {get {return Tags[(int)TagEnum.FStrRowCountColumn];}}

		/// <summary>"#RecordPos"</summary>
		public static string StrRowPosColumn {get {return Tags[(int)TagEnum.FStrRowPosColumn];}}

		/// <summary>"Show"</summary>
		public static string StrShow {get {return Tags[(int)TagEnum.FStrShow];}}

		/// <summary>"Hide"</summary>
		public static string StrHide {get {return Tags[(int)TagEnum.FStrHide];}}

		/// <summary>"Autofit"</summary>
		public static string StrAutofit {get {return Tags[(int)TagEnum.FStrAutofit];}}

		/// <summary>"All"</summary>
		public static string StrAutofitOn {get {return Tags[(int)TagEnum.FStrAutofitOn];}}

		/// <summary>"Selected"</summary>
		public static string StrAutofitOff {get {return Tags[(int)TagEnum.FStrAutofitOff];}}

		/// <summary>"Keep"</summary>
		public static string StrKeepAutofit {get {return Tags[(int)TagEnum.FStrKeepAutofit];}}

		/// <summary>"Fixed"</summary>
		public static string StrDontKeepAutofit {get {return Tags[(int)TagEnum.FStrDontKeepAutofit];}}

		/// <summary>"First"</summary>
		public static string StrAutofitModeFirst {get {return Tags[(int)TagEnum.FStrAutofitModeFirst];}}

		/// <summary>"Last"</summary>
		public static string StrAutofitModeLast {get {return Tags[(int)TagEnum.FStrAutofitModeLast];}}

		/// <summary>"None"</summary>
		public static string StrAutofitModeNone {get {return Tags[(int)TagEnum.FStrAutofitModeNone];}}

		/// <summary>"Balanced"</summary>
		public static string StrAutofitModeBalanced {get {return Tags[(int)TagEnum.FStrAutofitModeBalanced];}}

		/// <summary>"Local"</summary>
		public static string StrDefinedLocal {get {return Tags[(int)TagEnum.FStrDefinedLocal];}}

        /// <summary>"Global"</summary>
        public static string StrDefinedGlobal { get { return Tags[(int)TagEnum.FStrDefinedGlobal]; } }

        /// <summary>"Relative"</summary>
        public static string StrRelativeDelete { get { return Tags[(int)TagEnum.FStrRelativeDelete]; } }

		/// <summary>"Full"</summary>
		public static string StrFullDelete { get { return Tags[(int)TagEnum.FStrFullDelete]; } }

		/// <summary>"Static"</summary>
		public static string StrStaticInclude { get { return Tags[(int)TagEnum.FStrStaticInclude]; } }

		/// <summary>"Dynamic"</summary>
		public static string StrDynamicInclude { get { return Tags[(int)TagEnum.FStrDynamicInclude]; } }

		/// <summary>"R"</summary>
		public static string StrCopyRows { get { return Tags[(int)TagEnum.FStrCopyRows]; } }

		/// <summary>"C"</summary>
		public static string StrCopyCols { get { return Tags[(int)TagEnum.FStrCopyCols]; } }
		
		/// <summary>"RC"</summary>
		public static string StrCopyRowsAndCols { get { return Tags[(int)TagEnum.FStrCopyRowsAndCols]; } }

		/// <summary>"LEFT"</summary>
		public static string StrAlignLeft { get { return Tags[(int)TagEnum.FStrAlignLeft]; } }

		/// <summary>"Center"</summary>
		public static string StrAlignCenter { get { return Tags[(int)TagEnum.FStrAlignCenter]; } }

		/// <summary>"Right"</summary>
		public static string StrAlignRight { get { return Tags[(int)TagEnum.FStrAlignRight]; } }

		/// <summary>"Top"</summary>
		public static string StrAlignTop { get { return Tags[(int)TagEnum.FStrAlignTop]; } }

		/// <summary>"Bottom"</summary>
		public static string StrAlignBottom { get { return Tags[(int)TagEnum.FStrAlignBottom]; } }

		/// <summary>"InRow"</summary>
		public static string StrInRow { get { return Tags[(int)TagEnum.FStrInRow]; } }

		/// <summary>"InCol"</summary>
		public static string StrInCol { get { return Tags[(int)TagEnum.FStrInCol]; } }

		/// <summary>"DontGrow"</summary>
		public static string StrDontGrow { get { return Tags[(int)TagEnum.FStrDontGrow]; } }

		/// <summary>"DontShrink"</summary>
		public static string StrDontShrink { get { return Tags[(int)TagEnum.FStrDontShrink]; } }

		/// <summary>"Sum"</summary>
		public static string StrAggSum { get { return Tags[(int)TagEnum.FStrAggSum]; } }

		/// <summary>"Avg"</summary>
		public static string StrAggAvg { get { return Tags[(int)TagEnum.FStrAggAvg]; } }

		/// <summary>"Max"</summary>
		public static string StrAggMax { get { return Tags[(int)TagEnum.FStrAggMax]; } }

		/// <summary>"Min"</summary>
		public static string StrAggMin { get { return Tags[(int)TagEnum.FStrAggMin]; } }

		#endregion

		#region Ranges
		/// <summary>Range delimiter.</summary>
		public static string RowRange1 {get {return Tags[(int)TagEnum.FRowRange1];}}

		/// <summary>Range delimiter.</summary>
		public static string RowFull1 {get {return Tags[(int)TagEnum.FRowFull1];}}

		/// <summary>Range delimiter.</summary>
		public static string RowRange2 {get {return Tags[(int)TagEnum.FRowRange2];}}

		/// <summary>Range delimiter.</summary>
		public static string RowFull2 {get {return Tags[(int)TagEnum.FRowFull2];}}

		/// <summary>Range delimiter.</summary>
		public static string ColRange1 {get {return Tags[(int)TagEnum.FColRange1];}}

		/// <summary>Range delimiter.</summary>
		public static string ColFull1 {get {return Tags[(int)TagEnum.FColFull1];}}

		/// <summary>Range delimiter.</summary>
		public static string ColRange2 {get {return Tags[(int)TagEnum.FColRange2];}}

		/// <summary>Range delimiter.</summary>
		public static string ColFull2 {get {return Tags[(int)TagEnum.FColFull2];}}

		/// <summary>Range delimiter.</summary>
		public static string CrossTabRange1 {get {return Tags[(int)TagEnum.FCrossTabRange1];}}

		/// <summary>Range delimiter.</summary>
		public static string CrossTabRange2 {get {return Tags[(int)TagEnum.FCrossTabRange2];}}

		/// <summary>Range delimiter.</summary>
		public static string CrossTabFull1 {get {return Tags[(int)TagEnum.FCrossTabFull1];}}

		/// <summary>Range delimiter.</summary>
		public static string CrossTabFull2 {get {return Tags[(int)TagEnum.FCrossTabFull2];}}

        /// <summary>
        /// Named Range starting with KeepRows_
        /// </summary>
        public static string KeepRowsTogether { get { return Tags[(int)TagEnum.FKeepRowsTogether]; } }

        /// <summary>
        /// Named Range starting with KeepColumns_
        /// </summary>
        public static string KeepColsTogether { get { return Tags[(int)TagEnum.FKeepColsTogether]; } }

		/// <summary>
		/// End of tag for keeprowstogether and keepcolstogether
		/// </summary>
		public static string StrEndKeepTogether { get { return Tags[(int)TagEnum.FStrEndKeepTogether]; } }

		#endregion

		#region Tag names
		/// <summary>All available tags.</summary>
		public static bool TryGetTag(string TagId, out TValueType ResultValue)
		{
			return FTagTable.TryGetValue(TagId, out ResultValue);
		}

		/// <summary>
		/// List of tag ids.
		/// </summary>
		public static ICollection<string> TagTableKeys
		{
			get 
			{
				return FTagTable.Keys;
			}
		}

		/// <summary>Number of params for all available tags.</summary>
		public static bool TryGetTagParams(string key, out int ResultValue)
		{
    		return FTagParams.TryGetValue(key, out ResultValue);
		}
    
 
		private static readonly string FStrEqual = GetStrEqual(); //STATIC* Not optimum, but it works and is thread safe..  
		/// <summary>=</summary>
		public static string StrEqual {get{return FStrEqual;}}

		private static readonly string FStrConfigSheet = GetStrConfig(); //STATIC* Set on resources. 
		/// <summary>CONFIG</summary>
		public static string StrConfigSheet {get {return FStrConfigSheet;}} 

		private static readonly string FStrDebug = GetStrDebug(); //STATIC* Set on resources. 

		/// <summary>DEBUG</summary>
		public static string StrDebug {get {return FStrDebug;}}

        private static readonly string FStrColWithRowCount = GetStrColWithRowCount();
        internal static string StrColWithRowCount { get { return FStrColWithRowCount; } } 

		
		private static readonly string FStrErrorsInResultFile = GetStrErrorsInResultFile(); //STATIC* Set on resources. 
		/// <summary>ERRORSINRESULTFILE</summary>
		public static string StrErrorsInResultFile {get {return FStrErrorsInResultFile;}} 

		private static readonly string FormatDelete = GetFormatDelete(); //STATIC* Set on resources. 
		private static readonly string FormatAdd = GetFormatAdd(); //STATIC* Set on resources. 
		#endregion

		#region Config Tag names
		/// <summary>Returns a particular configuration tag.</summary>
		public static string ConfigTag(ConfigTagEnum tag)
		{
			return ConfigTags[(int)tag];
		}

		/// <summary>Modifies a format.(For example "Font-Name").</summary>
		internal static void ApplyFormatTag(string tag, string fullExpression, TFlxApplyFormat fmt, ref bool exteriorBorders)
		{
			bool add = true;
			if(tag.StartsWith(FormatDelete, StringComparison.InvariantCultureIgnoreCase))
			{
				tag = tag.Remove(0, FormatDelete.Length);
				add = false;
			}

			if(tag.StartsWith(FormatAdd, StringComparison.InvariantCultureIgnoreCase))
			{
				tag = tag.Remove(0, FormatAdd.Length);
			}

            ApplyFormatTagEnum r;
            if (!ApplyFormatTags.TryGetValue(tag, out r)) 
            {
                FlxMessages.ThrowException(FlxErr.ErrInvalidApplyTag, tag, fullExpression);
            }

            switch (r)
            {
				case ApplyFormatTagEnum.All:
					fmt.SetAllMembers(add);
					break;
                case ApplyFormatTagEnum.Border:
                    fmt.Borders.SetAllMembers(add);
                    break;
                case ApplyFormatTagEnum.BorderLeft:
                    fmt.Borders.Left = add;
                    break;
                case ApplyFormatTagEnum.BorderRight:
                    fmt.Borders.Right = add;
                    break;
                case ApplyFormatTagEnum.BorderTop:
                    fmt.Borders.Top = add;
                    break;
				case ApplyFormatTagEnum.BorderBottom:
					fmt.Borders.Bottom = add;
					break;
				case ApplyFormatTagEnum.BorderExterior:
					exteriorBorders = add;
					break;
				case ApplyFormatTagEnum.Font:
                    fmt.Font.SetAllMembers(add);
                    break;
                case ApplyFormatTagEnum.FontFamily:
                    fmt.Font.Name = add;
                    break;
                case ApplyFormatTagEnum.FontSize:
                    fmt.Font.Size20 = add;
                    break;
                case ApplyFormatTagEnum.FontColor:
                    fmt.Font.Color = add;
                    break;
                case ApplyFormatTagEnum.FontStyle:
                    fmt.Font.Style = add;
                    break;
                case ApplyFormatTagEnum.FontUnderline:
                    fmt.Font.Underline = add;
                    break;
                case ApplyFormatTagEnum.NumericFormat:
                    fmt.Format = add;
                    break;
                case ApplyFormatTagEnum.Background:
                    fmt.FillPattern.SetAllMembers(add);
                    break;
                case ApplyFormatTagEnum.BackgroundPattern:
                    fmt.FillPattern.Pattern = add;
                    break;
                case ApplyFormatTagEnum.BackgroundColor:
                    fmt.FillPattern.FgColor = add;
                    fmt.FillPattern.BgColor = add;
                    fmt.FillPattern.Gradient = add;
                    break;
                case ApplyFormatTagEnum.TextAlign:
                    fmt.HAlignment = add;
                    fmt.VAlignment = add;
                    break;
                case ApplyFormatTagEnum.TextAlignHoriz:
                    fmt.HAlignment = add;
                    break;
                case ApplyFormatTagEnum.TextAlignVert:
                    fmt.VAlignment = add;
                    break;
                case ApplyFormatTagEnum.Locked:
                    fmt.Locked = add;
                    break;
                case ApplyFormatTagEnum.Hidden:
                    fmt.Hidden = add;
                    break;
                case ApplyFormatTagEnum.TextWrap:
                    fmt.WrapText = add;
                    break;
                case ApplyFormatTagEnum.ShrinkToFit:
                    fmt.ShrinkToFit = add;
                    break;
                case ApplyFormatTagEnum.Rotation:
                    fmt.Rotation = add;
                    break;
                case ApplyFormatTagEnum.TextIndent:
                    fmt.Indent = add;
                    break;
                default:
					FlxMessages.ThrowException(FlxErr.ErrUndefinedId, tag, "TagNames.resx");
					break;
            }

		}

		/// <summary>Returns the parameters for a particular configuration tag.</summary>
		public static string ConfigTagParams(ConfigTagEnum tag)
		{
			return ConfigTagsParams[(int)tag];
		}
 
		#endregion
		
		#region String loaders
		#region Tags
        private static Dictionary<string, TValueType> CreateTagTable()
		{
			ResourceManager rm = new ResourceManager("FlexCel.Report.TagNames", Assembly.GetExecutingAssembly());
			Dictionary<string, TValueType> Result = new Dictionary<string,TValueType>(StringComparer.InvariantCultureIgnoreCase);
			Result.Add(rm.GetString("DELETEROW"), TValueType.DeleteRow);
			Result.Add(rm.GetString("DELETERANGE"), TValueType.DeleteRange);
			Result.Add(rm.GetString("DELETECOLUMN"), TValueType.DeleteCol);

			Result.Add(rm.GetString("FORMATCELL"), TValueType.FormatCell);
			Result.Add(rm.GetString("FORMATROW"), TValueType.FormatRow);
			Result.Add(rm.GetString("FORMATCOLUMN"), TValueType.FormatCol);
			Result.Add(rm.GetString("FORMATRANGE"), TValueType.FormatRange);

			Result.Add(rm.GetString("STRIF"), TValueType.IF);
			Result.Add(rm.GetString("STREQUAL"), TValueType.Equal);
			Result.Add(rm.GetString("INCLUDE"), TValueType.Include);

			Result.Add(rm.GetString("HPAGEBREAK"), TValueType.HPageBreak);
			Result.Add(rm.GetString("VPAGEBREAK"), TValueType.VPageBreak);

			Result.Add(rm.GetString("CONFIG"), TValueType.ConfigSheet);
			Result.Add(rm.GetString("DELETESHEET"), TValueType.DeleteSheet);

			Result.Add(rm.GetString("COMMENT"), TValueType.Comment);

			Result.Add(rm.GetString("EVALUATE"), TValueType.Evaluate);
			Result.Add(rm.GetString("IMGSIZE"), TValueType.ImgSize);
			Result.Add(rm.GetString("IMGPOS"), TValueType.ImgPos);
			Result.Add(rm.GetString("IMGFIT"), TValueType.ImgFit);
			Result.Add(rm.GetString("IMGDELETE"), TValueType.ImgDelete);

			Result.Add(rm.GetString("LOOKUP"), TValueType.Lookup);
			Result.Add(rm.GetString("ARRAY"), TValueType.Array);
			
			Result.Add(rm.GetString("REGEX"), TValueType.Regex);

			Result.Add(rm.GetString("MERGERANGE"), TValueType.MergeRange);

			Result.Add(rm.GetString("FORMULA"), TValueType.Formula);
			Result.Add(rm.GetString("COLUMNWIDTH"), TValueType.ColumnWidth);
			Result.Add(rm.GetString("ROWHEIGHT"), TValueType.RowHeight);
			Result.Add(rm.GetString("AUTOFITSETTINGS"), TValueType.AutofitSettings);
			Result.Add(rm.GetString("HTML"), TValueType.Html);
			Result.Add(rm.GetString("REF"), TValueType.Ref);
            Result.Add(rm.GetString("DEFINED"), TValueType.Defined);
            Result.Add(rm.GetString("DEFINEDFORMAT"), TValueType.DefinedFormat);
            Result.Add(rm.GetString("PREPROCESS"), TValueType.Preprocess);
			Result.Add(rm.GetString("AUTOPAGEBREAKS"), TValueType.AutoPageBreaks);
			Result.Add(rm.GetString("AGGREGATE"), TValueType.Aggregate);
			Result.Add(rm.GetString("LIST"), TValueType.List);
			Result.Add(rm.GetString("DBVALUE"), TValueType.DbValue);

			return Result;
            
		}

		private static string GetStrEqual()
		{
			ResourceManager rm = new ResourceManager("FlexCel.Report.TagNames", Assembly.GetExecutingAssembly());
			return rm.GetString("STREQUAL");
		}

		private static string GetStrConfig()
		{
			ResourceManager rm = new ResourceManager("FlexCel.Report.TagNames", Assembly.GetExecutingAssembly());
			return rm.GetString("CONFIG");
		}

		private static string GetStrDebug()
		{
			ResourceManager rm = new ResourceManager("FlexCel.Report.TagNames", Assembly.GetExecutingAssembly());
			return rm.GetString("DEBUG");
		}

        private static string GetStrColWithRowCount()
        {
            ResourceManager rm = new ResourceManager("FlexCel.Report.TagNames", Assembly.GetExecutingAssembly());
            return rm.GetString("FLEXCELCOUNT");
        }

		private static string GetStrErrorsInResultFile()
		{
			ResourceManager rm = new ResourceManager("FlexCel.Report.TagNames", Assembly.GetExecutingAssembly());
			return rm.GetString("ERRORSINRESULTFILE");
		}

		private static string GetFormatAdd()
		{
			ResourceManager rm = new ResourceManager("FlexCel.Report.ConfigTagNames", Assembly.GetExecutingAssembly());
			return rm.GetString("Format.ADD");
		}

		private static string GetFormatDelete()
		{
			ResourceManager rm = new ResourceManager("FlexCel.Report.ConfigTagNames", Assembly.GetExecutingAssembly());
			return rm.GetString("Format.DELETE");
		}

        private static Dictionary<string, int> CreateTagParams()
		{
			ResourceManager rm = new ResourceManager("FlexCel.Report.TagNames", Assembly.GetExecutingAssembly());

			Dictionary<string, int> Result = new Dictionary<string,int>();
			Result.Add(rm.GetString("DELETEROW"),0);
			Result.Add(rm.GetString("DELETERANGE"),2);
			Result.Add(rm.GetString("DELETECOLUMN"),0);

			Result.Add(rm.GetString("FORMATCELL"),1);
			Result.Add(rm.GetString("FORMATROW"),1);
			Result.Add(rm.GetString("FORMATCOLUMN"),1);
			Result.Add(rm.GetString("FORMATRANGE"),2);

			Result.Add(rm.GetString("STRIF"),3);
			Result.Add(rm.GetString("STREQUAL"),1);
			Result.Add(rm.GetString("INCLUDE"),3);

			Result.Add(rm.GetString("HPAGEBREAK"),0);
			Result.Add(rm.GetString("VPAGEBREAK"),0);

			Result.Add(rm.GetString("CONFIG"),0);
			Result.Add(rm.GetString("DELETESHEET"),0);

			Result.Add(rm.GetString("COMMENT"),1);

			Result.Add(rm.GetString("EVALUATE"),1);

			Result.Add(rm.GetString("IMGSIZE"),2);
			Result.Add(rm.GetString("IMGPOS"),2);
			Result.Add(rm.GetString("IMGFIT"),2);
			Result.Add(rm.GetString("IMGDELETE"),0);

			Result.Add(rm.GetString("LOOKUP"),4);
			Result.Add(rm.GetString("ARRAY"),1);
			Result.Add(rm.GetString("REGEX"),1);
			Result.Add(rm.GetString("MERGERANGE"),1);

			Result.Add(rm.GetString("FORMULA"),0);
			Result.Add(rm.GetString("COLUMNWIDTH"),1);
			Result.Add(rm.GetString("ROWHEIGHT"),1);
			Result.Add(rm.GetString("AUTOFITSETTINGS"), 3);
			Result.Add(rm.GetString("HTML"),1);
			Result.Add(rm.GetString("REF"),1);

			Result.Add(rm.GetString("DEFINED"),1);
			Result.Add(rm.GetString("DEFINEDFORMAT"), 1);
			Result.Add(rm.GetString("PREPROCESS"), 0);
            
			Result.Add(rm.GetString("AUTOPAGEBREAKS"), 1);
			Result.Add(rm.GetString("AGGREGATE"), 2);
			Result.Add(rm.GetString("LIST"), 1);
			Result.Add(rm.GetString("DBVALUE"), 3);

			return Result;
		}
	

		#endregion

		#region Config Tags
		private static string[] CreateConfigTags()
		{
			ResourceManager rm = new ResourceManager("FlexCel.Report.ConfigTagNames", Assembly.GetExecutingAssembly());
			string[] Result = new string[64]; //bigger just in case
			foreach (ConfigTagEnum tag in TCompactFramework.EnumGetValues(typeof(ConfigTagEnum)))
			{
				Result[(int)tag] = rm.GetString(tag.ToString().ToUpper(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
				Debug.Assert(Result[(int)tag] != null);
			}
			return Result;
		}

		private static string[] CreateConfigTagsParams()
		{
			ResourceManager rm = new ResourceManager("FlexCel.Report.ConfigTagNames", Assembly.GetExecutingAssembly());
			string[] Result = new string[64]; //bigger just in case
			foreach (ConfigTagEnum tag in TCompactFramework.EnumGetValues(typeof(ConfigTagEnum)))
			{
				Result[(int)tag] = rm.GetString("P"+tag.ToString().ToUpper(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
				Debug.Assert(Result[(int)tag] != null);
			}
            return Result;
		}

		private static Dictionary<string, ApplyFormatTagEnum> CreateApplyFormatTags()
		{
			ResourceManager rm = new ResourceManager("FlexCel.Report.ConfigTagNames", Assembly.GetExecutingAssembly());
			Dictionary<string, ApplyFormatTagEnum> Result = new Dictionary<string, ApplyFormatTagEnum>(StringComparer.InvariantCultureIgnoreCase);
			foreach (ApplyFormatTagEnum tag in TCompactFramework.EnumGetValues(typeof(ApplyFormatTagEnum)))
			{
				string s = rm.GetString("Format." + tag.ToString().ToUpper(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
				Debug.Assert(s != null);
				Result[s] = tag;
			}
			return Result;
		}


		#endregion

		#region Tokens
		private static void LoadString(ResourceManager rm, string Id, ref string[] Tmp, TagEnum Tag)
		{
			string s = rm.GetString(Id); 
			if (s == null || s.Length == 0) FlxMessages.ThrowException(FlxErr.ErrUndefinedId, Id, "TagTokens.resx");
			Tmp[(int) Tag] =s; 
		}

		private static string[] CreateTags()
		{
			ResourceManager rm = new ResourceManager("FlexCel.Report.TagTokens", Assembly.GetExecutingAssembly());
			string[] Tmp= new string[128]; //bigger just in case

			LoadString(rm, "StrOpen", ref Tmp, TagEnum.FStrOpen);
			LoadString(rm, "StrOpenHLink", ref Tmp, TagEnum.FStrOpenHLink);
			LoadString(rm, "StrOpenHLink2", ref Tmp, TagEnum.FStrOpenHLink2);
			LoadString(rm, "StrClose", ref Tmp, TagEnum.FStrClose);
			LoadString(rm, "StrCloseHLink", ref Tmp, TagEnum.FStrCloseHLink);
			LoadString(rm, "ParamDelim", ref Tmp, TagEnum.FParamDelim);
			LoadString(rm, "StrOpenParen", ref Tmp, TagEnum.FStrOpenParen);
			LoadString(rm, "StrCloseParen", ref Tmp, TagEnum.FStrCloseParen);
			LoadString(rm, "StrQuote", ref Tmp, TagEnum.FStrQuote);
			LoadString(rm, "DbSeparator", ref Tmp, TagEnum.FDbSeparator);
			LoadString(rm, "RowRange1", ref Tmp, TagEnum.FRowRange1);
			LoadString(rm, "RowFull1", ref Tmp, TagEnum.FRowFull1);
			LoadString(rm, "RowRange2", ref Tmp, TagEnum.FRowRange2);
			LoadString(rm, "RowFull2", ref Tmp, TagEnum.FRowFull2);
			LoadString(rm, "ColRange1", ref Tmp, TagEnum.FColRange1);
			LoadString(rm, "ColFull1", ref Tmp, TagEnum.FColFull1);
			LoadString(rm, "ColRange2", ref Tmp, TagEnum.FColRange2);
			LoadString(rm, "ColFull2", ref Tmp, TagEnum.FColFull2);
			LoadString(rm, "CrossTabRange1", ref Tmp, TagEnum.FCrossTabRange1);
			LoadString(rm, "CrossTabRange2", ref Tmp, TagEnum.FCrossTabRange2);
			LoadString(rm, "CrossTabFull1", ref Tmp, TagEnum.FCrossTabFull1);
			LoadString(rm, "CrossTabFull2", ref Tmp, TagEnum.FCrossTabFull2);
			LoadString(rm, "StrFullDs", ref Tmp, TagEnum.FStrFullDs);
			LoadString(rm, "StrFullDsCaptions", ref Tmp, TagEnum.FStrFullDsCaptions);
			LoadString(rm, "StrDeleteLastRow", ref Tmp, TagEnum.FStrDeleteLastRow);
			LoadString(rm, "StrDontInsertRanges", ref Tmp, TagEnum.FStrDontInsertRanges);
			LoadString(rm, "StrExcludeSheet", ref Tmp, TagEnum.FStrExcludeSheet);
			LoadString(rm, "RelationshipSeparator", ref Tmp, TagEnum.FRelationshipSeparator);
			LoadString(rm, "InternalDB", ref Tmp, TagEnum.FInternalDB);
			LoadString(rm, "StrRowCountColumn", ref Tmp, TagEnum.FStrRowCountColumn);
			LoadString(rm, "StrRowPosColumn", ref Tmp, TagEnum.FStrRowPosColumn);

			LoadString(rm, "StrShow", ref Tmp, TagEnum.FStrShow);
			LoadString(rm, "StrHide", ref Tmp, TagEnum.FStrHide);
			LoadString(rm, "StrAutofit", ref Tmp, TagEnum.FStrAutofit);
             
			LoadString(rm, "StrAutofitOn", ref Tmp, TagEnum.FStrAutofitOn);
			LoadString(rm, "StrAutofitOff", ref Tmp, TagEnum.FStrAutofitOff);
			LoadString(rm, "StrKeepAutofit", ref Tmp, TagEnum.FStrKeepAutofit);
			LoadString(rm, "StrDontKeepAutofit", ref Tmp, TagEnum.FStrDontKeepAutofit);
			LoadString(rm, "StrAutofitModeFirst", ref Tmp, TagEnum.FStrAutofitModeFirst);
			LoadString(rm, "StrAutofitModeLast", ref Tmp, TagEnum.FStrAutofitModeLast);
			LoadString(rm, "StrAutofitModeNone", ref Tmp, TagEnum.FStrAutofitModeNone);
			LoadString(rm, "StrAutofitModeBalanced", ref Tmp, TagEnum.FStrAutofitModeBalanced);
			LoadString(rm, "StrDefinedLocal", ref Tmp, TagEnum.FStrDefinedLocal);
            LoadString(rm, "StrDefinedGlobal", ref Tmp, TagEnum.FStrDefinedGlobal);
            LoadString(rm, "StrRelativeDelete", ref Tmp, TagEnum.FStrRelativeDelete);
			LoadString(rm, "StrFullDelete", ref Tmp, TagEnum.FStrFullDelete);
			LoadString(rm, "StrStaticInclude", ref Tmp, TagEnum.FStrStaticInclude);
			LoadString(rm, "StrDynamicInclude", ref Tmp, TagEnum.FStrDynamicInclude);

			LoadString(rm, "StrCopyRows", ref Tmp, TagEnum.FStrCopyRows);
			LoadString(rm, "StrCopyCols", ref Tmp, TagEnum.FStrCopyCols);
            LoadString(rm, "StrCopyRowsAndCols", ref Tmp, TagEnum.FStrCopyRowsAndCols);

            LoadString(rm, "StrKeepRowsTogether", ref Tmp, TagEnum.FKeepRowsTogether);
			LoadString(rm, "StrKeepColsTogether", ref Tmp, TagEnum.FKeepColsTogether);
			LoadString(rm, "StrEndKeepTogether", ref Tmp, TagEnum.FStrEndKeepTogether);

			LoadString(rm, "StrAlignLeft", ref Tmp, TagEnum.FStrAlignLeft);
			LoadString(rm, "StrAlignCenter", ref Tmp, TagEnum.FStrAlignCenter);
			LoadString(rm, "StrAlignRight", ref Tmp, TagEnum.FStrAlignRight);
			LoadString(rm, "StrAlignTop", ref Tmp, TagEnum.FStrAlignTop);
			LoadString(rm, "StrAlignBottom", ref Tmp, TagEnum.FStrAlignBottom);

			LoadString(rm, "StrInRow", ref Tmp, TagEnum.FStrInRow);
			LoadString(rm, "StrInCol", ref Tmp, TagEnum.FStrInCol);

			LoadString(rm, "StrDontGrow", ref Tmp, TagEnum.FStrDontGrow);
			LoadString(rm, "StrDontShrink", ref Tmp, TagEnum.FStrDontShrink);

			LoadString(rm, "StrAggSum", ref Tmp, TagEnum.FStrAggSum);
			LoadString(rm, "StrAggAvg", ref Tmp, TagEnum.FStrAggAvg);
			LoadString(rm, "StrAggMax", ref Tmp, TagEnum.FStrAggMax);
			LoadString(rm, "StrAggMin", ref Tmp, TagEnum.FStrAggMin);

			return Tmp;
        }
		#endregion
        #endregion

    }
    #region Pseudo columns
    /// <summary>
    /// Defines special meta-columns available for any table.
    /// </summary>
    internal enum TPseudoColumn 
    {
        RowCount = -1,
        RowPos = -2
    }
    #endregion

	#region DisposeMode

	/// <summary>
	/// Indicates if FlexCel must dispose a table after it is done using it.
	/// </summary>
	public enum TDisposeMode
	{
		/// <summary>
		/// FlexCel will not dispose the table. Use this option for example when adding a table that was created on the designer.
		/// </summary>
		DoNotDispose,

		/// <summary>
		/// FlexCel will dispose this table after it has finish using it. Use this option when adding temporary datasets, so you do not have to take care of disposing
		/// the table yourself.
		/// </summary>
		DisposeAfterRun
	}
	#endregion

}
