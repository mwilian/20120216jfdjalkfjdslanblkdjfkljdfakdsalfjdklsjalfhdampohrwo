#region Using directives

using System;
using System.Text;
using System.Globalization;
using System.Resources;
using System.Reflection;
using FlexCel.Core;

#endregion

	/*
     *   Tokens needed to create a PDF file.
	 *   Resources on this unit should NOT be localized. */

namespace FlexCel.Pdf
{
// If this enum is not public, it will be replaced by the obfuscator and fail. It can be added to
// the obfuscator list of exceptions too, but this is another way.

	/// <summary>
	/// Tokens for creating a PDF file. Internal use.
	/// </summary>
    public enum TPdfToken
    {
		/// <summary>
		/// Pdf header for acrobat 5.
		/// </summary>
		Header14,

		/// <summary>
		/// Pdf header for acrobat 7.
		/// </summary>
		Header16,

        /// <summary>
        /// Comment with more than 4 bytes>128.
        /// </summary>
        HeaderComment,

        /// <summary>
        /// (
        /// </summary>
        OpenString,

        /// <summary>
        /// )
        /// </summary>
        CloseString,

        /// <summary>
        /// \
        /// </summary>
        EscapeString,

        /// <summary>
        /// &lt;&lt;
        /// </summary>
        StartDictionary,

        /// <summary>
        /// &gt;&gt;
        /// </summary>
        EndDictionary,

        /// <summary>
        /// null
        /// </summary>
        NullText,

        /// <summary>
        /// XRef section.
        /// </summary>
        XRef,

        /// <summary>
        /// Begin of a trailer section.
        /// </summary>
        Trailer,

        /// <summary>
        /// XRef subsection inside trailer.
        /// </summary>
        StartXRef,

        /// <summary>
        /// Last line of the file.
        /// </summary>
        Eof,

        /// <summary>
        /// Size keyword.
        /// </summary>
        SizeName,

        /// <summary>
        /// Root keyword.
        /// </summary>
        RootName,

        /// <summary>
        /// Info Keyword.
        /// </summary>
        InfoName,

        /// <summary>
        /// /Type
        /// </summary>
        TypeName,

        /// <summary>
        /// /Kids
        /// </summary>
        KidsName,

        /// <summary>
        /// /Count
        /// </summary>
        CountName,

		/// <summary>
		/// /First
		/// </summary>
		FirstName,

		/// <summary>
		/// /Last
		/// </summary>
		LastName,

        /// <summary>
        /// Indirect object
        /// </summary>
        Obj,

        /// <summary>
        /// End of an indirect object.
        /// </summary>
        EndObj,

        /// <summary>
        /// A call to an indirect object.
        /// </summary>
        CallObj,

        /// <summary>
        /// /length
        /// </summary>
        LengthName,

        /// <summary></summary>
        Length1Name,

        /// <summary>
        /// stream
        /// </summary>
        Stream,

        /// <summary>
        /// endstream
        /// </summary>
        EndStream,

        /// <summary>
        /// Pages
        /// </summary>
        PagesName,

        /// <summary>
        /// Page
        /// </summary>
        PageName,

        /// <summary>
        /// /Contents
        /// </summary>
        ContentsName,

        /// <summary></summary>
        AnnotsName,
        /// <summary></summary>
        AnnotName,

        /// <summary>
        /// /Parent
        /// </summary>
        ParentName,

        /// <summary>
        /// Catalog
        /// </summary>
        CatalogName,

        /// <summary>
        /// [
        /// </summary>
        OpenArray,

        /// <summary>
        /// ]
        /// </summary>
        CloseArray,

        /// <summary></summary>
        TitleName,
        /// <summary></summary>
        AuthorName,
        /// <summary></summary>
        SubjectName,
        /// <summary></summary>
        KeywordsName,
        /// <summary></summary>
        CreatorName,
        /// <summary></summary>
        ProducerName,
        /// <summary></summary>
        CreationDateName,

        /// <summary>
        /// Producer of the document.
        /// </summary>
        Producer,

		/// <summary></summary>
		FilterName,
		/// <summary></summary>
		SubFilterName,

        /// <summary></summary>
        FlateDecodeName,

        /// <summary></summary>
        ResourcesName,

        /// <summary></summary>
        FontName,

        /// <summary></summary>
        FontPrefix,

        /// <summary></summary>
        ImgPrefix,

        /// <summary></summary>
        PatternPrefix,

        /// <summary></summary>
        GradientPrefix,

        /// <summary></summary>
        GStatePrefix,

        /// <summary></summary>
        SubtypeName,
        /// <summary></summary>
        TrueTypeName,
        /// <summary></summary>
        BaseFontName,
        /// <summary></summary>
        FontNameName,
        /// <summary></summary>
        FontBBoxName,
        /// <summary></summary>
        FirstCharName,
        /// <summary></summary>
        LastCharName,
        /// <summary></summary>
        EncodingName,
        /// <summary></summary>
        WinAnsiEncodingName,

        /// <summary></summary>
        Bold,
        /// <summary></summary>
        Italic,
        /// <summary></summary>
        BoldItalic,

        /// <summary></summary>
        PsBold,
        /// <summary></summary>
        PsItalic,
        /// <summary></summary>
        PsBoldItalic,
        /// <summary></summary>
        PsOblique,
        /// <summary></summary>
        PsBoldOblique,

        /// <summary></summary>
        PsCourier,
        /// <summary></summary>
        PsHelvetica,
        /// <summary></summary>
        PsTimes,
        /// <summary></summary>
        PsRoman,
        /// <summary></summary>
        PsSymbol,

        /// <summary></summary>
        StArial,
        /// <summary></summary>
        StTimesNewRoman,
        /// <summary></summary>
        StCourier,
        /// <summary></summary>
        StCourierNew,
        /// <summary></summary>
        StMicrosoftSansSerif,
        /// <summary></summary>
        StMicrosoftSerif,

        /// <summary></summary>
        StFixedFonts,
        /// <summary></summary>
        StSerifFonts,
        /// <summary></summary>
        StSymbolFonts,

        /// <summary></summary>
        CIDFontType2Name,
        /// <summary></summary>
        Type0Name,
        /// <summary></summary>
        Type1Name,
        /// <summary></summary>
        DescendantFontsName,
        /// <summary></summary>
        IdentityHName,
        /// <summary></summary>
        CIDSystemInfo,
        /// <summary></summary>
        CIDToGIDMap,

        /// <summary></summary>
        CommandFont,
        /// <summary></summary>
        CommandBeginText,
        /// <summary></summary>
        CommandEndText,
        /// <summary></summary>
        CommandSetBrushColor,
        /// <summary></summary>
        CommandSetPenColor,
        /// <summary></summary>
        CommandSetAlphaBrush,
        /// <summary></summary>
        CommandSetAlphaPen,
        /// <summary></summary>
        CommandTextMove,
        /// <summary></summary>
        CommandTextWrite,
        /// <summary></summary>
        CommandTextRendering,
        /// <summary></summary>
        CommandTextKerningWrite,
        /// <summary></summary>
        CommandMove,
        /// <summary></summary>
        CommandLineToAndStroke,
        /// <summary></summary>
        CommandLineTo,
        /// <summary></summary>
        CommandBezier,
        /// <summary></summary>
        CommandLineWidth,
        /// <summary></summary>
        CommandSetLineStyle,
        /// <summary></summary>
        CommandFillPath,
        /// <summary></summary>
        CommandFillAndStroke,
        /// <summary></summary>
        CommandStroke,
        /// <summary></summary>
        CommandClipPath,
        /// <summary></summary>
        CommandClipPathEvenOddRule,
        /// <summary></summary>
        CommandRectangle,
        /// <summary></summary>
        CommandClosePath,

        /// <summary></summary>
        CommandDo,
        /// <summary></summary>
        Commandscn,
        /// <summary></summary>
        Commandcs,

        /// <summary></summary>
        Commandgs,

        /// <summary></summary>
        XObjectName,
        /// <summary></summary>
        ImageName,
        /// <summary></summary>
        FormName,
        /// <summary></summary>
        PatternName,
        /// <summary></summary>
        ExtGStateName,
        /// <summary></summary>
        WidthName,
        /// <summary></summary>
        WidthsName,
        /// <summary></summary>
        HeightName,
        /// <summary></summary>
        DCTDecodeName,
        /// <summary></summary>
        BitsPerComponentName,

        /// <summary></summary>
        WName,
        /// <summary></summary>
        DWName,

        /// <summary></summary>
        ColorSpaceName,
        /// <summary></summary>
        DeviceRGBName,
        /// <summary></summary>
        DeviceGrayName,

        /// <summary></summary>
        MatteName,

        /// <summary></summary>
        MediaBoxName,
        /// <summary></summary>
        FontDescriptorName,
        /// <summary></summary>
        FlagsName,
        /// <summary></summary>
        FontAscentName,
        /// <summary></summary>
        FontDescentName,
        /// <summary></summary>
        FontFile2Name,
        /// <summary></summary>
        ItalicAngleName,
        /// <summary></summary>
        CapHeightName,
        /// <summary></summary>
        StemVName,

        /// <summary></summary>
        ToUnicodeName,
        /// <summary></summary>
        ToUnicodeData,
        /// <summary></summary>
        ToUnicodeData2,

        /// <summary></summary>
        PredictorName,
        /// <summary></summary>
        IndexedName,
        /// <summary></summary>
        ColorsName,
        /// <summary></summary>
        ColumnsName,
        /// <summary></summary>
        DecodeParmsName,

        /// <summary></summary>
        MaskName,
        /// <summary></summary>
        SMaskName,
		/// <summary></summary>
		ImageMaskName,

        /// <summary>
        /// Windows font folder.
        /// </summary>
        FontFolder,
        /// <summary></summary>
        UpDir,
        /// <summary></summary>
        LinuxFontFolder,

        /// <summary></summary>
        TTFExtension,
        /// <summary></summary>
        TTCExtension,
        /// <summary></summary>
        FamilyItalic,
        /// <summary></summary>
        FamilyOblique,
        /// <summary></summary>
        FamilyBold,

        /// <summary></summary>
        beginbfchar,
        /// <summary></summary>
        endbfchar,
        /// <summary></summary>
        beginbfrange,
        /// <summary></summary>
        endbfrange,

        /// <summary></summary>
        LinkName,
        /// <summary></summary>
        RectName,
        /// <summary></summary>
        AName,
        /// <summary></summary>
        MName,
        /// <summary></summary>
        LinkData,
        /// <summary></summary>
        BorderName,
        /// <summary></summary>
        Border0,

        /// <summary></summary>
        NameName,
        /// <summary></summary>
        TextName,

        /// <summary></summary>
        CAName,
        /// <summary></summary>
        ICName,
        /// <summary></summary>
        BSName,

        /// <summary></summary>
        PatternTypeName,
        /// <summary></summary>
        PaintTypeName,
        /// <summary></summary>
        TilingTypeName,
        /// <summary></summary>
        XStepName,
        /// <summary></summary>
        YStepName,
        /// <summary></summary>
        BBoxName,

        /// <summary></summary>
        PatternColorSpacePrefix,
        /// <summary></summary>
        MatrixName,
        /// <summary></summary>
        ShadingName,
        /// <summary></summary>
        ShadingTypeName,
        /// <summary></summary>
        CoordsName,
        /// <summary></summary>
        FunctionName,
        /// <summary></summary>
        FunctionTypeName,
        /// <summary></summary>
        DomainName, 
        /// <summary></summary>
        RangeName,
        /// <summary></summary>
        C0Name,
        /// <summary></summary>
        C1Name,
        /// <summary></summary>
        NName,
        /// <summary></summary>
        ExtendName,
        /// <summary></summary>
        FunctionsName,
        /// <summary></summary>
        BoundsName,
        /// <summary></summary>
        EncodeName,

        /// <summary></summary>
        AlphaName,
        /// <summary></summary>
        GName,
        /// <summary></summary>
        SName,
        /// <summary></summary>
        GroupName,
        /// <summary></summary>
        TransparencyName,

		/// <summary></summary>
		OutlinesName,
		/// <summary></summary>
		PrevName,
		/// <summary></summary>
		NextName,

		/// <summary></summary>
		CName,
		/// <summary></summary>
		FName,

		/// <summary></summary>
		PageModeName,
		/// <summary></summary>
		UseOutlinesName,
		/// <summary></summary>
		UseThumbsName,
		/// <summary></summary>
		FullScreenName,

		/// <summary></summary>
		DestName,
		/// <summary></summary>
		FitName,
		/// <summary></summary>
		FitHName,
		/// <summary></summary>
		FitVName,
		/// <summary></summary>
		XYZName,

		/// <summary></summary>
		TrueText,
		/// <summary></summary>
		FalseText,

		/// <summary></summary>
		AcroFormName,
		/// <summary></summary>
		FieldsName,
		/// <summary></summary>
		SigFlagsName,

        /// <summary></summary>
        SigName,
        /// <summary></summary>
        VName,

		/// <summary></summary>
		WidgetName,

		/// <summary></summary>
		PName,
		/// <summary></summary>
		TName,
		/// <summary></summary>
		FTName,
		/// <summary></summary>
		FfName,

		/// <summary></summary>
		LocationName,
		/// <summary></summary>
		ReasonName,
		/// <summary></summary>
		ContactInfoName,

		/// <summary></summary>
		APName,

        /// <summary></summary>
        ByteRangeName,

        /// <summary></summary>
        adbe_pkcs7_detachedName,
        /// <summary></summary>
        Adobe_PPKLiteName,

        /// <summary></summary>
        ReferenceName,

        /// <summary></summary>
        TransformMethodName,
        /// <summary></summary>
        TransformParamsName,
        /// <summary></summary>
        DocMDPName,
        /// <summary></summary>
        FieldMDPName,
        /// <summary></summary>
        MD5Name,
        /// <summary></summary>
        SigRefName,

        /// <summary></summary>
        V1_2Name,

        /// <summary></summary>
        DigestMethodName,

        /// <summary></summary>
        PermsName,

        /// <summary></summary>
        DigestValueName,
        /// <summary></summary>
        DigestLocationName,

        /// <summary></summary>
        IDName,

        /// <summary></summary>
        FormTypeName,

        /// <summary></summary>
        ImageCName,
        /// <summary></summary>
        ImageIName,
        /// <summary></summary>
        ImageBName,
        /// <summary></summary>
        ProcSetName,
        /// <summary></summary>
        PDFName,
        /// <summary></summary>
        FRMName,
        /// <summary></summary>
        n0Name,
        /// <summary></summary>
        n2Name,

        /// <summary></summary>
        BitsPerSampleName,

        /// <summary></summary>
        None
    }

    internal sealed class TPdfTokens
    {
        
        private static readonly string[] Tokens = InitTokens(); //STATIC*  
        internal static readonly string NewLine = ((char)10).ToString();

        private TPdfTokens() { }

        /// <summary>
        /// Returns a token from the TPdfToken enumerator.
        /// </summary>
        /// <param name="Code">Code to search for.</param>
        /// <returns>The associated string.</returns>
        public static string GetString(TPdfToken Code)
        {
            return Tokens[(int)Code];
        }

		private static string[] InitTokens()
		{               
			ResourceManager rm = new ResourceManager("FlexCel.Pdf.pdftokens", Assembly.GetExecutingAssembly());
			Array TokenList = TCompactFramework.EnumGetValues(typeof(TPdfToken));
			string[] Result = new string[300];
			foreach (TPdfToken Token in TokenList)
			{
				string s = rm.GetString(Token.ToString());
				Result[(int)Token]=s;
			}
			return Result;
		}
    }
}
