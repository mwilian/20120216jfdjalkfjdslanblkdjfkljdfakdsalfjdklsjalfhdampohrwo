using System;
using System.Text;
using System.Globalization;

using System.Collections.Generic;

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
    #region Supporting classes
#if(FRAMEWORK20)
    internal sealed class TPositionTags: List<THtmlTag>
    {
        internal bool LastIsBlock;
        internal bool LastIsParagraph;

        internal THtmlTag[] ConvertToArray()
        {
            return ToArray();
        }
    }

#else
    
    internal sealed class TPositionTags: ArrayList
    {
		internal bool LastIsBlock;
		internal bool LastIsParagraph;

        internal THtmlTag[] ConvertToArray()
        {
            return (THtmlTag[]) ToArray(typeof(THtmlTag));
        }

    }

#endif
    #endregion

	#region Enumerations
	/// <summary>
	/// Defines the HTML version that will be used when exporting.
	/// </summary>
	public enum THtmlVersion
	{
		/// <summary>
		/// The HTML generated will be 4.01 strict. See http://www.w3.org/TR/html401/
		/// </summary>
		Html_401,

		/// <summary>
		/// HTML generated will be XHTML 1.0. See http://www.w3.org/TR/xhtml1/
		/// </summary>
		XHTML_10,

        /// <summary>
        /// Html generated will be compatible with 3.2, and not have any CSS applied. Useful for when you want simple html markup and don't
        /// worry too much about exact looking. Many properties in FlexCelHtmlExport (like ie6 transparent png support) will be ignored in this mode.
        /// </summary>
        Html_32

	}

	/// <summary>
	/// Defines the way html is generated.
	/// </summary>
	public enum THtmlStyle
	{
		/// <summary>
		/// Only HTML 3.2 tags (like &lt;b&gt;) will be used.
		/// </summary>
		Simple,

		/// <summary>
		/// Cascading style sheets will be used.
		/// </summary>
		Css
	}

	/// <summary>
	/// How the html page will be saved.
	/// </summary>
	public enum THtmlFileFormat
	{
		/// <summary>
		/// File will be saved as an html file with external images.
		/// </summary>
		Html,

		/// <summary>
		/// File will be saved in MHTML (Mime HTML) file format with all the data inside a single file.
		/// This format is the one used to encode emails.
		/// </summary>
		MHtml
	}

	/// <summary>
	/// Defines how characters will be converted when encoding a string as Html.
	/// </summary>
	public enum TEnterStyle
	{
		/// <summary>
		/// Enter characters in the input string will be converted to &lt;br&gt; tags. Multiple spaces will be converted to &amp;nbsp; entities.
		/// </summary>
		Br,

		/// <summary>
		/// Enter characters in the input string will be converted to &amp;#0A; entities. Multiple spaces will be not converted.
		/// </summary>
		Char10,

		/// <summary>
		/// Enter and multiple spaces will be ignored.
		/// </summary>
		Ignore
	}

    /// <summary>
    /// Defines how images will be automatically named by FlexCel, when you do not supply a better name.
    /// </summary>
    public enum TImageNaming
    {
        /// <summary>
        /// The image will be named using standard naming, in a format similar to "filename_image_n.png"
        /// </summary>
        Default,

        /// <summary>
        /// The image will be named using a GUID. This ensures that any image will be unique even if you have many users
        /// requesting the same file at the same time. (Default naming will use the same name for all the users, so images would be overwritten).
        /// As a downside, everytime an image is called a new file will be created, so you can get a lot of images just from an user refreshing a page.
        /// </summary>
        Guid,
    }
	#endregion

	#region HtmlFont Event Handler
	/// <summary>
	/// Arguments passed on <see cref="FlexCel.Render.FlexCelHtmlExport.OnHtmlFont"/>, 
	/// </summary>
	public class HtmlFontEventArgs: EventArgs
	{
		private readonly ExcelFile FWorkbook;

		private readonly TFlxFont FCellFont;
		private string FFontFamily;

		/// <summary>
		/// Creates a new Argument.
		/// </summary>
		/// <param name="aWorkbook">See <see cref="Workbook"/></param>
		/// <param name="aCellFont">See <see cref="CellFont"/></param>
		/// <param name="aFontFamily">See <see cref="FontFamily"/></param>
		public HtmlFontEventArgs(ExcelFile aWorkbook, TFlxFont aCellFont, string aFontFamily)
		{
			FWorkbook = aWorkbook;
			FCellFont = aCellFont;
			FFontFamily = aFontFamily;
		}

		/// <summary>
		/// ExcelFile with the cell we are exporting.
		/// </summary>
		public ExcelFile Workbook {get {return FWorkbook;}}

		/// <summary>
		/// Font we want to process.
		/// </summary>
		public TFlxFont CellFont {get {return FCellFont;}}

		/// <summary>
		/// Use this property to return the new font you want for the cell, if you need to replace it.
		/// Note that you can return more than one font here, and the format for this string is the format on a font selector "font-family" in a CSS stylesheet.
		/// You could for example return the string @"Baskerville, "Heisi Mincho W3", Symbol, serif" here. Look for a complete description of 
		/// the "font-family" descriptor in the CSS reference. (http://www.w3.org/TR/REC-CSS2/)
		/// </summary>
		public string FontFamily {get {return FFontFamily;} set{FFontFamily = value;}}

	}

	/// <summary>
	/// Delegate used to specify which fonts to use on a page.
	/// </summary>
	public delegate void HtmlFontEventHandler(object sender, HtmlFontEventArgs e);


	/// <summary>
	/// Provides a method to customize the fonts used in the HTML methods.
	/// </summary>
	public interface IHtmlFontEvent
	{
		/// <summary>
		/// Method to be called each time a new font is used.
		/// </summary>
		/// <param name="e"></param>
		void DoHtmlFont(HtmlFontEventArgs e);
	}
	#endregion

	/// <summary>
	/// Defines special fixes to the generated files to workaround browser bugs.
	/// </summary>
	public struct THtmlFixes
	{
		#region Privates
		private bool FIe6TransparentPngSupport;
		private bool FOutlook2007CssSupport;
		private bool FWordWrapSupport;
		#endregion

		#region Properties
		/// <summary>
		/// By default, Internet explorer does not support transparent PNGs. Normally this is not an issue, since Excel does not use 
		/// much transparency. But if you rely on transparent images and don't want to use gif images instead of png, you can set this
		/// property to true. It will add special code to the HTML file to support transparent images in IE6.
		/// </summary>
		public bool IE6TransparentPngSupport {get {return FIe6TransparentPngSupport;} set{FIe6TransparentPngSupport = value;}}

		/// <summary>
		/// Outlook 2007 renders HTML worse than previous versions, since it switched to the Word 2007 rendering engine instead of
		/// Internet Explorer to show HTML emails. If you apply this fix, some code will be added to the generated HTML file to improve
		/// the display in Outlook 2007. Other browsers will not be affected and will still render the original file. Turn this option on if
		/// you plan to email the generated file as an HTML email or to edit them in Word 2007. Note that the pages will not validate with the
		/// w3c validator if this option is on.
		/// </summary>
		public bool Outlook2007CssSupport {get {return FOutlook2007CssSupport;} set{FOutlook2007CssSupport = value;}}

		/// <summary>
		/// Some older browsers (and Word 2007) might not support the CSS white-space tag. In this case, if a line longer than a cell cannot be expanded to the right
		/// (because there is data in the next cell) it will wrap down instead of being cropped. This fix will cut the text on this cell to the displayable
		/// characters. If a letter was displayed by the half on the right, after applying this fix it will not display.
		/// This fix is automatically applied when <see cref="Outlook2007CssSupport"/> is selected, so there is normally no reason to apply it. You might get 
		/// a smaller file with this fix (if you have a lot of hidden text), but the display will not be as accurate as when it is off, so it is recomended to keep it off.
		/// </summary>
		public bool WordWrapSupport {get {return FWordWrapSupport;} set{FWordWrapSupport = value;}}
		#endregion

        #region Equality
        /// <summary>
        /// True if both objects are equal.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            if (!(obj is THtmlFixes)) return false;
            THtmlFixes o2 = (THtmlFixes)obj;
            return (o2.FIe6TransparentPngSupport == FIe6TransparentPngSupport 
                && o2.FOutlook2007CssSupport == FOutlook2007CssSupport
                && o2.FWordWrapSupport == FWordWrapSupport);
        }

        /// <summary></summary>
        public static bool operator==(THtmlFixes b1, THtmlFixes b2)
        {
            return b1.Equals(b2);
        }

        /// <summary></summary>
        public static bool operator!=(THtmlFixes b1, THtmlFixes b2)
        {
            return !(b1 == b2);
        }

        /// <summary>
        /// Hash code for the struct.
        /// </summary>
        /// <returns>hashcode.</returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHash(
                FIe6TransparentPngSupport.GetHashCode(),
                FOutlook2007CssSupport.GetHashCode(),
                FWordWrapSupport.GetHashCode());
        }

        #endregion
    }



    /// <summary>
    /// Contains a list of HTML entities and their values
    /// </summary>
    public sealed class THtmlEntities
    {
        #region Privates
		private static readonly StringIntHashtable FNameToCode = CreateNameToCode();//STATIC*
		private static readonly IntStringHashtable FCodeToName = CreateCodeToName();//STATIC*
        #endregion

        #region Constructors
        private THtmlEntities()
        {
        }
        #endregion

        #region Public Methods
        /// <summary>
        /// Converts an Html entity like "amp" into the unicode code for the character. The input string can also 
        /// be a # code, in decimal or hexadecimal. (for example &amp;#64).
        /// </summary>
        /// <param name="EntityName">Name of the entity without the starting "&amp;" and the trailing ";"</param>
        /// <param name="Code">Unicode representation of the entity.</param>
        /// <returns>True if the code exists, false otherwise.</returns>
        public static bool TryNameToCode(string EntityName, out int Code)
        {
            if (FNameToCode.TryGetValue(EntityName, out Code)) return true;      

            if (EntityName == null || EntityName.Length< 2) return false;
            if (EntityName[0] != '#') return false;
            
            if (EntityName[1] == 'x' || EntityName[1] == 'X') //Hexadecimal
                return TryHexToCode(EntityName, ref Code);

            return TryDecToCode(EntityName, ref Code);
        }

        /// <summary>
        /// Returns the maximum length for a name. This includes decimal entities, that can have up to 7 numbers.
        /// </summary>
        public static int MaxNameLength {get {return 10;}}


        /// <summary>
        /// Returns the identifier of a tag.
        /// </summary>
        /// <param name="HtmlString">String containing the tag.</param>
        /// <param name="i">Position of the open "&lt;"</param>
        /// <returns>The tag name.</returns>
        public static string GetTag(string HtmlString, int i)
        {
            int k = i+1;
            while (k < HtmlString.Length && Char.IsLetterOrDigit(HtmlString[k]))
                k++;

            return HtmlString.Substring(i+1, k-i-1);
        }


		/// <summary>
		/// Converts a normal string into a string that can be used inside an HTML file. This includes converting characters
		/// to entities, replacing carriage returns by &lt;br&gt; tags, replacing multiple spaces by nbsp and more. EnterStyle is TEnterStyle.Br if omitted.
		/// </summary>
		/// <param name="originalString">String we want to convert.</param>
		/// <param name="htmlVersion">Version of html we are targeting. In Html 4 &lt;br&gt; is valid and &lt;br/&gt; is not. In XHtml the inverse is true.</param>
		/// <param name="encoding">Code page used to encode the string. Normally this is UTF-8</param>
		/// <returns></returns>
		public static string EncodeAsHtml(string originalString, THtmlVersion htmlVersion, Encoding encoding)
		{
			return EncodeAsHtml(originalString, htmlVersion, encoding, TEnterStyle.Br);
		}


		/// <summary>
		/// Converts a normal string into a string that can be used inside an HTML file. This includes converting characters
		/// to entities, replacing carriage returns by &lt;br&gt; tags, replacing multiple spaces by nbsp and more. EnterStyle is TEnterStyle.Br if omitted.
		/// </summary>
		/// <param name="originalString">String we want to convert.</param>
		/// <param name="htmlVersion">Version of html we are targeting. In Html 4 &lt;br&gt; is valid and &lt;br/&gt; is not. In XHtml the inverse is true.</param>
		/// <param name="EnterStyle">How to convert enters and multiple spaces in the text.</param>
		/// <param name="encoding">Code page used to encode the string. Normally this is UTF-8</param>
		/// <returns></returns>
		public static string EncodeAsHtml(string originalString, THtmlVersion htmlVersion, Encoding encoding, TEnterStyle EnterStyle)
		{
			if (originalString == null) return string.Empty;
			string EndOfTag = THtmlEntities.EndOfTag(htmlVersion);
			StringBuilder Result = new StringBuilder(originalString.Length);
			for (int i = 0; i < originalString.Length; i++)
			{
				if (originalString[i] == '\n') 
				{
					switch (EnterStyle)
					{
						case TEnterStyle.Br:
							Result.Append("<br" + EndOfTag);
							break;
						case TEnterStyle.Char10:
							Result.Append("&#x0A;");
							break;
						case TEnterStyle.Ignore:
							Result.Append(" ");
							break;
					}
					continue;
				}
				if (originalString[i] == '\r') 
				{
					continue;
				}

				if (i>0 && originalString[i] == ' ' && originalString[i-1] == ' ')
				{
					if (EnterStyle == TEnterStyle.Br) Result.Append("&nbsp;"); else Result.Append(" ");
					continue;
				}

				string ent;
				if (FCodeToName.TryGetValue(originalString[i], out ent)) 
				{
					Result.Append("&"+ent+";"); //Needs to go first to verify &amp is converted.
					continue;
				}
                
				Result.Append(originalString[i]); //We are using UTF-8 here, so there should be no issues.

			}

			return Result.ToString();
		}

		#endregion

        #region Implementation
		internal static string EndOfTag(THtmlVersion htmlVersion)
		{
		 return htmlVersion != THtmlVersion.XHTML_10 ? ">" : " />"; //add a space before /> for compatibility with HTML 4. http://www.w3.org/TR/xhtml1/#guidelines
		}

        private static bool TryHexToCode(string EntityName, ref int Code)
        {
            if (EntityName.Length < 3 || EntityName.Length > 2 + 6) return false; //no number or number too big

            Int64 aCode = 0;
            int pw = 0;
            for (int i = EntityName.Length - 1; i >= 2; i--)
            {
                Int64 digit = 0;
                if (EntityName[i] >= '0' && EntityName[i] <= '9')
                {
                    digit = ((int)EntityName[i]) - (int)'0';
                }
                else
                    if (EntityName[i] >= 'A' && EntityName[i] <= 'F')
                {
                    digit = 10 + ((int)EntityName[i]) - (int)'A';
                }
                else
                    if (EntityName[i] >= 'a' && EntityName[i] <= 'f')
                {
                    digit = 10 + ((int)EntityName[i]) - (int)'a';
                }
                else return false;
                                
                aCode += digit << pw;
                pw += 4;
            }
            if (aCode < 0 || aCode > 0x10FFFF) return false; //outside unicode limits.
            Code = (int)aCode;
            return true;
        }

        private static bool TryDecToCode(string EntityName, ref int Code)
        {
            if (EntityName.Length > 1 + 7) return false; //number too big

            Int64 aCode = 0;
            int pw = 1;
            for (int i = EntityName.Length - 1; i >= 1; i--)
            {
                Int64 digit = 0;
                if (EntityName[i] >= '0' && EntityName[i] <= '9')
                {
                    digit = ((int)EntityName[i]) - (int)'0';
                }
                else return false;
                                
                aCode += digit * pw;
                pw *= 10;
            }
            if (aCode < 0 || aCode > 0x10FFFF) return false; //outside unicode limits.
            Code = (int)aCode;
            return true;
        }

        #region Entities Table
        private static StringIntHashtable CreateNameToCode()
        {
            StringIntHashtable Ht = new StringIntHashtable();
            Ht.Add("nbsp", 160);
            Ht.Add("iexcl", 161);
            Ht.Add("cent", 162);
            Ht.Add("pound", 163);
            Ht.Add("curren", 164);
            Ht.Add("yen", 165);
            Ht.Add("brvbar", 166);
            Ht.Add("sect", 167);
            Ht.Add("uml", 168);
            Ht.Add("copy", 169);
            Ht.Add("ordf", 170);
            Ht.Add("laquo", 171);
            Ht.Add("not", 172);
            Ht.Add("shy", 173);
            Ht.Add("reg", 174);
            Ht.Add("macr", 175);
            Ht.Add("deg", 176);
            Ht.Add("plusmn", 177);
            Ht.Add("sup2", 178);
            Ht.Add("sup3", 179);
            Ht.Add("acute", 180);
            Ht.Add("micro", 181);
            Ht.Add("para", 182);
            Ht.Add("middot", 183);
            Ht.Add("cedil", 184);
            Ht.Add("sup1", 185);
            Ht.Add("ordm", 186);
            Ht.Add("raquo", 187);
            Ht.Add("frac14", 188);
            Ht.Add("frac12", 189);
            Ht.Add("frac34", 190);
            Ht.Add("iquest", 191);
            Ht.Add("Agrave", 192);
            Ht.Add("Aacute", 193);
            Ht.Add("Acirc", 194);
            Ht.Add("Atilde", 195);
            Ht.Add("Auml", 196);
            Ht.Add("Aring", 197);
            Ht.Add("AElig", 198);
            Ht.Add("Ccedil", 199);
            Ht.Add("Egrave", 200);
            Ht.Add("Eacute", 201);
            Ht.Add("Ecirc", 202);
            Ht.Add("Euml", 203);
            Ht.Add("Igrave", 204);
            Ht.Add("Iacute", 205);
            Ht.Add("Icirc", 206);
            Ht.Add("Iuml", 207);
            Ht.Add("ETH", 208);
            Ht.Add("Ntilde", 209);
            Ht.Add("Ograve", 210);
            Ht.Add("Oacute", 211);
            Ht.Add("Ocirc", 212);
            Ht.Add("Otilde", 213);
            Ht.Add("Ouml", 214);
            Ht.Add("times", 215);
            Ht.Add("Oslash", 216);
            Ht.Add("Ugrave", 217);
            Ht.Add("Uacute", 218);
            Ht.Add("Ucirc", 219);
            Ht.Add("Uuml", 220);
            Ht.Add("Yacute", 221);
            Ht.Add("THORN", 222);
            Ht.Add("szlig", 223);
            Ht.Add("agrave", 224);
            Ht.Add("aacute", 225);
            Ht.Add("acirc", 226);
            Ht.Add("atilde", 227);
            Ht.Add("auml", 228);
            Ht.Add("aring", 229);
            Ht.Add("aelig", 230);
            Ht.Add("ccedil", 231);
            Ht.Add("egrave", 232);
            Ht.Add("eacute", 233);
            Ht.Add("ecirc", 234);
            Ht.Add("euml", 235);
            Ht.Add("igrave", 236);
            Ht.Add("iacute", 237);
            Ht.Add("icirc", 238);
            Ht.Add("iuml", 239);
            Ht.Add("eth", 240);
            Ht.Add("ntilde", 241);
            Ht.Add("ograve", 242);
            Ht.Add("oacute", 243);
            Ht.Add("ocirc", 244);
            Ht.Add("otilde", 245);
            Ht.Add("ouml", 246);
            Ht.Add("divide", 247);
            Ht.Add("oslash", 248);
            Ht.Add("ugrave", 249);
            Ht.Add("uacute", 250);
            Ht.Add("ucirc", 251);
            Ht.Add("uuml", 252);
            Ht.Add("yacute", 253);
            Ht.Add("thorn", 254);
            Ht.Add("yuml", 255);
            Ht.Add("fnof", 402);
            Ht.Add("Alpha", 913);
            Ht.Add("Beta", 914);
            Ht.Add("Gamma", 915);
            Ht.Add("Delta", 916);
            Ht.Add("Epsilon", 917);
            Ht.Add("Zeta", 918);
            Ht.Add("Eta", 919);
            Ht.Add("Theta", 920);
            Ht.Add("Iota", 921);
            Ht.Add("Kappa", 922);
            Ht.Add("Lambda", 923);
            Ht.Add("Mu", 924);
            Ht.Add("Nu", 925);
            Ht.Add("Xi", 926);
            Ht.Add("Omicron", 927);
            Ht.Add("Pi", 928);
            Ht.Add("Rho", 929);
            Ht.Add("Sigma", 931);
            Ht.Add("Tau", 932);
            Ht.Add("Upsilon", 933);
            Ht.Add("Phi", 934);
            Ht.Add("Chi", 935);
            Ht.Add("Psi", 936);
            Ht.Add("Omega", 937);
            Ht.Add("alpha", 945);
            Ht.Add("beta", 946);
            Ht.Add("gamma", 947);
            Ht.Add("delta", 948);
            Ht.Add("epsilon", 949);
            Ht.Add("zeta", 950);
            Ht.Add("eta", 951);
            Ht.Add("theta", 952);
            Ht.Add("iota", 953);
            Ht.Add("kappa", 954);
            Ht.Add("lambda", 955);
            Ht.Add("mu", 956);
            Ht.Add("nu", 957);
            Ht.Add("xi", 958);
            Ht.Add("omicron", 959);
            Ht.Add("pi", 960);
            Ht.Add("rho", 961);
            Ht.Add("sigmaf", 962);
            Ht.Add("sigma", 963);
            Ht.Add("tau", 964);
            Ht.Add("upsilon", 965);
            Ht.Add("phi", 966);
            Ht.Add("chi", 967);
            Ht.Add("psi", 968);
            Ht.Add("omega", 969);
            Ht.Add("thetasym", 977);
            Ht.Add("upsih", 978);
            Ht.Add("piv", 982);
            Ht.Add("bull", 8226);
            Ht.Add("hellip", 8230);
            Ht.Add("prime", 8242);
            Ht.Add("Prime", 8243);
            Ht.Add("oline", 8254);
            Ht.Add("frasl", 8260);
            Ht.Add("weierp", 8472);
            Ht.Add("image", 8465);
            Ht.Add("real", 8476);
            Ht.Add("trade", 8482);
            Ht.Add("alefsym", 8501);
            Ht.Add("larr", 8592);
            Ht.Add("uarr", 8593);
            Ht.Add("rarr", 8594);
            Ht.Add("darr", 8595);
            Ht.Add("harr", 8596);
            Ht.Add("crarr", 8629);
            Ht.Add("lArr", 8656);
            Ht.Add("uArr", 8657);
            Ht.Add("rArr", 8658);
            Ht.Add("dArr", 8659);
            Ht.Add("hArr", 8660);
            Ht.Add("forall", 8704);
            Ht.Add("part", 8706);
            Ht.Add("exist", 8707);
            Ht.Add("empty", 8709);
            Ht.Add("nabla", 8711);
            Ht.Add("isin", 8712);
            Ht.Add("notin", 8713);
            Ht.Add("ni", 8715);
            Ht.Add("prod", 8719);
            Ht.Add("sum", 8721);
            Ht.Add("minus", 8722);
            Ht.Add("lowast", 8727);
            Ht.Add("radic", 8730);
            Ht.Add("prop", 8733);
            Ht.Add("infin", 8734);
            Ht.Add("ang", 8736);
            Ht.Add("and", 8743);
            Ht.Add("or", 8744);
            Ht.Add("cap", 8745);
            Ht.Add("cup", 8746);
            Ht.Add("int", 8747);
            Ht.Add("there4", 8756);
            Ht.Add("sim", 8764);
            Ht.Add("cong", 8773);
            Ht.Add("asymp", 8776);
            Ht.Add("ne", 8800);
            Ht.Add("equiv", 8801);
            Ht.Add("le", 8804);
            Ht.Add("ge", 8805);
            Ht.Add("sub", 8834);
            Ht.Add("sup", 8835);
            Ht.Add("nsub", 8836);
            Ht.Add("sube", 8838);
            Ht.Add("supe", 8839);
            Ht.Add("oplus", 8853);
            Ht.Add("otimes", 8855);
            Ht.Add("perp", 8869);
            Ht.Add("sdot", 8901);
            Ht.Add("lceil", 8968);
            Ht.Add("rceil", 8969);
            Ht.Add("lfloor", 8970);
            Ht.Add("rfloor", 8971);
            Ht.Add("lang", 9001);
            Ht.Add("rang", 9002);
            Ht.Add("loz", 9674);
            Ht.Add("spades", 9824);
            Ht.Add("clubs", 9827);
            Ht.Add("hearts", 9829);
            Ht.Add("diams", 9830);
            Ht.Add("quot", 34);
            Ht.Add("amp", 38);
            Ht.Add("lt", 60);
            Ht.Add("gt", 62);
            Ht.Add("OElig", 338);
            Ht.Add("oelig", 339);
            Ht.Add("Scaron", 352);
            Ht.Add("scaron", 353);
            Ht.Add("Yuml", 376);
            Ht.Add("circ", 710);
            Ht.Add("tilde", 732);
            Ht.Add("ensp", 8194);
            Ht.Add("emsp", 8195);
            Ht.Add("thinsp", 8201);
            Ht.Add("zwnj", 8204);
            Ht.Add("zwj", 8205);
            Ht.Add("lrm", 8206);
            Ht.Add("rlm", 8207);
            Ht.Add("ndash", 8211);
            Ht.Add("mdash", 8212);
            Ht.Add("lsquo", 8216);
            Ht.Add("rsquo", 8217);
            Ht.Add("sbquo", 8218);
            Ht.Add("ldquo", 8220);
            Ht.Add("rdquo", 8221);
            Ht.Add("bdquo", 8222);
            Ht.Add("dagger", 8224);
            Ht.Add("Dagger", 8225);
            Ht.Add("permil", 8240);
            Ht.Add("lsaquo", 8249);
            Ht.Add("rsaquo", 8250);
            Ht.Add("euro", 8364);

			return Ht;
        }
		private static IntStringHashtable CreateCodeToName()
		{
			IntStringHashtable Ht = new IntStringHashtable();
			foreach (string EntityName in FNameToCode.Keys)
			{
				Ht.Add(FNameToCode[EntityName], EntityName);
			}

			return Ht;
		}

		#endregion

        #endregion
    }

    
    /// <summary>
    /// Contains an HTML tag and its position on the string.
    /// </summary>
    public class THtmlTag
    {
        #region Privates
        private string FText;
        private int FPosition;
        #endregion

        #region Constructors
        /// <summary>
        /// Creates a new instance of a Tag.
        /// </summary>
        /// <param name="aPosition">Position of the final string where the tag is.</param>
        /// <param name="aText">Text of the tag. (not including brackets)</param>
        public THtmlTag(int aPosition, string aText)
        {
            FPosition = aPosition;
            FText = aText;
        }

        #endregion
        
        #region Properties
        /// <summary>
        /// The tag text.
        /// </summary>
        public string Text {get {return FText;} set {FText=value;}}

        /// <summary>
        /// Position of this tag inside the converted string. (the string without tags)
        /// </summary>
        public int Position {get {return FPosition;} set {FPosition=value;}}
        #endregion

    }

    /// <summary>
    /// An Html string parsed into a C# string and tags.
    /// </summary>
    public class THtmlParsedString
    {
        private string FText;
        private THtmlTag[] FTags;

        #region Constructors
        /// <summary>
        /// Creates a new empty THtmlParsedString.
        /// </summary>
        public THtmlParsedString()
        {
            FText = String.Empty;
            FTags = new THtmlTag[0];
        }

        /// <summary>
        /// Creates a new THtmlParsedString containing an existing Html string.
        /// </summary>
        /// <param name="HtmlString">Html string we want to parse.</param>
        public THtmlParsedString(string HtmlString): this()
        {
            Parse(HtmlString);
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// The parsed html text.
        /// </summary>
        public string Text {get{return FText;}}

        /// <summary>
        /// A list of the tags on the parsed string.
        /// </summary>
        public THtmlTag[] Tags {get{return FTags;}}
        #endregion

        #region Public Methods
        /// <summary>
        /// Parses an Html string and fills this instance with the parsed values.
        /// </summary>
        /// <param name="HtmlString">Html string we want to parse.</param>
        public void Parse(string HtmlString)
        {
            int InPre = 0;
            TPositionTags PositionTags = new TPositionTags();
            try
            {
                FText = String.Empty;
                bool PendingWhitespace = false;
                if (HtmlString==null) return;

                StringBuilder sb = new StringBuilder(HtmlString.Length);
                int i = 0;
                while (i < HtmlString.Length)
                {
                    switch (HtmlString[i])
                    {
                        case '&': 
                            AddEntity(PositionTags, HtmlString, sb, ref i, ref PendingWhitespace);
                            break;

                        case '<':
                            AddTag(PositionTags, HtmlString, sb, ref i, ref InPre, ref PendingWhitespace);
                            break;

                        case (char)10:
                        case (char)13:
                        case '\t':
                        case ' ':
                        case '\x200B':  //those are the whitespaces defined on html standard.
                            if (InPre > 0)
                            {
                                if (HtmlString[i] != '\r')
                                {
                                    Append(sb, HtmlString[i], PositionTags, ref PendingWhitespace);
                                }
                            }
                            else
                            {
                                if (sb.Length > 0 && sb[sb.Length - 1] != ' ' && sb[sb.Length - 1] != (char)10)
                                {
                                    PendingWhitespace = true;
                                }
                            }
                            break;

                        default:
                            {
                                Append(sb, HtmlString[i], PositionTags, ref PendingWhitespace);
                            }
                            break;
                    }

                    i++;
                }

                ReplaceNbsp(sb);  //We do it at the end so we can keep more than 2 spaces.
                FText = sb.ToString();
            }
            finally
            {
                FTags = PositionTags.ConvertToArray();
            }
        }
        #endregion

        #region Utilities
        private static void ReplaceNbsp(StringBuilder sb)
        {
            for (int i = 0; i < sb.Length; i++)
                if (sb[i] == (char)160) sb[i] = ' ';
        }

		private static void Append(StringBuilder sb, char c, TPositionTags PositionTags, ref bool PendingWhitespace)
		{
            if (PendingWhitespace) sb.Append(' ');
            PendingWhitespace = false;
			if (PositionTags.LastIsBlock)
			{
				if (sb.Length > 0 && sb[sb.Length - 1] != '\n') sb.Append('\n');
				PositionTags.LastIsBlock = false;
			}

			if (PositionTags.LastIsParagraph)
			{
				if (sb.Length > 1 && sb[sb.Length - 2] != '\n') sb.Append('\n');
				PositionTags.LastIsParagraph = false;
			}

			sb.Append(c);
		}

        private static void AddEntity(TPositionTags PositionTags, String HtmlString, StringBuilder sb, ref int i, ref bool PendingWhitespace)
        {
            int k = i + 1;
            while (k < HtmlString.Length && k - (i+1) < THtmlEntities.MaxNameLength && HtmlString[k] != ';')
            {
                k++;
            }
            int Code;
            if (k <= i+1 || k >= HtmlString.Length || HtmlString[k] != ';' || 
                !THtmlEntities.TryNameToCode(HtmlString.Substring(i+1, k-i-1), out Code))
            {
                Append(sb, HtmlString[i], PositionTags, ref PendingWhitespace);
                return;
            }

            Append(sb, (char)Code, PositionTags, ref PendingWhitespace);
            i=k;
        }

        private static void AddTag(TPositionTags PositionTags, String HtmlString, StringBuilder sb, ref int i, ref int InPre, ref bool PendingWhitespace)
        {
            int k = i + 1;
            while (k < HtmlString.Length && HtmlString[k] != '>')
            {
                k++;
            }

            if (k <= i + 1 || k >= HtmlString.Length || HtmlString[k] != '>')
            {
                Append(sb, HtmlString[i], PositionTags, ref PendingWhitespace);
                return;
            }

            if (IsTag(HtmlString, i, "BR"))
            {
                Append(sb, '\n', PositionTags, ref PendingWhitespace);
                i = k;
                return;
            }

            if (IsTag(HtmlString, i, "P") || IsTag(HtmlString, i, "/P"))
            {
                PositionTags.LastIsBlock = true;
                PositionTags.LastIsParagraph = true;
                if (sb.Length <= 0 || sb[sb.Length - 1] != '\n')
                {
                    PendingWhitespace = false; //no need to add the whitespace, as we are adding an enter anyway
                    sb.Append('\n'); //Do NOT use Append(sb) here, it would ruin the fun (and reset LastIsBlock).
                }
                i = k;
                return;
            }

            if (IsTag(HtmlString, i, "PRE"))
            {
                InPre++;
            }

            if (IsTag(HtmlString, i, "/PRE"))
            {
                InPre--;
            }

            if (IsTag(HtmlString, i, "DIV") ||
                IsTag(HtmlString, i, "H1") || IsTag(HtmlString, i, "H2") || IsTag(HtmlString, i, "H3") || IsTag(HtmlString, i, "H4") ||
                IsTag(HtmlString, i, "H5") || IsTag(HtmlString, i, "H6"))
            {
                PositionTags.LastIsBlock = true;
            }

            if (IsTag(HtmlString, i, "/DIV") ||
                IsTag(HtmlString, i, "/H1") || IsTag(HtmlString, i, "/H2") || IsTag(HtmlString, i, "/H3") || IsTag(HtmlString, i, "/H4") ||
                IsTag(HtmlString, i, "/H5") || IsTag(HtmlString, i, "/H6"))
            {
                PositionTags.LastIsBlock = true;
            }

            if (PendingWhitespace)
            {
                sb.Append(' ');
                PendingWhitespace = false;
            }
            PositionTags.Add(new THtmlTag(sb.Length, HtmlString.Substring(i + 1, k - i - 1)));
            i = k;
        }

        private static bool IsTag(string HtmlString, int i, string TagId)
        {
            if (i+1+ TagId.Length +1> HtmlString.Length) return false;
            for (int k = 0; k < TagId.Length; k++)
            {
                if (Char.ToLower(HtmlString[i+1+k], CultureInfo.InvariantCulture) != Char.ToLower(TagId[k], CultureInfo.InvariantCulture)) return false;
            }
            if (Char.IsLetterOrDigit(HtmlString[i+1+TagId.Length])) return false; //tag continues.
            return true;
        }

        #endregion
    }

    /// <summary>
    /// Converts between Color structs and HTML values.
    /// </summary>
    public sealed class THtmlColors
    {
        private THtmlColors(){}

        /// <summary>
        /// Returns a Color struct from an HTML string
        /// </summary>
        /// <param name="Value">String with color on HTML format. (one of the 16 named colors or #notation)</param>
        /// <returns>The corresponding Color.</returns>
        public static Color GetColor(string Value)
        {
            if (Value == null || Value.Length <= 0) return ColorUtil.Empty;

            if (Value[0] == '#')
            {
                int[] RGB = new int[3];
                int Pos = 1;
                for (int i = 0; i < 3; i++)
                {
                    for (int k = 4; k >= 0; k -= 4)
                    {
                        if (Pos >= Value.Length) break;
                        if (Char.IsDigit(Value[Pos])) 
                        {
                            RGB[i] += ((int)Value[Pos] - (int)'0') <<  k;
                        }
                        else
                            if (Value[Pos]>='A' && Value[Pos]<='F') 
                        {
                            RGB[i] += ((int)Value[Pos] - (int)'A' + 10) << k;
                        }
                        else
                            if (Value[Pos]>='a' && Value[Pos]<='f') 
                        {
                            RGB[i] += ((int)Value[Pos] - (int)'a' + 10) << k;
                        }
                        else return ColorUtil.Empty;

                        //Sometimes, colors are stored as #rgb, instead of #rrggbb.  In the #rgb, notation, both r are the same, both g are tthe same and both b are the same.
                        if (Value.Length != 4 || k != 4) Pos++;
                    }
                }
                    
                return ColorUtil.FromArgb(RGB[0], RGB[1], RGB[2]);
            }
            else
                switch (Value.ToLower(CultureInfo.InvariantCulture))
                {
                    case "black": return Colors.Black;
                    case "green": return Colors.Green;
                    case "silver": return Colors.Silver;
                    case "lime": return Colors.Lime;
                    case "gray": return Colors.Gray;
                    case "olive": return Colors.Olive;
                    case "white": return Colors.White;
                    case "yellow": return Colors.Yellow;
                    case "maroon": return Colors.Maroon;
                    case "navy": return Colors.Navy;
                    case "red": return Colors.Red;
                    case "blue": return Colors.Blue;
                    case "purple": return Colors.Purple;
                    case "teal": return Colors.Teal;
                    case "fuchsia": return Colors.Fuchsia;
                    case "aqua": return Colors.Aqua;
                }

            return ColorUtil.Empty;
        }

        /// <summary>
        /// Returns an HTML color string from a Color struct.
        /// </summary>
        /// <param name="Value">The color we want to convert.</param>
        /// <returns>String with color on HTML format. (one of the 16 named colors or #notation).</returns>
        public static string GetColor(Color Value)
        {
#if(WPF)
            if (Value == Colors.Black) return "black";
            if (Value == Colors.Green) return "green";
            if (Value == Colors.Silver) return "silver";
            if (Value == Colors.Lime) return "lime";
            if (Value == Colors.Gray) return "gray";
            if (Value == Colors.Olive) return "olive";
            if (Value == Colors.White) return "white";
            if (Value == Colors.Yellow) return "yellow";
            if (Value == Colors.Maroon) return "maroon";
            if (Value == Colors.Navy) return "navy";
            if (Value == Colors.Red) return "red";
            if (Value == Colors.Blue) return "blue";
            if (Value == Colors.Purple) return "purple";
            if (Value == Colors.Teal) return "teal";
            if (Value == Colors.Fuchsia) return "fuchsia";
            if (Value == Colors.Aqua) return "aqua";
#else
			int v = Value.ToArgb();
                    if (v == Colors.Black.ToArgb()) return "black";
                    if (v == Colors.Green.ToArgb()) return "green";
                    if (v == Colors.Silver.ToArgb()) return "silver";
                    if (v == Colors.Lime.ToArgb()) return "lime";
                    if (v == Colors.Gray.ToArgb()) return "gray";
                    if (v == Colors.Olive.ToArgb()) return "olive";
                    if (v == Colors.White.ToArgb()) return "white";
                    if (v == Colors.Yellow.ToArgb()) return "yellow";
                    if (v == Colors.Maroon.ToArgb()) return "maroon";
                    if (v == Colors.Navy.ToArgb()) return "navy";
                    if (v == Colors.Red.ToArgb()) return "red";
                    if (v == Colors.Blue.ToArgb()) return "blue";
                    if (v == Colors.Purple.ToArgb()) return "purple";
                    if (v == Colors.Teal.ToArgb()) return "teal";
                    if (v == Colors.Fuchsia.ToArgb()) return "fuchsia";
                    if (v == Colors.Aqua.ToArgb()) return "aqua";
#endif
#if (COMPACTFRAMEWORK || (!FRAMEWORK30 && !MONOTOUCH))
                    return "#" +
                        Value.R.ToString("x2", CultureInfo.InvariantCulture) +
                        Value.G.ToString("x2", CultureInfo.InvariantCulture) +
                        Value.B.ToString("x2", CultureInfo.InvariantCulture);
#else
                    return "#" +
                        Value.R().ToString("x2", CultureInfo.InvariantCulture) +
                        Value.G().ToString("x2", CultureInfo.InvariantCulture) +
                        Value.B().ToString("x2", CultureInfo.InvariantCulture);
#endif
        }
    }

	/// <summary>
	/// Creates html tags for different actions, and depending on the HTML style.
	/// </summary>
	public sealed class THtmlTagCreator
	{
		private THtmlTagCreator()
		{
		}

		private static string FormatStr(string s, params object[] p)
		{
			return String.Format(CultureInfo.InvariantCulture, s, p);
		}


		/// <summary>
		/// Returns a tag to change the font color. Remember to close it with <see cref="EndFontColor"/>
		/// </summary>
		/// <param name="aColor">Color to use.</param>
        /// <param name="htmlStyle">Style of the generated HTML.</param>
		/// <returns></returns>
		public static string StartFontColor(Color aColor, THtmlStyle htmlStyle)
		{
			if (htmlStyle == THtmlStyle.Simple)
			{
				return FormatStr("<font color = '{0}'>", THtmlColors.GetColor(aColor));
			}
			else
			{
				return FormatStr("<span style = 'color:{0};'>", THtmlColors.GetColor(aColor));
			}
		}

		/// <summary>
		/// Returns a tag to end changing a font color that was started with <see cref="StartFontColor"/>
		/// </summary>
		/// <param name="htmlStyle">Specifies if to use css or not.</param>
		/// <returns></returns>
		public static string EndFontColor(THtmlStyle htmlStyle)
		{
			if (htmlStyle == THtmlStyle.Simple)
			{
				return "</font>";
			}
			else
			{
				return "</span>";
			}
		}

		/// <summary>
		/// Returns the tags for a difference between one font and the next.
		/// </summary>
		/// <param name="lastFont"></param>
		/// <param name="nextFont"></param>
		/// <param name="htmlStyle"></param>
		/// <param name="htmlVersion"></param>
        /// <param name="originalFont"></param>
        /// <param name="xls"></param>
		/// <param name="tagsToClose">Tags that need to be closed. This method might decide to acumulate font tags or not, depending on the case.</param>
        /// <param name="OnHtmlFont"></param>
        /// <param name="MsFormat"></param>
		/// <returns></returns>
		internal static string DiffFont(ExcelFile xls, TFlxFont originalFont, TFlxFont lastFont, TFlxFont nextFont, THtmlVersion htmlVersion, THtmlStyle htmlStyle, ref StringBuilder tagsToClose, IHtmlFontEvent OnHtmlFont, bool MsFormat)
		{
			if (htmlStyle == THtmlStyle.Simple)
			{
				StringBuilder Result = new StringBuilder(tagsToClose.ToString());
				tagsToClose.Length = 0;
				if ((originalFont.Style & TFlxFontStyles.Bold)==0  && (nextFont.Style & TFlxFontStyles.Bold)!=0)
				{
					Result.Append("<b>");
					tagsToClose.Insert(0, "</b>");
				}
				if ((originalFont.Style & TFlxFontStyles.Italic)==0  && (nextFont.Style & TFlxFontStyles.Italic)!=0)
				{
					Result.Append("<i>");
					tagsToClose.Insert(0, "</i>");
				}
				if ((originalFont.Style & TFlxFontStyles.StrikeOut)==0  && (nextFont.Style & TFlxFontStyles.StrikeOut)!=0)
				{
					Result.Append("<strike>");
					tagsToClose.Insert(0, "</strike>");
				}
				if ((originalFont.Style & TFlxFontStyles.Subscript)==0  && (nextFont.Style & TFlxFontStyles.Subscript)!=0)
				{
					Result.Append("<sub>");
					tagsToClose.Insert(0, "</sub>");
				}
				if ((originalFont.Style & TFlxFontStyles.Superscript)==0  && (nextFont.Style & TFlxFontStyles.Superscript)!=0)
				{
					Result.Append("<sup>");
					tagsToClose.Insert(0, "</sup>");
				}
				if (originalFont.Underline == TFlxUnderline.None  && nextFont.Underline != TFlxUnderline.None)
				{
					Result.Append("<u>");
					tagsToClose.Insert(0, "</u>");
				}

				if (originalFont.Size20 != nextFont.Size20 || originalFont.Name != nextFont.Name || originalFont.Color != nextFont.Color)
				{
					Result.Append("<font");
					if (originalFont.Size20 != nextFont.Size20) 
					{
                        int size = 1;
                        if (MsFormat)
                        {
                            size = nextFont.Size20;
                        }
                        else
                        {
                            int[] SizesInPoints = { 8, 9, 12, 14, 18, 24, 34 };
                            float SizePoints = nextFont.Size20 / 20f;
                            for (int i = 1; i < SizesInPoints.Length; i++)
                            {
                                if (SizePoints < (SizesInPoints[i] + SizesInPoints[i - 1]) / 2f) break;
                                size = i + 1;
                            }
                        }
						Result.Append(FormatStr(" size = '{0}'", size));
					}
					if (originalFont.Name != nextFont.Name) Result.Append(FormatStr(" face = '{0}'", nextFont.Name));
					if (originalFont.Color != nextFont.Color) Result.Append(FormatStr(" color = '{0}'", THtmlColors.GetColor(nextFont.Color.ToColor(xls, Colors.Black)) ));

					Result.Append(">");
					tagsToClose.Insert(0, "</font>");
				}
				return Result.ToString();
			}
			else //*************************CSS
			{
				StringBuilder Result = new StringBuilder();
                DiffFontCss(xls, originalFont, nextFont, OnHtmlFont, Result);

				if (Result.Length > 0) 
				{
					String Result2 = tagsToClose.ToString() + FormatStr("<span style ='{0}'>", Result.ToString());
					tagsToClose.Length = 0;
					tagsToClose.Append("</span>");
					return Result2;
				}

				Result.Append(tagsToClose.ToString());
				tagsToClose.Length = 0;
				return Result.ToString();
			}
		}

        internal static void DiffFontCss(ExcelFile xls, TFlxFont originalFont, TFlxFont nextFont, IHtmlFontEvent OnHtmlFont, StringBuilder Result)
        {
            if ((originalFont.Style & TFlxFontStyles.Bold) != (nextFont.Style & TFlxFontStyles.Bold))
            {
                string FontBold = (nextFont.Style & TFlxFontStyles.Bold) == 0 ? "normal" : "bold";
                Result.Append(FormatStr("font-weight:{0};", FontBold));
            }
            if ((originalFont.Style & TFlxFontStyles.Italic) != (nextFont.Style & TFlxFontStyles.Italic))
            {
                string FontItalic = (nextFont.Style & TFlxFontStyles.Italic) == 0 ? "normal" : "italic";
                Result.Append(FormatStr("font-style:{0};", FontItalic));
            }

            //we cannot reset text-decoration. so it should always be off at the start of the cell.
            bool NeedsUl = nextFont.Underline != TFlxUnderline.None;
            bool NeedsSt = (nextFont.Style & TFlxFontStyles.StrikeOut) != 0;

            if (NeedsSt || NeedsUl)
            {
                string Decoration = string.Empty;
                if (NeedsSt) Decoration += "line-through";
                if (NeedsUl) Decoration += " underline";

                if (Decoration.Length > 0)
                    Result.Append(FormatStr("text-decoration:{0};", Decoration));
            }

            //Super and subscripts can't be both at the same time.
            if ((originalFont.Style & TFlxFontStyles.Subscript) != (nextFont.Style & TFlxFontStyles.Subscript)
                || (originalFont.Style & TFlxFontStyles.Superscript) != (nextFont.Style & TFlxFontStyles.Superscript))
            {
                string FontSub = "baseline";
                if ((nextFont.Style & TFlxFontStyles.Subscript) != 0) FontSub = "sub";
                else if ((nextFont.Style & TFlxFontStyles.Superscript) != 0) FontSub = "super";
                Result.Append(FormatStr("vertical-align:{0};", FontSub));
            }

            if (originalFont.Size20 != nextFont.Size20)
            {
                Result.Append(FormatStr("font-size:{0:##}pt;", nextFont.Size20 / 20f));
            }

            if (originalFont.Name != nextFont.Name)
            {
                HtmlFontEventArgs e = new HtmlFontEventArgs(xls, nextFont, nextFont.Name);
                if (e.FontFamily != null && e.FontFamily.IndexOf(" ") >= 0) e.FontFamily = "\"" + e.FontFamily + "\"";  //font names with spaces must be quoted, and they must be double quotes so they do not clash with the style tag.
                if (OnHtmlFont != null) OnHtmlFont.DoHtmlFont(e);
                Result.Append(FormatStr("font-family:{0};", e.FontFamily));
            }

            if (originalFont.Color != nextFont.Color)
            {
                Result.Append(FormatStr("color:{0};", THtmlColors.GetColor(nextFont.Color.ToColor(xls, Colors.Black))));
            }
        }

		internal static string CloseDiffFont(StringBuilder tagsToClose)
		{
			return tagsToClose.ToString();
		}
	}
}
