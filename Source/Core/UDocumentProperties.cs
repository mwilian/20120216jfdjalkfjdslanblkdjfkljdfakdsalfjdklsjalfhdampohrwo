using System;

namespace FlexCel.Core
{
    /// <summary>
    /// One custom property name and its id.
    /// </summary>
    internal class TOlePropertyName
    {
        private long FId;
        private string FName;

        public TOlePropertyName(long aId, string aName)
        {
            FId = aId;
            FName = aName;
        }

        /// <summary>
        /// Id of the property.
        /// </summary>
        public long Id {get {return FId;} set {FId=value;}}

        /// <summary>
        /// Name of the property.
        /// </summary>
        public string Name {get {return FName;} set {FName=value;}}
    }

    /// <summary>
    /// Standard properties of an ole file. There are two different sets of properties,
    /// the standard ones (properties that exist for any file) and the extended ones (properties that exist
    /// for Ms Office documents)
    /// </summary>
    public enum TPropertyId
    {
        /// <summary>This is not a valid property. It will do nothing.</summary>
        None = 0,

        /// <summary>Document Title. (Standard property)</summary>
        Title = 2,

        /// <summary>Document Subject. (Standard property)</summary>
        Subject = 3,

        /// <summary>Author. (Standard property)</summary>
        Author = 4,

        /// <summary>Keywords. (Standard property)</summary>
        Keywords = 5,

        /// <summary>Comments. (Standard property)</summary>
        Comments = 6,

        /// <summary>Template. (Standard property)</summary>
        Template = 7,

        /// <summary>Last user to save the file. (Standard property)</summary>
        LastSavedBy = 8,

        /// <summary>Revision number. (Standard property)</summary>
        RevisionNumber = 9,

        /// <summary>Total editing time. (Standard property)</summary>
        TotalEditingTime = 10,

        /// <summary>Last printed date. (Standard property)</summary>
        LastPrinted = 11,

        /// <summary>Date of created. (Standard property)</summary>
        CreateTimeDate = 12,

        /// <summary>Date last saved. (Standard property)</summary>
        LastSavedTimeDate = 13,

        /// <summary>Number of pages. (Standard property)</summary>
        NumberOfPages = 14,

        /// <summary>Number of words. (Standard property)</summary>
        NumberOfWords = 15,

        /// <summary>Number of characters. (Standard property)</summary>
        NumberOfCharacters = 16,

        /// <summary>Thumbnail image. (Standard property)</summary>
        Thumbnail = 17,

        /// <summary>Application that created the file. (Standard property)</summary>
        NameOfCreatingApplication = 18,

        /// <summary>Security. (Standard property)</summary>
        Security = 19,

        /// <summary>A text string typed by the user that indicates what category the file belongs to (memo, proposal, and so on). It is useful for finding files of same type. (Extended property)</summary>
        Category = 0xFFFF + 0x00000002,

        /// <summary>Target format for presentation (35mm, printer, video, and so on). (Extended property)</summary>
        PresentationTarget = 0xFFFF + 0x00000003,

        /// <summary>Number of bytes. (Extended property)</summary>
        Bytes = 0xFFFF + 0x00000004,

        /// <summary>Number of lines. (Extended property)</summary>
        Lines = 0xFFFF + 0x00000005,

        /// <summary>Number of paragraphs. (Extended property)</summary>
        Paragraphs = 0xFFFF + 0x00000006,

        /// <summary>Number of slides. (Extended property)</summary>
        Slides = 0xFFFF + 0x00000007,

        /// <summary>Number of pages that contain notes. (Extended property)</summary>
        Notes = 0xFFFF + 0x00000008,

        /// <summary>Number of slides that are hidden. (Extended property)</summary>
        HiddenSlides = 0xFFFF + 0x00000009,

        /// <summary>Number of sound or video clips. (Extended property)</summary>
        MMClips = 0xFFFF + 0x0000000A,

        /// <summary>Set to True (-1) when scaling of the thumbnail is desired. If not set, cropping is desired. (Extended property)</summary>
        ScaleCrop = 0xFFFF + 0x0000000B,

        /// <summary>Internally used property indicating the grouping of different document parts and the number of items in each group. The titles of the document parts are stored in the TitlesofParts property. The HeadingPairs property is stored as a vector of variants, in repeating pairs of VT_LPSTR (or VT_LPWSTR) and VT_I4 values. The VT_LPSTR value represents a heading name, and the VT_I4 value indicates the count of document parts under that heading. (Extended property)</summary>
        HeadingPairs = 0xFFFF + 0x0000000C,

        /// <summary>Names of document parts. (Extended property)</summary>
        TitlesofParts = 0xFFFF + 0x0000000D,

        /// <summary>Manager of the project. (Extended property)</summary>
        Manager = 0xFFFF + 0x0000000E,

        /// <summary>Company name. (Extended property)</summary>
        Company = 0xFFFF + 0x0000000F,

        /// <summary>Boolean value to indicate whether the custom links are hampered by excessive noise, for all applications. (Extended property)</summary>
        LinksUpToDate = 0xFFFF + 0x00000010,
    }

    /// <summary>
    /// Properties for an Excel sheet.
    /// </summary>
    public sealed class TDocumentProperties
    {
        #region Privates
        private ExcelFile FXls;
        #endregion

        #region Constructors
        /// <summary>
        /// Creates a new TDocumentProperties instance and initializes it.
        /// </summary>
        internal TDocumentProperties(ExcelFile xls)
        {
            FXls = xls;
        }

        /// <summary>
        /// Returns a standard document property (Like Author, Title, etc.). This method returns an object that might be:
        /// <list type="bullet">
        /// <item><b>null</b>: If the property is not assigned.</item>
        /// <item><b>Int16, Int32, Int64, Single, Double, Decimal</b>: If the property is a number.</item>
        /// <item><b>DateTime</b>: If the property is a Date or a DateTime.</item>
        /// <item><b>String</b>: If the property is a string.</item>
        /// <item><b>Boolean</b>: If the property is a boolean.</item>
        /// <item><b>An array of byte (byte[])</b>: If the property is a Blob.</item>
        /// <item><b>An array of object (object[])</b>: If the property is an array. Each one of the members of the array must be of one of the types specified here.</item> 
        /// </list>
        /// </summary>
        /// <param name="PropertyId"></param>
        /// <returns></returns>
        public object GetStandardProperty(TPropertyId PropertyId)
        {
            return FXls.GetStandardProperty(PropertyId);
        }
        #endregion
    }
}
