using System;
using System.IO;
using System.Text;
using System.Diagnostics;
using System.Xml;

using FlexCel.Core;
#if (FRAMEWORK30)
using System.IO.Packaging;
#endif
using System.Globalization;
using System.Collections.Generic;

namespace FlexCel.XlsAdapter
{
    internal class TSheetRelationship
    {
        internal readonly Uri Uri;
        internal readonly string ContentType;
        internal readonly string RelationshipType;
        internal readonly string StartElement;

        internal TSheetRelationship(Uri aUri, string aContentType, string aRelationshipType, string aStartElement)
        {
            Uri = aUri;
            ContentType = aContentType;
            RelationshipType = aRelationshipType;
            StartElement = aStartElement;
        }
    }

    internal struct TRelationshipCache
    {
        internal Uri relsFile;
        internal Dictionary<string, List<string>> RelsByType;
        internal Dictionary<string, string> RelsById;

        internal void Clear()
        {
            relsFile = null;
            RelsByType = new Dictionary<string, List<string>>();
            RelsById = new Dictionary<string,string>();
        }
    }

    internal class TOpenXmlManager
    {
        #region Namespaces
        public const string documentRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        public const string sharedStringsRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
        public const string stylesRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
        public const string externalRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath";
        public const string xlPathMissing = "http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlPathMissing";
        public const string imageRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
        public const string themeRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";

        public const string MainNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        public const string RelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        public const string MacroNamespace = "http://schemas.microsoft.com/office/excel/2006/main";
        public const string DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
        public const string SpreadsheetDrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

        public const string MarkupCompatNamespace = "http://schemas.openxmlformats.org/markup-compatibility/2006";

        public const string ChartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";

        internal readonly static Uri WorkbookURI = new Uri("/xl/workbook.xml", UriKind.Relative);
        internal const string WorkbookContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
        internal const string WorkbookMacroEnabledContentType = "application/vnd.ms-excel.sheet.macroEnabled.main+xml";

        internal readonly static Uri SSTURI = new Uri("/xl/sharedStrings.xml", UriKind.Relative);
        internal const string SSTContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";

        internal readonly static Uri ConnectionsURI = new Uri("/xl/connections.xml", UriKind.Relative);
        internal const string ConnectionsContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml";
        internal const string ConnectionsRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections";

        internal readonly static Uri StylesURI = new Uri("/xl/styles.xml", UriKind.Relative);
        internal const string StylesContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";

        internal readonly static Uri ThemeURI = new Uri("/xl/theme/theme1.xml", UriKind.Relative);
        internal const string ThemeContentType = "application/vnd.openxmlformats-officedocument.theme+xml";

        internal readonly static Uri VBAURI = new Uri("/xl/vbaProject.bin", UriKind.Relative);
        internal const string VBARelationshipType = "http://schemas.microsoft.com/office/2006/relationships/vbaProject";
        internal const string VBAContentType = "application/vnd.ms-office.vbaProject";

        internal readonly static Uri CoreFilePropsURI = new Uri("/docProps/core.xml", UriKind.Relative);
        internal const string CoreFilePropsRelationshipType = "http://schemas.openxmlformats.org/officedocument/2006/relationships/metadata/core-properties";
        internal const string CoreFilePropsContentType = "application/vnd.openxmlformats-package.core-properties+xml";
        internal const string CoreFilePropsNamespace = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties ";

        internal readonly static Uri CustomFilePropsURI = new Uri("/docProps/custom.xml", UriKind.Relative);
        internal const string CustomFilePropsRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";
        internal const string CustomFilePropsContentType = "application/vnd.openxmlformats-officedocument.custom-properties+xml";
        internal const string CustomFilePropsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";

        internal const string CustomXmlDataRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
        internal const string CustomXmlDataContentType = "application/xml";

        internal const string CustomXmlDataPropsRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps";
        internal const string CustomXmlDataPropsContentType = "application/vnd.openxmlformats-officedocument.customXmlProperties+xml";
        internal const string CustomXmlDataPropsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/customXmlDataProps";

        internal const string PrinterSettingsBaseURI = "/xl/printerSettings/printerSettings";
        internal const string PrinterSettingsRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings";
        internal const string PrinterSettingsContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings";

        internal const string PivotCacheDefBaseURI = "/xl/pivotCache/pivotCacheDefinition";
        internal const string PivotCacheDefRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";
        internal const string PivotCacheDefContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml";

        internal const string PivotTableBaseURI = "/xl/pivotTables/pivotTable";
        internal const string PivotTableRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";
        internal const string PivotTableContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml";

        internal const string HyperlinkRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

        internal const string CommentsBaseURI = "/xl/comments";
        internal const string CommentsRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
        internal const string CommentsContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";

        internal const string DrawingBaseURI = "/xl/drawings/";
        internal const string DrawingRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
        internal const string DrawingContentType = "application/vnd.openxmlformats-officedocument.drawing+xml";

        internal readonly static Uri ThemeManagerURI = new Uri("/xl/theme/theme/themeManager.xml", UriKind.Relative);
        internal const string ThemeManagerContentType = "application/vnd.openxmlformats-officedocument.themeManager+xml";

        internal const string WorksheetBaseURI = "/xl/worksheets/";
        internal const string WorksheetRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
        internal const string WorksheetContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";

        internal const string MacrosheetBaseURI = "/xl/macrosheets/";
        internal const string MacrosheetRelationshipType = "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet";
        internal const string MacrosheetContentType = "application/vnd.ms-excel.macrosheet+xml";

        internal const string IntMacrosheetBaseURI = "/xl/macrosheets/";
        internal const string IntMacrosheetRelationshipType = "http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet";
        internal const string IntMacrosheetContentType = "application/vnd.ms-excel.intlmacrosheet+xml";

        internal const string DialogsheetBaseURI = "/xl/dialogsheets/";
        internal const string DialogsheetRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet";
        internal const string DialogsheetContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml";

        internal const string ChartsheetBaseURI = "/xl/chartsheets/";
        internal const string ChartsheetRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
        internal const string ChartsheetContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml";

        internal const string ExternalLinkBaseURI = "/xl/externalLinks/";
        internal const string ExternalLinkRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink";
        internal const string ExternalLinkContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml";

        internal const string OleRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject";

        internal const int RelIdSST = 1; //starts at sheet.count
        internal const int RelIdStyles = 2; //starts at sheet.count
        internal const int RelIdThemes = 3; //starts at sheet.count

        internal const int RelIdExternalLinks = 4; //starts at sheet.count. ends at sheet.count+ externallinks.count
        internal const int RelIdVBA = 5; //starts at sheet.count + externallinks.count
        internal const int RelIdCustomXML = 6; //starts at sheet.count + externallinks.count  Ends at sheet.count + externallinks.count + customxml.count

        internal const int RelIdPivotCaches = 6; //starts at sheet.count + externallinks.count + customxml.count  Ends at sheet.count + externallinks.count + customxml.count + pivotcache.count
        internal const int RelIdConnections = 7; //starts at sheet.count + externallinks.count + customxml.count + pivotcache.count

        internal const int RelIdExtLinkFiles = 1; //this goes into every relationship, so it is always 1.
        internal const int RelIdThemeManager = 1;

        //null parent
        internal const int RelIdWorkbook = 1;
        internal const int RelIdCoreFileProps = 2;
        internal const int RelIdExtendedFileProps = 3;
        internal const int RelIdCustomFileProps = 4;

        //Legacy Drawing
        internal const string LegDrawMainNamespace = "urn:schemas-microsoft-com:vml";
        internal const string LegDrawOfficeNamespace = "urn:schemas-microsoft-com:office:office";
        internal const string LegDrawExcelNamespace = "urn:schemas-microsoft-com:office:excel";

        internal const string LegDrawBaseURI = "/xl/drawings/vmlDrawing";
        internal const string LegDrawRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
        internal const string LegDrawContentType = "application/vnd.openxmlformats-officedocument.vmlDrawing";

        #endregion

        internal static Uri ResolvePartUri(Uri a, Uri b)
        {
            if (FlxUtils.IsMonoRunning())
            {
                //Workarounds mono bug: https://bugzilla.novell.com/show_bug.cgi?id=675383
                //When fixed, this method can be replaced by simple "return PackUriHelper.ResolvePartUri(a, b);"
                if (!b.ToString().StartsWith("/")) return PackUriHelper.ResolvePartUri(a, b);

                return PackUriHelper.ResolvePartUri(a, new Uri(b.ToString().Substring(1), UriKind.Relative));
            }
            return PackUriHelper.ResolvePartUri(a, b);

        }
    }

    internal struct TXlState
    {
        internal XmlReader Reader;
        internal PackagePart Part;
        internal string DefaultNamespace;

        internal TXlState(XmlReader aReader, PackagePart aPart, string aDefaultNamespace)
        {
            Reader = aReader;
            Part = aPart;
            DefaultNamespace = aDefaultNamespace;
        }
    }

    /// <summary>
    /// A class for reading Open XML zipped files. We use the packaging API in this immplementation, so this requires .NET 3.0 or newer.
    /// </summary>
    internal class TOpenXmlReader : TOpenXmlManager, IDisposable
    {
        #region Variables
#if (FRAMEWORK30)
        private PackagePart FCurrentPart;
        private Package xlPackage;
        private TRelationshipCache CurrentPartRelationshipCache;
#endif
        private string MainFileName;
        private TExcelFileErrorActions ErrorActions = TExcelFileErrorActions.None; 

        private Uri workbookUri;
        private XmlReader xlReader;
        private Stack<TXlState> xlPendingReaders;
        internal bool NotXlsx;
        XmlReaderSettings XmlSettings;
        internal string DefaultNamespace;

        private Stream EncryptedStream;
        internal TEncryptionData Encryption;
        #endregion

        #region Constructor
        static readonly byte[] PkZipSignature = new byte[] { 0x50, 0x4b, 0x03, 0x04 };

        private TOpenXmlReader(Stream SimpleStream)
        {
            XmlSettings = new XmlReaderSettings();
            XmlSettings.IgnoreWhitespace = true;
            xlReader = XmlReader.Create(SimpleStream, XmlSettings);
        }

        internal static TOpenXmlReader CreateFromSimpleStream(Stream DataStream)
        {
            return new TOpenXmlReader(DataStream);
        }

        internal TOpenXmlReader(Stream DataStream, bool AvoidExceptions, TEncryptionData aEncryption, string aMainFileName, TExcelFileErrorActions aErrorActions)
        {
            ErrorActions = aErrorActions;
            MainFileName = aMainFileName;
            if (MainFileName == null) MainFileName = String.Empty;

            Encryption = aEncryption;
            long StreamPosition = DataStream.Position;

            Stream RealDataStream = DataStream;

            //EncryptedPackageEnvelope doesn't seem to work at this moment, and it isn't supported in mono.
            if (EncryptedDocReader.IsValidFile(DataStream))
            {
                DataStream.Position = StreamPosition;
                if (aEncryption == null) XlsMessages.ThrowException(XlsErr.ErrInvalidPassword);
                RealDataStream = EncryptedDocReader.Decrypt(DataStream, aEncryption);
                EncryptedStream = RealDataStream;

                DataStream.Position = StreamPosition; //DataStream won't be used anymore
                StreamPosition = 0;
            }

            if (AvoidExceptions) //verify the header.
            {
                byte[] Signature = new byte[PkZipSignature.Length];
                RealDataStream.Read(Signature, 0, Signature.Length);
                RealDataStream.Position = StreamPosition;
                if (!BitOps.CompareMem(Signature, PkZipSignature))
                {
                    NotXlsx = true;
                    return;
                }
            }

            // mono bug: https://bugzilla.novell.com/show_bug.cgi?id=675379
            // will cause this to fail when opening in mono. Sadly only way to fix it is to fix the mono code.
            // it *could* be workarounded by opening in readwrite mode, but really, you don't want to open a file
            // for reading in write mode. You could corrupt a file you shouldn't be touching.
            xlPackage = Package.Open(RealDataStream);

            //  Get the main document part (workbook.xml).
            workbookUri = GetUriForRelationship(documentRelationshipType);
            if (workbookUri == null) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);

            XmlSettings = new XmlReaderSettings();
            XmlSettings.IgnoreWhitespace = true;

            xlPendingReaders = new Stack<TXlState>();

        }
        #endregion

        #region Parts
        private PackagePart CurrentPart
        {
            get
            {
                return FCurrentPart;
            }
            set
            {
                FCurrentPart = value;
                CurrentPartRelationshipCache.Clear();
            }
        }

        Uri GetUriForRelationship(string relationshipType)
        {
            foreach (PackageRelationship relationship in xlPackage.GetRelationshipsByType(relationshipType))
            {
                //  There should only be one document part in the package. 
                return ResolvePartUri(new Uri("/", UriKind.Relative), relationship.TargetUri);
            }

            return null;
        }

        internal List<Uri> GetUrisForCurrentPartRelationship(string relationshipType)
        {
            List<Uri> Result = new List<Uri>();
            foreach (string TargetUriStr in GetCurrentPartRelationshipsByType(relationshipType))
            {
                Result.Add(ResolvePartUri(CurrentPart.Uri, new Uri(TargetUriStr, UriKind.RelativeOrAbsolute)));
            }

            return Result;
        }

        internal void SelectWorkbook(out bool MacroEnabled)
        {
            CurrentPart = xlPackage.GetPart(workbookUri);
            MacroEnabled = CurrentPart.ContentType == WorkbookMacroEnabledContentType;
            if (xlReader != null) xlReader.Close();
            xlReader = XmlReader.Create(CurrentPart.GetStream(), XmlSettings);
            DefaultNamespace = MainNamespace;
        }

        private void SelectWorkbookPart(string RelationshipType, string aDefaultNamespace)
        {
            if (xlReader != null) xlReader.Close();
            xlReader = null;
            CurrentPart = null;

            PackagePart workbookPart = xlPackage.GetPart(workbookUri);
            PackageRelationshipCollection Rels = workbookPart.GetRelationshipsByType(RelationshipType);

            foreach (PackageRelationship relationship in Rels)
            {
                //Just one relationship
                Uri sharedStringsUri = ResolvePartUri(workbookUri, relationship.TargetUri);
                CurrentPart = xlPackage.GetPart(sharedStringsUri);

                xlReader = XmlReader.Create(CurrentPart.GetStream(), XmlSettings);
                DefaultNamespace = aDefaultNamespace;
                return; //always use the first.
            }
        }

        internal void SelectSheet(string RelationshipId)
        {
            SelectPart(RelationshipId, MainNamespace);
        }

        internal void SelectPart(string RelationshipId, string aDefaultNamespace)
        {
            PackagePart workbookPart = xlPackage.GetPart(workbookUri);
            PackageRelationship sheetRelation = workbookPart.GetRelationship(RelationshipId);
            Uri sheetUri = ResolvePartUri(workbookUri, sheetRelation.TargetUri);
            CurrentPart = xlPackage.GetPart(sheetUri);

            if (xlReader != null) xlReader.Close();
            xlReader = XmlReader.Create(CurrentPart.GetStream(), XmlSettings);
            DefaultNamespace = aDefaultNamespace;
        }

        internal void SelectMasterPart(string RelationshipType, string aDefaultNamespace)
        {
            if (xlReader != null) xlReader.Close();
            xlReader = null;
            CurrentPart = null;

            Uri PartUri = GetUriForRelationship(RelationshipType);
            if (PartUri == null) return;
            CurrentPart = xlPackage.GetPart(PartUri);

            xlReader = XmlReader.Create(CurrentPart.GetStream(), XmlSettings);
            DefaultNamespace = aDefaultNamespace;
        }

        internal void SelectMasterPart(Uri PartUri, string aDefaultNamespace)
        {
            if (xlReader != null) xlReader.Close();
            xlReader = null;
            CurrentPart = null;

            if (PartUri == null) return;
            CurrentPart = xlPackage.GetPart(PartUri);

            xlReader = XmlReader.Create(CurrentPart.GetStream(), XmlSettings);
            DefaultNamespace = aDefaultNamespace;
        }

        internal void SelectFromCurrentPartAndPush(string RelationshipId, string aDefaultNamespace, bool Legacy)
        {
            string childRelation = GetCurrentPartRelationship(RelationshipId);
            Uri childUri = ResolvePartUri(CurrentPart.Uri, new Uri(childRelation, UriKind.RelativeOrAbsolute));
            PushPart();
            CurrentPart = xlPackage.GetPart(childUri);

            if (Legacy)
            {
                //Sadly the new XmlReader that is created with XmlReader.Create can't read malformed xml, and we need to do so when reading legacy shapes.
                //As XmlTextReader is obsolete (and not available in silverlight) we will have to end up writing our own xlm parser :(
                xlReader = new XmlTextReader(CurrentPart.GetStream());
                ((XmlTextReader)xlReader).WhitespaceHandling = WhitespaceHandling.Significant;
            }
            else
            {
                xlReader = XmlReader.Create(CurrentPart.GetStream(), XmlSettings);
            }
            DefaultNamespace = aDefaultNamespace;
        }

        private string GetCurrentPartRelationship(string RelationshipId)
        {
  /*          try
            {
                return CurrentPart.GetRelationship(RelationshipId).TargetUri.OriginalString;
            }
            catch (UriFormatException)
            {*/
                return ManuallyReadRelationshipById(RelationshipId);
          //  }
        }

        private IEnumerable<string> GetCurrentPartRelationshipsByType(string RelationshipType)
        {
            /*          try
                      {
                          return CurrentPart.GetRelationship(RelationshipId).TargetUri.OriginalString;
                      }
                      catch (UriFormatException)
                      {*/
            return ManuallyReadRelationshipByType(RelationshipType);
            //  }
        }

        private string ManuallyReadRelationshipById(string RelationshipId)
        {
            Uri relsUri = PackUriHelper.GetRelationshipPartUri(CurrentPart.Uri);
            if (CurrentPartRelationshipCache.relsFile != relsUri)
            {
                CacheCurrentRelations(relsUri);
            }

            return CurrentPartRelationshipCache.RelsById[RelationshipId];
        }

        private IEnumerable<string> ManuallyReadRelationshipByType(string RelationshipType)
        {
            Uri relsUri = PackUriHelper.GetRelationshipPartUri(CurrentPart.Uri);
            if (CurrentPartRelationshipCache.relsFile != relsUri)
            {
                CacheCurrentRelations(relsUri);
            }

            if (!CurrentPartRelationshipCache.RelsByType.ContainsKey(RelationshipType)) return new List<string>(0);
            return CurrentPartRelationshipCache.RelsByType[RelationshipType];
        }

        private void CacheCurrentRelations(Uri relsUri)
        {
            CurrentPartRelationshipCache.Clear();
            CurrentPartRelationshipCache.relsFile = relsUri;
            if (!xlPackage.PartExists(relsUri)) return;
            PackagePart pPart = xlPackage.GetPart(relsUri);
            {
                using (Stream xl1 = pPart.GetStream())
                {
                    XmlDocument doc = new XmlDocument();
                    doc.Load(xl1);
                    foreach (XmlNode node in doc.GetElementsByTagName("Relationship"))
                    {
                        XmlAttribute id = node.Attributes["Id"];
                        if (id != null)
                        {
                            CurrentPartRelationshipCache.RelsById[id.Value] = node.Attributes["Target"].Value;
                        }

                        XmlAttribute linktype = node.Attributes["Type"];
                        if (linktype != null)
                        {
                            if (!CurrentPartRelationshipCache.RelsByType.ContainsKey(linktype.Value))
                            {
                                CurrentPartRelationshipCache.RelsByType[linktype.Value] = new List<string>();
                            }

                            CurrentPartRelationshipCache.RelsByType[linktype.Value].Add(node.Attributes["Target"].Value);
                        }
                    }
                }
            }
        }

        internal void PushPart() //Always call PopPart here.
        {
            xlPendingReaders.Push(new TXlState(xlReader, CurrentPart, DefaultNamespace));
            xlReader = null;
            CurrentPart = null;
            DefaultNamespace = null;
        }

        internal void PopPart()
        {
            if (xlReader != null) xlReader.Close();
            TXlState state = xlPendingReaders.Pop();
            xlReader = state.Reader;
            CurrentPart = state.Part;
            DefaultNamespace = state.DefaultNamespace;
        }

        internal void SelectSST()
        {
            SelectWorkbookPart(sharedStringsRelationshipType, MainNamespace);
        }

        internal string GetExternalLink(string id)
        {
           return GetCurrentPartRelationship(id);
        }

        internal void SelectStyles()
        {
            SelectWorkbookPart(stylesRelationshipType, MainNamespace);
        }

        internal void SelectConnections()
        {
            SelectWorkbookPart(ConnectionsRelationshipType, MainNamespace);
        }

        internal void SelectTheme()
        {
            SelectWorkbookPart(themeRelationshipType, DrawingNamespace);
        }

        internal void SelectCoreProps()
        {
            SelectMasterPart(CoreFilePropsRelationshipType, CoreFilePropsNamespace);
        }

        internal void SelectCustomFileProps()
        {
            SelectMasterPart(CustomFilePropsRelationshipType, CustomFilePropsNamespace);
        }

        internal bool Eof
        {
            get
            {
                return xlReader == null || xlReader.EOF;
            }
        }

        internal byte[] ReadOtherParts()
        {
            return null;
        }

        internal string CurrentPartContentType
        {
            get
            {
                return CurrentPart.ContentType;
            }
        }
        #endregion

        #region Escape
        private static string UnescapeString(string s)
        {
            //return XmlConvert.DecodeName(s); //this has issues with "_X" that shouldn't be recognized (only "_x" should)

            if (String.IsNullOrEmpty(s)) return s;
            //This should be faster since "_x" is almost never found in strings. But it is slower...
            //int Start = s.IndexOf("_x", StringComparison.Ordinal);
            //if (Start < 0) return s;
            int Start = 0;

            StringBuilder r = null;
            int rCopied = 0;
            for (int i = Start; i < s.Length; i++)
            {
                if (s[i] == '_' && i + 6 < s.Length && s[i + 1] == 'x' && s[i + 6] == '_')
                {
                    bool HasHexDigits = true;
                    int spv = 0;
                    int shift = 12;
                    for (int k = i + 2; k < i + 6; k++) //this won't be so common as to do a loop unroll.
                    {
                        if (s[k] >= '0' && s[k] <= '9') { spv += (s[k] - '0') << shift; shift -= 4; continue; }
                        if (s[k] >= 'a' && s[k] <= 'f') { spv += ((s[k] - 'a') + 10) << shift; shift -= 4; continue; }
                        if (s[k] >= 'A' && s[k] <= 'F') { spv += ((s[k] - 'A') + 10) << shift; shift -= 4; continue; }
                        HasHexDigits = false;
                        break;
                    }

                    if (!HasHexDigits) continue;

                    if (r == null) r = new StringBuilder(s.Length + 30);
                    r.Append(s, rCopied, i - rCopied);
                    r.Append((char) spv);
                    i += 6;
                    rCopied = i + 1;

                }

            }

            if (r == null) return s;

            r.Append(s, rCopied, s.Length - rCopied);
            return r.ToString();
        }

        #endregion

        #region Xml

        internal bool NextTag()
        {
            if (xlReader == null) return false;
            bool ok = xlReader.Read();
            while (ok && xlReader.NodeType != XmlNodeType.Element && xlReader.NodeType != XmlNodeType.EndElement)
            {
                ok = xlReader.Read();
            }
            return ok;
        }

        internal bool NextTagNoSkip()
        {
            return xlReader.Read();
        }

        internal void FinishTag()
        {
            string CurrentTag = xlReader.Name;
            if (!NextTag()) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
            if (AtEndElement() && xlReader.Name == CurrentTag)
            {
                if (!NextTag()) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
            }
        }

        internal void FinishTagAndIgnoreChildren()
        {
            string CurrentTag = xlReader.Name;
            if (!IsSimpleTag)
            {
                while (!AtEndElement() || xlReader.Name != CurrentTag)
                {
                    if (!NextTag()) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
                }
            }
            if (!NextTag()) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);

        }

        internal string GetXml()
        {
            return xlReader.ReadOuterXml();
        }

        internal string RecordName()
        {
            if (xlReader.NamespaceURI == DefaultNamespace)
            {
                return xlReader.LocalName;
            }
            return xlReader.NamespaceURI + ":" + xlReader.LocalName;
        }

        internal string RecordLocalName()
        {
            return xlReader.LocalName;
        }

        internal string RecordNamespace()
        {
            return xlReader.NamespaceURI;
        }

        internal string ReadValueAsString()
        {
            return UnescapeString(xlReader.ReadElementContentAsString());
        }

        internal string ReadInnerXml()
        {
            return xlReader.ReadInnerXml();
        }

        internal string ReadLegacyValue()
        {
#if (!FRAMEWORK40) 
            try
            {
              // XmlReader xlReader2 = xlReader.ReadSubtree();
              // return xlReader2.ReadInnerXml();
            }
            catch 
            {
            }
#endif
            // NET 3.5 can have issues with this, but there is no other way to read malformed xml.
            string StartNode = xlReader.Name;
            char[] buff = new char[4096]; //a buffer of 1 here will cause issues in mono. 
            int Read;
            StringBuilder Result = new StringBuilder();
            while ((Read = ((XmlTextReader)xlReader).ReadChars(buff, 0, buff.Length)) > 0)
            {
                Result.Append(buff, 0, Read);
                if (xlReader.Name != StartNode) break;
            }

            return Result.ToString();
        }


        internal double ReadValueAsDouble()
        {
            return xlReader.ReadElementContentAsDouble();
        }

        internal int ReadValueAsInt()
        {
            return xlReader.ReadElementContentAsInt();
        }

        internal TRichString ReadValueAsRichString(IFlexCelFontList FontList)
        {
            return TxSSTRecord.LoadRichStringFromXml(this, FontList);
        }

        internal bool AtEndElement()
        {
            if (xlReader == null) return true;
            return xlReader.NodeType == XmlNodeType.EndElement;
        }

        internal bool AtEndElement(string StartElement)
        {
            if (xlReader == null) return true;
            bool Result = xlReader.NodeType == XmlNodeType.EndElement;
            if (Result) CheckEndElement(StartElement);
            return Result;
        }

        internal void CheckEndElement(string StartElement)
        {
            if (RecordName() != StartElement) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
            NextTag();
        }

        internal bool IsSimpleTag
        {
            get
            {
                return xlReader.IsEmptyElement;
            }
        }

        internal bool HasAttribute(string AttrName)
        {
            return xlReader.GetAttribute(AttrName) != null;
        }

        internal TXlsxAttribute[] GetAttributes()
        {
            TXlsxAttribute[] Result = new TXlsxAttribute[xlReader.AttributeCount];
            int i = 0;
            if (xlReader.HasAttributes)
            {
                while (xlReader.MoveToNextAttribute())
                {
                    Result[i] = new TXlsxAttribute(xlReader.NamespaceURI, xlReader.LocalName, xlReader.Value);
                    i++;
                } 
            }
            // Move the reader back to the element node.
            xlReader.MoveToElement();

            return Result;
        }

        internal string GetAttribute(string AttrName)
        {
            return UnescapeString(xlReader.GetAttribute(AttrName));
        }

        internal string GetAttribute(string AttrName, string AttNamespace)
        {
            return UnescapeString(xlReader.GetAttribute(AttrName, AttNamespace));
        }

        internal string GetRelationship(string relName)
        {
            return UnescapeString(xlReader.GetAttribute(relName, RelationshipNamespace));
        }

        internal byte[] GetPart(string RelationshipType, int Start, int ExtraLen)
        {
            PackagePart workbookPart = xlPackage.GetPart(workbookUri);
            PackageRelationshipCollection Rels = workbookPart.GetRelationshipsByType(RelationshipType);

            foreach (PackageRelationship relationship in Rels)
            {
                string ct;
                return GetRelationshipData(ExtraLen, Start, workbookPart.Uri, relationship.TargetUri, out ct);
            }

            return null;
        }

        internal byte[] GetRelationshipData(string relName, int ExtraLen, int Start, out string ContentType, out string FileName)
        {
            return GetRelationshipData(relName, ExtraLen, Start, RelationshipNamespace, out ContentType, out FileName);
        }

        internal byte[] GetRelationshipData(string relName, int ExtraLen, int Start, string TargetNamespace,
            out string ContentType, out string FileName)
        {
            ContentType = null;
            FileName = null;

            string RelId = xlReader.GetAttribute(relName, TargetNamespace);
            if (RelId == null) return null;
            FileName = GetCurrentPartRelationship(RelId);
            Uri TargetUri = new Uri(FileName, UriKind.RelativeOrAbsolute);
            return GetRelationshipData(ExtraLen, Start, CurrentPart.Uri, TargetUri, out ContentType);
        }

        private byte[] GetRelationshipData(int ExtraLen, int Start, Uri SourceUri, Uri TargetUri, out string ContentType)
        {
            Uri PartUri = ResolvePartUri(SourceUri, TargetUri);
            if (!xlPackage.PartExists(PartUri))
            {
                if ((ErrorActions & TExcelFileErrorActions.ErrorOnXlsxMissingPart) != 0)
                {
                    XlsMessages.ThrowException(XlsErr.ErrMissingPart, MainFileName, PartUri.ToString());
                }
                ContentType = null;
                if (FlexCelTrace.Enabled)
                {
                    FlexCelTrace.Write(new TXlsxMissingPartError(
                            XlsMessages.GetString(XlsErr.ErrMissingPart, MainFileName, PartUri.ToString()), MainFileName, PartUri.ToString()));
                }
                return null;
            }

            PackagePart pp1 = xlPackage.GetPart(PartUri);
            ContentType = pp1.ContentType;
            using (Stream xl1 = pp1.GetStream())
            {
                byte[] Result = new byte[xl1.Length + ExtraLen];
                xl1.Read(Result, Start, (int)xl1.Length);
                return Result;
            }
        }
        

        internal bool GetAttributeAsBool(string AttrName, bool DefaultValue)
        {
            string s = xlReader.GetAttribute(AttrName);
            if (s == null || s.Length == 0) return DefaultValue;
            return FlxConvert.ToXlsxBoolean(s);
        }

        internal double GetAttributeAsDouble(string AttrName, double DefaultValue)
        {
            string s = xlReader.GetAttribute(AttrName);
            if (s == null || s.Length == 0) return DefaultValue;
            return Convert.ToDouble(s, CultureInfo.InvariantCulture);
        }

        internal int GetAttributeAsInt(string AttrName, int DefaultValue)
        {
            string s = xlReader.GetAttribute(AttrName);
            if (s == null || s.Length == 0) return DefaultValue;
            return Convert.ToInt32(s, CultureInfo.InvariantCulture);
        }

        internal long GetAttributeAsLong(string AttrName, int DefaultValue)
        {
            string s = xlReader.GetAttribute(AttrName);
            if (s == null || s.Length == 0) return DefaultValue;
            return Convert.ToInt64(s, CultureInfo.InvariantCulture);
        }

        internal long GetAttributeAsHex(string AttrName, long DefaultValue)
        {
            string s = xlReader.GetAttribute(AttrName);
            if (s == null || s.Length == 0) return DefaultValue;
            
            return Convert.ToInt64(s, 16);
        }

        internal double GetAttributeAsAngle(string AttrName, int DefaultValue)
        {
            return GetAttributeAsInt(AttrName, DefaultValue) / 60000.0;
        }

        internal double GetAttributeAsPercent(string AttrName, double DefaultValue)
        {
            string s = xlReader.GetAttribute(AttrName);
            if (s == null) return 0;
            if (s.EndsWith("%"))
            {
                return Double.Parse(s.Substring(0, s.Length - 1), CultureInfo.InvariantCulture) / 100.0;
            }

            return Double.Parse(s, CultureInfo.InvariantCulture) / 100000.0;
        }

        internal TDrawingCoordinate GetAttributeAsDrawingCoord(string AttrName, TDrawingCoordinate DefaultValue)
        {
            string s = xlReader.GetAttribute(AttrName);
            if (s == null) return DefaultValue;
            return GetAttributeAsDrawingCoord(s);
        }

        internal static TDrawingCoordinate GetAttributeAsDrawingCoord(string s)
        {
            if (s.Length > 2)
            {
                switch (s.Substring(s.Length - 2))
                {
                    case "cm": return TDrawingCoordinate.FromCm(Double.Parse(s.Substring(0, s.Length - 2), CultureInfo.InvariantCulture));
                    case "mm": return TDrawingCoordinate.FromMm(Double.Parse(s.Substring(0, s.Length - 2), CultureInfo.InvariantCulture));
                    case "in": return TDrawingCoordinate.FromInches(Double.Parse(s.Substring(0, s.Length - 2), CultureInfo.InvariantCulture));
                    case "pt": return TDrawingCoordinate.FromPoints(Double.Parse(s.Substring(0, s.Length - 2), CultureInfo.InvariantCulture));
                    case "pc": return TDrawingCoordinate.FromPc(Double.Parse(s.Substring(0, s.Length - 2), CultureInfo.InvariantCulture));
                    case "pi": return TDrawingCoordinate.FromPi(Double.Parse(s.Substring(0, s.Length - 2), CultureInfo.InvariantCulture));
                }
            }

            return new TDrawingCoordinate(long.Parse(s, CultureInfo.InvariantCulture));
        }

        internal TCellAddress GetAttributeAsAddress(string AttrName)
        {
            string s = xlReader.GetAttribute(AttrName);
            if (s == null) return null;
            return new TCellAddress(s);
        }

        internal TXlsCellRange GetAttributeAsRange(string AttrName, bool ReturnOneBased)
        {
            string s = xlReader.GetAttribute(AttrName);
            if (s == null) return null;
            return GetOneRangeAtt(s, ReturnOneBased);
        }

        private TXlsCellRange GetOneRangeAtt(string s, bool ReturnOneBased)
        {
            TXlsCellRange Result = new TXlsCellRange(s);
            if (ReturnOneBased || Result == null) return Result;
            return Result.Dec();
        }

        internal TXlsCellRange[] GetAttributeAsSeriesOfRanges(string AttrName, bool ReturnOneBased)
        {
            List<TXlsCellRange> Result = new List<TXlsCellRange>();
            string s = xlReader.GetAttribute(AttrName);
            if (s == null) return null;
            string[] sa = s.Split(' ');
            foreach (string s1 in sa)
            {
                if (string.IsNullOrEmpty(s1)) continue;
                string s2 = s1.Trim();
                if (s2.Length == 0) continue;
                Result.Add(GetOneRangeAtt(s2, ReturnOneBased));
            }

            return Result.ToArray();
        }


        #endregion

        #region IDisposable Members

        public void Dispose()
        {
            if (xlPackage != null) xlPackage.Close();
            if (xlReader != null) xlReader.Close();
            if (xlPendingReaders != null) //might be null if we exited from the constructor.
            {
                while (xlPendingReaders.Count > 0)
                {
                    TXlState x = xlPendingReaders.Pop();
                    if (x.Reader != null) x.Reader.Close();
                }
            }

            if (EncryptedStream != null)
            {
                EncryptedStream.Dispose();
                EncryptedStream = null;
            }

            GC.SuppressFinalize(this);
        }

        #endregion


    }

    internal sealed class TOpenXmlWriter : TOpenXmlManager, IDisposable
    {
        #region Variables
        internal const string FlexCelRid = "flId"; //we keep it different from Excel "rId" so there are no clashes in unknown relationships preserved.
        private XmlWriter xlWriter;
        private Package xlPackage;
        private PackagePart CurrentPart;

        private List<TTagDef> PendingTags;

        private Dictionary<string, int> UsedSheetIds;
        internal string DefaultNamespacePrefix;
        private XmlWriterSettings WriterSettings;
        #endregion

        #region Constructor
        public TOpenXmlWriter(Stream DataStream, bool AllowOverwritingFiles)
        {
            FileMode fm = FileMode.CreateNew;   
            if (AllowOverwritingFiles) fm = FileMode.Create;

            if (!DataStream.CanRead) XlsMessages.ThrowException(XlsErr.ErrStreamNeedsReadAccess);
            xlPackage = Package.Open(DataStream, fm);

            UsedSheetIds = new Dictionary<string, int>();
            PendingTags = new List<TTagDef>();

            WriterSettings = new XmlWriterSettings();
            //WriterSettings.NamespaceHandling = NamespaceHandling.OmitDuplicates;
        }
        #endregion

        #region Sheet Relationships
        internal TSheetRelationship GetSheetRelationship(TSheetType SheetType, bool International)
        {
            switch (SheetType)
            {
                case TSheetType.Worksheet:
                    return new TSheetRelationship(GetSheetFile(WorksheetBaseURI + "sheet"), WorksheetContentType, WorksheetRelationshipType, "worksheet");

                case TSheetType.Chart:
                    return new TSheetRelationship(GetSheetFile(ChartsheetBaseURI + "sheet"), ChartsheetContentType, ChartsheetRelationshipType, "chartsheet");

                case TSheetType.Dialog:
                    return new TSheetRelationship(GetSheetFile(DialogsheetBaseURI + "sheet"), DialogsheetContentType, DialogsheetRelationshipType, "dialogsheet");

                case TSheetType.Macro:
                    if (International)
                    {
                        return new TSheetRelationship(GetSheetFile(IntMacrosheetBaseURI + "intlsheet"), IntMacrosheetContentType, IntMacrosheetRelationshipType, "macrosheet");
                    }
                    return new TSheetRelationship(GetSheetFile(MacrosheetBaseURI + "sheet"), MacrosheetContentType, MacrosheetRelationshipType, "macrosheet");

                case TSheetType.Other:
                default:
                    FlxMessages.ThrowException(FlxErr.ErrInternal);
                    break;
            }

            return null; //to compile
        }

        private Uri GetSheetFile(string BaseFile)
        {
            int NewSheetId;
            if (UsedSheetIds.TryGetValue(BaseFile, out NewSheetId))
            {
                UsedSheetIds[BaseFile]++;
            }
            else
            {
                NewSheetId = 1;
                UsedSheetIds[BaseFile] = 2;
            }

            return GetFileUri(BaseFile, NewSheetId);
        }

        internal static Uri GetFileUri(string BaseFile, int Id)
        {
            return new Uri(BaseFile + Id.ToString(CultureInfo.InvariantCulture) + ".xml", UriKind.Relative);
        }

        #endregion

        #region Parts
        internal void CreatePart(Uri PartUri, string ContentType)
        {
            CurrentPart = xlPackage.CreatePart(PartUri, ContentType, CompressionOption.Maximum);

            if (xlWriter != null)
            {
                xlWriter.Close();
                xlWriter = null;
            }

            xlWriter = XmlWriter.Create(CurrentPart.GetStream(), WriterSettings);
        }

        internal void WritePart(Uri PartUri, string ContentType, byte[] Data)
        {
            WritePart(PartUri, ContentType, Data, 0);
        }

        internal void WritePart(Uri PartUri, string ContentType, byte[] Data, int DataOfs)
        {
            PackagePart Part = xlPackage.CreatePart(PartUri, ContentType, CompressionOption.Maximum);
            {
                using (Stream st = Part.GetStream())
                {
                    if (Data != null) st.Write(Data, DataOfs, Data.Length - DataOfs);
                }
            }
        }

        internal void CreateRelationshipFromUri(Uri SourceUri, string RelationshipType, int Id)
        {
            CreateRelationshipFromUri(SourceUri, RelationshipType, GetRId(Id));
        }

        public static string GetRId(int Id)
        {
            return FlexCelRid + Id.ToString(CultureInfo.InvariantCulture);
        }

        internal void CreateRelationshipFromUri(Uri SourceUri, string RelationshipType, string RelId)
        {
            if (SourceUri == null)
            {
                Uri TargetUri = MakeRelativeUri(String.Empty, CurrentPart.Uri);
                xlPackage.CreateRelationship(TargetUri, TargetMode.Internal, RelationshipType, RelId);
            }
            else
            {
                PackagePart SourcePart = xlPackage.GetPart(SourceUri);
                Uri TargetUri = MakeRelativeUri(SourceUri.OriginalString, CurrentPart.Uri);
                SourcePart.CreateRelationship(TargetUri, TargetMode.Internal, RelationshipType, RelId);
            }
        }

        internal void CreateRelationshipToUri(Uri TargetUri, TargetMode aTargetMode, string RelationshipType, string RelId)
        {
            CurrentPart.CreateRelationship(TargetUri, aTargetMode, RelationshipType, RelId);
        }

        internal void CreateRelationshipToUri(Uri SourceUri, Uri TargetUri, TargetMode aTargetMode, string RelationshipType, string RelId)
        {
            PackagePart SourcePart = xlPackage.GetPart(SourceUri); 
            SourcePart.CreateRelationship(TargetUri, aTargetMode, RelationshipType, RelId);
        }



        internal Uri CurrentUri
        {
            get { return CurrentPart.Uri; }
        }

        private static Uri MakeRelativeUri(string SourceUri, Uri TargetUri)
        {
            //there should be a better way :(
            Uri AbsSource = new Uri("http://root" + SourceUri);
            Uri AbsTarget = new Uri("http://root" + TargetUri.OriginalString);
            return AbsSource.MakeRelativeUri(AbsTarget);

        }
        #endregion

        #region Escape
        private static string EscapeString(string s, out bool NeedsPreserveWhitespace)
        {
            NeedsPreserveWhitespace = false;
            if (String.IsNullOrEmpty(s)) return s;
            //return XmlConvert.EncodeName(s);  //EncodeName encodes way too much, it will encode characters Excel won't understand.

            StringBuilder r = null;
            int rCopied = 0;
            bool LastWasWhiteSpace = true;
            for (int i = 0; i < s.Length; i++)
            {
                if (s[i] > 0xFFFD || s[i] <= 0x20)
                {
                    bool IsSpace = s[i] == ' ' || s[i] == '\t' || s[i] == '\n' || s[i] == '\r';
                    if (IsSpace)
                    {
                        if (LastWasWhiteSpace) NeedsPreserveWhitespace = true;
                        else LastWasWhiteSpace = true;
                        continue;
                    }

                    LastWasWhiteSpace = false;
                }
                else
                {
                    LastWasWhiteSpace = false;
                    if (s[i] == '_' && i + 6 < s.Length && s[i+1] == 'x' && s[i + 6] == '_')
                    {
                        bool HasHexDigits = true;
                        for (int k = i + 2; k < i + 6; k++) //this won't be so common as to do a loop unroll.
                        {
                            if (s[k] >= '0' && s[k] <= '9') continue;
                            if (s[k] >= 'a' && s[k] <= 'f') continue;
                            if (s[k] >= 'A' && s[k] <= 'F') continue;
                            HasHexDigits = false;
                            break;
                        }

                        if (!HasHexDigits) continue;
                    }
                    else continue;
                }

                if (r == null) r = new StringBuilder(s.Length + 30);
                r.Append(s, rCopied, i - rCopied);
                r.Append("_x");
                r.Append(((int)s[i]).ToString("X4", CultureInfo.InvariantCulture));
                r.Append("_");
                rCopied = i + 1;

            }

            if (LastWasWhiteSpace) NeedsPreserveWhitespace = true;
            if (r == null) return s;

            r.Append(s, rCopied, s.Length - rCopied);
            return r.ToString();
        }
        #endregion

        #region Xml
        private void SetWhitespacePreserve()
        {
            xlWriter.WriteAttributeString("xml", "space", null, "preserve");
        }

        internal void WriteStartDocument(string localName, bool IncludeRelationshipNs)
        {
            xlWriter.WriteStartDocument(true);
            xlWriter.WriteStartElement(localName, MainNamespace);
            if (IncludeRelationshipNs) xlWriter.WriteAttributeString("xmlns", "r", null, RelationshipNamespace);
        }

        internal void WriteStartDocument(string localName, string prefix, string Namespace)
        {
            xlWriter.WriteStartDocument(true);
            xlWriter.WriteStartElement(prefix, localName, Namespace);
        }

        internal void WriteEndDocument()
        {
            xlWriter.WriteEndElement();
            xlWriter.WriteEndDocument();
        }

        internal void WriteStartElement(string localName)
        {
            WriteStartElement(localName, true);
        }

        internal void WriteStartElement(string localName, bool SkipIfEmpty)
        {
            WriteStartElement(localName, DefaultNamespacePrefix, SkipIfEmpty);
        }

        internal void WriteStartElement(string localName, string prefix, bool SkipIfEmpty)
        {
            PendingTags.Add(new TTagDef(localName, prefix));

            if (!SkipIfEmpty)
            {
                ActuallyWriteStartElement();
            }
        }

        private void ActuallyWriteStartElement()
        {
            for (int i = 0; i < PendingTags.Count; i++)
            {
                if (PendingTags[i] != null)
                {
                    if (PendingTags[i].DefaultNamespacePrefix != null) xlWriter.WriteStartElement(PendingTags[i].DefaultNamespacePrefix, PendingTags[i].Tag, null);
                    else xlWriter.WriteStartElement(PendingTags[i].Tag);

                    PendingTags[i] = null;
                }
            }
        }

        internal void WriteEndElement()
        {
            TTagDef ptag = PendingTags[PendingTags.Count - 1];
            PendingTags.RemoveAt(PendingTags.Count - 1);
            if (ptag == null)
            {
                xlWriter.WriteEndElement();
            }
        }

        private void CheckPendingTag()
        {
            if (PendingTags.Count == 0 || PendingTags[PendingTags.Count - 1] == null) return;
            ActuallyWriteStartElement();
        }

        internal void WriteElement(string ElementName, string ElementValue)
        {
            WriteElement(ElementName, ElementValue, true);
        }

        internal void WriteElement(string ElementName, string ElementValue, bool AddPreserveWhiteSpace)
        {
            CheckPendingTag();
            bool PreserveWhiteSpace;
            string s = EscapeString(ElementValue, out PreserveWhiteSpace);
            if (PreserveWhiteSpace && AddPreserveWhiteSpace)
            {
                xlWriter.WriteStartElement(DefaultNamespacePrefix, ElementName, null);
                SetWhitespacePreserve();
                xlWriter.WriteString(s);
                xlWriter.WriteEndElement();
            }
            else
            {
                xlWriter.WriteElementString(DefaultNamespacePrefix, ElementName, null, s);
            }
        }

        internal void WriteString(string ElementValue)
        {
            CheckPendingTag();
            bool PreserveWhiteSpace;
            string s = EscapeString(ElementValue, out PreserveWhiteSpace);
            if (PreserveWhiteSpace)
            {
                SetWhitespacePreserve();
            }
            xlWriter.WriteString(s);
        }


        internal void WriteElement(string ElementName, uint ElementValue)
        {
            CheckPendingTag();
            xlWriter.WriteElementString(DefaultNamespacePrefix, ElementName, null, ElementValue.ToString(CultureInfo.InvariantCulture));
        }

        internal void WriteElement(string ElementName, bool ElementValue)
        {
            CheckPendingTag();
            string value = ElementValue ? "1" : "0";
            xlWriter.WriteElementString(DefaultNamespacePrefix, ElementName, null, value);
        }

        internal void WriteElement(string ElementName, double ElementValue)
        {
            CheckPendingTag();
            xlWriter.WriteElementString(DefaultNamespacePrefix, ElementName, null, ElementValue.ToString(CultureInfo.InvariantCulture));
        }

        internal void WriteRichText(TExcelString XS, IFlexCelFontList FontList)
        {
            TXlsxRichStringWriter.WriteRichOrPlainText(this, FontList, XS);
        }

        internal void WriteRichText(TRichString XS, IFlexCelFontList FontList)
        {
            TXlsxRichStringWriter.WriteRichOrPlainText(this, FontList, XS);
        }

        internal void WriteAtt(string name, bool value, bool defaultvalue)
        {
            if (value == defaultvalue) return;
            WriteAtt(name, value);
        }

        internal void WriteAtt(string name, bool value)
        {
            string svalue = value ? "1" : "0";
            CheckPendingTag();
            xlWriter.WriteAttributeString(name, svalue);
        }

        internal void WriteAtt(string name, int value)
        {
            string svalue = Convert.ToString(value, CultureInfo.InvariantCulture);
            CheckPendingTag();
            xlWriter.WriteAttributeString(name, svalue);
        }
        
        internal void WriteAtt(string name, int value, int defaultValue)
        {
            if (value != defaultValue) WriteAtt(name, value);
        }

        internal void WriteAtt(string name, long value)
        {
            string svalue = Convert.ToString(value, CultureInfo.InvariantCulture);
            CheckPendingTag();
            xlWriter.WriteAttributeString(name, svalue);
        }

        internal void WriteAtt(string name, long value, long defaultValue)
        {
            if (value != defaultValue) WriteAtt(name, value);
        }

        internal void WriteAtt(string name, double value)
        {
            string svalue = Convert.ToString(value, CultureInfo.InvariantCulture);
            CheckPendingTag();
            xlWriter.WriteAttributeString(name, svalue);
        }

        internal void WriteAtt(string name, string value)
        {
            if (string.IsNullOrEmpty(value)) return;
            bool nws; //This is not needed here.
            string svalue = EscapeString(value, out nws);
            CheckPendingTag();
            xlWriter.WriteAttributeString(name, svalue);
        }

        internal void WriteAttRaw(string attnamespace, string name, string value)
        {
            CheckPendingTag();
            xlWriter.WriteAttributeString(name, attnamespace, value);
        }

        internal void WriteAtt(string name, string value, bool SkipIfEmpty)
        {
            if (!SkipIfEmpty && string.IsNullOrEmpty(value))
            {
                WriteEmptyAtt(name);
                return;
            }

            WriteAtt(name, value);
        }

        internal void WriteEmptyAtt(string name)
        {
            CheckPendingTag();
            xlWriter.WriteAttributeString(name, string.Empty);
        }

        internal void WriteAtt(string name, string ns, string value)
        {
            if (string.IsNullOrEmpty(value)) return;
            bool nws; //This is not needed here.
            string svalue = EscapeString(value, out nws);
            CheckPendingTag();
            xlWriter.WriteAttributeString(name, ns, svalue);
        }

        internal void WriteAtt(string prefix, string localname, string ns, string value)
        {
            if (string.IsNullOrEmpty(value)) return;
            bool nws; //This is not needed here.
            string svalue = EscapeString(value, out nws);
            CheckPendingTag();
            xlWriter.WriteAttributeString(prefix, localname, ns, svalue);
        }

        internal void WriteAttHex(string name, long value, int pad)
        {
            string svalue = Convert.ToString(value, 16).ToUpper(CultureInfo.InvariantCulture);
            if (svalue.Length < pad) svalue = svalue.PadLeft(pad, '0');
            
            CheckPendingTag();
            xlWriter.WriteAttributeString(name, svalue);
        }

        internal void WriteAttPercent(string name, double value)
        {
            string svalue = Convert.ToString((int)Math.Round(value * 100000), CultureInfo.InvariantCulture);
            CheckPendingTag();
            xlWriter.WriteAttributeString(name, svalue);
        }

        internal void WriteAttAsAngle(string name, double value, double defaultValue)
        {
            if (value != defaultValue)
            {
                WriteAttAsAngle(name, value);
            }
        }

        internal void WriteAttAsAngle(string name, double value)
        {
            WriteAtt(name, (int)Math.Round(value * 60000.0));
        }

        internal void WriteAttAsAddress(string name, TCellAddress CellAddress)
        {
            if (CellAddress == null) return;
            WriteAtt(name, CellAddress.CellRef);
        }

        internal void WriteAttAsRange(string name, TXlsCellRange r, bool RIsOneBased)
        {
            if (r == null) return;
            WriteAtt(name, GetOneWriteRangeAtt(r, RIsOneBased));
        }

        private string GetOneWriteRangeAtt(TXlsCellRange r, bool RIsOneBased)
        {
            int offs = RIsOneBased ? 0 : 1;
            TCellAddress a1 = new TCellAddress(r.Top + offs, r.Left + offs);
            TCellAddress a2 = new TCellAddress(r.Bottom + offs, r.Right + offs);
            string a1CellRef = a1.CellRef;
            string a2CellRef = a2.CellRef;
            if (a1CellRef == a2CellRef) return a1CellRef;
            return a1CellRef + TFormulaMessages.TokenString(TFormulaToken.fmRangeSep) + a2CellRef;
        }

        internal void WriteAttAsSeriesOfRanges(string name, TXlsCellRange[] SelectedRange, bool RIsOneBased)
        {
            if (SelectedRange == null) return;
            StringBuilder sb = new StringBuilder();

            foreach (TXlsCellRange r in SelectedRange)
            {
                if (r != null)
                {
                    if (sb.Length > 0) sb.Append(" ");
                    sb.Append(GetOneWriteRangeAtt(r, RIsOneBased));
                }
            }

            WriteAtt(name, sb.ToString());
        }

        internal static string ConvertFromDrawingCoord(TDrawingCoordinate Coord)
        {
            return Coord.Emu.ToString(CultureInfo.InvariantCulture);
        }


        internal void WriteRelationship(string name, int relId)
        {
            string svalue = FlexCelRid + Convert.ToString(relId, CultureInfo.InvariantCulture);
            CheckPendingTag();
            xlWriter.WriteAttributeString(name, RelationshipNamespace, svalue);
        }

        internal void WriteRaw(string xml)
        {
            CheckPendingTag();
            xlWriter.WriteRaw(xml);
        }

        internal void WriteFutureStorage(TFutureStorage futureList)
        {
            if (futureList == null) return;
            for (int i = 0; i < futureList.Count; i++)
            {
                xlWriter.WriteRaw(futureList[i].Xml);                
            }
        }

        #endregion

        #region IDisposable Members

        public void Dispose()
        {
            if (xlWriter != null) xlWriter.Close();
            if (xlPackage != null) xlPackage.Close();
            GC.SuppressFinalize(this);
        }

        #endregion


    }

    class TTagDef
    {
        internal string DefaultNamespacePrefix;
        internal string Tag;

        internal TTagDef(string aTag, string aDefaultNamespace)
        {
            DefaultNamespacePrefix = aDefaultNamespace;
            Tag = aTag;
        }
    }
}
