using System;
using System.Collections.Generic;
using System.Text;
using FlexCel.Core;
using System.Globalization;
using System.Reflection;
using System.IO.Packaging;
using System.IO;

namespace FlexCel.XlsAdapter
{
    class TXlsxRecordWriter
    {
        #region Variables
        private TOpenXmlWriter DataStream;
        private ExcelFile xls;
        private TWorkbook Workbook;
        private TFileProps FileProps;

        private TXlsxDrawingWriter DrawingWriter;
        #endregion

        internal TXlsxRecordWriter(TOpenXmlWriter aDataStream, ExcelFile axls, TWorkbook aWorkbook, TFileProps aFileProps)
        {
            DataStream = aDataStream;
            xls = axls;
            Workbook = aWorkbook;
            DrawingWriter = new TXlsxDrawingWriter(aDataStream, axls, aWorkbook);
            FileProps = aFileProps;
        }

        internal TWorkbookGlobals Globals
        {
            get
            {
                return Workbook.Globals;
            }
        }

        internal void Save(TProtection Protection, byte[] OtherXlsxParts, bool MacroEnabled)
        {
            Workbook.PrepareToSave();
            Globals.SST.FixRefs();

            WriteWorkbook(MacroEnabled);
            WriteCustomXML();
            WriteSST();
            DrawingWriter.WriteTheme(false);
            WritePivotCacheParts();
            WriteStyles();
            WriteSheets();

            WriteExternalLinks();
            WriteConnections();

            WriteMacros(MacroEnabled);
            WriteCoreFileProperties();
            WriteCustomFileProperties();
        }


        #region Workbook
        internal void WriteWorkbook(bool MacroEnabled)
        {
            string ContentType = MacroEnabled? TOpenXmlWriter.WorkbookMacroEnabledContentType : TOpenXmlWriter.WorkbookContentType;
            DataStream.CreatePart(TOpenXmlWriter.WorkbookURI, ContentType);
            DataStream.CreateRelationshipFromUri(null, TOpenXmlWriter.documentRelationshipType, TOpenXmlManager.RelIdWorkbook);
             
            DataStream.WriteStartDocument("workbook", true);
            WriteActualWorkbook();
            DataStream.WriteEndDocument();
        }

        private void WriteActualWorkbook()
        {
            WriteFileVersion();
            WriteFileSharing();
            WriteWorkbookPr();
            WriteWorkbookProtection();
            WriteBookViews();
            WriteBoundSheets();
            WriteFunctionGroups();
            WriteExternalReferences();
            WriteDefinedNames();
            WriteCalcPr();
            WriteOleSize();
            WriteCustomWorkbookViews();
            WritePivotCaches();
            WriteSmartTagPr();
            WriteSmartTagTypes();
            WriteWebPublishing();
            WriteFileRecoveryPr();
            WriteWebPublishObjects();
            WriteWorkbook_ExtLst();
        }

        private void WriteFileVersion()
        {
            DataStream.WriteStartElement("fileVersion");
            DataStream.WriteAtt("appName", "xl"); //We can't write anything here. If not xl, Excel (2007, not 2010) will change the font sizes in legacy drawing. We can't fix one without breaking the other, so we need to say we are "xl" here.
            DataStream.WriteAtt("rupBuild", "4506"); //If we are Excel, we need a valid rupbuild.
            //DataStream.WriteAtt("rupBuild", Assembly.GetExecutingAssembly().GetName().Version.ToString());

            string fVer = "4";
            /*This is we wanted to identify as Excel 2010. We don't want to yet.
             * switch (Globals.sBOF.BiffFileFormat())
            {
                case TExcelFileFormat.v2010:
                    fVer = "5"; //5 in xlsx, 6 in biff for Excel 2010
                    break;
            }*/
            DataStream.WriteAtt("lastEdited", fVer); //If this or the next is omitted, and appname is xl, Excel 2010 will crash.
            DataStream.WriteAtt("lowestEdited", fVer);

            DataStream.WriteAtt("codeName", Globals.CodeName07);
            DataStream.WriteEndElement();
        }

        private void WriteFileSharing()
        {
            if (Globals.FileEncryption.FileSharing == null) return;
            DataStream.WriteStartElement("fileSharing");
            DataStream.WriteAtt("readOnlyRecommended", Globals.FileEncryption.FileSharing.RecommendReadOnly, false);
            string User = Globals.FileEncryption.FileSharing.User;
            if (Workbook.Globals.FileEncryption.WriteProt != null)
            {
                DataStream.WriteAttHex("reservationPassword", Globals.FileEncryption.FileSharing.HashedPass, 4);
                if (String.IsNullOrEmpty(User)) User = " "; //if we have a pass, we need to have an user, or we will crash Excel.
            }
            DataStream.WriteAtt("userName", User);
            DataStream.WriteEndElement();
        }

        private void WriteWorkbookPr()
        {
            DataStream.WriteStartElement("workbookPr", false);
            //if (DataStream.GetAttributeAsBool("allowRefreshQuery", false)) Globals;
            DataStream.WriteAtt("autoCompressPictures", Globals.AutoCompressPictures, true);
            DataStream.WriteAtt("backupFile", Globals.Backup, false);
            DataStream.WriteAtt("checkCompatibility", Globals.CheckCompatibility, false);
            DataStream.WriteAtt("codeName", Globals.CodeName);
            DataStream.WriteAtt("date1904", Globals.Dates1904, false);

            DataStream.WriteAtt("defaultThemeVersion", Globals.Theme.ThemeVersion, 0);
            if (Globals.BookExt != null)
            {
                DataStream.WriteAtt("filterPrivacy", Globals.BookExt.FilterPrivacy, false);
                DataStream.WriteAtt("hidePivotFieldList", Globals.BookExt.HidePivotList, false);
                DataStream.WriteAtt("promptedSolutions", Globals.BookExt.BuggedUserAboutSolution, false);
                DataStream.WriteAtt("publishItems", Globals.BookExt.PublishedBookItems, false);
            }
            DataStream.WriteAtt("refreshAllConnections", Globals.RefreshAll, false);

            DataStream.WriteAtt("saveExternalLinkValues", Globals.SaveExternalLinkValues, true);
            DataStream.WriteAtt("showBorderUnselectedTables", !Globals.HideBorderUnselLists, true);

            if (Globals.BookExt != null)
            {
                DataStream.WriteAtt("showInkAnnotation", Globals.BookExt.ShowInkAnnotation, true);
            }

            switch (Globals.HideObj)
            {
                case THideObj.ShowPlaceholder:
                    DataStream.WriteAtt("showObjects", "placeholders");
                    break;
                case THideObj.HideAll:
                    DataStream.WriteAtt("showObjects", "none");
                    break;
            }

            if (Globals.BookExt != null)
            {
                DataStream.WriteAtt("showPivotChartFilter", Globals.BookExt.ShowPivotChartFilter, false);
            }

            switch (Globals.UpdateLinks)
            {
                case TUpdateLinkOption.SilentlyUpdate: DataStream.WriteAtt("updateLinks", "always"); break;
                case TUpdateLinkOption.DontUpdate: DataStream.WriteAtt("updateLinks", "never"); break;
            }

            DataStream.WriteEndElement();
        }

        private void WriteWorkbookProtection()
        {
            if (Globals.WorkbookProtection.Protect == null &&
                Globals.WorkbookProtection.WindowProtect == null &&
                Globals.WorkbookProtection.Prot4Rev == null &&
                Globals.WorkbookProtection.Password == null)
                return;

            DataStream.WriteStartElement("workbookProtection");

            if (Globals.WorkbookProtection.Password != null && Globals.WorkbookProtection.Password.GetHash() != 0)
            {
                DataStream.WriteAttHex("workbookPassword", Globals.WorkbookProtection.Password.GetHash(), 0);
            }

            DataStream.WriteAtt("lockStructure", Globals.WorkbookProtection.Protect == null ? false : Globals.WorkbookProtection.Protect.Protected, false);
            DataStream.WriteAtt("lockWindows", Globals.WorkbookProtection.WindowProtect == null ? false : Globals.WorkbookProtection.WindowProtect.Protected, false);
            DataStream.WriteAtt("lockRevision", Globals.WorkbookProtection.Prot4Rev == null ? false : Globals.WorkbookProtection.Prot4Rev.Protected, false);

            DataStream.WriteEndElement();
        }

        private void WriteBookViews()
        {
            DataStream.WriteStartElement("bookViews");
            foreach (TWindow1Record w1 in Globals.Window1)
            {
                if (w1 != null) WriteBookView(w1);
            }
            DataStream.WriteEndElement();
        }

        private void WriteBookView(TWindow1Record w1)
        {
            DataStream.WriteStartElement("workbookView");

            if (w1.ActiveSheet > 0) DataStream.WriteAtt("activeTab", w1.ActiveSheet);
            if (w1.FirstSheetVisible > 0) DataStream.WriteAtt("firstSheet", w1.FirstSheetVisible);

            DataStream.WriteAtt("xWindow", w1.xWin);
            DataStream.WriteAtt("yWindow", w1.yWin);
            DataStream.WriteAtt("windowWidth", w1.dxWin);
            DataStream.WriteAtt("windowHeight", w1.dyWin);

            if (w1.TabRatio != 600) DataStream.WriteAtt("tabRatio", w1.TabRatio);

            DataStream.WriteAtt("minimized", (w1.Options & 0x02) != 0, false);

            DataStream.WriteAtt("showHorizontalScroll", (w1.Options & 0x08) != 0, true);
            DataStream.WriteAtt("showVerticalScroll", (w1.Options & 0x10) != 0, true);
            DataStream.WriteAtt("showSheetTabs", (w1.Options & 0x20) != 0, true);
            DataStream.WriteAtt("autoFilterDateGrouping", (w1.Options & 0x40) == 0, true);

            if ((w1.Options & 0x04) != 0) DataStream.WriteAtt("visibility", "veryHidden");
            else if ((w1.Options & 0x01) != 0) DataStream.WriteAtt("visibility", "hidden");

            DataStream.WriteFutureStorage(w1.FutureStorage);
            DataStream.WriteEndElement();
        }

        private void WriteBoundSheets()
        {
            DataStream.WriteStartElement("sheets");
            for (int i = 0; i < Globals.SheetCount; i++)
            {
                WriteBoundSheet(Globals.BoundSheets.BoundSheets[i], Globals.BoundSheets.GetTabId(i), i);
            }
            DataStream.WriteEndElement();
        }

        private void WriteBoundSheet(TBoundSheetRecord sheet, int sheetID, int sheetPos)
        {
            DataStream.WriteStartElement("sheet", false);
            DataStream.WriteAtt("name", sheet.SheetName);
            DataStream.WriteAtt("sheetId", sheetID);
            DataStream.WriteAtt("id", TOpenXmlManager.RelationshipNamespace, TOpenXmlWriter.GetRId(sheetPos + 1));

            if ((sheet.OptionFlags & 0x02) != 0) DataStream.WriteAtt("state", "veryHidden");
            else if ((sheet.OptionFlags & 0x01) != 0) DataStream.WriteAtt("state", "hidden");
            DataStream.WriteEndElement();
        }

        private void WriteFunctionGroups()
        {
        }

        private void WriteExternalReferences()
        {
            DataStream.WriteStartElement("externalReferences");
            for (int i = 0; i < Globals.References.Supbooks.Count; i++)
            {
                TSupBookRecord SupBook = Globals.References.Supbooks[i];
                if (SupBookNeedsSaving(SupBook))
                {
                    DataStream.WriteStartElement("externalReference");
                    DataStream.WriteRelationship("id", Globals.SheetCount + TOpenXmlWriter.RelIdExternalLinks + i);
                    DataStream.WriteEndElement();
                }
            }
            DataStream.WriteEndElement();
        }

        private void WriteDefinedNames()
        {
            DataStream.WriteStartElement("definedNames");

            for (int i = 0; i < Globals.Names.Count; i++)
            {
                string name = Globals.Names[i].Name;

                //The name filter_database *must* be present if there is an autofilter. If there isn't Excel keeps the name in xls and doesn't save it in xlsx, and so do we.
                if (name != null && name.Length == 1 && name[0] == (char)InternalNameRange.Filter_DataBase && !HasAutoFilter(Globals.Names[i])) continue;
                DataStream.WriteStartElement("definedName");
                WriteDefinedName(Globals.Names[i]);
                DataStream.WriteEndElement();
            }
            DataStream.WriteEndElement();
        }

        private bool HasAutoFilter(TNameRecord Name)
        {
            int FilterSheet = Name.RangeSheet;
            if (FilterSheet < 0 || FilterSheet >= Workbook.Sheets.Count) return false;
            return Workbook.Sheets[FilterSheet].HasAutoFilter();
        }

        private void WriteDefinedName(TNameRecord aName)
        {
            if (aName.IsAddin) return;
            string FormulaText = TFormulaConvertInternalToText.AsString(aName.Data, 0, 0, null, Globals, FlxConsts.Max_FormulaStringConstant, true);
            if (FormulaText == null || FormulaText.Length == 0) return;  //names used as addin holders will have formulatext = empty, so we won't write them in xlsx

            DataStream.WriteAtt("name", TXlsNamedRange.GetXlsxInternal(aName.Name));

            int opt = aName.OptionFlags;

            DataStream.WriteAtt("hidden", (opt & 0x01) != 0, false);
            DataStream.WriteAtt("function", (opt & 0x02) != 0, false);
            DataStream.WriteAtt("vbProcedure", (opt & 0x04) != 0, false);
            DataStream.WriteAtt("xlm", (opt & 0x08) != 0, false);

            int fid = (opt >> 6) & 0x3F;
            if (fid > 0)
            {
                DataStream.WriteAtt("functionGroupId", fid);
            }

            DataStream.WriteAtt("publishToServer", (opt & 0x2000) != 0, false);
            DataStream.WriteAtt("workbookParameter", (opt & 0x4000) != 0, false);

            DataStream.WriteAtt("shortcutKey", aName.KeyboardShortcut);
            if (aName.RangeSheet >= 0) DataStream.WriteAtt("localSheetId", aName.RangeSheet);
            DataStream.WriteAtt("customMenu", aName.FMenu);
            DataStream.WriteAtt("description", aName.FDescription);
            DataStream.WriteAtt("help", aName.FHelp);
            DataStream.WriteAtt("statusBar", aName.FStatusBar);
            DataStream.WriteAtt("comment", aName.Comment);

            DataStream.WriteString(FormulaText);

        }

        private void WriteCalcPr()
        {
            DataStream.WriteStartElement("calcPr");
            switch (Globals.CalcOptions.CalcMode)
            {
                case TSheetCalcMode.Manual:
                    DataStream.WriteAtt("calcMode", "manual");
                    break;
                case TSheetCalcMode.AutomaticExceptTables:
                    DataStream.WriteAtt("calcMode", "autoNoTable");
                    break;
            }

            DataStream.WriteAtt("calcOnSave", Globals.CalcOptions.SaveRecalc, true);

            DataStream.WriteAtt("concurrentCalc", Globals.MultithreadRecalc != 0, true);

            if (Globals.MultithreadRecalc > 0)
            {
                DataStream.WriteAtt("concurrentManualCount", Globals.MultithreadRecalc);
            }

            DataStream.WriteAtt("forceFullCalc", Globals.ForceFullRecalc, false);
            DataStream.WriteAtt("fullPrecision", !Globals.PrecisionAsDisplayed, true);

            DataStream.WriteAtt("iterate", Globals.CalcOptions.IterationEnabled, false);
            if (Globals.CalcOptions.CalcCount != 100) DataStream.WriteAtt("iterateCount", Globals.CalcOptions.CalcCount);
            if (Globals.CalcOptions.Delta != 0.001) DataStream.WriteAtt("iterateDelta", Globals.CalcOptions.Delta);

            if (!Globals.CalcOptions.A1RefMode)
            {
                DataStream.WriteAtt("refMode", "R1C1");
            }
            DataStream.WriteEndElement();
        }

        private void WriteOleSize()
        {
        }

        private void WriteCustomWorkbookViews()
        {
        }

        private void WritePivotCaches()
        {
            DataStream.WriteStartElement("pivotCaches");
            foreach (TXlsxPivotCache pc in Globals.XlsxPivotCache.List)
            {
                DataStream.WriteStartElement("pivotCache");
                DataStream.WriteAtt("cacheId", pc.CacheId);
                DataStream.WriteAtt("id", TOpenXmlManager.RelationshipNamespace,
                    TOpenXmlWriter.GetRId(Globals.SheetCount + Globals.References.Supbooks.Count + Globals.CustomXMLData.Count + TOpenXmlManager.RelIdPivotCaches));
                DataStream.WriteEndElement();
            }
            DataStream.WriteEndElement();
        }

        private void WriteSmartTagPr()
        {
        }

        private void WriteSmartTagTypes()
        {
        }

        private void WriteWebPublishing()
        {
        }

        private void WriteFileRecoveryPr()
        {
        }

        private void WriteWebPublishObjects()
        {
        }

        private void WriteWorkbook_ExtLst()
        {
            DataStream.WriteFutureStorage(Globals.FutureStorage);
        }
        #endregion

        #region Custom XML
        private void WriteCustomXML()
        {
            int WorkbookRelId = Globals.SheetCount + Globals.References.Supbooks.Count + TOpenXmlManager.RelIdCustomXML;
            foreach (TCustomXMLDataStorage ds in Globals.CustomXMLData)
            {
                DataStream.CreatePart(ds.PartUri, TOpenXmlManager.CustomXmlDataContentType);
                DataStream.CreateRelationshipFromUri(TOpenXmlManager.WorkbookURI, TOpenXmlManager.CustomXmlDataRelationshipType, WorkbookRelId);
                WorkbookRelId++;
                DataStream.WriteRaw(ds.XML);

                WriteCustomXMLProps(ds);
            }
        }

        private void WriteCustomXMLProps(TCustomXMLDataStorage ds)
        {
            int RelId = 1;
            foreach (TCustomXMLDataStorageProp dsp in ds.CustomXMLDataStorageProps)
            {
                DataStream.CreatePart(dsp.PartUri, TOpenXmlManager.CustomXmlDataPropsContentType);
                DataStream.CreateRelationshipFromUri(ds.PartUri, TOpenXmlManager.CustomXmlDataPropsRelationshipType, RelId);
                RelId++;
                DataStream.WriteRaw(dsp.XML);
            }
        }
        #endregion

        #region SST
        internal void WriteSST()
        {
            if (Globals.SST.Count == 0) return;
            DataStream.CreatePart(TOpenXmlWriter.SSTURI, TOpenXmlWriter.SSTContentType);
            DataStream.CreateRelationshipFromUri(TOpenXmlWriter.WorkbookURI, TOpenXmlWriter.sharedStringsRelationshipType, Globals.SheetCount + TOpenXmlManager.RelIdSST);

            bool Repeatable = false;
            UInt32 TotalRefs;
            IEnumerator<KeyValuePair<TSSTEntry, TSSTEntry>> myEnumerator;
            TSSTEntry[] SortedEntries;
            Globals.SST.PrepareToSave(Repeatable, out TotalRefs, out myEnumerator, out SortedEntries);

            DataStream.WriteStartDocument("sst", false);
            DataStream.WriteAtt("count", TotalRefs);
            DataStream.WriteAtt("uniqueCount", Globals.SST.Count);
            WriteActualSST(myEnumerator, SortedEntries);
            DataStream.WriteEndDocument();
        }

        private void WriteActualSST(IEnumerator<KeyValuePair<TSSTEntry, TSSTEntry>> myEnumerator, TSSTEntry[] SortedEntries)
        {
            if (SortedEntries != null)
            {
                foreach (TSSTEntry Se in SortedEntries)
                {
                    WriteSi(Se);
                }
            }
            else
            {
                myEnumerator.Reset();
                while (myEnumerator.MoveNext())
                    WriteSi(myEnumerator.Current.Key);
            }

            WriteSST_ExtLst();
        }

        private void WriteSi(TSSTEntry Se)
        {
            if (Se.Data.Length < 0 || Se.Data.Length > FlxConsts.Max_StringLenInCell) XlsMessages.ThrowException(XlsErr.ErrStringTooLong, Se.Data, FlxConsts.Max_StringLenInCell);
            DataStream.WriteStartElement("si", false);

            TXlsxRichStringWriter.WriteRichOrPlainText(DataStream, xls, Se);
            DataStream.WriteEndElement();
        }

        private void WriteSST_ExtLst()
        {
            DataStream.WriteFutureStorage(Globals.SST.FutureStorage);
        }
        #endregion

        #region Pivot Caches
        private void WritePivotCacheParts()
        {
            int WorkbookRelId = Globals.SheetCount + Globals.References.Supbooks.Count + Globals.CustomXMLData.Count + TOpenXmlManager.RelIdPivotCaches;
            int i = 0;
            foreach (TXlsxPivotCache pc in Globals.XlsxPivotCache.List)
            {
                pc.LastSavedUri = new Uri(TOpenXmlWriter.PivotCacheDefBaseURI + (i + 1).ToString(CultureInfo.InvariantCulture) + ".xml", UriKind.Relative);
                DataStream.CreatePart(pc.LastSavedUri, TOpenXmlManager.PivotCacheDefContentType);
               
                DataStream.CreateRelationshipFromUri(TOpenXmlManager.WorkbookURI, TOpenXmlManager.PivotCacheDefRelationshipType, WorkbookRelId + i);
                pc.SaveToXlsx(DataStream, "pivotCacheDefinition", true);
                i++;
            }            
        }

        #endregion

        #region Styles
        private void WriteStyles()
        {
            DataStream.CreatePart(TOpenXmlWriter.StylesURI, TOpenXmlWriter.StylesContentType);
            DataStream.CreateRelationshipFromUri(TOpenXmlWriter.WorkbookURI, TOpenXmlWriter.stylesRelationshipType, Globals.SheetCount + TOpenXmlManager.RelIdStyles);

            DataStream.WriteStartDocument("styleSheet", false);
            WriteActualStyles();
            DataStream.WriteEndDocument();
        }

        private void WriteActualStyles()
        {
            WriteNumFmts();
            WriteFonts();
            WriteFills();
            WriteBorders();
            WriteCellStyleXfs();
            WriteCellXfs();
            WriteCellStyles();
            WriteDxfs();
            WriteTableStyles();
            WriteColors();
            WriteStyles_ExtLst();
        }

        private void WriteNumFmts()
        {
            if (Globals.Formats.IsEmpty) return;
            DataStream.WriteStartElement("numFmts");

            int aCount = 0;
            for (int i = 0; i < Globals.Formats.Count; i++)
            {
                if (Globals.Formats[i] != null) aCount++;
            }

            DataStream.WriteAtt("count", aCount);
            for (int i = 0; i < Globals.Formats.Count; i++)
            {
                TFormatRecord format = Globals.Formats[i];
                if (format == null) continue;
                DataStream.WriteStartElement("numFmt", false);
                DataStream.WriteAtt("numFmtId", format.FormatId);
                DataStream.WriteAtt("formatCode", format.FormatDef);
                DataStream.WriteEndElement();
            }
            DataStream.WriteEndElement();
        }

        private void WriteFonts()
        {
            DataStream.WriteStartElement("fonts");
            DataStream.WriteAtt("count", Globals.Fonts.Count);
            for (int i = 0; i < Globals.Fonts.Count; i++)
            {
                DataStream.WriteStartElement("font", false);
                TXlsxFontReaderWriter.WriteFont(DataStream, Globals.Fonts[i].FlxFont(), true); //no font4 issues here.
                DataStream.WriteEndElement();
            }
            DataStream.WriteEndElement();
        }

        private void WriteFills()
        {
            if (Globals.Patterns.Count == 0) return;
            DataStream.WriteStartElement("fills");
            DataStream.WriteAtt("count", Globals.Patterns.Count);

            for (int i = 0; i < Globals.Patterns.Count; i++)
            {
                TFlxFillPattern pat = Globals.Patterns[i];
                DataStream.WriteStartElement("fill", false);
                TXlsxFillReaderWriter.SaveToXml(DataStream, pat);
                DataStream.WriteEndElement();
            }
            DataStream.WriteEndElement();
        }

        private void WriteBorders()
        {
            if (Globals.Borders.Count == 0) return;
            DataStream.WriteStartElement("borders");
            DataStream.WriteAtt("count", Globals.Borders.Count);

            for (int i = 0; i < Globals.Borders.Count; i++)
            {
                TFlxBorders border = Globals.Borders[i];
                DataStream.WriteStartElement("border", false);
                TXlsxBorderReaderWriter.SaveToXml(DataStream, border);
                DataStream.WriteEndElement();
            }

            DataStream.WriteEndElement();
        }

        private void WriteCellStyleXfs()
        {
            DataStream.WriteStartElement("cellStyleXfs");
            WriteXFS(Globals.StyleXF, true);
            DataStream.WriteEndElement();
        }

        private void WriteCellXfs()
        {
            DataStream.WriteStartElement("cellXfs");
            WriteXFS(Globals.CellXF, false);
            DataStream.WriteEndElement();
        }

        private void WriteXFS(TXFRecordList XFList, bool IsStyle)
        {
            DataStream.WriteAtt("count", XFList.Count);
            for (int i = 0; i < XFList.Count; i++)
            {
                WriteXF(XFList[i], IsStyle);
            }
        }

        private void WriteXF(TXFRecord XF, bool IsStyle)
        {
            DataStream.WriteStartElement("xf", false);

            WriteXFAttrs(XF, IsStyle, IsStyle);
            WriteAlignment(XF);
            WriteProtection(XF);
            DataStream.WriteFutureStorage(XF.FutureStorage);

            DataStream.WriteEndElement();
        }

        private void WriteXFAttrs(TXFRecord XF, bool IsStyle, bool DefaultApply)
        {
            WriteLinkedStyle(DataStream, "applyNumberFormat", XF.LinkedStyle.LinkedNumericFormat, DefaultApply);
            WriteLinkedStyle(DataStream, "applyFont", XF.LinkedStyle.LinkedFont, DefaultApply);
            WriteLinkedStyle(DataStream, "applyFill", XF.LinkedStyle.LinkedFill, DefaultApply);
            WriteLinkedStyle(DataStream, "applyBorder", XF.LinkedStyle.LinkedBorder, DefaultApply);
            WriteLinkedStyle(DataStream, "applyAlignment", XF.LinkedStyle.LinkedAlignment, DefaultApply);
            WriteLinkedStyle(DataStream, "applyProtection", XF.LinkedStyle.LinkedProtection, DefaultApply);

            DataStream.WriteAtt("fontId", FixFont4(XF.FontIndex));
            DataStream.WriteAtt("numFmtId", XF.FormatIndex);
            DataStream.WriteAtt("fillId", XF.FillPattern);
            DataStream.WriteAtt("borderId", XF.Borders);

            if (!IsStyle)
            {
                DataStream.WriteAtt("xfId", XF.Parent);
            }

            DataStream.WriteAtt("quotePrefix", XF.Lotus123Prefix, false);
            DataStream.WriteAtt("pivotButton", XF.SxButton, false);
        }

        private void WriteLinkedStyle(TOpenXmlWriter DataStream, string Tag, bool Value, bool DefaultApply)
        {
            //In normal cells, "Apply" means it is not linked to the style. In styles, "Apply" means it applies to the parent;
            if (!DefaultApply) Value = !Value;
            DataStream.WriteAtt(Tag, Value, DefaultApply);
        }

        private int FixFont4(int p)
        {
            if (p >= 4) p -= 1;
            if (p >= Globals.Fonts.Count) return 0; //Corrupt file.
            return p;
        }

        private void WriteAlignment(TXFRecord XF)
        {
            if (XF.HAlignment == THFlxAlignment.general && XF.VAlignment == TVFlxAlignment.bottom
                && XF.Rotation == 0 && XF.WrapText == false && XF.Indent == 0 && XF.JustLast == false
                && XF.ShrinkToFit == false && XF.IReadOrder == 0) return;

            DataStream.WriteStartElement("alignment");

            switch (XF.HAlignment)
            {
                case THFlxAlignment.left: DataStream.WriteAtt("horizontal", "left"); break;
                case THFlxAlignment.center: DataStream.WriteAtt("horizontal", "center"); break;
                case THFlxAlignment.right: DataStream.WriteAtt("horizontal", "right"); break;
                case THFlxAlignment.fill: DataStream.WriteAtt("horizontal", "fill"); break;
                case THFlxAlignment.justify: DataStream.WriteAtt("horizontal", "justify"); break;
                case THFlxAlignment.center_across_selection: DataStream.WriteAtt("horizontal", "centerContinuous"); break;
                case THFlxAlignment.distributed: DataStream.WriteAtt("horizontal", "distributed"); break;
            }

            switch (XF.VAlignment)
            {
                case TVFlxAlignment.top: DataStream.WriteAtt("vertical", "top"); break;
                case TVFlxAlignment.center: DataStream.WriteAtt("vertical", "center"); break;
                case TVFlxAlignment.justify: DataStream.WriteAtt("vertical", "justify"); break;
                case TVFlxAlignment.distributed: DataStream.WriteAtt("vertical", "distributed"); break;
            }

            if (XF.Rotation != 0) DataStream.WriteAtt("textRotation", XF.Rotation);
            DataStream.WriteAtt("wrapText", XF.WrapText, false);
            if (XF.Indent != 0) DataStream.WriteAtt("indent", XF.Indent);
            DataStream.WriteAtt("justifyLastLine", XF.JustLast, false);
            DataStream.WriteAtt("shrinkToFit", XF.ShrinkToFit, false);
            if (XF.IReadOrder != 0) DataStream.WriteAtt("readingOrder", XF.IReadOrder);

            DataStream.WriteEndElement();

        }

        private void WriteProtection(TXFRecord XF)
        {
            if (!XF.Hidden && XF.Locked) return;
            DataStream.WriteStartElement("protection");
            DataStream.WriteAtt("hidden", XF.Hidden, false);
            DataStream.WriteAtt("locked", XF.Locked, true);
            DataStream.WriteEndElement();
        }


        private void WriteCellStyles()
        {
            DataStream.WriteStartElement("cellStyles");
            DataStream.WriteAtt("count", Globals.Styles.Count);
            for (int i = 0; i < Globals.Styles.Count; i++)
            {
                WriteStyle((TStyleRecord)Globals.Styles[i]);
            }

            DataStream.WriteEndElement();
        }

        private void WriteStyle(TStyleRecord Style)
        {
            DataStream.WriteStartElement("cellStyle", false);
            DataStream.WriteAtt("name", Style.Name);
            DataStream.WriteAtt("xfId", Style.XF);
            if (Style.IsBuiltInStyle()) DataStream.WriteAtt("builtinId", Style.BuiltinId);
            if (Style.iLevel >= 0) DataStream.WriteAtt("iLevel", Style.iLevel);
            DataStream.WriteAtt("hidden", Style.Hidden, false);
            DataStream.WriteAtt("customBuiltin", Style.CustomBuiltin, false);

            DataStream.WriteFutureStorage(Style.FutureStorage);
            DataStream.WriteEndElement();
        }

        private void WriteDxfs()
        {
            //DataStream.WriteStartElement("dxfs");
            DataStream.WriteFutureStorage(Globals.DXF.Xlsx);
            //DataStream.WriteEndElement();
        }

        private void WriteTableStyles()
        {
            //DataStream.WriteStartElement("tableStyles");
            DataStream.WriteFutureStorage(Globals.TableStyles.Xlsx);
            //DataStream.WriteEndElement();
        }

        private void WriteColors()
        {
            DataStream.WriteStartElement("colors");
            WriteIndexedColors();
            WriteMRUColors();
            DataStream.WriteEndElement();
        }

        private void WriteIndexedColors()
        {
            if (Globals.Palette == null) return;
            if (Globals.Palette.IsStandard()) return;

            DataStream.WriteStartElement("indexedColors", false);

            for (int i = 0; i < 8; i++) //backward compat colors
            {
                WriteIndexedColor(Globals.Palette.GetRgbColor(i));
            }

            for (int i = 0; i < TPaletteRecord.Count; i++)
            {
                WriteIndexedColor(Globals.Palette.GetRgbColor(i));
            }
            DataStream.WriteEndElement();
        }

        private void WriteIndexedColor(System.Drawing.Color aColor)
        {
            DataStream.WriteStartElement("rgbColor", false);
            unchecked
            {
                DataStream.WriteAttHex("rgb", (UInt32)aColor.ToArgb(), 6);
            }
            DataStream.WriteEndElement();
        }

        private void WriteMRUColors()
        {
            DataStream.WriteFutureStorage(Globals.MruColors);
        }

        private void WriteStyles_ExtLst()
        {
            DataStream.WriteFutureStorage(Globals.StylesFutureStorage);
        }

        #endregion

        #region Worksheets
        private void WriteSheets()
        {
            int LegDrawingId = 1;
            for (int i = 0; i < Workbook.Sheets.Count; i++)
            {
                WriteSheet(Workbook.Sheets[i], i + 1, ref LegDrawingId);
            }
        }

        private void WriteSheet(TSheet aSheet, int SheetId, ref int LegDrawingId)
        {
            if (aSheet.SheetType == TSheetType.Other) return;
            TSheetRelationship SheetRel = DataStream.GetSheetRelationship(aSheet.SheetType, aSheet.International);
            DataStream.CreatePart(SheetRel.Uri, SheetRel.ContentType);
            DataStream.CreateRelationshipFromUri(TOpenXmlWriter.WorkbookURI, SheetRel.RelationshipType, SheetId);

            if (aSheet.SheetType == TSheetType.Chart || aSheet.SheetType == TSheetType.Other)
            {
                WriteSheet_ExtLst(aSheet);
                return;
            }

            TNoteAuthorList CommentAuthors = aSheet.Notes.GetAuthors();

            DataStream.WriteStartDocument(SheetRel.StartElement, true);
            WriteActualSheet(aSheet, SheetId, CommentAuthors);
            DataStream.WriteEndDocument();

            WriteRelatedParts(aSheet, SheetId, ref LegDrawingId, SheetRel, CommentAuthors);

        }

        private void WriteRelatedParts(TSheet aSheet, int SheetId, ref int LegDrawingId, TSheetRelationship SheetRel, TNoteAuthorList CommentAuthors)
        {
            int CurrentSheetLastRel = 0;

            CurrentSheetLastRel += aSheet.HLinks.Count;

            if (aSheet.PageSetup.Pls != null)
            {
                CurrentSheetLastRel++;
                Uri PrinterSettingsURI = new Uri(TOpenXmlManager.PrinterSettingsBaseURI + SheetId.ToString(CultureInfo.InvariantCulture) + ".bin", UriKind.Relative);
                DataStream.WritePart(PrinterSettingsURI, TOpenXmlWriter.PrinterSettingsContentType, 
                    GetPrinterSettings(aSheet.PageSetup.Pls), 2);
                DataStream.CreateRelationshipToUri(SheetRel.Uri, PrinterSettingsURI, TargetMode.Internal, TOpenXmlWriter.PrinterSettingsRelationshipType,
                    TOpenXmlWriter.GetRId(CurrentSheetLastRel));
            }


            if (HasDrawings(aSheet.Drawing))
            {
                CurrentSheetLastRel++;
                DrawingWriter.WriteDrawing(CurrentSheetLastRel, aSheet, SheetId, SheetRel.Uri);
            }

            if (CommentAuthors.Count > 0  || aSheet.Drawing.LegacyCount > 0)
            {
                WriteLegacyDrawingPart(false, aSheet, ref LegDrawingId, SheetRel, ref CurrentSheetLastRel);
            }

            if (aSheet.HeaderImages != null && aSheet.HeaderImages.DrawingCount > 0)
            {
                WriteLegacyDrawingPart(true, aSheet, ref LegDrawingId, SheetRel, ref CurrentSheetLastRel);
            }

            if (CommentAuthors.Count > 0)
            {
                WriteCommentPart(aSheet, SheetId, SheetRel, ref CurrentSheetLastRel, CommentAuthors);
            }

            //After here, CurrentSheetLastRel won't update in sync with refs in the sheet, as Pivot tables have no ref in the sheet.
            //So put all sheetrels that need sync above this comment.
            if (aSheet.XlsxPivotTables.List.Count > 0)
            {
                WritePivotTableParts(aSheet, SheetId, SheetRel, ref CurrentSheetLastRel);
            }
        }

        private bool HasDrawings(TDrawing Drawing)
        {
            if (Drawing == null) return false;
            for (int i = 0; i < Drawing.ObjectCount; i++)
            {
                TEscherOPTRecord opt = Drawing.GetOPT(i);
                if (TXlsxDrawingWriter.DrawingMustBeSaved(opt.GetObj().ObjType, opt)) return true;

            }
            return false;
        }

        private byte[] GetPrinterSettings(TPlsRecord PlsRecord)
        {
            if (PlsRecord.Continue == null) return PlsRecord.Data;
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(PlsRecord.Data, 0, PlsRecord.Data.Length);

                TContinueRecord cont = PlsRecord.Continue;
                while (cont != null)
                {
                    ms.Write(cont.Data, 0, cont.Data.Length);
                    cont = cont.Continue;
                }

                return ms.ToArray();
            }
        }

        private void WriteActualSheet(TSheet aSheet, int SheetId, TNoteAuthorList CommentAuthors)
        {
            int CurrentSheetLastRel = 0;
            TSheetType SheetType = aSheet.SheetType;
            
            WriteSheetPr(aSheet);
            if (SheetType != TSheetType.Dialog) WriteDimension(aSheet);
            WriteSheetViews(aSheet);
            WriteSheetFormatPr(aSheet);
            if (SheetType != TSheetType.Dialog)
            {
                WriteCols(aSheet);
                WriteSheetData(aSheet);
                WriteSheetCalcPr(aSheet);
            }
            WriteSheetProtection(aSheet);
            if (SheetType != TSheetType.Dialog)
            {
                WriteProtectedRanges(aSheet);
                WriteScenarios(aSheet);
                WriteAutoFilter(aSheet, SheetId);
                WriteSortState(aSheet);
                WriteDataConsolidate(aSheet);
            }
            
            WriteCustomSheetViews(aSheet);

            if (SheetType != TSheetType.Dialog)
            {
                WriteMergeCells(aSheet);
                WritePhoneticPr(aSheet);
                WriteConditionalFormatting(aSheet);
                WriteDataValidations(aSheet);
                WriteHyperlinks(aSheet, ref CurrentSheetLastRel);
            }

            WritePrintOptions(aSheet);
            WritePageMargins(aSheet);
            WritePageSetup(aSheet, ref CurrentSheetLastRel);
            WriteHeaderFooter(aSheet);

            if (SheetType != TSheetType.Dialog)
            {
                WriteRowBreaks(aSheet);
                WriteColBreaks(aSheet);
                WriteCustomProperties(aSheet);
                WriteCellWatches(aSheet);
                WriteIgnoredErrors(aSheet);
                WriteSmartTags(aSheet);
            }

            WriteDrawing(aSheet, ref CurrentSheetLastRel);
            WriteDrawingHF(aSheet, ref CurrentSheetLastRel);

            if (CommentAuthors.Count > 0 || aSheet.Drawing.LegacyCount > 0)
            {
                WriteLegacyDrawing(aSheet, ref CurrentSheetLastRel);
            }

            if (aSheet.HeaderImages != null && aSheet.HeaderImages.DrawingCount > 0)
            {
                WriteLegacyDrawingHF(aSheet, ref CurrentSheetLastRel);
            }

            if (SheetType != TSheetType.Dialog) WritePicture(aSheet);

            WriteOleObjects(aSheet);

            if (SheetType != TSheetType.Dialog)
            {
                WriteControls(aSheet);
                WriteWebPublishItems(aSheet);
                WriteTableParts(aSheet);
            }
            WriteSheet_ExtLst(aSheet);
        }

        private void WriteSheetPr(TSheet aSheet)
        {
            DataStream.WriteStartElement("sheetPr");
            WriteActualSheetPr(aSheet);
            DataStream.WriteEndElement();
        }

        private void WriteActualSheetPr(TSheet aSheet)
        {
            if (aSheet.SheetGlobals.Sync != null)
            {
                DataStream.WriteAtt("syncRef", new TCellAddress(aSheet.SheetGlobals.Sync.Row + 1, aSheet.SheetGlobals.Sync.Col + 1).CellRef);
            }

            DataStream.WriteAtt("syncHorizontal", aSheet.SheetGlobals.WsBool.SyncHoriz, false);
            DataStream.WriteAtt("syncVertical", aSheet.SheetGlobals.WsBool.SyncVert, false);

            DataStream.WriteAtt("transitionEntry", aSheet.SheetGlobals.WsBool.AltFormulaEntry, false);
            DataStream.WriteAtt("transitionEvaluation", aSheet.SheetGlobals.WsBool.AltExprEval, false);

            if (aSheet.SheetExt != null)
            {
                DataStream.WriteAtt("published", !aSheet.SheetExt.NotPublished, true);
                DataStream.WriteAtt("enableFormatConditionsCalculation", aSheet.SheetExt.CondFmtCalc, true);
            }

            if (!String.IsNullOrEmpty(aSheet.CodeName) && aSheet.SheetType != TSheetType.Dialog) DataStream.WriteAtt("codeName", aSheet.CodeName);

            DataStream.WriteAtt("filterMode", aSheet.SortAndFilter.FilterMode, false);

            if (aSheet.SheetExt != null && !aSheet.SheetExt.SheetColor.IsAutomatic)
            {
                DataStream.WriteStartElement("tabColor");
                TXlsxColorReaderWriter.WriteColor(DataStream, aSheet.SheetExt.SheetColor);
                DataStream.WriteEndElement();
            }


            DataStream.WriteStartElement("outlinePr");
            DataStream.WriteAtt("applyStyles", aSheet.SheetGlobals.WsBool.ApplyStyles, false);
            DataStream.WriteAtt("summaryBelow", aSheet.SheetGlobals.WsBool.RowSumsBelow, true);
            DataStream.WriteAtt("summaryRight", aSheet.SheetGlobals.WsBool.ColSumsRight, true);
            DataStream.WriteAtt("showOutlineSymbols", aSheet.SheetGlobals.WsBool.DspGuts, true);
            DataStream.WriteEndElement();


            DataStream.WriteStartElement("pageSetUpPr");
            DataStream.WriteAtt("autoPageBreaks", aSheet.SheetGlobals.WsBool.ShowAutoBreaks, true);
            DataStream.WriteAtt("fitToPage", aSheet.SheetGlobals.WsBool.FitToPage, false);
            DataStream.WriteEndElement();
        }

        private void WriteDimension(TSheet aSheet)
        {
            DataStream.WriteStartElement("dimension");
            TXlsCellRange UsedRange = aSheet.Cells.UsedRange();
            TCellAddress a1 = new TCellAddress(UsedRange.Top + 1, UsedRange.Left + 1);
            TCellAddress a2 = new TCellAddress(UsedRange.Bottom + 1, UsedRange.Right + 1);

            if (a2.Row < a1.Row || a2.Col < a1.Col || (a1.CellRef == a2.CellRef)) DataStream.WriteAtt("ref", a1.CellRef);
            else DataStream.WriteAtt("ref", a1.CellRef + TFormulaMessages.TokenString(TFormulaToken.fmRangeSep) + a2.CellRef);

            DataStream.WriteEndElement();
        }

        private void WriteSheetViews(TSheet aSheet)
        {
            DataStream.WriteStartElement("sheetViews");
            WriteSheetView(aSheet.Window);
            DataStream.WriteEndElement();
        }

        private void WriteSheetView(TWindow Window)
        {
            DataStream.WriteStartElement("sheetView");
            DataStream.WriteAtt("windowProtection", Globals.WorkbookProtection.WindowProtect == null ? false : Globals.WorkbookProtection.WindowProtect.Protected, false);
            TSheetOptions so = Window.Window2.Options;
            DataStream.WriteAtt("showFormulas", Window.Window2.ShowFormulaText, false);
            DataStream.WriteAtt("showGridLines", Window.Window2.ShowGridLines, true);
            DataStream.WriteAtt("showRowColHeaders", (so & TSheetOptions.ShowRowAndColumnHeaders) != 0, true);
            DataStream.WriteAtt("showZeros", (so & TSheetOptions.ZeroValues) != 0, true);
            DataStream.WriteAtt("rightToLeft", (so & TSheetOptions.RightToLeft) != 0, false);
            DataStream.WriteAtt("tabSelected", Window.Window2.Selected, false);
            if (Window.Plv != null)
            {
                DataStream.WriteAtt("showRuler", Window.Plv.ShowRuler, true);
            }

            DataStream.WriteAtt("showOutlineSymbols", (so & TSheetOptions.OutlineSymbols) != 0, true);
            DataStream.WriteAtt("defaultGridColor", (so & TSheetOptions.AutomaticGridLineColors) != 0, true);

            if (Window.Plv != null)
            {
                DataStream.WriteAtt("showWhiteSpace", Window.Plv.ShowWhiteSpace, true);
            }

            if ((so & TSheetOptions.PageBreakView) != 0)
            {
                DataStream.WriteAtt("view", "pageBreakPreview");
            }
            else
            {
                if (Window.Plv != null && Window.Plv.PageLayoutPreview)
                {
                    DataStream.WriteAtt("view", "pageLayout");
                }
            }
                

            DataStream.WriteAttAsAddress("topLeftCell", new TCellAddress(Window.Window2.FirstRow + 1, Window.Window2.FirstCol + 1));

            DataStream.WriteAtt("colorId", Window.Window2.GetGridLinesColor(xls).GetBiff8ColorIndex(xls, TAutomaticColor.DefaultForeground), 64);
            if (Window.Scl != null) DataStream.WriteAtt("zoomScale", Window.Scl.Zoom, 100);
            DataStream.WriteAtt("zoomScaleNormal", Window.Window2.ScaleInNormalView, 0);
            if (Window.Plv != null)
            {
                DataStream.WriteAtt("zoomScaleSheetLayoutView", Window.Plv.Zoom, 0);
            }
            DataStream.WriteAtt("zoomScalePageLayoutView", Window.Window2.ScaleInPageBreakPreview, 0);

            DataStream.WriteAtt("workbookViewId", 0);

            WritePane(Window.Window2, Window.Pane);
            for (int i = 0; i < 4; i++)
            {
                if (Window.Selection.MustSavePane(i, Window))
                {
                    WriteSelection(Window.Selection, (TPanePosition)i);
                }
            }
            //WritePivotSelection();
            WriteSheetView_ExtLst(Window);

            DataStream.WriteEndElement();
        }

        private void WritePane(TWindow2Record Window2, TPaneRecord Pane)
        {
            if (Pane == null) return;
            DataStream.WriteStartElement("pane");
                   DataStream.WriteAtt("xSplit", Pane.ColSplit,0);
                   DataStream.WriteAtt("ySplit", Pane.RowSplit,0);

            DataStream.WriteAttAsAddress("topLeftCell", new TCellAddress(Pane.FirstVisibleRow + 1, Pane.FirstVisibleCol + 1));
       
            string ActivePane = GetPanePos((TPanePosition)Pane.ActivePane);
            if (ActivePane != "topLeft") DataStream.WriteAtt("activePane", ActivePane);

            string Split = GetSplit(Window2);
            if (Split != "split") DataStream.WriteAtt("state", Split);
            DataStream.WriteEndElement();
        }

        private string GetSplit(TWindow2Record Window2)
        {
            if (Window2.IsFrozenButNoSplit) return "frozen";
            if (Window2.IsFrozen) return "frozenSplit";
            return "split";
        }

        private string GetPanePos(TPanePosition PanePos)
        {
            switch (PanePos)
            {
                case TPanePosition.LowerRight: return "bottomRight";
                case TPanePosition.UpperRight: return "topRight";
                case TPanePosition.LowerLeft: return "bottomLeft";
                default:
                    return "topLeft";
            }
        }

        private void WriteSelection(TSheetSelection Selection, TPanePosition PanePosition)
        {
            if (Selection == null) return;
            TXlsCellRange[] SelectedRange = Selection.GetSelection(PanePosition);
            if (SelectedRange == null) return;
            DataStream.WriteStartElement("selection");
            string PanePos = GetPanePos(PanePosition);
            if (PanePos != "topLeft") DataStream.WriteAtt("pane", PanePos);
            DataStream.WriteAttAsAddress("activeCell", Selection.GetActiveCellBase1(PanePosition));
            DataStream.WriteAtt("activeCellId", Selection.ActiveCellId(PanePosition), 0);
            DataStream.WriteAttAsSeriesOfRanges("sqref", SelectedRange, false);
            DataStream.WriteEndElement();
        }

        private void WriteSheetView_ExtLst(TWindow Window)
        {
            DataStream.WriteFutureStorage(Window.FutureStorage);
        }

        private void WriteSheetFormatPr(TSheet aSheet)
        {
            DataStream.WriteStartElement("sheetFormatPr");
            DataStream.WriteAtt("defaultColWidth", aSheet.Columns.DefColWidth / 256.0);


            DataStream.WriteAtt("defaultRowHeight", aSheet.DefRowHeight / 20.0);

            aSheet.FixGuts();
            if (aSheet.SheetGlobals.Guts.RowLevel > 0)
            {
                DataStream.WriteAtt("outlineLevelRow", aSheet.SheetGlobals.Guts.RowLevel - 1);
            }

            if (aSheet.SheetGlobals.Guts.ColLevel > 0)
            {
                DataStream.WriteAtt("outlineLevelCol", aSheet.SheetGlobals.Guts.ColLevel - 1);
            }

            int flags = aSheet.SheetGlobals.DefRowHeight.Flags;
            DataStream.WriteAtt("customHeight", (flags & 0x01) != 0, false);
            DataStream.WriteAtt("zeroHeight", (flags & 0x02) != 0, false);
            DataStream.WriteAtt("thickTop", (flags & 0x04) != 0, false);
            DataStream.WriteAtt("thickBottom", (flags & 0x08) != 0, false);

            DataStream.WriteEndElement();
        }

        private void WriteCols(TSheet aSheet)
        {
            DataStream.WriteStartElement("cols");

            int LastColumn = aSheet.Columns.ColCount;
            for (int i = 0; i < LastColumn; i++)
            {
                TColInfo ci = aSheet.Columns[i];
                if (ci == null) continue;
                int k = i + 1;
                while (k < LastColumn && ci.IsEqual(aSheet.Columns[k])) { k++; }

                WriteOneColumn(DataStream, ci, i, k);
                i = k - 1;
            }

            DataStream.WriteEndElement();
        }

        private void WriteOneColumn(TOpenXmlWriter DataStream, TColInfo ci, int i, int k)
        {
            DataStream.WriteStartElement("col");
            DataStream.WriteAtt("min", i + 1);
            DataStream.WriteAtt("max", k);

            int opt = ci.Options;

            DataStream.WriteAtt("width", ci.Width / 256.0);

            if (ci.XF != 0) DataStream.WriteAtt("style", ci.XF);


            DataStream.WriteAtt("hidden", (opt & 0x01) != 0, false);
            DataStream.WriteAtt("bestFit", (opt & 0x04) != 0, false);
            DataStream.WriteAtt("customWidth", (opt & 0x02) != 0, false);

            DataStream.WriteAtt("phonetic", (opt & 0x08) != 0, false);

            if (ci.GetColOutlineLevel() > 0) DataStream.WriteAtt("outlineLevel", ci.GetColOutlineLevel());
            DataStream.WriteAtt("collapsed", (opt & 0x1000) != 0, false);


            DataStream.WriteEndElement();
        }

        private void WriteSheetData(TSheet aSheet)
        {
            DataStream.WriteStartElement("sheetData", false);
            bool Dates1904 = Workbook.Globals.Dates1904;
            for (int i = 0; i < aSheet.Cells.CellList.Count; i++)
            {
                TXlsxCellWriter.WriteRow(DataStream, aSheet, aSheet.Cells, i, Dates1904);
            }
            DataStream.WriteEndElement();
        }

        private void WriteSheetCalcPr(TSheet aSheet)
        {
            DataStream.WriteStartElement("sheetCalcPr");
            DataStream.WriteEndElement();
        }

        private void WriteSheetProtection(TSheet aSheet)
        {
            TSheetProtectionOptions spo = aSheet.GetSheetProtectionOptions();
            if (!spo.Contents) return;

            DataStream.WriteStartElement("sheetProtection");

            if (aSheet.SheetProtection.Password != null && aSheet.SheetProtection.Password.GetHash() != 0)
            {
                DataStream.WriteAttHex("password", aSheet.SheetProtection.Password.GetHash(), 0);
            }


            DataStream.WriteAtt("sheet", spo.Contents, false);
            DataStream.WriteAtt("objects", spo.Objects, false);
            DataStream.WriteAtt("scenarios", spo.Scenarios, false);
            DataStream.WriteAtt("formatCells", !spo.CellFormatting, true);
            DataStream.WriteAtt("formatColumns", !spo.ColumnFormatting, true);
            DataStream.WriteAtt("formatRows", !spo.RowFormatting, true);
            DataStream.WriteAtt("insertColumns", !spo.InsertColumns, true);
            DataStream.WriteAtt("insertRows", !spo.InsertRows, true);
            DataStream.WriteAtt("insertHyperlinks", !spo.InsertHyperlinks, true);
            DataStream.WriteAtt("deleteColumns", !spo.DeleteColumns, true);
            DataStream.WriteAtt("deleteRows", !spo.DeleteRows, true);
            DataStream.WriteAtt("selectLockedCells", !spo.SelectLockedCells, false);
            DataStream.WriteAtt("sort", !spo.SortCellRange, true);
            DataStream.WriteAtt("autoFilter", !spo.EditAutoFilters, true);
            DataStream.WriteAtt("pivotTables", !spo.EditPivotTables, true);
            DataStream.WriteAtt("selectUnlockedCells", !spo.SelectUnlockedCells, false);
            DataStream.WriteEndElement();
        }

        private void WriteProtectedRanges(TSheet aSheet)
        {
            if (aSheet.SheetGlobals.ProtectedRanges.Count == 0) return;
            DataStream.WriteStartElement("protectedRanges");
            for (int i = 0; i < aSheet.SheetGlobals.ProtectedRanges.Count; i++)
            {
                WriteProtectedRange(aSheet.SheetGlobals.ProtectedRanges[i]);
            }
            DataStream.WriteEndElement();
        }

        private void WriteProtectedRange(TProtectedRange Range)
        {
            DataStream.WriteStartElement("protectedRange");
            DataStream.WriteAttAsSeriesOfRanges("sqref", Range.Ranges, false);
            DataStream.WriteAtt("name", Range.Name);
            DataStream.WriteAtt("password", Range.PasswordHash, true);
            DataStream.WriteEndElement();
        }

        private void WriteScenarios(TSheet aSheet)
        {
            DataStream.WriteStartElement("scenarios");
            DataStream.WriteEndElement();
        }

        private void WriteAutoFilter(TSheet aSheet, int SheetId)
        {
            TXlsCellRange cr = aSheet.GetAutoFilterRange(SheetId - 1);
            if (cr == null) return;
            DataStream.WriteStartElement("autoFilter");
            DataStream.WriteAttAsRange("ref", cr, false);

            WriteFilterColumn();
            WriteSortState();               
            WriteAutoFilter_ExtLst(aSheet.SortAndFilter.AutoFilter);
            DataStream.WriteEndElement();

        }

        private void WriteFilterColumn()
        {
            DataStream.WriteStartElement("filterColumn");
            DataStream.WriteEndElement();
        }

        private void WriteSortState()
        {
            DataStream.WriteStartElement("sortState");
            DataStream.WriteEndElement();
        }

        private void WriteAutoFilter_ExtLst(TAutoFilter AutoFilter)
        {
            if (AutoFilter != null) DataStream.WriteFutureStorage(AutoFilter.FutureStorage);
        }

        private void WriteSortState(TSheet aSheet)
        {
            DataStream.WriteStartElement("sortState");
            DataStream.WriteEndElement();
        }

        private void WriteDataConsolidate(TSheet aSheet)
        {
            DataStream.WriteStartElement("dataConsolidate");
            DataStream.WriteEndElement();
        }

        private void WriteCustomSheetViews(TSheet aSheet)
        {
            DataStream.WriteStartElement("customSheetViews");
            DataStream.WriteEndElement();
        }

        private void WriteMergeCells(TSheet aSheet)
        {
            if (aSheet.MergedCells.Count == 0) return;
            DataStream.WriteStartElement("mergeCells");
            for (int i = 0; i < aSheet.MergedCells.Count; i++)
            {
                TMergedCells Me = (TMergedCells)aSheet.MergedCells[i];
                for (int k = 0; k < Me.MergedCount(); k++)
                {
                    TXlsCellRange r = Me.MergedCell(k);
                    if (r != null)
                    {
                        DataStream.WriteStartElement("mergeCell");
                        DataStream.WriteAttAsRange("ref", r, false);
                        DataStream.WriteEndElement();
                    }
                }
            }
            DataStream.WriteEndElement();
        }

        private void WritePhoneticPr(TSheet aSheet)
        {
            DataStream.WriteStartElement("phoneticPr");
            DataStream.WriteEndElement();
        }

        private void WriteConditionalFormatting(TSheet aSheet)
        {
            DataStream.WriteStartElement("conditionalFormatting");
            DataStream.WriteEndElement();
        }

        private void WriteDataValidations(TSheet aSheet)
        {
            if (aSheet.DataValidation.Count == 0) return;

            DataStream.WriteStartElement("dataValidations");
            DataStream.WriteAtt("xWindow", aSheet.DataValidation.xWindow, 0);
            DataStream.WriteAtt("yWindow", aSheet.DataValidation.yWindow, 0);
            DataStream.WriteAtt("disablePrompts", aSheet.DataValidation.DisablePrompts, false);
            DataStream.WriteAtt("count", aSheet.DataValidation.Count);

            for (int i = 0; i < aSheet.DataValidation.Count; i++)
            {
                WriteDataValidation(aSheet.DataValidation.GetDataValidation(i, aSheet.Cells.CellList, true), aSheet.DataValidation.GetDataValidationRange(i, false));
            }
            DataStream.WriteEndElement();
        }

        private void WriteDataValidation(TDataValidationInfo dv, TXlsCellRange[] ranges)
        {
            DataStream.WriteStartElement("dataValidation", false);

            DataStream.WriteAtt("type", GetDataValidationType(dv.ValidationType));
            DataStream.WriteAtt("errorStyle", GetDataValidationIcon(dv.ErrorIcon));

            DataStream.WriteAtt("imeMode", GetDataValidationImeMode(dv.ImeMode));
            DataStream.WriteAtt("operator", GetDataValidationOperator(dv.Condition));

            DataStream.WriteAtt("allowBlank", dv.IgnoreEmptyCells, false);
            DataStream.WriteAtt("showDropDown", !dv.InCellDropDown, false);
            DataStream.WriteAtt("showInputMessage", dv.ShowInputBox, false);
            DataStream.WriteAtt("showErrorMessage", dv.ShowErrorBox, false);
            DataStream.WriteAtt("errorTitle", TruncStr(dv.ErrorBoxCaption, FlxConsts.Max_DvErrorTitleLen));
            DataStream.WriteAtt("error", TruncStr(dv.ErrorBoxText, FlxConsts.Max_DvErrorTextLen));
            DataStream.WriteAtt("promptTitle", TruncStr(dv.InputBoxCaption, FlxConsts.Max_DvInputTitleLen));
            DataStream.WriteAtt("prompt", TruncStr(dv.InputBoxText, FlxConsts.Max_DvInputTextLen));
            DataStream.WriteAttAsSeriesOfRanges("sqref", ranges, false);

            if (!String.IsNullOrEmpty(dv.FirstFormula))
            {
                DataStream.WriteElement("formula1", MakeDvExplicit(dv.ExplicitList, SupressStartFmla(dv.FirstFormula)));
            }
            if (!string.IsNullOrEmpty(dv.SecondFormula)) //docs say it is required, excel doesn't write it if not needed.
            {
                DataStream.WriteElement("formula2", SupressStartFmla(dv.SecondFormula));
            }

            DataStream.WriteEndElement();
        }

        private string TruncStr(string s, int MaxLen)
        {
            if (s == null || s.Length <= MaxLen) return s;
            return s.Substring(0, MaxLen);
        }

        private string MakeDvExplicit(bool IsExplicit, string fmla)
        {
            if (string.IsNullOrEmpty(fmla)) return fmla;
            if (!IsExplicit) return fmla;
            return fmla.Replace((char)0, ',');

        }

        private string SupressStartFmla(string fmla)
        {
            if (string.IsNullOrEmpty(fmla)) return fmla;
            if (fmla.StartsWith(TFormulaMessages.TokenString(TFormulaToken.fmStartFormula), false, CultureInfo.InvariantCulture))
            {
                return fmla.Substring(1);
            }

            return fmla;
        }

        private string GetDataValidationOperator(TDataValidationConditionType dvc)
        {
            switch (dvc)
            {
                case TDataValidationConditionType.Between: return null;
                case TDataValidationConditionType.NotBetween: return "notBetween";
                case TDataValidationConditionType.EqualTo: return "equal";
                case TDataValidationConditionType.NotEqualTo: return "notEqual";
                case TDataValidationConditionType.LessThan: return "lessThan";
                case TDataValidationConditionType.LessThanOrEqualTo: return "lessThanOrEqual";
                case TDataValidationConditionType.GreaterThan: return "greaterThan";
                case TDataValidationConditionType.GreaterThanOrEqualTo: return "greaterThanOrEqual";
            }
            return null;
        }

        private string GetDataValidationImeMode(TDataValidationImeMode vim)
        {
            switch (vim)
            {
                case TDataValidationImeMode.NoControl: return null;
                case TDataValidationImeMode.Off: return "off";
                case TDataValidationImeMode.On: return "on";
                case TDataValidationImeMode.Disabled: return "disabled";
                case TDataValidationImeMode.Hiragana: return "hiragana";
                case TDataValidationImeMode.FullKatakana: return "fullKatakana";
                case TDataValidationImeMode.HalfKatakana: return "halfKatakana";
                case TDataValidationImeMode.FullAlpha: return "fullAlpha";
                case TDataValidationImeMode.HalfAlpha: return "halfAlpha";
                case TDataValidationImeMode.FullHangul: return "fullHangul";
                case TDataValidationImeMode.HalfHangul: return "halfHangul";
            }

            return null;
        }

        private string GetDataValidationIcon(TDataValidationIcon dvi)
        {
            switch (dvi)
            {
                case TDataValidationIcon.Stop: return null;
                case TDataValidationIcon.Warning: return "warning";
                case TDataValidationIcon.Information: return "information";
            }
            return null;
        }

        private string GetDataValidationType(TDataValidationDataType dvt)
        {
            switch (dvt)
            {
                case TDataValidationDataType.AnyValue: return null;
                case TDataValidationDataType.WholeNumber: return "whole";
                case TDataValidationDataType.Decimal: return "decimal";
                case TDataValidationDataType.List: return "list";
                case TDataValidationDataType.Date: return "date";
                case TDataValidationDataType.Time: return "time";
                case TDataValidationDataType.TextLenght: return "textLength";
                case TDataValidationDataType.Custom: return "custom";
            }

            return null;
        }

        private void WriteHyperlinks(TSheet aSheet, ref int CurrentSheetLastRel)
        {
            DataStream.WriteStartElement("hyperlinks");
            for (int i = 0; i < aSheet.HLinks.Count; i++)
            {
                WriteHyperlink(aSheet.HLinks[i], ref CurrentSheetLastRel);
            }
            DataStream.WriteEndElement();
        }

        private void WriteHyperlink(THLinkRecord HLink, ref int CurrentSheetLastRel)
        {
            DataStream.WriteStartElement("hyperlink");
            DataStream.WriteAttAsRange("ref", HLink.GetCellRange(), false);

            CurrentSheetLastRel++;
            if (!string.IsNullOrEmpty(HLink.Text))
            {
                string HLinkText = HLink.Text;
                if (HLink.LinkType == THyperLinkType.UNC) HLinkText =  "file:///" + HLinkText;

                Uri TargetUri;
                try
                {
                    TargetUri = new Uri(HLinkText, UriKind.RelativeOrAbsolute);
                }
                catch (UriFormatException ex)
                {
                    if (FlexCelTrace.HasListeners) FlexCelTrace.Write(new TMalformedUrlError(ex.Message, HLinkText));
                    TargetUri = null; //Sadly we can't write something like http://www.test,com with this shitty System.IO.Packaging implementation.
                }

                if (TargetUri != null)
                {
                    DataStream.CreateRelationshipToUri(TargetUri, TargetMode.External, TOpenXmlManager.HyperlinkRelationshipType, TOpenXmlWriter.FlexCelRid + CurrentSheetLastRel.ToString(CultureInfo.InvariantCulture));
                    DataStream.WriteRelationship("id", CurrentSheetLastRel);
                }
            }
            

            DataStream.WriteAtt("location", HLink.TextMark);
            if (HLink.Hint != null)
            {
                DataStream.WriteAtt("tooltip", HLink.Hint.Text);
            }
            DataStream.WriteAtt("display", HLink.Description);
            DataStream.WriteEndElement();
        }

        private void WritePrintOptions(TSheet WorkSheet)
        {
            DataStream.WriteStartElement("printOptions");
            DataStream.WriteAtt("horizontalCentered", WorkSheet.PageSetup.HCenter, false);
            DataStream.WriteAtt("verticalCentered", WorkSheet.PageSetup.VCenter, false);
            DataStream.WriteAtt("headings", WorkSheet.SheetGlobals.PrintHeaders, false);
            DataStream.WriteAtt("gridLines", WorkSheet.SheetGlobals.PrintGridLines, false);
            DataStream.WriteAtt("gridLinesSet", WorkSheet.SheetGlobals.GridSet, true);

            DataStream.WriteEndElement();
        }

        private void WritePageMargins(TSheet WorkSheet)
        {

            //this are required, but if we fill them with "-1" Excel will complain
            if (WorkSheet.PageSetup.LeftMargin >= 0 && WorkSheet.PageSetup.RightMargin >= 0
                && WorkSheet.PageSetup.TopMargin >= 0 && WorkSheet.PageSetup.BottomMargin >= 0
                && WorkSheet.PageSetup.Setup.HeaderMargin >= 0 && WorkSheet.PageSetup.Setup.FooterMargin >= 0)
            {
                DataStream.WriteStartElement("pageMargins");
                DataStream.WriteAtt("left", WorkSheet.PageSetup.LeftMargin);
                DataStream.WriteAtt("right", WorkSheet.PageSetup.RightMargin);
                DataStream.WriteAtt("top", WorkSheet.PageSetup.TopMargin);
                DataStream.WriteAtt("bottom", WorkSheet.PageSetup.BottomMargin);
                DataStream.WriteAtt("header", WorkSheet.PageSetup.Setup.HeaderMargin);
                DataStream.WriteAtt("footer", WorkSheet.PageSetup.Setup.FooterMargin);
                DataStream.WriteEndElement();
            }
        }

        private void WritePageSetup(TSheet WorkSheet, ref int CurrentSheetLastRel)
        {
            DataStream.WriteStartElement("pageSetup");

            if (WorkSheet.PageSetup.Setup.PaperSize != 1) DataStream.WriteAtt("paperSize", WorkSheet.PageSetup.Setup.PaperSize);

            if (WorkSheet.PageSetup.Setup.Scale != 100) DataStream.WriteAtt("scale", WorkSheet.PageSetup.Setup.Scale);
            unchecked
            {
                if (WorkSheet.PageSetup.Setup.PageStart != null) DataStream.WriteAtt("firstPageNumber", (UInt32)WorkSheet.PageSetup.Setup.PageStart);
            }
            if (WorkSheet.PageSetup.Setup.FitWidth != 1) DataStream.WriteAtt("fitToWidth", WorkSheet.PageSetup.Setup.FitWidth);
            if (WorkSheet.PageSetup.Setup.FitHeight != 1) DataStream.WriteAtt("fitToHeight", WorkSheet.PageSetup.Setup.FitHeight);

            int opt = WorkSheet.PageSetup.Setup.GetPrintOptions(true);
            if ((opt & 0x01) != 0) DataStream.WriteAtt("pageOrder", "overThenDown");

                DataStream.WriteAtt("usePrinterDefaults", (opt &0x04) == 0, true);
                DataStream.WriteAtt("blackAndWhite", (opt &0x08) != 0, false);
                DataStream.WriteAtt("draft", (opt &0x10) != 0, false);

                if ((opt & 0x40) == 0)
                {
                    if ((opt & 0x02) != 0) DataStream.WriteAtt("orientation", "portrait");
                    else DataStream.WriteAtt("orientation", "landscape");
                }
             
            DataStream.WriteAtt("useFirstPageNumber", (opt &0x80) != 0, false);
            
            if ((opt & 0x20) != 0)
            {
                if ((opt & 0x200) != 0)  DataStream.WriteAtt("cellComments", "atEnd");
                else DataStream.WriteAtt("cellComments", "asDisplayed");
            }

            switch ((opt >> 10) & 0x03)
            {
                case 1: DataStream.WriteAtt("errors", "blank"); break;
                case 2: DataStream.WriteAtt("errors", "dash"); break;
                case 3: DataStream.WriteAtt("errors", "NA"); break;
            }

            if (WorkSheet.PageSetup.Setup.HPrintRes!=600) DataStream.WriteAtt("horizontalDpi", WorkSheet.PageSetup.Setup.HPrintRes);
            if (WorkSheet.PageSetup.Setup.VPrintRes!=600) DataStream.WriteAtt("verticalDpi", WorkSheet.PageSetup.Setup.VPrintRes);
            if (WorkSheet.PageSetup.Setup.Copies!=1)   DataStream.WriteAtt("copies", WorkSheet.PageSetup.Setup.Copies);

            if (WorkSheet.PageSetup.Pls != null)
            {
                CurrentSheetLastRel++;
                DataStream.WriteRelationship("id", CurrentSheetLastRel);
            }


            DataStream.WriteEndElement();
        }

        private void WriteHeaderFooter(TSheet WorkSheet)
        {
            DataStream.WriteStartElement("headerFooter");
            THeaderAndFooter ha = WorkSheet.PageSetup.HeaderAndFooter;
            DataStream.WriteAtt("differentOddEven", ha.DiffEvenPages, false);
            DataStream.WriteAtt("differentFirst", ha.DiffFirstPage, false);
            DataStream.WriteAtt("scaleWithDoc", ha.ScaleWithDoc, true);
            DataStream.WriteAtt("alignWithMargins", ha.AlignMargins, true);

            if (ha.DefaultHeader != null && ha.DefaultHeader.Length > 0) DataStream.WriteElement("oddHeader", ha.DefaultHeader);
            if (ha.DefaultFooter != null && ha.DefaultFooter.Length > 0) DataStream.WriteElement("oddFooter", ha.DefaultFooter);

            if (ha.DiffEvenPages) //we might have to write it even if null, if oddheader is not.
            {
                DataStream.WriteElement("evenHeader", ha.EvenHeader);
                DataStream.WriteElement("evenFooter", ha.EvenFooter);
            }

            if (ha.DiffFirstPage)
            {
                DataStream.WriteElement("firstHeader", ha.FirstHeader);
                DataStream.WriteElement("firstFooter", ha.FirstFooter);
            }

            DataStream.WriteEndElement();
        }

        private void WriteRowBreaks(TSheet aSheet)
        {
            DataStream.WriteStartElement("rowBreaks");
            WriteBreaks(aSheet.SheetGlobals.HPageBreaks);
            DataStream.WriteEndElement();
        }

        private void WriteColBreaks(TSheet aSheet)
        {
            DataStream.WriteStartElement("colBreaks");
            WriteBreaks(aSheet.SheetGlobals.VPageBreaks);
            DataStream.WriteEndElement();
        }

        private void WriteBreaks(IPageBreakList PageBreakList)
        {
            int Count = PageBreakList.RealCount();
            if (Count <= 0) return;

            DataStream.WriteAtt("count", Count);
            DataStream.WriteAtt("manualBreakCount", Count);
            for (int i = 0; i < Count; i++)
            {
                TPageBreak brk = PageBreakList.GetItem(i);
                DataStream.WriteStartElement("brk");

                DataStream.WriteAtt("man", true, false);
                DataStream.WriteAtt("id", brk.Id);
                if (brk.Min > 0) DataStream.WriteAtt("min", brk.Min);
                if (brk.Max > 0) DataStream.WriteAtt("max", brk.Max);

                DataStream.WriteEndElement();
            }
        }

        private void WriteCustomProperties(TSheet aSheet)
        {
            DataStream.WriteStartElement("customProperties");
            DataStream.WriteEndElement();
        }

        private void WriteCellWatches(TSheet aSheet)
        {
            DataStream.WriteStartElement("cellWatches");
            DataStream.WriteEndElement();
        }

        private void WriteIgnoredErrors(TSheet aSheet)
        {
            DataStream.WriteStartElement("ignoredErrors");
            DataStream.WriteEndElement();
        }

        private void WriteSmartTags(TSheet aSheet)
        {
            DataStream.WriteStartElement("smartTags");
            DataStream.WriteEndElement();
        }

        private void WriteDrawing(TSheet aSheet, ref int CurrentSheetLastRel)
        {
            DataStream.WriteStartElement("drawing");
            if (HasDrawings(aSheet.Drawing))
            {
                CurrentSheetLastRel++;
                DataStream.WriteRelationship("id", CurrentSheetLastRel);
            }
            DataStream.WriteEndElement();
        }

        private void WriteDrawingHF(TSheet aSheet, ref int CurrentSheetLastRel)
        {
            DataStream.WriteStartElement("drawingHF");
            DataStream.WriteEndElement();
        }

        private void WriteLegacyDrawing(TSheet aSheet, ref int CurrentSheetLastRel)
        {
            DataStream.WriteStartElement("legacyDrawing");
            CurrentSheetLastRel++;
            DataStream.WriteRelationship("id", CurrentSheetLastRel);
            DataStream.WriteEndElement();
        }

        private void WriteLegacyDrawingHF(TSheet aSheet, ref int CurrentSheetLastRel)
        {
            DataStream.WriteStartElement("legacyDrawingHF");
            CurrentSheetLastRel++;
            DataStream.WriteRelationship("id", CurrentSheetLastRel);
            DataStream.WriteEndElement();
        }

        private void WritePicture(TSheet aSheet)
        {
            DataStream.WriteStartElement("picture");
            DataStream.WriteEndElement();
        }

        private void WriteOleObjects(TSheet aSheet)
        {
            DataStream.WriteStartElement("oleObjects");
            DataStream.WriteEndElement();
        }

        private void WriteControls(TSheet aSheet)
        {
            DataStream.WriteStartElement("controls");
            DataStream.WriteEndElement();
        }

        private void WriteWebPublishItems(TSheet aSheet)
        {
            DataStream.WriteStartElement("webPublishItems");
            DataStream.WriteEndElement();
        }

        private void WriteTableParts(TSheet aSheet)
        {
            DataStream.WriteStartElement("tableParts");
            DataStream.WriteEndElement();
        }

        private void WriteSheet_ExtLst(TSheet aSheet)
        {
            DataStream.WriteFutureStorage(aSheet.FutureStorage);
        }
        #endregion

        #region External Links
        private void WriteExternalLinks()
        {
            for (int i = 0; i < Globals.References.Supbooks.Count; i++)
            {
                WriteExternalLink(i, Globals.References.Supbooks[i]);
            }
        }

        private void WriteExternalLink(int i, TSupBookRecord SupBookRecord)
        {
            if (!SupBookNeedsSaving(SupBookRecord)) return;

            Uri PartUri = TOpenXmlWriter.GetFileUri(TOpenXmlWriter.ExternalLinkBaseURI + "externalLink", i + 1);
            DataStream.CreatePart(PartUri, TOpenXmlWriter.ExternalLinkContentType);
            DataStream.CreateRelationshipFromUri(TOpenXmlWriter.WorkbookURI, TOpenXmlWriter.ExternalLinkRelationshipType, Globals.SheetCount + TOpenXmlWriter.RelIdExternalLinks + i);

            DataStream.WriteStartDocument("externalLink", false);
            WriteActualExternalLink(SupBookRecord);
            DataStream.WriteFutureStorage(SupBookRecord.FutureStorage);
            DataStream.WriteEndDocument();
        }

        private static bool SupBookNeedsSaving(TSupBookRecord SupBookRecord)
        {
            if (SupBookRecord.IsLocal) return false; // this one won't be saved
            if (string.IsNullOrEmpty(SupBookRecord.OleOrDdeLink) && string.IsNullOrEmpty(SupBookRecord.BookName())) return false;
            return true;
        }

        private void WriteActualExternalLink(TSupBookRecord SupBookRecord)
        {
            string OleOrDdeLink = SupBookRecord.OleOrDdeLink;
            if (OleOrDdeLink != null)
            {
                WriteOleOrDDELink(OleOrDdeLink, SupBookRecord);
                return;
            }

            WriteExternalBooks(SupBookRecord);
        }

        private void WriteExternalBooks(TSupBookRecord SupBookRecord)
        {
            DataStream.WriteStartElement("externalBook");
            DataStream.WriteRelationship("id", 1);

            string FileName = SupBookRecord.BookName();//.Replace('\\', '/');

            if (FileName == TFormulaMessages.ErrString(TFlxFormulaErrorValue.ErrRef))
            {
                Uri DestUri = new Uri(FileName, UriKind.RelativeOrAbsolute);   //CreateRelationship uses OriginalString, so we need to have that one right.
                DataStream.CreateRelationshipToUri(DestUri, TargetMode.External, TOpenXmlManager.xlPathMissing, TOpenXmlWriter.FlexCelRid + "1");
            }
            else
            {
                StringBuilder sb = new StringBuilder();
                //Absolute paths should start with file:///
                bool IsAbsolute = FileName.StartsWith("\\\\") || (FileName.Length >= 2 && FileName[1] == ':');
                if (IsAbsolute) { sb.Append("file:///"); }
                foreach (char c in FileName)
                {
                    if (c == ' ' || c == '%' || c == '#' || c == '{' || c == '}' || c == '^' || c == '`') { sb.Append("%"); sb.Append(((int)c).ToString("x2", CultureInfo.InvariantCulture)); }
                    else if (!IsAbsolute && c == '\\') sb.Append("/");
                    else sb.Append(c); //OOXML doesn't like Hex-escaped unicode chars. We can't use Uri.EscapeString here.
                }
                Uri DestUri = new Uri(sb.ToString(), UriKind.RelativeOrAbsolute);   //CreateRelationship uses OriginalString, so we need to have that one right.
                DataStream.CreateRelationshipToUri(DestUri, TargetMode.External, TOpenXmlManager.externalRelationshipType, TOpenXmlWriter.FlexCelRid + "1");
            }

            WriteExternalSheetNames(SupBookRecord);//order is important...
            WriteExternalDefinedNames(SupBookRecord);

            DataStream.WriteEndElement();

        }

        private void WriteExternalDefinedNames(TSupBookRecord SupBook)
        {
            DataStream.WriteStartElement("definedNames");

            for (int i = 0; i < SupBook.FExternNameList.Count; i++)
            {
                TExternNameRecord ex = SupBook.FExternNameList[i];
                DataStream.WriteStartElement("definedName");
                DataStream.WriteAtt("name", ex.Name);
                if (ex.SheetIndexInOtherFile > 0) DataStream.WriteAtt("sheetId", ex.SheetIndexInOtherFile - 1);
                DataStream.WriteEndElement();
            }

            DataStream.WriteEndElement();
        }

        private void WriteExternalSheetNames(TSupBookRecord SupBook)
        {
            if (SupBook.IsAddin) return;
            DataStream.WriteStartElement("sheetNames");
            
            int SheetCount = SupBook.SheetCount();
            for (int i = 0; i < SheetCount; i++)
            {
                DataStream.WriteStartElement("sheetName");
                DataStream.WriteAtt("val", SupBook.SheetName(i, Globals));
                DataStream.WriteEndElement();
            }

            DataStream.WriteEndElement();
        }


         
        private void WriteOleOrDDELink(string OleOrDdeLink, TSupBookRecord SupBookRecord)
        {
            if (SupBookRecord.FExternNameList.Count == 0) return; //can't figure if it is a DDE or OLE link
            TExternNameRecord en = SupBookRecord.FExternNameList[0];

            if (en.IsDdeLink) WriteDdeLink(OleOrDdeLink, SupBookRecord); else WriteOleLink(OleOrDdeLink, SupBookRecord);
        }

        private void WriteDdeLink(string OleOrDdeLink, TSupBookRecord SupBookRecord)
        {
            DataStream.WriteStartElement("ddeLink");
            
            int Pos3 = OleOrDdeLink.IndexOf((char)3);
            string FileName = string.Empty; if (Pos3 + 1 < OleOrDdeLink.Length) FileName = OleOrDdeLink.Substring(Pos3 + 1);
            DataStream.WriteAtt("ddeTopic", FileName);

            if (Pos3 < 0) Pos3 = 0;
            string ddeService = OleOrDdeLink.Substring(0, Pos3);
            DataStream.WriteAtt("ddeService", ddeService);

            DataStream.WriteStartElement("ddeItems");

            for (int i = 0; i < SupBookRecord.FExternNameList.Count; i++)
            {
                TExternNameRecord en = SupBookRecord.FExternNameList[i];

                DataStream.WriteStartElement("ddeItem");

                int opt = en.OptionFlags;

                DataStream.WriteAtt("name", en.Name);
                DataStream.WriteAtt("advise", (opt & 0x02) != 0, false);
                DataStream.WriteAtt("preferPic", (opt & 0x04) != 0, false);
                DataStream.WriteAtt("ole", (opt & 0x8) != 0, false);

                DataStream.WriteEndElement();
            }

            DataStream.WriteEndElement();
            DataStream.WriteEndElement();
        }

        private void WriteOleLink(string OleOrDdeLink, TSupBookRecord SupBookRecord)
        {
            DataStream.WriteStartElement("oleLink");
            WriteOleOrDdeRelationship(OleOrDdeLink, "progId");

            DataStream.WriteStartElement("oleItems");

            for (int i = 0; i < SupBookRecord.FExternNameList.Count; i++)
            {
                TExternNameRecord en = SupBookRecord.FExternNameList[i];

                DataStream.WriteStartElement("oleItem");

                int opt = en.OptionFlags;

                DataStream.WriteAtt("name", en.Name);
                DataStream.WriteAtt("advise", (opt & 0x02) != 0, false);
                DataStream.WriteAtt("preferPic", (opt & 0x04) != 0, false);
                DataStream.WriteAtt("icon", (opt & 0x8000) != 0, false);

                DataStream.WriteEndElement();
            }

            DataStream.WriteEndElement();
            DataStream.WriteEndElement();

        }

        private void WriteOleOrDdeRelationship(string OleOrDdeLink, string ProgIdName)
        {
            DataStream.WriteRelationship("id", 1);

            int Pos3 = OleOrDdeLink.IndexOf((char)3);
            string FileName = string.Empty; if (Pos3 + 1 < OleOrDdeLink.Length) FileName = OleOrDdeLink.Substring(Pos3 + 1);
            DataStream.CreateRelationshipToUri(new Uri(FileName, UriKind.RelativeOrAbsolute), TargetMode.External, TOpenXmlManager.OleRelationshipType, TOpenXmlWriter.FlexCelRid + "1");

            if (Pos3 < 0) Pos3 = 0;
            string ProgId = OleOrDdeLink.Substring(0, Pos3);
            DataStream.WriteAtt(ProgIdName, ProgId);
        }
        #endregion

        #region Connections
        internal void WriteConnections()
        {
            if (String.IsNullOrEmpty(Globals.XlsxConnections)) return;
            DataStream.CreatePart(TOpenXmlWriter.ConnectionsURI, TOpenXmlWriter.ConnectionsContentType);
            DataStream.CreateRelationshipFromUri(TOpenXmlWriter.WorkbookURI, TOpenXmlWriter.ConnectionsRelationshipType,
                Globals.SheetCount + Globals.References.Supbooks.Count + Globals.CustomXMLData.Count +
                Globals.XlsxPivotCache.List.Count + TOpenXmlManager.RelIdConnections);

            DataStream.WriteRaw(Globals.XlsxConnections);

        }
        #endregion

        #region Macros
        private void WriteMacros(bool MacroEnabled)
        {
            if (!xls.HasMacros() || !MacroEnabled) return;
            byte[] MacroData = xls.GetMacroData();
            if (MacroData == null || MacroData.Length == 0) return;
            DataStream.WritePart(TOpenXmlWriter.VBAURI, TOpenXmlWriter.VBAContentType, MacroData);
            DataStream.CreateRelationshipToUri(TOpenXmlWriter.WorkbookURI, TOpenXmlWriter.VBAURI, TargetMode.Internal, TOpenXmlWriter.VBARelationshipType,
                TOpenXmlWriter.GetRId(Globals.SheetCount + Globals.References.Supbooks.Count + TOpenXmlManager.RelIdVBA));
        }
        #endregion

        #region File Properties
        private void WriteCustomFileProperties()
        {
            if (FileProps.Custom == null) return;

            DataStream.CreatePart(TOpenXmlWriter.CustomFilePropsURI, TOpenXmlWriter.CustomFilePropsContentType);
            DataStream.CreateRelationshipFromUri(null, TOpenXmlWriter.CustomFilePropsRelationshipType, TOpenXmlManager.RelIdCustomFileProps);
            DataStream.WriteRaw(FileProps.Custom);

        }

        private void WriteCoreFileProperties()
        {
            //DataStream.CreatePart(TOpenXmlWriter.CoreFilePropsURI, TOpenXmlWriter.CoreFilePropsContentType);
            //DataStream.CreateRelationshipFromUri(null, TOpenXmlWriter.CoreFilePropsRelationshipType, TOpenXmlManager.RelIdCoreFileProps);

        }

        #endregion

        #region Comments
        private bool WriteCommentPart(TSheet aSheet, int SheetId, TSheetRelationship SheetRel, ref int CurrentSheetLastRel, TNoteAuthorList CommentAuthors)
        {
            CurrentSheetLastRel++; 
            Uri CommentsURI = new Uri(TOpenXmlManager.CommentsBaseURI + SheetId.ToString(CultureInfo.InvariantCulture) + ".xml", UriKind.Relative);
            DataStream.CreatePart(CommentsURI, TOpenXmlManager.CommentsContentType);
            DataStream.CreateRelationshipFromUri(SheetRel.Uri, TOpenXmlWriter.CommentsRelationshipType, CurrentSheetLastRel);

            DataStream.WriteStartDocument("comments", false);

            WriteActualComments(aSheet, CommentAuthors);

            DataStream.WriteEndDocument();
            return true;
        }

        private void WriteActualComments(TSheet aSheet, TNoteAuthorList Authors)
        {
            WriteAuthors(Authors);
            WriteCommentList(aSheet, Authors);
            WriteComments_ExtLst(aSheet.Notes);
        }

        private void WriteAuthors(TNoteAuthorList Authors)
        {
            DataStream.WriteStartElement("authors", false);
            WriteActualAuthors(Authors);
            DataStream.WriteEndElement();
        }

        private void WriteActualAuthors(TNoteAuthorList Authors)
        {
            foreach (string author in Authors.AuthorsById())
            {
                DataStream.WriteElement("author", author);
            }
        }

        private void WriteCommentList(TSheet aSheet, TNoteAuthorList Authors)
        {
            DataStream.WriteStartElement("commentList", false);

            for (int r = 0; r < aSheet.Notes.Count; r++)
            {
                TNoteRecordList nl = aSheet.Notes[r];
                for (int cIndex = 0; cIndex < nl.Count; cIndex++)
                {
                    TNoteRecord nr = nl[cIndex];
                    WriteOneComment(Authors, r, nr);
                }
            }

            DataStream.WriteEndElement();
        }

        private void WriteOneComment(TNoteAuthorList Authors, int row, TNoteRecord nr)
        {
            DataStream.WriteStartElement("comment");
            DataStream.WriteAttAsAddress("ref", new TCellAddress(row + 1, nr.Col + 1));
            DataStream.WriteAtt("authorId", Authors.GetId(nr.GetAuthor()));

            WriteCommentText(nr);
            WriteCommentPr(nr);

            DataStream.WriteEndElement();
        }

        private void WriteCommentText(TNoteRecord nr)
        {
            DataStream.WriteStartElement("text", false);
            DataStream.WriteRichText(nr.GetText(), xls);
            DataStream.WriteEndElement();
        }

        private void WriteCommentPr(TNoteRecord nr)
        {
            DataStream.WriteStartElement("commentPr");
            DataStream.WriteEndElement();
        }

        private void WriteComments_ExtLst(TNoteList Notes)
        {
            DataStream.WriteFutureStorage(Notes.FutureStorage);
        }
        #endregion

        #region Legacy Drawing
        private void WriteLegacyDrawingPart(bool HF, TSheet aSheet, ref int DrawingId, TSheetRelationship SheetRel, ref int CurrentSheetLastRel)
        {
            CurrentSheetLastRel++;
            Uri LegDrawURI = new Uri(TOpenXmlManager.LegDrawBaseURI + DrawingId.ToString(CultureInfo.InvariantCulture) + ".vml", UriKind.Relative);
            DrawingId++;
            DataStream.CreatePart(LegDrawURI, TOpenXmlManager.LegDrawContentType);
            DataStream.CreateRelationshipFromUri(SheetRel.Uri, TOpenXmlWriter.LegDrawRelationshipType, CurrentSheetLastRel);

            DataStream.WriteStartDocument("xml", null, null);
            DataStream.WriteAtt("xmlns", "v", null, TOpenXmlManager.LegDrawMainNamespace);
            DataStream.WriteAtt("xmlns", "o", null, TOpenXmlManager.LegDrawOfficeNamespace);
            DataStream.WriteAtt("xmlns", "x", null, TOpenXmlManager.LegDrawExcelNamespace);

            if (HF)
            {
                WriteLegacyDrawingHFObjects(aSheet);
            }
            else WriteLegacyDrawingObjects(aSheet);
            DataStream.WriteEndDocument();
        }

        private void WriteLegacyDrawingObjects(TSheet aSheet)
        {
            string SaveDefaultNamespacePrefix = DataStream.DefaultNamespacePrefix;
            DataStream.DefaultNamespacePrefix = "v";
            try
            {
                WriteLegacyComments(aSheet);
                WriteLegacyObjects(aSheet);
            }
            finally
            {
                DataStream.DefaultNamespacePrefix = SaveDefaultNamespacePrefix;
            }
        }

        private void WriteLegacyDrawingHFObjects(TSheet aSheet)
        {
            string SaveDefaultNamespacePrefix = DataStream.DefaultNamespacePrefix;
            DataStream.DefaultNamespacePrefix = "v";
            try
            {
                WriteLegacyDrawingHFPreamble(aSheet);
                WriteLegacyHFObjects(aSheet);
            }
            finally
            {
                DataStream.DefaultNamespacePrefix = SaveDefaultNamespacePrefix;
            }
        }

        private void WriteLegacyComments(TSheet aSheet)
        {
            for (int r = 0; r < aSheet.Notes.Count; r++)
            {
                TNoteRecordList nl = aSheet.Notes[r];
                for (int cIndex = 0; cIndex < nl.Count; cIndex++)
                {
                    TNoteRecord nr = nl[cIndex];
                    WriteOneLegacyDrawingComment(r, nr, aSheet);
                }
            }
        }

        private void WriteOneLegacyDrawingComment(int r, TNoteRecord nr, TSheet aSheet)
        {
            TEscherOPTRecord Opt = nr.GetOpt();
            if (Opt == null) return;
            TShapeOptionList ShapeOptions = Opt.ShapeOptions();

            WriteLegacyObject(r, nr.Col, TObjectType.Comment, aSheet, Opt, nr.GetAnchor(r, aSheet), nr.GetDwg(), nr.GetClientTextBox(), ShapeOptions,
                ColorUtil.BgrToRgb(TCommentProperties.DefaultFillColorRGB), TCommentProperties.DefaultFillColorSystem, -1, TCommentProperties.DefaultLineColorSystem);

        }

        private void WriteLegacyObjects(TSheet aSheet)
        {
            for (int i = 0; i < aSheet.Drawing.ObjectCount; i++)
            {
                if (aSheet.Drawing.IsSpecialDropdown(i)) continue;
                TShapeProperties ShapeProps = aSheet.Drawing.GetObjectProperties(i, true);
                switch (ShapeProps.ObjectType)
                {
                    case TObjectType.CheckBox:
                    case TObjectType.OptionButton:
                    case TObjectType.GroupBox:
                    case TObjectType.Button:
                    case TObjectType.ComboBox:
                    case TObjectType.ListBox:
                    case TObjectType.Label:
                    case TObjectType.Spinner:
                    case TObjectType.ScrollBar:
                        WriteLegacyObject(-1, -1, ShapeProps.ObjectType, aSheet, aSheet.Drawing.GetOPT(i), ShapeProps.Anchor,
                            aSheet.Drawing.GetDwg(i), aSheet.Drawing.GetClientTextBox(i), ShapeProps.ShapeOptions,
                            -1, TSystemColor.Window, -1, TSystemColor.WindowText);
                        break;
                }

            }
        }


        private void WriteLegacyDrawingHFPreamble(TSheet aSheet)
        {
            DataStream.WriteStartElement("shapelayout", "o", false);
              DataStream.WriteAtt("v", "ext", null, "edit");

              DataStream.WriteStartElement("idmap", "o", false);
                DataStream.WriteAtt("v", "ext", null, "edit");
                DataStream.WriteAtt("data", 1);
              DataStream.WriteEndElement(); //idmap

            DataStream.WriteEndElement();  //shapelayout

            DataStream.WriteStartElement("shapetype", false);
              DataStream.WriteAtt("id", (char)0 + "t75");
              DataStream.WriteAtt("coordsize", "21600,21600");
              DataStream.WriteAtt("o", "spt", null, "75");
              DataStream.WriteAtt("o", "preferrelative", null, "t");
              DataStream.WriteAtt("path", "m@4@5l@4@11@9@11@9@5xe");
              DataStream.WriteAtt("filled", "f");
              DataStream.WriteAtt("stroked", "f");
            
              DataStream.WriteStartElement("stroke", false);
                DataStream.WriteAtt("joinstyle", "miter");
              DataStream.WriteEndElement(); //stroke

              DataStream.WriteStartElement("formulas", false);
                WriteEqn("if lineDrawn pixelLineWidth 0");
                WriteEqn("sum @0 1 0");
                WriteEqn("sum 0 0 @1");
                WriteEqn("prod @2 1 2");
                WriteEqn("prod @3 21600 pixelWidth");
                WriteEqn("prod @3 21600 pixelHeight");
                WriteEqn("sum @0 0 1");
                WriteEqn("prod @6 1 2");
                WriteEqn("prod @7 21600 pixelWidth");
                WriteEqn("sum @8 21600 0");
                WriteEqn("prod @7 21600 pixelHeight");
                WriteEqn("sum @10 21600 0"); 
              DataStream.WriteEndElement(); //formulas

              DataStream.WriteStartElement("path", false);
                DataStream.WriteAtt("o", "extrusionok", null, "f");
                DataStream.WriteAtt("gradientshapeok", "t");
                DataStream.WriteAtt("o", "connecttype", null, "rect");
              DataStream.WriteEndElement(); //path

              DataStream.WriteStartElement("lock", "o", false);
                DataStream.WriteAtt("v", "ext", null, "edit");
                DataStream.WriteAtt("aspectratio", "t");
              DataStream.WriteEndElement(); //lock

            DataStream.WriteEndElement(); //shapetype
        }

        private void WriteEqn(string eqn)
        {
            DataStream.WriteStartElement("f", false);
            DataStream.WriteAtt("eqn", eqn);
            DataStream.WriteEndElement();
        }

        private void WriteLegacyHFObjects(TSheet aSheet)
        {
            for (int i = 0; i < aSheet.HeaderImages.DrawingCount; i++)
            {
                TEscherOPTRecord blip = aSheet.HeaderImages.GetBlip(i);
                if (blip == null) continue;

                
                TXlsImgType imageType = TXlsImgType.Unknown;
                using (MemoryStream ms = new MemoryStream())
                {
                    aSheet.HeaderImages.GetDrawingFromStream(i, null, ms, ref imageType, false);
                    DrawingWriter.WriteBlipData(DataStream, null, TXlsxDrawingWriter.GetContentType(imageType), ms.ToArray(), false);
                }

                WriteLegacyHFImage(blip, i, DrawingWriter.GetCurrentMediaRelId);

            }
        }

        private void WriteLegacyHFImage(TEscherOPTRecord blip, int BlipPos, int RelId)
        {
            DataStream.WriteStartElement("shape");
            DataStream.WriteAtt("id", blip.ShapeName);
            DataStream.WriteAtt("o", "spid", null, (char)0 + "s" + blip.ShapeId().ToString(CultureInfo.InvariantCulture));
            DataStream.WriteAtt("type", "#" + (char)0 + "t75");
            THeaderOrFooterAnchor Anchor = blip.GetHeaderAnchor();
            DataStream.WriteAtt("style", String.Format(CultureInfo.InvariantCulture, 
                "position:absolute;margin-left:0;margin-top:0;width:{0}pt;height:{1}pt;z-index:{2}", Anchor.Width * 72.0 / 96.0, Anchor.Height * 72.0 / 96.0, BlipPos + 1));

            if (!blip.PreferRelativeSize) DataStream.WriteAtt("o", "preferrelative", null, "f");
            DataStream.WriteStartElement("imagedata");
            DataStream.WriteAtt("o", "relid", null, TOpenXmlWriter.GetRId(RelId));
            DataStream.WriteAtt("o", "title", null, blip.FileName);

            TCropArea Crop = blip.CropArea;
            WriteAttFF("croptop", Crop.CropFromTop, 0);
            WriteAttFF("cropbottom", Crop.CropFromBottom, 0);
            WriteAttFF("cropleft", Crop.CropFromLeft, 0);
            WriteAttFF("cropright", Crop.CropFromRight, 0);
            WriteAttFF("blacklevel", blip.Brightness, FlxConsts.DefaultBrightness);
            WriteAttFF("gain", blip.Contrast, FlxConsts.DefaultContrast);
            WriteAttFF("gamma", blip.Gamma, FlxConsts.DefaultGamma);

            if (blip.TransparentColor != FlxConsts.NoTransparentColor)
            {
                unchecked
                {
                    string FillColor = THtmlColors.GetColor(ColorUtil.FromArgb((int)blip.TransparentColor));
                    DataStream.WriteAtt("chromakey", FillColor);
                }
            }

            if (blip.Grayscale) DataStream.WriteAtt("grayscale", "t");
            if (blip.BiLevel) DataStream.WriteAtt("bilevel", "t");


            DataStream.WriteEndElement();

            DataStream.WriteStartElement("lock", "o", false);
            DataStream.WriteAtt("v", "ext", null, "edit");
            DataStream.WriteAtt("rotation", "t");
            if (!blip.LockAspectRatio) DataStream.WriteAtt("aspectratio", "f");
            DataStream.WriteEndElement(); //lock

            DataStream.WriteEndElement();
        }

        private void WriteAttFF(string AttName, int value, int DefaultValue)
        {
            if (value == DefaultValue) return;
            DataStream.WriteAtt(AttName, value.ToString(CultureInfo.InvariantCulture) + "f");
        }

        private void WriteLegacyObject(int aRow, int aCol, TObjectType ObjType, TSheet aSheet, TEscherOPTRecord Opt, TClientAnchor Anchor, 
            TEscherClientDataRecord Dwg, TEscherClientTextBoxRecord Tb, TShapeOptionList ShapeOptions,
            long DefFillCol, TSystemColor DefFillSys, long DefLineCol, TSystemColor DefLineSys)
        {
            DataStream.WriteStartElement("shape", false);

            if (!string.IsNullOrEmpty(Opt.ShapeName))
            {
                DataStream.WriteAtt("id", Opt.ShapeName);
            }

            string AltText = ShapeOptions.AsUnicodeString(TShapeOption.wzDescription, null);
            if (AltText != null)
            {
                DataStream.WriteAtt("alt", AltText);
            }

            bool HasFill = ShapeOptions.AsBool(TShapeOption.fNoFillHitTest, true, 4);
            if (!HasFill) DataStream.WriteAtt("filled", "f");
            string FillColor = GetColorString(ShapeOptions, TShapeOption.fillColor, DefFillCol, DefFillSys);
            if (FillColor != null) DataStream.WriteAtt("fillcolor", FillColor);

            bool HasBorder = ShapeOptions.AsBool(TShapeOption.fNoLineDrawDash, true, 3);
            if (!HasBorder) DataStream.WriteAtt("stroked", "f");
            string LineColor = GetColorString(ShapeOptions, TShapeOption.lineColor, DefLineCol, DefLineSys);
            if (LineColor != null) DataStream.WriteAtt("strokecolor", LineColor);

            string StyleStr = "position:absolute";
            if (!Opt.Visible) StyleStr += ";visibility:hidden";
            DataStream.WriteAtt("style", StyleStr);

            //common props
            DataStream.WriteAtt("coordsize", "21600,21600");
            DataStream.WriteAtt("path", "m,l,21600r21600,l21600,xe");
            if (ObjType == TObjectType.Comment)
            {
                DataStream.WriteAtt("o", "spt", null, "202");
            }
            else
            {
                DataStream.WriteAtt("o", "spt", null, "201"); //without this, text in checkboxes can't be edited. button is 201 too
            }
            DataStream.WriteAtt("o", "insetmode", null, "auto");

            DataStream.WriteStartElement("stroke");
            DataStream.WriteAtt("joinstyle", "miter");
            DataStream.WriteEndElement();

            if (ObjType == TObjectType.Comment)
            {
                DataStream.WriteStartElement("shadow");
                DataStream.WriteAtt("on", "t");
                DataStream.WriteAtt("color", "black");
                DataStream.WriteAtt("obscured", "t");
                DataStream.WriteEndElement();
            }


            DataStream.WriteStartElement("path");
            if (ObjType == TObjectType.Comment)
            {
                DataStream.WriteAtt("gradientshapeok", "t");
                DataStream.WriteAtt("o", "connecttype", null, "none");
            }
            else
            {
                DataStream.WriteAtt("shadowok", GetLegacyBool(ShapeOptions.AsBool(TShapeOption.fFillOK, false, 5)));
                DataStream.WriteAtt("strokeok", GetLegacyBool(ShapeOptions.AsBool(TShapeOption.fFillOK, false, 3)));
                DataStream.WriteAtt("fillok", GetLegacyBool(ShapeOptions.AsBool(TShapeOption.fFillOK, false, 0)));
                DataStream.WriteAtt("o", "extrusionok", null, "f");
                DataStream.WriteAtt("o", "connecttype", null, "rect");
            }

            DataStream.WriteEndElement();

            long lar = ShapeOptions.AsLong(TShapeOption.fLockAgainstGrouping, 0);
            if ((lar & 0x80) != 0 && (lar & 0x800000) != 0)
            {
                WriteLegacyLockAspectRatio();
            }

            DataStream.WriteStartElement("textbox");

            string Style = "mso-direction-alt:auto";
            long ft = ShapeOptions.AsLong(TShapeOption.fFitTextToShape, 0);
            if ((ft & 0x2) != 0 && (ft & 0x20000) != 0) Style += ";mso-fit-shape-to-text:t";

            if (Tb != null)
            {
                TTextRotation TextRotation = Tb.TextRotation;
                switch (TextRotation)
                {
                    case TTextRotation.Normal:
                        break;
                    case TTextRotation.Rotated90Degrees:
                        Style += ";layout-flow:vertical;mso-layout-flow-alt:bottom-to-top";
                        break;
                    case TTextRotation.RotatedMinus90Degrees:
                        Style += ";layout-flow:vertical";
                        break;
                    case TTextRotation.Vertical:
                        Style += ";layout-flow:vertical;mso-layout-flow-alt:top-to-bottom";
                        break;
                }
            }

            DataStream.WriteAtt("style", Style);

            if (Tb != null)
            {
                TTXO txo = Tb.ClientData as TTXO;
                if (txo != null)
                {
                    WriteTextboxDiv(txo.GetText(), ObjType);
                }
            }

            DataStream.WriteEndElement();

            WriteClientData(aRow, aCol, ObjType, Anchor, Dwg, Tb, aSheet);
            DataStream.WriteEndElement();
        }

        private static string GetLegacyBool(bool p)
        {
            if (p) return "t"; else return "f";
        }

        private string GetColorString(TShapeOptionList ShapeOptions, TShapeOption ShpOpt, long DefCol, TSystemColor DefSysCol)
        {
            string FillColor = null;
            unchecked
            {
                TFillStyle FillStyle = TSheet.GetObjectBackground(ShapeOptions, ShpOpt, DefCol, DefSysCol);

                TSolidFill SolidFill = FillStyle as TSolidFill;
                if (SolidFill != null)
                {
                    switch (SolidFill.Color.ColorType)
                    {
                        case TDrawingColorType.System:
                            string Fc = ColorUtil.GetSystemColorName(SolidFill.Color.System);
                            if (Fc != null) FillColor = Fc + " [" + (56 + ColorUtil.GetSysColor(SolidFill.Color.System)).ToString(CultureInfo.InvariantCulture) + "]";
                            break;

                        default:
                            FillColor = THtmlColors.GetColor(SolidFill.Color.ToColor(xls));
                            break;
                    }
                }
            }
            return FillColor;
        }

        private void WriteLegacyLockAspectRatio()
        {
            string SaveDefaultNamespacePrefix = DataStream.DefaultNamespacePrefix;
            DataStream.DefaultNamespacePrefix = "o" ;
            try
            {
                DataStream.WriteStartElement("lock");
                DataStream.WriteAtt("aspectratio", "t");
                DataStream.WriteAtt("ext", TOpenXmlManager.LegDrawMainNamespace, "edit");
                DataStream.WriteEndElement();
            }
            finally
            {
                DataStream.DefaultNamespacePrefix = SaveDefaultNamespacePrefix;
            }
        }

        private void WriteTextboxDiv(TRichString Text, TObjectType ObjType)
        {
            string SaveDefaultNamespacePrefix = DataStream.DefaultNamespacePrefix;
            DataStream.DefaultNamespacePrefix = null;
            try
            {
                DataStream.WriteStartElement("div");
                DataStream.WriteAtt("style", "text-align:left");
                if (Text != null && ObjType != TObjectType.Comment)
                {
                    DataStream.WriteRaw(Text.ToHtml(xls, xls.GetDefaultFormat, THtmlVersion.Html_401, THtmlStyle.Simple, Encoding.UTF8, null, true));
                }
                DataStream.WriteEndElement();
            }
            finally
            {
                DataStream.DefaultNamespacePrefix = SaveDefaultNamespacePrefix;
            }
        }

        private void WriteClientData(int aRow, int aCol, TObjectType ObjType, TClientAnchor Anchor, TEscherClientDataRecord Dwg, TEscherClientTextBoxRecord Tb, TSheet aSheet)
        {
            string SaveDefaultNamespacePrefix = DataStream.DefaultNamespacePrefix;
            DataStream.DefaultNamespacePrefix = "x";
            try
            {
                if (Dwg == null) return;
                TMsObj MsObj = Dwg.ClientData as TMsObj;
                if (MsObj == null) return;

                DataStream.WriteStartElement("ClientData", false);
                switch (ObjType)
                {
                    case TObjectType.Comment: DataStream.WriteAtt("ObjectType", "Note"); break;
                    case TObjectType.CheckBox: DataStream.WriteAtt("ObjectType", "Checkbox"); break;
                    case TObjectType.OptionButton: DataStream.WriteAtt("ObjectType", "Radio"); break;
                    case TObjectType.ComboBox: DataStream.WriteAtt("ObjectType", "Drop"); break;
                    case TObjectType.Button: DataStream.WriteAtt("ObjectType", "Button"); break;
                    case TObjectType.ListBox: DataStream.WriteAtt("ObjectType", "List"); break;
                    case TObjectType.GroupBox: DataStream.WriteAtt("ObjectType", "GBox"); break;
                    case TObjectType.Label: DataStream.WriteAtt("ObjectType", "Label"); break;
                    case TObjectType.Spinner: DataStream.WriteAtt("ObjectType", "Spin"); break;
                    case TObjectType.ScrollBar: DataStream.WriteAtt("ObjectType", "Scroll"); break;
                    default: return;
                }

                IRowColSize rc = new RowColSize(xls.HeightCorrection, xls.WidthCorrection, aSheet);

                if (Anchor != null)
                {
                    if (Anchor.AnchorType != TFlxAnchorType.MoveAndDontResize && Anchor.AnchorType != TFlxAnchorType.MoveAndResize) DataStream.WriteElement("MoveWithCells", string.Empty);
                    if (Anchor.AnchorType != TFlxAnchorType.MoveAndResize) DataStream.WriteElement("SizeWithCells", string.Empty);

                    TClientAnchor Anchor1 = Anchor.Inc(); //we need this one 1-based so dxpix are calc ok.
                    DataStream.WriteElement("Anchor", String.Format(CultureInfo.InvariantCulture, "{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}",
                    Anchor.Col1, Anchor1.Dx1Pix(rc), Anchor.Row1, Anchor1.Dy1Pix(rc), Anchor.Col2, Anchor1.Dx2Pix(rc), Anchor.Row2, Anchor1.Dy2Pix(rc)));
                }

                if (!MsObj.IsLocked) DataStream.WriteElement("Locked", "False");
                if (MsObj.IsDefaultSize) DataStream.WriteElement("DefaultSize", string.Empty);
                if (!MsObj.IsPrintable) DataStream.WriteElement("PrintObject", "False");
                if (MsObj.IsDisabled) DataStream.WriteElement("Disabled", string.Empty);
                //if (MsObj.IsPublished) DataStream.WriteElement("Published", string.Empty);
                if (!MsObj.IsAutoFill) DataStream.WriteElement("AutoFill", "False");
                if (!MsObj.IsAutoLine) DataStream.WriteElement("AutoLine", "False");

                if (ObjType == TObjectType.CheckBox || ObjType == TObjectType.OptionButton)
                {
                    DataStream.WriteElement("Checked", (uint)MsObj.GetCheckbox());
                }

                
                string FmlaLink = MsObj.GetFmlaLinkXlsx(aSheet.Cells.CellList);
                if (FmlaLink != null) DataStream.WriteElement("FmlaLink", FmlaLink);

                if (MsObj.HasSpinProps())
                {
                    double h1 = Anchor.CalcImageHeightInternal(new RowColSize(1, 1, aSheet));
                    int sv = MsObj.GetObjectSpinValue(true, h1);
                    TSpinProperties sd = MsObj.GetSpinProps();
                    DataStream.WriteElement("Val", sv);
                    DataStream.WriteElement("Min", sd.Min);
                    DataStream.WriteElement("Max", sd.Max);
                    DataStream.WriteElement("Inc", sd.Incr);
                    DataStream.WriteElement("Page", sd.Page);
                    DataStream.WriteElement("Dx", sd.Dx);
                }

                string FmlaRange = MsObj.GetFmlaRangeXlsx(aSheet.Cells.CellList);
                if (FmlaRange != null) DataStream.WriteElement("FmlaRange", FmlaRange);

                TComboBoxProperties cb = MsObj.GetComboProps();
                if (cb != null)
                {
                    DataStream.WriteElement("Sel", MsObj.GetObjectSelection());
                    if (cb.DropLines > 0) DataStream.WriteElement("DropLines", cb.DropLines);
                    if (cb.SelectionType != TListBoxSelectionType.Single) DataStream.WriteElement("SelType", GetSelType(cb.SelectionType));
                }

                string Macro = MsObj.GetFmlaMacroXlsx(aSheet.Cells.CellList);
                if (Macro != null) DataStream.WriteElement("FmlaMacro", Macro);

                if (!MsObj.GetObject3D())
                {
                    if (ObjType == TObjectType.ListBox || ObjType == TObjectType.ComboBox)
                    {
                        DataStream.WriteElement("NoThreeD2", null);
                    }
                    else
                    {
                        DataStream.WriteElement("NoThreeD", null);
                    }
                }


                if (ObjType == TObjectType.OptionButton && !MsObj.GetRbFirstInGroup()) DataStream.WriteElement("FirstButton", null);

                if (Tb != null)
                {
                    WriteHAlign(Tb.HAlign);
                    WriteVAlign(Tb.VAlign);
                    if (!Tb.LockText) DataStream.WriteElement("LockText", "False");
                }

                if (aRow >= 0) DataStream.WriteElement("Row", (UInt32)aRow);
                if (aCol >= 0) DataStream.WriteElement("Column", (UInt32)aCol);

                DataStream.WriteEndElement();
            }
            finally
            {
                DataStream.DefaultNamespacePrefix = SaveDefaultNamespacePrefix;
            }
        }

        private string GetSelType(TListBoxSelectionType st)
        {
            switch (st)
            {
                case TListBoxSelectionType.Multi:
                    return "Multi";

                case TListBoxSelectionType.Extend:
                    return "Extend";

                default:
                    return "Single";
            }
        }

        private void WriteVAlign(TVFlxAlignment align)
        {
            switch (align)
            {
                case TVFlxAlignment.center:
                    DataStream.WriteElement("TextVAlign", "Center");
                    break;

                case TVFlxAlignment.bottom:
                    DataStream.WriteElement("TextVAlign", "Bottom");
                    break;

                case TVFlxAlignment.justify:
                    DataStream.WriteElement("TextVAlign", "Justify");
                    break;

                case TVFlxAlignment.distributed:
                    DataStream.WriteElement("TextVAlign", "Distributed");
                    break;
            }
        }

        private void WriteHAlign(THFlxAlignment align)
        {
            switch (align)
            {
                case THFlxAlignment.center:
                    DataStream.WriteElement("TextHAlign", "Center");
                    break;

                case THFlxAlignment.right:
                    DataStream.WriteElement("TextHAlign", "Right");
                    break;
                
                case THFlxAlignment.justify:
                    DataStream.WriteElement("TextHAlign", "Justify");
                    break;

                case THFlxAlignment.distributed:
                    DataStream.WriteElement("TextHAlign", "Distributed");
                    break;
            }
        }

        #endregion

        #region Pivot Tables
        private void WritePivotTableParts(TSheet aSheet, int SheetId, TSheetRelationship SheetRel, ref int CurrentSheetLastRel)
        {
            for (int i = 0; i < aSheet.XlsxPivotTables.List.Count; i++)
            {
                TXlsxPivotTable pt = aSheet.XlsxPivotTables.List[i];
                if (pt.Cache == null) continue;
                CurrentSheetLastRel++;
                Uri PivotURI = new Uri(TOpenXmlManager.PivotTableBaseURI + (i+1).ToString(CultureInfo.InvariantCulture) + ".xml", UriKind.Relative);
                DataStream.CreatePart(PivotURI, TOpenXmlManager.PivotTableContentType);
                DataStream.CreateRelationshipFromUri(SheetRel.Uri, TOpenXmlWriter.PivotTableRelationshipType, CurrentSheetLastRel);
                DataStream.CreateRelationshipToUri(pt.Cache.LastSavedUri, TargetMode.Internal, TOpenXmlManager.PivotCacheDefRelationshipType, TOpenXmlWriter.FlexCelRid + "1");

                pt.SaveToXlsx(DataStream, "pivotTableDefinition", false);
            }
        }

        #endregion
    }
}
