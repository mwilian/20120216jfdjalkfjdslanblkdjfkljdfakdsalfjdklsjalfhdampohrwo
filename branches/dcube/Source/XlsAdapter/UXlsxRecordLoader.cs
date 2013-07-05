using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Diagnostics;

using FlexCel.Core;
using System.Drawing;
using System.Globalization;

namespace FlexCel.XlsAdapter
{
    /// <summary>
    /// A class to load xlsx files
    /// </summary>
    internal class TXlsxRecordLoader
    {
        #region Variables
        private ExcelFile xls;
        private TOpenXmlReader DataStream;
        private TXlsxDrawingLoader DrawingLoader;
        internal TSST SST;
        internal TVirtualReader VirtualReader;
        #endregion

        #region Constructors
        internal TXlsxRecordLoader(TOpenXmlReader aDataStream, ExcelFile axls, TSST aSST, TVirtualReader aVirtualReader)
        {
            DataStream = aDataStream;
            xls = axls;
            DrawingLoader = new TXlsxDrawingLoader(DataStream, xls);
            SST = aSST;
            VirtualReader = aVirtualReader;
        }
        #endregion

        #region Utils
        internal TEncryptionData Encryption
        {
            get
            {
                if (DataStream == null) return null;
                return DataStream.Encryption;
            }
        }
        #endregion

        #region Workbook
        internal void ReadWorkbook(TWorkbookGlobals Globals, List<string> ExternalRefs, List<string> NameDefinitions, out bool MacroEnabled)
        {
            DataStream.SelectWorkbook(out MacroEnabled);
            while (DataStream.NextTag())
            {
                switch (DataStream.RecordName())
                {
                    case "workbook":
                        ReadActualWorkbook(Globals, ExternalRefs, NameDefinitions);
                        break;

                    default: //shouldn't happen
                        DataStream.GetXml();
                        break;
                }
            }
        }

        private void ReadActualWorkbook(TWorkbookGlobals Globals, List<string> ExternalRefs, List<string> NameDefinitions)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "bookViews": ReadBookViews(Globals); break;
                    case "calcPr": ReadCalcPr(Globals); break;
                    case "customWorkbookViews": ReadCustomWorkbookViews(Globals); break;
                    case "definedNames": ReadDefinedNames(Globals, NameDefinitions); break;
                    case "externalReferences": ReadExternalReferences(Globals, ExternalRefs); break;
                    case "fileRecoveryPr": ReadFileRecoveryPr(Globals); break;
                    case "fileSharing": ReadFileSharing(Globals); break;
                    case "fileVersion": ReadFileVersion(Globals); break;
                    case "functionGroups": ReadFunctionGroups(Globals); break;
                    case "oleSize": ReadOleSize(Globals); break;
                    case "pivotCaches": ReadPivotCaches(Globals); break;
                    case "sheets": ReadSheets(Globals); break;
                    case "smartTagPr": ReadSmartTagPr(Globals); break;
                    case "smartTagTypes": ReadSmartTagTypes(Globals); break;
                    case "webPublishing": ReadWebPublishing(Globals); break;
                    case "webPublishObjects": ReadWebPublishObjects(Globals); break;
                    case "workbookPr": ReadWorkbookPr(Globals); break;
                    case "workbookProtection": ReadWorkbookProtection(Globals); break;

                    case "extLst":
                    default:
                        Globals.AddFutureStorage(new TFutureStorageRecord(DataStream.GetXml()));
                        break;
                }
            }
        }

        internal void LoadNameDefinitions(TWorkbookGlobals Globals, List<string> Definitions)
        {

            for (int i = 0; i < Definitions.Count; i++)
            {
                string FormulaText = TFormulaMessages.TokenString(TFormulaToken.fmStartFormula) + Definitions[i];

                    int DefaultSheet = 0;
                    int ds = Globals.Names[i].RangeSheet;
                    if (ds >= 0)
                        DefaultSheet = ds;
                    string DefaultSheetName = Globals.GetSheetName(DefaultSheet);
                    bool ThrowExceptions = (Globals.Workbook.ErrorActions & TExcelFileErrorActions.ErrorOnXlsxInvalidName) != 0;

                    TFormulaConvertTextToInternal Ps = new TFormulaConvertTextToInternal(xls, -1, true, FormulaText, true, true, true, DefaultSheetName, TFmReturnType.Ref, false);
                    Ps.SetReadingXlsx();
                    try
                    {
                        Ps.Parse();
                        TParsedTokenList Data = Ps.GetTokens();
                        Globals.Names[i].FormulaData = Data;
                    }
                    catch (FlexCelException)
                    {
                        if (FlexCelTrace.Enabled) FlexCelTrace.Write(new TXlsxInvalidNameError(
                            XlsMessages.GetString(XlsErr.ErrBadName, Globals.Names[i].Name, 0), Globals.Workbook.ActiveFileName, Globals.Names[i].Name, Definitions[i]));
                        if (ThrowExceptions) throw;

                        Globals.Names[i].FormulaData = new TParsedTokenList(new TBaseParsedToken[] { new TErrDataToken(TFlxFormulaErrorValue.ErrRef) });
                    }
            }
        }

        private void ReadBookViews(TWorkbookGlobals Globals)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "workbookView":
                        ReadWindow1(Globals);
                        break;

                    default:
                        DataStream.GetXml();
                        break;
                }
            }
        }

        private void ReadWindow1(TWorkbookGlobals Globals)
        {
            Globals.AddNewWindow1();
            TWindow1Record w1 = new TWindow1Record();
            Globals.Window1[Globals.Window1.Length - 1] = w1;

            w1.ActiveSheet = DataStream.GetAttributeAsInt("activeTab", 0);
            w1.FirstSheetVisible = DataStream.GetAttributeAsInt("firstSheet", 0);

            w1.xWin = DataStream.GetAttributeAsInt("xWindow", 0);
            w1.yWin = DataStream.GetAttributeAsInt("yWindow", 0);
            w1.dxWin = DataStream.GetAttributeAsInt("windowWidth", 0);
            w1.dyWin = DataStream.GetAttributeAsInt("windowHeight", 0);

            w1.TabRatio = DataStream.GetAttributeAsInt("tabRatio", 600);

            w1.Options = BitOps.GetBool(
                false,
                DataStream.GetAttributeAsBool("minimized", false),
                false,
                DataStream.GetAttributeAsBool("showHorizontalScroll", true),
                DataStream.GetAttributeAsBool("showVerticalScroll", true),
                DataStream.GetAttributeAsBool("showSheetTabs", true),
                !DataStream.GetAttributeAsBool("autoFilterDateGrouping", true));


            switch (DataStream.GetAttribute("visibility"))
            {
                case "hidden": w1.Options |= 0x01; break;
                case "veryHidden": w1.Options |= 0x04; break;
            }


            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                 w1.AddFutureStorage(new TFutureStorageRecord(DataStream.GetXml()));
            }

        }

        private void ReadCalcPr(TWorkbookGlobals Globals)
        {
            //Globals.RecalcId = DataStream.GetAttributeAsInt("calcId", 0);

            switch(DataStream.GetAttribute("calcMode"))
            {
                case "autoNoTable": Globals.CalcOptions.CalcMode = TSheetCalcMode.AutomaticExceptTables; break;
                case "manual": Globals.CalcOptions.CalcMode = TSheetCalcMode.Manual; break;
            }

            Globals.CalcOptions.SaveRecalc = DataStream.GetAttributeAsBool("calcOnSave", true);

            if (DataStream.GetAttributeAsBool("concurrentCalc", true))
            {
                Globals.MultithreadRecalc = DataStream.GetAttributeAsInt("concurrentManualCount", -1);
            }
            else
                Globals.MultithreadRecalc = 0;

            Globals.ForceFullRecalc = DataStream.GetAttributeAsBool("forceFullCalc", false);
            Globals.PrecisionAsDisplayed = !DataStream.GetAttributeAsBool("fullPrecision", true);

            Globals.CalcOptions.IterationEnabled = DataStream.GetAttributeAsBool("iterate", false);
            Globals.CalcOptions.CalcCount = DataStream.GetAttributeAsInt("iterateCount", 100);
            Globals.CalcOptions.Delta = DataStream.GetAttributeAsDouble("iterateDelta", 0.001);
            Globals.CalcOptions.A1RefMode = DataStream.GetAttribute("refMode") != "R1C1";

            DataStream.FinishTag();
        }

        private void ReadCustomWorkbookViews(TWorkbookGlobals Globals)
        {
            DataStream.GetXml(); //IMPLEMENT
        }

        private void ReadDefinedNames(TWorkbookGlobals Globals, List<string> Definitions)
        {
            //Read them all first, so we can handle circular references / unordered references
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "definedName": ReadDefinedName(Globals, Definitions); break;
                    default:
                        DataStream.GetXml();
                        break;
                }
            }
        }

        private void ReadDefinedName(TWorkbookGlobals Globals, List<string> Definitions)
        {
            bool IsInternal;
            string name = TXlsNamedRange.GetInternal(DataStream.GetAttribute("name"), out IsInternal);

            int OptionFlags =
                (DataStream.GetAttributeAsBool("hidden", false) ? 0x01 : 0) |
                (DataStream.GetAttributeAsBool("function", false) ? 0x02 : 0) |
                (DataStream.GetAttributeAsBool("vbProcedure", false) ? 0x04 : 0) |
                (DataStream.GetAttributeAsBool("xlm", false) ? 0x08 : 0) |
                0 | //will be set later, when we know if this is an array fmla or not.
                (IsInternal ? 0x20 : 0) |
                DataStream.GetAttributeAsInt("functionGroupId", 0) << 6 |

                (DataStream.GetAttributeAsBool("publishToServer", false) ? 0x2000 : 0) |
                (DataStream.GetAttributeAsBool("workbookParameter", false) ? 0x4000 : 0);




            string KeybStr = DataStream.GetAttribute("shortcutKey");
            TNameRecord NameDef = new TNameRecord(
                (int)xlr.NAME,
                null, //Will be added later
                OptionFlags,
                name,
                KeybStr,
                DataStream.GetAttributeAsInt("localSheetId", -1), //localsheetid is 0-based(!!)
                DataStream.GetAttribute("customMenu"),
                DataStream.GetAttribute("description"),
                DataStream.GetAttribute("help"),
                DataStream.GetAttribute("statusBar"),
                DataStream.GetAttribute("comment")
                );

            Globals.Names.Add(NameDef);

            Definitions.Add(DataStream.ReadValueAsString());
        }

        private void ReadExternalReferences(TWorkbookGlobals Globals, List<string> ExternalRefs)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "externalReference":
                        ExternalRefs.Add(DataStream.GetRelationship("id"));
                        DataStream.FinishTag();
                        break;
                    default:
                        DataStream.GetXml();
                        break;
                }
            }
        }

        private void ReadFileRecoveryPr(TWorkbookGlobals Globals)
        {
            DataStream.GetXml(); //IMPLEMENT
        }

        private void ReadFileSharing(TWorkbookGlobals Globals)
        {
            string ModifyHash = DataStream.GetAttribute("reservationPassword");
            Globals.FileEncryption.FileSharing = new TFileSharingRecord(DataStream.GetAttributeAsBool("readOnlyRecommended", false), ModifyHash, DataStream.GetAttribute("userName"), true);
            if (ModifyHash != null) Globals.FileEncryption.WriteProt = new TWriteProtRecord();
            DataStream.FinishTag();
        }

        private void ReadFileVersion(TWorkbookGlobals Globals)
        {
            Globals.CodeName07 = DataStream.GetAttribute("codeName");
            Globals.sBOF = TBOFRecord.CreateEmptyWorkbook(DataStream.GetAttributeAsInt("lastEdited", 4));
            DataStream.FinishTag();
        }

        private void ReadFunctionGroups(TWorkbookGlobals Globals)
        {
            DataStream.GetXml(); //IMPLEMENT
        }

        private void ReadOleSize(TWorkbookGlobals Globals)
        {
            DataStream.GetXml(); //IMPLEMENT
        }

        private void ReadPivotCaches(TWorkbookGlobals Globals)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "pivotCache": Globals.XlsxPivotCache.Load(DataStream); break;
                    default:
                        DataStream.GetXml();
                        break;
                }
            }
        }

        private void ReadSheets(TWorkbookGlobals Globals)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "sheet": ReadSheet(Globals); break;
                    default:
                        DataStream.GetXml();
                        break;
                }
            }
        }

        private void ReadSheet(TWorkbookGlobals Globals)
        {
            int OptionFlags = 0;

            switch (DataStream.GetAttribute("state"))
            {
                case "hidden": OptionFlags |= 0x01; break;
                case "veryHidden": OptionFlags |= 0x02; break;
            }
            //Sheet type will be set later.

            Globals.AddBoundSheetFromFile(Convert.ToInt32(DataStream.GetAttribute("sheetId")), new TBoundSheetRecord(OptionFlags, DataStream.GetAttribute("name"), DataStream.GetRelationship("id")));
            DataStream.FinishTag();
        }

        private void ReadSmartTagPr(TWorkbookGlobals Globals)
        {
            DataStream.GetXml(); //IMPLEMENT
        }

        private void ReadSmartTagTypes(TWorkbookGlobals Globals)
        {
            DataStream.GetXml(); //IMPLEMENT
        }

        private void ReadWebPublishing(TWorkbookGlobals Globals)
        {
            DataStream.GetXml(); //IMPLEMENT
        }

        private void ReadWebPublishObjects(TWorkbookGlobals Globals)
        {
            DataStream.GetXml(); //IMPLEMENT
        }

        private void ReadWorkbookPr(TWorkbookGlobals Globals)
        {
            
            //if (DataStream.GetAttributeAsBool("allowRefreshQuery", false)) Globals;
            Globals.AutoCompressPictures = DataStream.GetAttributeAsBool("autoCompressPictures", true);
            Globals.Backup = DataStream.GetAttributeAsBool("backupFile", false);
            Globals.CheckCompatibility = DataStream.GetAttributeAsBool("checkCompatibility", false);
            Globals.CodeName = DataStream.GetAttribute("codeName");
            Globals.Dates1904 = DataStream.GetAttributeAsBool("date1904", false);

            Globals.Theme.ThemeVersion = DataStream.GetAttributeAsInt("defaultThemeVersion", 0);
            if (Globals.BookExt == null) Globals.BookExt = new TBookExtRecord();
            Globals.BookExt.FilterPrivacy = DataStream.GetAttributeAsBool("filterPrivacy", false);
            Globals.BookExt.HidePivotList = DataStream.GetAttributeAsBool("hidePivotFieldList", false);
            Globals.BookExt.BuggedUserAboutSolution = DataStream.GetAttributeAsBool("promptedSolutions", false);
            Globals.BookExt.PublishedBookItems = DataStream.GetAttributeAsBool("publishItems", false);
            Globals.RefreshAll = DataStream.GetAttributeAsBool("refreshAllConnections", false);

            Globals.SaveExternalLinkValues = DataStream.GetAttributeAsBool("saveExternalLinkValues", true);
            Globals.HideBorderUnselLists = !DataStream.GetAttributeAsBool("showBorderUnselectedTables", true);
            Globals.BookExt.ShowInkAnnotation = DataStream.GetAttributeAsBool("showInkAnnotation", true);
            
            string ShowObjects = DataStream.GetAttribute("showObjects");
            switch (ShowObjects)
            {
                case "none":
                    Globals.HideObj = THideObj.HideAll;
                    break;

                case "placeholders":
                    Globals.HideObj = THideObj.ShowPlaceholder;
                    break;

                case "all":
                default:
                    Globals.HideObj = THideObj.ShowAll;
                    break;
            }


            Globals.BookExt.ShowPivotChartFilter = DataStream.GetAttributeAsBool("showPivotChartFilter", false);
            
            string UpdateLinks = DataStream.GetAttribute("updateLinks");
            switch (UpdateLinks)
            {
                case "always":
                    Globals.UpdateLinks = TUpdateLinkOption.SilentlyUpdate;
                    break;

                case "never":
                    Globals.UpdateLinks = TUpdateLinkOption.DontUpdate;
                    break;
                
                case "userSet":
                default:
                    Globals.UpdateLinks = TUpdateLinkOption.PromptUser;
                    break;

            }
            
            DataStream.FinishTag();              
        }

        private void ReadWorkbookProtection(TWorkbookGlobals Globals)
        {
            Globals.WorkbookProtection.Protect = new TProtectRecord(DataStream.GetAttributeAsBool("lockStructure", false));
            if (DataStream.GetAttributeAsBool("lockWindows", false)) Globals.WorkbookProtection.WindowProtect = new TWindowProtectRecord(true);
            if (DataStream.GetAttributeAsBool("lockRevision", false)) Globals.WorkbookProtection.Prot4Rev = new TProt4RevRecord(true);
            
            int PassHash = (int)DataStream.GetAttributeAsHex("workbookPassword", 0);
            if (PassHash != 0) Globals.WorkbookProtection.Password = new TPasswordRecord(PassHash);
            
            DataStream.FinishTag();
        }

        #endregion

        #region SST
        internal void ReadSST()
        {
            DataStream.SelectSST();
            while (DataStream.NextTag())
            {
                switch (DataStream.RecordName())
                {
                    case "sst":
                        ReadActualSST();
                        break;

                    default: //shouldn't happen
                        DataStream.GetXml();
                        break;
                }
            }
        }

        private void ReadActualSST()
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "si":
                        SST.LoadXml(TxSSTRecord.LoadFromXml(DataStream, xls));
                        break;
                    
                    case "extLst":
                    default:
                        SST.AddFutureStorage(new TFutureStorageRecord(DataStream.GetXml()));
                        break;
                }
            }
        }
        #endregion

        #region Connections
        internal void ReadConnections(TWorkbookGlobals Globals)
        {
            DataStream.SelectConnections();
            if (DataStream.Eof) return;
            DataStream.NextTag();

            Globals.XlsxConnections = DataStream.GetXml();
        }

        #endregion
        #region Pivot Tables
        internal void ReadPivotTables(TWorkbookGlobals Globals, TSheet aSheet)
        {
            List<Uri> PivotUris = DataStream.GetUrisForCurrentPartRelationship(TOpenXmlManager.PivotTableRelationshipType);
            foreach (Uri PivotUri in PivotUris)
            {
                ReadPivotTablePart(Globals, aSheet, PivotUri);
            }
        }

        private void ReadPivotTablePart(TWorkbookGlobals Globals, TSheet aSheet, Uri PivotUri)
        {
            DataStream.SelectMasterPart(PivotUri, TOpenXmlManager.MainNamespace);
            while (DataStream.NextTag())
            {
                switch (DataStream.RecordName())
                {
                    case "pivotTableDefinition":
                        TXlsxPivotTable Table = aSheet.XlsxPivotTables.Add();
                        ReadActualPivotTable(Globals, Table);
                        break;

                    default: //shouldn't happen
                        DataStream.GetXml();
                        break;
                }
            }
        }        

        private void ReadActualPivotTable(TWorkbookGlobals Globals, TXlsxPivotTable Table)
        {
            Table.ReadPivotTable(DataStream, Globals.XlsxPivotCache);
        }
        #endregion

        #region Styles
        internal void ReadStyles(TWorkbookGlobals Globals)
        {
            DataStream.SelectStyles();
            while (DataStream.NextTag())
            {
                switch (DataStream.RecordName())
                {
                    case "styleSheet":
                        ReadActualStyles(Globals);
                        break;

                    default: //shouldn't happen
                        DataStream.GetXml();
                        break;
                }
            }
        }

        private void ReadActualStyles(TWorkbookGlobals Globals)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "borders": ReadBorders(Globals); break;
                    case "cellStyles": ReadCellStyles(Globals); break;
                    case "cellStyleXfs": ReadCellStyleXfs(Globals); break;
                    case "cellXfs": ReadCellXfs(Globals); break;
                    case "colors": ReadColors(Globals); break;
                    case "dxfs": ReadDxfs(Globals); break;
                    case "fills": ReadFills(Globals); break;
                    case "fonts": ReadFonts(Globals); break;
                    case "numFmts": ReadNumFmts(Globals); break;
                    case "tableStyles": ReadTableStyles(Globals); break;

                    default:
                        Globals.AddStylesFutureStorage(new TFutureStorageRecord(DataStream.GetXml()));
                        break;
                }
            }
        }

        private void ReadCellStyles(TWorkbookGlobals Globals)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())                
                {
                    case "cellStyle":
                        Globals.Styles.Add(ReadStyleRecord());
                        break;

                    default:
                        DataStream.GetXml();
                        break;
                }
            }
        }

        private TStyleRecord ReadStyleRecord()
        {
            TStyleRecord Result = ReadStyleAttributes();

            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return Result; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return Result;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "extLst":
                    default:
                        Result.AddFutureStorage(new TFutureStorageRecord(DataStream.GetXml()));
                        break;
                }
            }

            return Result;
        }

        private TStyleRecord ReadStyleAttributes()
        {
            return TStyleRecord.CreateFromXlsx(
                DataStream.GetAttribute("name"),
                DataStream.GetAttributeAsInt("xfId", 0),
                DataStream.GetAttributeAsInt("builtinId", -1),
                DataStream.GetAttributeAsInt("iLevel", -1),
                DataStream.GetAttributeAsBool("hidden", false),
                DataStream.GetAttributeAsBool("customBuiltin", false));

        }

        private void ReadCellStyleXfs(TWorkbookGlobals Globals)
        {
            ReadXFS(Globals.StyleXF, true);
        }

        private void ReadCellXfs(TWorkbookGlobals Globals)
        {
            ReadXFS(Globals.CellXF, false);
        }

        private void ReadXFS(TXFRecordList XFList, bool IsStyle)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "xf":
                        XFList.Add(ReadXF(IsStyle));
                        break;

                    default:
                        DataStream.GetXml();
                        break;
                }
            }
        }

        private TXFRecord ReadXF(bool IsStyle)
        {
            TXFRecord Result = ReadXFAttributes(IsStyle, IsStyle);

            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return Result; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return Result;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "alignment":
                        ReadAlignment(Result);
                        DataStream.FinishTag();
                        break;

                    case "protection": 
                        ReadProtection(Result);
                        DataStream.FinishTag();
                        break;
                    
                    case "extLst":
                    default:
                        Result.AddFutureStorage(new TFutureStorageRecord(DataStream.GetXml()));
                        break;
                }
            }

            return Result;
        }

        private TXFRecord ReadXFAttributes(bool IsStyle, bool DefaultApply)
        {
            TLinkedStyle LinkedStyle = new TLinkedStyle();
            LinkedStyle.LinkedNumericFormat = GetLinkedStyle(DataStream, "applyNumberFormat", DefaultApply);
            LinkedStyle.LinkedFont = GetLinkedStyle(DataStream, "applyFont", DefaultApply);
            LinkedStyle.LinkedFill = GetLinkedStyle(DataStream, "applyFill", DefaultApply);
            LinkedStyle.LinkedBorder = GetLinkedStyle(DataStream, "applyBorder", DefaultApply);
            LinkedStyle.LinkedAlignment = GetLinkedStyle(DataStream, "applyAlignment", DefaultApply);
            LinkedStyle.LinkedProtection = GetLinkedStyle(DataStream, "applyProtection", DefaultApply);

            return new TXFRecord(
                FixFont4(DataStream.GetAttributeAsInt("fontId", 0)),
                DataStream.GetAttributeAsInt("numFmtId", 0),
                DataStream.GetAttributeAsInt("fillId", 0),
                DataStream.GetAttributeAsInt("borderId", 0),
                IsStyle,
                DataStream.GetAttributeAsInt("xfId", 0),
                THFlxAlignment.general,
                TVFlxAlignment.bottom,
                true,
                false,
                false,
                false,
                DataStream.GetAttributeAsBool("quotePrefix", false),
                0,
                0,
                DataStream.GetAttributeAsBool("pivotButton", false),
                LinkedStyle);
        }

        private bool GetLinkedStyle(TOpenXmlReader DataStream, string attName, bool DefaultApply)
        {
            //In normal cells, "Apply" means it is not linked to the style. In styles, "Apply" means it applies to the parent;
            if (DefaultApply) return DataStream.GetAttributeAsBool(attName, DefaultApply);
            return !DataStream.GetAttributeAsBool(attName, DefaultApply);
        }

        private static int FixFont4(int p)
        {
            if (p >= 4) return p + 1;
            return p;
        }

        private void ReadAlignment(TXFRecord Result)
        {
            switch (DataStream.GetAttribute("horizontal"))
            {
                case "left": Result.HAlignment = THFlxAlignment.left; break;
                case "center": Result.HAlignment = THFlxAlignment.center; break;
                case "right": Result.HAlignment = THFlxAlignment.right; break;
                case "fill": Result.HAlignment = THFlxAlignment.fill; break;
                case "justify": Result.HAlignment = THFlxAlignment.justify; break;
                case "centerContinuous": Result.HAlignment = THFlxAlignment.center_across_selection; break;
                case "distributed": Result.HAlignment = THFlxAlignment.distributed; break;
                case "general": Result.HAlignment = THFlxAlignment.general; break;
            }

            switch (DataStream.GetAttribute("vertical"))
            {
                case "top": Result.VAlignment = TVFlxAlignment.top; break;
                case "center": Result.VAlignment = TVFlxAlignment.center; break;
                case "bottom": Result.VAlignment = TVFlxAlignment.bottom; break;
                case "justify": Result.VAlignment = TVFlxAlignment.justify; break;
                case "distributed": Result.VAlignment = TVFlxAlignment.distributed; break;
            }

            Result.Rotation = (byte)DataStream.GetAttributeAsInt("textRotation", 0);
            Result.WrapText = DataStream.GetAttributeAsBool("wrapText", false);
            Result.Indent = (byte)DataStream.GetAttributeAsInt("indent", 0);
            Result.JustLast = DataStream.GetAttributeAsBool("justifyLastLine", false);
            Result.ShrinkToFit = DataStream.GetAttributeAsBool("shrinkToFit", false);
            Result.IReadOrder = (byte)DataStream.GetAttributeAsInt("readingOrder", 0);

        }

        private void ReadProtection(TXFRecord Result)
        {
            Result.Hidden = DataStream.GetAttributeAsBool("hidden", false);
            Result.Locked = DataStream.GetAttributeAsBool("locked", true);
            
        }   


        private void ReadColors(TWorkbookGlobals Globals)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "indexedColors":
                        ReadIndexedColors(Globals, DataStream);
                        break;

                    case "mruColors":
                        TFutureStorage.Add(ref Globals.MruColors, ReadMruColors(DataStream));
                        break;
                        
                    default:
                        DataStream.GetXml();
                        break;
                }
            }
        }

        private static void ReadIndexedColors(TWorkbookGlobals Globals, TOpenXmlReader DataStream)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;
            int ColorIndex = 0;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "rgbColor":
                        if (ColorIndex - 8 >= 0 && ColorIndex - 8 < XlsConsts.HighColorPaletteRange) //discard first 8 colors, they are for backward compatibility with who knows which version of Excel, probably Excel 1.5.
                        {
                            unchecked                            
                            {
                                long rgb = DataStream.GetAttributeAsHex("rgb", 0);
                                Globals.SetColorPalette(ColorIndex - 8, Color.FromArgb((int)(0xFF000000 | rgb)));
                            }                            
                        }
                        DataStream.FinishTag();
                        ColorIndex++;
                        break;

                    default:
                        DataStream.GetXml();
                        break;
                }
            }
        }

        private TFutureStorageRecord ReadMruColors(TOpenXmlReader DataStream)
        {
            return new TFutureStorageRecord(DataStream.GetXml());
        }


        private void ReadDxfs(TWorkbookGlobals Globals)
        {
            TFutureStorage.Add(ref Globals.DXF.Xlsx, new TFutureStorageRecord(DataStream.GetXml())); //PARSE
        }

        private void ReadFills(TWorkbookGlobals Globals)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "fill":
                        Globals.Patterns.AddForced(TXlsxFillReaderWriter.LoadFromXml(DataStream));
                        break;

                    default:
                        DataStream.GetXml();
                        break;
                }
            }
        }

        private void ReadNumFmts(TWorkbookGlobals Globals)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "numFmt":
                        Globals.Formats.SetBucket(new TFormatRecord(DataStream.GetAttribute("formatCode"), DataStream.GetAttributeAsInt("numFmtId", -1)));
                        DataStream.FinishTag();
                        break;

                    default:
                        DataStream.GetXml();
                        break;
                }
            }
        }

        private void ReadTableStyles(TWorkbookGlobals Globals)
        {
            TFutureStorage.Add(ref Globals.TableStyles.Xlsx, new TFutureStorageRecord(DataStream.GetXml())); //PARSE
        }

        private void ReadFonts(TWorkbookGlobals Globals)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "font":
                        TFlxFont fnt = new TFlxFont();
                        TXlsxFontReaderWriter.ReadFont(DataStream, fnt);
                        Globals.Fonts.Add(new TFontRecord(fnt));
                        break;

                    default:
                        DataStream.GetXml();
                        break;
                }
            }
        }

        private void ReadBorders(TWorkbookGlobals Globals)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "border":
                        Globals.Borders.AddForced(TXlsxBorderReaderWriter.LoadFromXml(DataStream));
                        break;

                    default:
                        DataStream.GetXml();
                        break;
                }
            }
        }

        #endregion

        #region Sheet
        internal void ReadSheet(string RelId, TSheetList Sheets, TWorkbookGlobals Globals)
        {
            if (RelId == null) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
            TSheet NewSheet = null;

            List<String> LegacyDrawingRels = new List<string>();
            List<String> LegacyDrawingHFRels = new List<string>();
            if (RelId.Trim().Length == 0)
            {
                NewSheet = Sheets[Sheets.Add(new TWorkSheet(Globals))]; //this can happen in xl4macros and xlsx files (not xlsm). 
            }
            else
            {
                DataStream.SelectSheet(RelId);
                while (DataStream.NextTag())
                {
                    switch (DataStream.CurrentPartContentType)
                    {
                        case TOpenXmlManager.ChartsheetContentType:
                            if (DataStream.RecordName() != "chartsheet") XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
                            NewSheet = Sheets[Sheets.Add(new TFlxChart(Globals, false))];
                            ReadChartSheet((TFlxChart)NewSheet);
                            break;

                        case TOpenXmlManager.WorksheetContentType:
                            if (DataStream.RecordName() != "worksheet") XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
                            NewSheet = Sheets[Sheets.Add(new TWorkSheet(Globals))];
                            ReadSheet(NewSheet, Sheets.Count, LegacyDrawingRels, LegacyDrawingHFRels);
                            break;

                        case TOpenXmlManager.DialogsheetContentType:
                            if (DataStream.RecordName() != "dialogsheet") XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
                            NewSheet = Sheets[Sheets.Add(new TWorkSheet(Globals))];
                            NewSheet.SheetGlobals.WsBool.Dialog = true;
                            NewSheet.Columns.AllowStandardWidth = false;
                            NewSheet.Columns.IsDialog = true;
                            ReadSheet(NewSheet, Sheets.Count, LegacyDrawingRels, LegacyDrawingHFRels);
                            break;

                        case TOpenXmlManager.IntMacrosheetContentType:
                        case TOpenXmlManager.MacrosheetContentType:
                            if (DataStream.RecordLocalName() != "macrosheet" || DataStream.RecordNamespace() != TOpenXmlManager.MacroNamespace) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
                            NewSheet = Sheets[Sheets.Add(new TMacroSheet(Globals))];
                            ReadSheet(NewSheet, Sheets.Count, LegacyDrawingRels, LegacyDrawingHFRels);
                            break;

                        default:
                            NewSheet = Sheets[Sheets.Add(new TFlxUnsupportedSheet(Globals))];
                            ReadUnsupportedSheet((TFlxUnsupportedSheet)NewSheet);
                            break;
                    }
                }
                ReadPivotTables(Globals, NewSheet); //must only run if a sheet was selected.
            }

            ReadHeaderImages(NewSheet.HeaderImages, LegacyDrawingHFRels);

            Dictionary<TOneCellRef, TObjectProperties> CommentProperties = GetCommentPropertiesAndLoadLegacyObjs(LegacyDrawingRels, Sheets.Count, NewSheet);
            if (CommentProperties.Count > 0)
            {
                ReadComments(NewSheet, CommentProperties);
            }

            NewSheet.EnsureRequiredRecords();
        }

        #endregion

        #region Chart Sheet
        private void ReadChartSheet(TFlxChart aChart)
        {
            //TXlsxChartReader ChartReader = new TXlsxChartReader(DataStream, xls);
            //ChartReader.ReadChart(aChart);
            aChart.AddFutureStorage(new TFutureStorageRecord(DataStream.GetXml()));//IMPLEMENT
        }
        #endregion

        #region UnspportedSheets
        private void ReadUnsupportedSheet(TFlxUnsupportedSheet sheet)
        {
            sheet.AddFutureStorage(new TFutureStorageRecord(DataStream.GetXml()));
        }
        #endregion

        #region Worksheet

        private void ReadSheet(TSheet WorkSheet, int WorkingSheet, List<String> LegacyDrawingRels, List<String> LegacyDrawingHFRels)
        {
            WorkSheet.InitXlsx();

            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "autoFilter": ReadAutoFilter(WorkSheet, WorkingSheet); break;
                    case "cellWatches": ReadCellWatches(WorkSheet); break;
                    case "colBreaks": ReadColBreaks(WorkSheet); break;
                    case "cols": ReadCols(WorkSheet); break;
                    case "conditionalFormatting": ReadConditionalFormatting(WorkSheet); break;
                    case "controls": ReadControls(WorkSheet); break;
                    case "customProperties": ReadCustomProperties(WorkSheet); break;
                    case "customSheetViews": ReadCustomSheetViews(WorkSheet); break;
                    case "dataConsolidate": ReadDataConsolidate(WorkSheet); break;
                    case "dataValidations": ReadDataValidations(WorkSheet); break;
                    case "dimension": ReadDimension(WorkSheet); break;
                    case "drawing": ReadDrawing(WorkSheet, WorkingSheet); break;
                    case "headerFooter": ReadHeaderFooter(WorkSheet); break;
                    case "hyperlinks": ReadHyperlinks(WorkSheet); break;
                    case "ignoredErrors": ReadIgnoredErrors(WorkSheet); break;
                    case "legacyDrawing": ReadLegacyDrawing(WorkSheet, LegacyDrawingRels); break;
                    case "legacyDrawingHF": ReadLegacyDrawingHF(WorkSheet, LegacyDrawingHFRels); break;
                    case "mergeCells": ReadMergeCells(WorkSheet); break;
                    case "oleObjects": ReadOleObjects(WorkSheet); break;
                    case "pageMargins": ReadPageMargins(WorkSheet); break;
                    case "pageSetup": ReadPageSetup(WorkSheet); break;
                    case "phoneticPr": ReadPhoneticPr(WorkSheet); break;
                    case "picture": ReadPicture(WorkSheet); break;
                    case "printOptions": ReadPrintOptions(WorkSheet); break;
                    case "protectedRanges": ReadProtectedRanges(WorkSheet); break;
                    case "rowBreaks": ReadRowBreaks(WorkSheet); break;
                    case "scenarios": ReadScenarios(WorkSheet); break;
                    case "sheetCalcPr": ReadSheetCalcPr(WorkSheet); break;
                    case "sheetData": ReadSheetData(WorkSheet, WorkingSheet); break;
                    case "sheetFormatPr": ReadSheetFormatPr(WorkSheet); break;
                    case "sheetPr": ReadSheetPr(WorkSheet); break;
                    case "sheetProtection": ReadSheetProtection(WorkSheet); break;
                    case "sheetViews": ReadSheetViews(WorkSheet); break;
                    case "smartTags": ReadSmartTags(WorkSheet); break;
                    case "sortState": ReadSortState(WorkSheet); break;
                    case "tableParts": ReadTableParts(WorkSheet); break;
                    case "webPublishItems": ReadWebPublishItems(WorkSheet); break;

                    case "extLst":
                    default: WorkSheet.AddFutureStorage(new TFutureStorageRecord(DataStream.GetXml()));
                        break;

                }
            }
        }

        private void ReadAutoFilter(TSheet WorkSheet, int SheetId)
        {
            TXlsCellRange cr = DataStream.GetAttributeAsRange("ref", false);
            if (cr != null)
            {
                WorkSheet.SetAutoFilter(SheetId - 1, cr.Top, cr.Left, cr.Right);
            }

            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "filterColumn":
                    case "sortState":
                    default:
                        TFutureStorageRecord AutoFilter = new TFutureStorageRecord(DataStream.GetXml());
                        TFutureStorage.Add(ref WorkSheet.SortAndFilter.AutoFilter.FutureStorage, AutoFilter);
                        break;
                }
            }
        }

        private void ReadCellWatches(TSheet WorkSheet)
        {
            DataStream.GetXml(); //IMPLEMENT
        }

        private void ReadColBreaks(TSheet WorkSheet)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "brk":
                        ReadBreak(WorkSheet.SheetGlobals.VPageBreaks);
                        DataStream.FinishTag();
                        break;

                    default:
                        DataStream.GetXml();
                        break;
                }
            }
        }

        private void ReadBreak(IPageBreakList PageBreakList)
        {
            if (DataStream.GetAttributeAsBool("man", false))
            {
                PageBreakList.AddBreak(DataStream.GetAttributeAsInt("id", 0), DataStream.GetAttributeAsInt("min", 0),
                    DataStream.GetAttributeAsInt("max", 0), false);
            }
        }


        private void ReadCols(TSheet WorkSheet)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "col": 
                        ReadCol(WorkSheet);
                        DataStream.FinishTag();
                        break;

                    default:
                        DataStream.GetXml();
                        break;
                }
            }
        }

        private void ReadCol(TSheet WorkSheet)
        {
            int FirstColumn = DataStream.GetAttributeAsInt("min", -1);
            if (FirstColumn <= 0) return; //this is required

            int LastColumn = DataStream.GetAttributeAsInt("max", -1);
            if (LastColumn <= 0) return; //this is required

            for (int i = FirstColumn; i <= LastColumn; i++)
                WorkSheet.Columns.Add(i - 1, new TColInfo(
                    ConvertWidth(DataStream.GetAttributeAsDouble("width", 8.43)),  //default val doesn't matter,as if it isn't set, customWidth will be set false.
                    DataStream.GetAttributeAsInt("style", 0),
                    BitOps.GetBool(
                           DataStream.GetAttributeAsBool("hidden", false),
                           DataStream.GetAttributeAsBool("customWidth", false) && DataStream.HasAttribute("width"),
                           DataStream.GetAttributeAsBool("bestFit", false),
                           DataStream.GetAttributeAsBool("phonetic", false)
                           )

                       | ((DataStream.GetAttributeAsInt("outlineLevel", 0) & 0x7) << 8)
                       | (DataStream.GetAttributeAsBool("collapsed", false)? 0x1000: 0)
                           
                   ,                          
                   true));

        }

        private static int ConvertWidth(double w)
        {
            return Convert.ToInt32(w * 256);
        }

        private void ReadConditionalFormatting(TSheet WorkSheet)
        {
            DataStream.GetXml(); //IMPLEMENT
        }

        private void ReadControls(TSheet WorkSheet)
        {
            DataStream.GetXml(); //IMPLEMENT
        }

        private void ReadCustomProperties(TSheet WorkSheet)
        {
            DataStream.GetXml(); //IMPLEMENT
        }

        private void ReadCustomSheetViews(TSheet WorkSheet)
        {
            DataStream.GetXml(); //IMPLEMENT
        }

        private void ReadDataConsolidate(TSheet WorkSheet)
        {
            DataStream.GetXml(); //IMPLEMENT
        }

        private void ReadDataValidations(TSheet WorkSheet)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();

            WorkSheet.DataValidation.DisablePrompts = DataStream.GetAttributeAsBool("disablePrompts", false);
            WorkSheet.DataValidation.xWindow = DataStream.GetAttributeAsInt("xWindow", 0);
            WorkSheet.DataValidation.yWindow = DataStream.GetAttributeAsInt("yWindow", 0);

            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "dataValidation":
                        ReadDataValidation(WorkSheet);
                        break;

                    default:
                        DataStream.GetXml();
                        break;
                }
            }
        }

        private void ReadDataValidation(TSheet WorkSheet)
        {
            TDataValidationInfo dv = new TDataValidationInfo();

            dv.ValidationType = GetValidationType(DataStream.GetAttribute("type"));
            dv.ErrorIcon = GetValidationIcon(DataStream.GetAttribute("errorStyle"));

            dv.ImeMode = GetImeMode(DataStream.GetAttribute("imeMode"));

            dv.Condition = GetValidationCondition(DataStream.GetAttribute("operator"));

            dv.IgnoreEmptyCells = DataStream.GetAttributeAsBool("allowBlank", false);
            dv.InCellDropDown = !DataStream.GetAttributeAsBool("showDropDown", false);  //this is so wrong...
            dv.ShowInputBox = DataStream.GetAttributeAsBool("showInputMessage", false);
            dv.ShowErrorBox = DataStream.GetAttributeAsBool("showErrorMessage", false);
            dv.ErrorBoxCaption = DataStream.GetAttribute("errorTitle");
            dv.ErrorBoxText = DataStream.GetAttribute("error");
            dv.InputBoxCaption = DataStream.GetAttribute("promptTitle");
            dv.InputBoxText = DataStream.GetAttribute("prompt");

            TXlsCellRange[] ranges = DataStream.GetAttributeAsSeriesOfRanges("sqref", false);

            ReadDataValFormulas(dv);

            if (dv.ExplicitList) MakeDvExplicit(dv);

            foreach (TXlsCellRange range in ranges)
            {
                WorkSheet.DataValidation.AddRange(range, dv, WorkSheet.Cells.CellList, false, true);                
            }
        }

        private void MakeDvExplicit(TDataValidationInfo dv)
        {
            //this is a silly substitution. Quoted commas are delimitors too. No way in Excel to have a "," inside a list :(
            if (!string.IsNullOrEmpty(dv.FirstFormula))
            {
                dv.FirstFormula = dv.FirstFormula.Replace(',', (char)0);
            }
        }

        private void ReadDataValFormulas(TDataValidationInfo dv)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();

            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "formula1":
                        dv.FirstFormula = DataStream.ReadValueAsString();
                        break;

                    case "formula2":
                        dv.SecondFormula = DataStream.ReadValueAsString();
                        break;

                    default:
                        DataStream.GetXml();
                        break;
                }
            }

            //before formula1 is modified.
            dv.ExplicitList = dv.ValidationType == TDataValidationDataType.List && !string.IsNullOrEmpty(dv.FirstFormula) && dv.FirstFormula.StartsWith(TFormulaMessages.TokenString(TFormulaToken.fmStr));

            if (!string.IsNullOrEmpty(dv.FirstFormula)) dv.FirstFormula = TFormulaMessages.TokenString(TFormulaToken.fmStartFormula) + dv.FirstFormula;
            if (!string.IsNullOrEmpty(dv.SecondFormula)) dv.SecondFormula = TFormulaMessages.TokenString(TFormulaToken.fmStartFormula) + dv.SecondFormula;
        }

        private TDataValidationConditionType GetValidationCondition(string vc)
        {
            switch (vc)
            {
                case "between": return TDataValidationConditionType.Between;
                case "notBetween": return TDataValidationConditionType.NotBetween;
                case "equal": return TDataValidationConditionType.EqualTo;
                case "notEqual": return TDataValidationConditionType.NotEqualTo;
                case "lessThan": return TDataValidationConditionType.LessThan;
                case "lessThanOrEqual": return TDataValidationConditionType.LessThanOrEqualTo;
                case "greaterThan": return TDataValidationConditionType.GreaterThan;
                case "greaterThanOrEqual": return TDataValidationConditionType.GreaterThanOrEqualTo;
            }

            return TDataValidationConditionType.Between;
        }

        private TDataValidationIcon GetValidationIcon(string vi)
        {
            switch (vi)
            {
                case "stop": return TDataValidationIcon.Stop;
                case "warning": return TDataValidationIcon.Warning;
                case "information": return TDataValidationIcon.Information;
            }

            return TDataValidationIcon.Stop;
        }

        private TDataValidationDataType GetValidationType(string vt)
        {
            switch (vt)
            {
                case "none": return TDataValidationDataType.AnyValue;
                case "whole": return TDataValidationDataType.WholeNumber;
                case "decimal": return TDataValidationDataType.Decimal;
                case "list": return TDataValidationDataType.List;
                case "date": return TDataValidationDataType.Date;
                case "time": return TDataValidationDataType.Time;
                case "textLength": return TDataValidationDataType.TextLenght;
                case "custom": return TDataValidationDataType.Custom;
            }

            return TDataValidationDataType.AnyValue;
        }

        private TDataValidationImeMode GetImeMode(string im)
        {
            switch (im)
            {
                case "noControl": return TDataValidationImeMode.NoControl;
                case "off": return TDataValidationImeMode.Off;
                case "on": return TDataValidationImeMode.On;
                case "disabled": return TDataValidationImeMode.Disabled;
                case "hiragana": return TDataValidationImeMode.Hiragana;
                case "fullKatakana": return TDataValidationImeMode.FullKatakana;
                case "halfKatakana": return TDataValidationImeMode.HalfKatakana;
                case "fullAlpha": return TDataValidationImeMode.FullAlpha;
                case "halfAlpha": return TDataValidationImeMode.HalfAlpha;
                case "fullHangul": return TDataValidationImeMode.FullHangul;
                case "halfHangul": return TDataValidationImeMode.HalfHangul;
            }

            return TDataValidationImeMode.NoControl;
        }


        private void ReadDimension(TSheet WorkSheet)
        {
            DataStream.GetXml(); //We won't use it.
        }

        private void ReadDrawing(TSheet WorkSheet, int WorkingSheet)
        {
            DrawingLoader.ReadDrawing(DataStream.GetRelationship("id"), WorkSheet, WorkingSheet);
            DataStream.FinishTag();
        }

        private void ReadHeaderFooter(TSheet WorkSheet)
        {
            WorkSheet.PageSetup.HeaderAndFooter.DiffEvenPages = DataStream.GetAttributeAsBool("differentOddEven", false);
            WorkSheet.PageSetup.HeaderAndFooter.DiffFirstPage = DataStream.GetAttributeAsBool("differentFirst", false);
            WorkSheet.PageSetup.HeaderAndFooter.ScaleWithDoc = DataStream.GetAttributeAsBool("scaleWithDoc", true);
            WorkSheet.PageSetup.HeaderAndFooter.AlignMargins = DataStream.GetAttributeAsBool("alignWithMargins", true);

            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "oddHeader":
                        WorkSheet.PageSetup.HeaderAndFooter.DefaultHeader = DataStream.ReadValueAsString();
                        break;

                    case "oddFooter":
                        WorkSheet.PageSetup.HeaderAndFooter.DefaultFooter = DataStream.ReadValueAsString();
                        break;

                    case "evenHeader":
                        WorkSheet.PageSetup.HeaderAndFooter.EvenHeader = DataStream.ReadValueAsString();
                        break;

                    case "evenFooter":
                        WorkSheet.PageSetup.HeaderAndFooter.EvenFooter = DataStream.ReadValueAsString();
                        break;

                    case "firstHeader":
                        WorkSheet.PageSetup.HeaderAndFooter.FirstHeader = DataStream.ReadValueAsString();
                        break;

                    case "firstFooter":
                        WorkSheet.PageSetup.HeaderAndFooter.FirstFooter = DataStream.ReadValueAsString();
                        break;

                    default:
                        DataStream.GetXml();
                        break;
                }
            }
        }

        private void ReadHyperlinks(TSheet WorkSheet)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "hyperlink":
                        WorkSheet.HLinks.Add(ReadHyperlink());
                        break;

                    default:
                        DataStream.GetXml();
                        break;
                }
            }
            
        }

        private THyperLinkType GetHyperLinkType(THyperLink HLink, bool UriAbsolute)
        {
            if (string.IsNullOrEmpty(HLink.Text)) return THyperLinkType.CurrentWorkbook;

            const string UNCStart = "file:///";
            if (HLink.Text.StartsWith(UNCStart, StringComparison.InvariantCultureIgnoreCase))
            {
                HLink.Text = HLink.Text.Remove(0, UNCStart.Length);
                return THyperLinkType.UNC;
            }

            if (!UriAbsolute) return THyperLinkType.LocalFile;

            return THyperLinkType.URL;
        }

        private THLinkRecord ReadHyperlink()
        {
            TXlsCellRange range = DataStream.GetAttributeAsRange("ref", false);

            THyperLink HLink = new THyperLink();

            bool UriAbsolute = true; 
            HLink.Text = null;

            string Rid = DataStream.GetRelationship("id");
            if (Rid != null)
            {
                string HLinkUri = DataStream.GetExternalLink(Rid); 
                HLink.Text = HLinkUri;
                UriAbsolute = IsAbsoluteUri(HLinkUri);
            }
            HLink.TextMark = DataStream.GetAttribute("location");
            HLink.Hint = DataStream.GetAttribute("tooltip");
            HLink.Description = DataStream.GetAttribute("display");

            HLink.LinkType = GetHyperLinkType(HLink, UriAbsolute);

            DataStream.FinishTag();
            return THLinkRecord.CreateNew(range, HLink);
        }

        private static bool IsAbsoluteUri(string s)
        {
            return s.IndexOf(Uri.SchemeDelimiter) > 0 || s.StartsWith(Uri.UriSchemeMailto + ":");
        }

        private void ReadIgnoredErrors(TSheet WorkSheet)
        {
            DataStream.GetXml(); //IMPLEMENT
        }

        private void ReadLegacyDrawing(TSheet WorkSheet, List<String> LegacyDrawingRels) //We will only read comments from here.
        {
            LegacyDrawingRels.Add(DataStream.GetRelationship("id")); //actually there should be only one according to docs, but we play safe.
            DataStream.FinishTag();
        }

        private void ReadLegacyDrawingHF(TSheet WorkSheet, List<String> LegacyDrawingHFRels)
        {
            LegacyDrawingHFRels.Add(DataStream.GetRelationship("id")); //actually there should be only one according to docs, but we play safe.
            DataStream.FinishTag();
        }

        private void ReadMergeCells(TSheet WorkSheet)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "mergeCell":
                        TXlsCellRange range = DataStream.GetAttributeAsRange("ref", true);
                        if (range != null)
                        {
                            TMergedCells Mc = new TMergedCells();
                            //no need for premerge here, we will assume cells are ok in a saved file.
                            Mc.MergeCells(range.Top - 1, range.Left - 1, range.Bottom -1, range.Right -1);
                            WorkSheet.MergedCells.Add(Mc);
                            
                        }
                        DataStream.FinishTag();
                        break;

                    default:
                        DataStream.GetXml();
                        break;
                }
            }
        }

        private void ReadOleObjects(TSheet WorkSheet)
        {
            DataStream.GetXml(); //IMPLEMENT
        }

        private void ReadPageMargins(TSheet WorkSheet)
        {
            WorkSheet.PageSetup.LeftMargin = DataStream.GetAttributeAsDouble("left", 0);
            WorkSheet.PageSetup.RightMargin = DataStream.GetAttributeAsDouble("right", 0);
            WorkSheet.PageSetup.TopMargin = DataStream.GetAttributeAsDouble("top", 0);
            WorkSheet.PageSetup.BottomMargin = DataStream.GetAttributeAsDouble("bottom", 0);
            WorkSheet.PageSetup.Setup.HeaderMargin = DataStream.GetAttributeAsDouble("header", 0);
            WorkSheet.PageSetup.Setup.FooterMargin = DataStream.GetAttributeAsDouble("footer", 0);

            DataStream.FinishTag();

        }

        private void ReadPageSetup(TSheet WorkSheet)
        {
            WorkSheet.PageSetup.Setup.PaperSize = DataStream.GetAttributeAsInt("paperSize", 1);

            //DataStream.GetAttribute("paperHeight");
            //DataStream.GetAttribute("paperWidth");

            WorkSheet.PageSetup.Setup.Scale = DataStream.GetAttributeAsInt("scale", 100);
            unchecked
            {
                WorkSheet.PageSetup.Setup.PageStart = (Int16)DataStream.GetAttributeAsLong("firstPageNumber", 1);
            }
            WorkSheet.PageSetup.Setup.FitWidth = DataStream.GetAttributeAsInt("fitToWidth", 1);
            WorkSheet.PageSetup.Setup.FitHeight = DataStream.GetAttributeAsInt("fitToHeight", 1);

            WorkSheet.PageSetup.Setup.SetPrintOptions(
                BitOps.GetBool(
                DataStream.GetAttribute("pageOrder") == "overThenDown",
                DataStream.GetAttribute("orientation") == "portrait",

                !DataStream.GetAttributeAsBool("usePrinterDefaults", true),

                DataStream.GetAttributeAsBool("blackAndWhite", false),
                DataStream.GetAttributeAsBool("draft", false),
                DataStream.GetAttribute("cellComments") != "none" && !string.IsNullOrEmpty(DataStream.GetAttribute("cellComments")),
                DataStream.GetAttribute("orientation") == "default",

                DataStream.GetAttributeAsBool("useFirstPageNumber", false),
                
                false,
                DataStream.GetAttribute("cellComments") == "atEnd",

                GetPrintErrors("errors", true),
                GetPrintErrors("errors", false)));

            long hdpi = DataStream.GetAttributeAsLong("horizontalDpi", 600);
            if (hdpi < int.MaxValue && hdpi >= 0) WorkSheet.PageSetup.Setup.HPrintRes = (int)hdpi;
           
            long vdpi = DataStream.GetAttributeAsLong("verticalDpi", 600);
            if (vdpi < int.MaxValue && vdpi >= 0) WorkSheet.PageSetup.Setup.VPrintRes = (int)vdpi;

            WorkSheet.PageSetup.Setup.Copies = DataStream.GetAttributeAsInt("copies", 1);


            string Ct; string Fn;
            byte[] Pls = DataStream.GetRelationshipData("id", 2, 2, out Ct, out Fn); 
            if (Pls != null) 
            {
                WorkSheet.PageSetup.Pls = TPlsRecord.FromLongData(Pls);
            }

            DataStream.FinishTag();
        }

        private bool GetPrintErrors(string errTag, bool firstBit)
        {
            switch(DataStream.GetAttribute(errTag))
            {
                case "displayed": return false;
                case "blank": return firstBit;
                case "dash": return !firstBit;
                case "NA": return true;
            }
            return false;
        }

        private void ReadPhoneticPr(TSheet WorkSheet)
        {
            DataStream.GetXml(); //IMPLEMENT
        }

        private void ReadPicture(TSheet WorkSheet)
        {
            DataStream.GetXml(); //IMPLEMENT
        }

        private void ReadPrintOptions(TSheet WorkSheet)
        {
            WorkSheet.PageSetup.HCenter = DataStream.GetAttributeAsBool("horizontalCentered", false);
            WorkSheet.PageSetup.VCenter = DataStream.GetAttributeAsBool("verticalCentered", false);
            WorkSheet.SheetGlobals.PrintHeaders = DataStream.GetAttributeAsBool("headings", false);
            WorkSheet.SheetGlobals.PrintGridLines = DataStream.GetAttributeAsBool("gridLines", false);
            WorkSheet.SheetGlobals.GridSet = DataStream.GetAttributeAsBool("gridLinesSet", true);
        
            DataStream.FinishTag();
        }

        private void ReadProtectedRanges(TSheet WorkSheet)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "protectedRange":
                        WorkSheet.SheetGlobals.ProtectedRanges.Add(ReadProtectedRange());
                        break;

                    default:
                        DataStream.GetXml();
                        break;
                }
            }
        }

        private TProtectedRange ReadProtectedRange()
        {
            TProtectedRange Result = new TProtectedRange();
            Result.Ranges = DataStream.GetAttributeAsSeriesOfRanges("sqref", false);
            Result.Name = DataStream.GetAttribute("name");
            Result.Password = DataStream.GetAttribute("password");
            DataStream.FinishTag();
            return Result;
        }

        private void ReadRowBreaks(TSheet WorkSheet)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "brk":
                        ReadBreak(WorkSheet.SheetGlobals.HPageBreaks);
                        DataStream.FinishTag();
                        break;

                    default:
                        DataStream.GetXml();
                        break;
                }
            }
        }

        private void ReadScenarios(TSheet WorkSheet)
        {
            DataStream.GetXml(); //IMPLEMENT
        }

        private void ReadSheetCalcPr(TSheet WorkSheet)
        {
            DataStream.GetXml(); //IMPLEMENT
        }

        private void ReadSheetData(TSheet WorkSheet, int WorkingSheet)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            Dictionary<int, TSharedFmlaData> SharedFormulas = new Dictionary<int, TSharedFmlaData>();
            int CurrentRow = -1;
            bool HasMultiCellArrayFmlas = false;

            //object rowTag = DataStream.xlReader.NameTable.Add("row"); Using a Nametable really doesn't change anything, and we must chnge the pattern not to use "switch"
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "row":
                        TXlsxCellReader.ReadRow(xls, SST, VirtualReader, DataStream, WorkSheet, WorkingSheet, ref CurrentRow, SharedFormulas, ref HasMultiCellArrayFmlas, xls.OptionsDates1904);
                        break;

                    default:
                        DataStream.GetXml();
                        break;
                }
            }
        }

        private void ReadSheetFormatPr(TSheet WorkSheet)
        {

            WorkSheet.SheetGlobals.DefRowHeight = new TDefaultRowHeightRecord();
            double BaseWidth = DataStream.GetAttributeAsDouble("baseColWidth", 8);
            double FontSize = ExcelMetrics.GetFont0Width(xls);
            WorkSheet.Columns.DefColWidth = (int)(DataStream.GetAttributeAsDouble("defaultColWidth", GetDefColWidth(FontSize, BaseWidth)) * 256);

            //"Right" way would be to use this code:
            //if (DataStream.GetAttributeAsBool("customHeight", false))
            //{
            //    WorkSheet.DefRowHeight = (int)(DataStream.GetAttributeAsDouble("defaultRowHeight", 15) * 20);
            //}
            //else
            //{
            //    WorkSheet.DefRowHeight = ExcelMetrics.GetRowHeightInPixels(xls.GetDefaultFont);
            //}
            //But as always, we can't calcualte Exactly the row height Excel does, so we will use their value.
            //This can be problematic with hand-crafted files.

            int DefH = DataStream.HasAttribute("defaultRowHeight") ? 15 : ExcelMetrics.GetRowHeightInPixels(xls.GetDefaultFont); //we don't wan't to call ExcelMetrics unless really necessary.
            WorkSheet.DefRowHeight = (int)(DataStream.GetAttributeAsDouble("defaultRowHeight", DefH) * 20);

            //Won't load Guts as it will be calculated when saving.

            WorkSheet.SheetGlobals.DefRowHeight.Flags = BitOps.GetBool(
                DataStream.GetAttributeAsBool("customHeight", false),
                DataStream.GetAttributeAsBool("zeroHeight", false),
                DataStream.GetAttributeAsBool("thickTop", false),
                DataStream.GetAttributeAsBool("thickBottom", false)); //here is bottom, in rows is both :(

            DataStream.FinishTag();
        }

        private static double GetDefColWidth(double FontSize, double BaseWidth)
        {
            //Here we have some function f(BaseWidth, FontSize) that is not really documented anywhere,
            //and that seems to be using modulo 8 integer math.
            //While I couldn't figure out a single formula, the code below should provide correct result in all cases.
            //See columnwidth_study.xls for more information on how this was deduced.

            //while those are ints, converting them from doubles isn't probably worth.
            double K0 = 8 * (Math.Truncate((FontSize + 3) / 16) + 1);
            double i = 4 - ((FontSize + 4) % 8);
            double totali = i * (BaseWidth % 8);

            double Max = 10 + Math.Truncate((FontSize - 1) / 4) * 2;
            double Min = Max - 7;

            double K = K0 + totali;
            while (K < Min) K += 8;
            while (K > Max) K -= 8;

            return Math.Truncate((BaseWidth * FontSize + K) / FontSize * 256) / 256.0;  
        }

        private void ReadSheetPr(TSheet WorkSheet)
        {
            string Sync = DataStream.GetAttribute("syncRef");
            if (Sync != null)
            {
                TCellAddress addr = new TCellAddress(Sync);
                WorkSheet.SheetGlobals.Sync = new TSyncRecord(addr.Row - 1, addr.Col - 1);
            }

            WorkSheet.SheetGlobals.WsBool.SyncHoriz = DataStream.GetAttributeAsBool("syncHorizontal", false);
            WorkSheet.SheetGlobals.WsBool.SyncVert = DataStream.GetAttributeAsBool("syncVertical", false);

            WorkSheet.SheetGlobals.WsBool.AltFormulaEntry = DataStream.GetAttributeAsBool("transitionEntry", false);
            WorkSheet.SheetGlobals.WsBool.AltExprEval = DataStream.GetAttributeAsBool("transitionEvaluation", false);

            if (!DataStream.GetAttributeAsBool("published", true))
            {
                if (WorkSheet.SheetExt == null) WorkSheet.SheetExt = new TSheetExtRecord(TExcelColor.Automatic);
                WorkSheet.SheetExt.NotPublished = true;
            }

            if (!DataStream.GetAttributeAsBool("enableFormatConditionsCalculation", true))
            {
                if (WorkSheet.SheetExt == null) WorkSheet.SheetExt = new TSheetExtRecord(TExcelColor.Automatic);
                WorkSheet.SheetExt.CondFmtCalc = false;
            }

            WorkSheet.CodeName = DataStream.GetAttribute("codeName");

            WorkSheet.SortAndFilter.FilterMode = DataStream.GetAttributeAsBool("filterMode", false);

            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "tabColor": ReadSheetColor(WorkSheet); break;
                    case "outlinePr": ReadOutlinePr(WorkSheet); break;
                    case "pageSetUpPr": ReadSetupPr(WorkSheet); break;

                    case "extLst":
                    default:
                        DataStream.GetXml();
                        break;
                }
            }
        }

        private void ReadSheetColor(TSheet WorkSheet)
        {
            if (WorkSheet.SheetExt == null) WorkSheet.SheetExt = new TSheetExtRecord(TExcelColor.Automatic);
            WorkSheet.SheetExt.SheetColor = TXlsxColorReaderWriter.GetColor(DataStream);
            DataStream.FinishTag();
        }

        private void ReadSetupPr(TSheet WorkSheet)
        {
            WorkSheet.SheetGlobals.WsBool.ShowAutoBreaks = DataStream.GetAttributeAsBool("autoPageBreaks", true);
            WorkSheet.SheetGlobals.WsBool.FitToPage = DataStream.GetAttributeAsBool("fitToPage", false);
            DataStream.FinishTag();
        }

        private void ReadOutlinePr(TSheet WorkSheet)
        {
            WorkSheet.SheetGlobals.WsBool.ApplyStyles = DataStream.GetAttributeAsBool("applyStyles", false);
            WorkSheet.SheetGlobals.WsBool.RowSumsBelow = DataStream.GetAttributeAsBool("summaryBelow", true);
            WorkSheet.SheetGlobals.WsBool.ColSumsRight = DataStream.GetAttributeAsBool("summaryRight", true);
            WorkSheet.SheetGlobals.WsBool.DspGuts = DataStream.GetAttributeAsBool("showOutlineSymbols", true);
            DataStream.FinishTag();
        }



        private void ReadSheetProtection(TSheet WorkSheet)
        {
            TSheetProtectionOptions spo = new TSheetProtectionOptions();
            /*<xsd:attribute name="algorithmName" type="s:ST_Xstring" use="optional"/>  2913 
     <xsd:attribute name="hashValue" type="xsd:base64Binary" use="optional"/>  2914 
     <xsd:attribute name="saltValue" type="xsd:base64Binary" use="optional"/>  2915 
     <xsd:attribute name="spinCount" type="xsd:unsignedInt" use="optional"/>  2916 */

            int PassHash = (int)DataStream.GetAttributeAsHex("password", 0);
            if (PassHash != 0) WorkSheet.SheetProtection.Password = new TPasswordRecord(PassHash);

            spo.Contents = DataStream.GetAttributeAsBool("sheet", false);
            spo.Objects = DataStream.GetAttributeAsBool("objects", false);
            spo.Scenarios = DataStream.GetAttributeAsBool("scenarios", false);
            spo.CellFormatting = !DataStream.GetAttributeAsBool("formatCells", true);
            spo.ColumnFormatting = !DataStream.GetAttributeAsBool("formatColumns", true);
            spo.RowFormatting = !DataStream.GetAttributeAsBool("formatRows", true);
            spo.InsertColumns = !DataStream.GetAttributeAsBool("insertColumns", true);
            spo.InsertRows = !DataStream.GetAttributeAsBool("insertRows", true);
            spo.InsertHyperlinks = !DataStream.GetAttributeAsBool("insertHyperlinks", true);
            spo.DeleteColumns = !DataStream.GetAttributeAsBool("deleteColumns", true);
            spo.DeleteRows = !DataStream.GetAttributeAsBool("deleteRows", true);
            spo.SelectLockedCells = !DataStream.GetAttributeAsBool("selectLockedCells", false);
            spo.SortCellRange = !DataStream.GetAttributeAsBool("sort", true);
            spo.EditAutoFilters = !DataStream.GetAttributeAsBool("autoFilter", true);
            spo.EditPivotTables = !DataStream.GetAttributeAsBool("pivotTables", true);
            spo.SelectUnlockedCells = !DataStream.GetAttributeAsBool("selectUnlockedCells", false);

            WorkSheet.SetSheetProtectionOptions(spo);

            DataStream.FinishTag();
        }

        private void ReadSheetViews(TSheet WorkSheet)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "sheetView":
                        ReadSheetView(WorkSheet.Window);
                        break;

                    default:
                        WorkSheet.Window.AddFutureStorage(new TFutureStorageRecord(DataStream.GetXml())); break;
                }
            }         
        }

        private void ReadSheetView(TWindow Window)
        {
            Window.Window2 = new TWindow2Record(false);

            //before options, since SetGridLinesColor will change it.
            Window.Window2.SetGridLinesColor(TExcelColor.FromBiff8ColorIndex(DataStream.GetAttributeAsInt("colorId", 64)));

            Window.Window2.RawOptions = BitOps.GetBool(
                        DataStream.GetAttributeAsBool("showFormulas", false),
                        DataStream.GetAttributeAsBool("showGridLines", true),
                        DataStream.GetAttributeAsBool("showRowColHeaders", true),
                        false, //frozen
                        DataStream.GetAttributeAsBool("showZeros", true),
                        DataStream.GetAttributeAsBool("defaultGridColor", true),
                        DataStream.GetAttributeAsBool("rightToLeft", false),
                        DataStream.GetAttributeAsBool("showOutlineSymbols", true),
                        false, //frozennosplit
                        DataStream.GetAttributeAsBool("tabSelected", false),
                        false, //paged,
                        DataStream.GetAttribute("view") == "pageBreakPreview");

            TCellAddress TopLeftCell = DataStream.GetAttributeAsAddress("topLeftCell");
            if (TopLeftCell != null)
            {
                Window.Window2.FirstRow = TopLeftCell.Row - 1;
                Window.Window2.FirstCol = TopLeftCell.Col - 1;
            }

            //Window.Window2.WorkbookViewId = DataStream.GetAttributeAsInt("workbookViewId", 0); //For more than one Window2.

            byte[] PlvData = new byte[16]; //2007 always includes this record, we will too when loading xlsx}
            BitOps.SetWord(PlvData, 0, 0x088B);            
            BitOps.SetWord(PlvData, 12, DataStream.GetAttributeAsInt("zoomScaleSheetLayoutView", 0));
            PlvData[14] = (byte)BitOps.GetBool(
                        DataStream.GetAttribute("view") == "pageLayout",
                        DataStream.GetAttributeAsBool("showRuler", true),
                        !DataStream.GetAttributeAsBool("showWhiteSpace", true)
                        );
            Window.Plv = new TPlvRecord((int)xlr.PLV, PlvData);
       
            //Window.Window2.windowProtection = DataStream.GetAttributeAsBool("windowProtection", false);

            if (DataStream.GetAttribute("zoomScale") != null)
            {
                Window.Scl = new TSCLRecord(DataStream.GetAttributeAsInt("zoomScale", 100));
            }

            Window.Window2.ScaleInNormalView = DataStream.GetAttributeAsInt("zoomScaleNormal", 0);
            Window.Window2.ScaleInPageBreakPreview = DataStream.GetAttributeAsInt("zoomScalePageLayoutView", 0);


            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "pane": ReadPane(Window); break;
                    case "selection": ReadSelection(Window); break;
                    case "pivotSelection": ReadPivotSelection(Window); break;

                    default:
                        Window.AddFutureStorage(new TFutureStorageRecord(DataStream.GetXml())); break;
                }
            }
        }

        private void ReadPane(TWindow Window)
        {
            Window.Pane = new TPaneRecord();
            Window.Pane.ColSplit = (int)DataStream.GetAttributeAsDouble("xSplit", 0);
            Window.Pane.RowSplit = (int)DataStream.GetAttributeAsDouble("ySplit", 0);

            TCellAddress TopLeftCell = DataStream.GetAttributeAsAddress("topLeftCell");
            if (TopLeftCell != null)
            {
                Window.Pane.FirstVisibleRow = TopLeftCell.Row - 1;
                Window.Pane.FirstVisibleCol = TopLeftCell.Col - 1;
            }

            Window.Pane.ActivePane = (int)GetPanePosition(DataStream.GetAttribute("activePane"));

            switch (DataStream.GetAttribute("state"))
            {
                case "frozen": Window.Window2.IsFrozen = true; Window.Window2.IsFrozenButNoSplit = true; break;
                case "frozenSplit": Window.Window2.IsFrozen = true; break;
                case "split": 
                default: 
                       break;
            }

            DataStream.FinishTag();
        }

        private void ReadSelection(TWindow Window)
        {
            if (Window.Selection == null) Window.Selection = new TSheetSelection();
            int ActiveRow = -1;
            int ActiveCol = -1;
            TCellAddress addr = DataStream.GetAttributeAsAddress("activeCell");
            if (addr != null)
            {
                ActiveRow = addr.Row - 1;
                ActiveCol = addr.Col - 1;
            }
            Window.Selection.Select(GetPanePosition(DataStream.GetAttribute("pane")), 
                DataStream.GetAttributeAsSeriesOfRanges("sqref", false), ActiveRow, ActiveCol,
                DataStream.GetAttributeAsInt("activeCellId", 0));

            DataStream.FinishTag();
        }

        private TPanePosition GetPanePosition(string p)
        {
            switch (p)
            {
                case "bottomRight": return TPanePosition.LowerRight;
                case "topRight": return TPanePosition.UpperRight;
                case "bottomLeft": return TPanePosition.LowerLeft;
                case "topLeft":
                default:
                    return TPanePosition.UpperLeft;
            }
        }

        private void ReadPivotSelection(TWindow Window)
        {
            DataStream.GetXml(); //IMPLEMENT
        }

        private void ReadSmartTags(TSheet WorkSheet)
        {
            DataStream.GetXml(); //IMPLEMENT
        }

        private void ReadSortState(TSheet WorkSheet)
        {
            DataStream.GetXml(); //IMPLEMENT
        }

        private void ReadTableParts(TSheet WorkSheet)
        {
            DataStream.GetXml(); //IMPLEMENT
        }

        private void ReadWebPublishItems(TSheet WorkSheet)
        {
            DataStream.GetXml(); //IMPLEMENT
        }

        #endregion

        #region Legacy Drawing
        private Dictionary<TOneCellRef, TObjectProperties> GetCommentPropertiesAndLoadLegacyObjs(List<string> LegacyDrawingRels, int WorkingSheet, TSheet aSheet)
        {
            Dictionary<TOneCellRef, TObjectProperties> Result = new Dictionary<TOneCellRef, TObjectProperties>();

            if (LegacyDrawingRels.Count == 0) return Result;

            foreach (string ld in LegacyDrawingRels)
            {
                DataStream.SelectFromCurrentPartAndPush(ld, null, true);
                while (DataStream.NextTag())
                {
                    switch (DataStream.RecordName())
                    {
                        case ":xml": ReadLegacyDrawingXml(Result, WorkingSheet, aSheet); break;

                        default:
                            DataStream.GetXml();
                            break;
                    }
                }

                DataStream.PopPart();
            }

            return Result;
        }

        private void ReadHeaderImages(TDrawing Drawing, List<string> LegacyDrawingHFRels)
        {
            if (LegacyDrawingHFRels.Count == 0) return;

            foreach (string ld in LegacyDrawingHFRels)
            {
                DataStream.SelectFromCurrentPartAndPush(ld, null, true);
                while (DataStream.NextTag())
                {
                    switch (DataStream.RecordName())
                    {
                        case ":xml": ReadLegacyDrawingHFXml(Drawing); break;

                        default:
                            DataStream.GetXml();
                            break;
                    }
                }

                DataStream.PopPart();
            }
        }

        private void ReadLegacyDrawingXml(Dictionary<TOneCellRef, TObjectProperties> Result, int WorkingSheet, TSheet aSheet)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case TOpenXmlManager.LegDrawMainNamespace + ":shape":
                        ReadLegacyShape(Result, WorkingSheet, aSheet);
                        break;

                    default: //we will ignore the rest
                        DataStream.GetXml(); break;
                }
            }
        }

        private void ReadLegacyDrawingHFXml(TDrawing Drawing)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case TOpenXmlManager.LegDrawMainNamespace + ":shape":
                        ReadLegacyShapeHF(Drawing);
                        break;

                    default: //we will ignore the rest
                        DataStream.GetXml(); break;
                }
            }
        }

        private void ReadLegacyShape(Dictionary<TOneCellRef, TObjectProperties> CommentProperties, int WorkingSheet, TSheet aSheet)
        {
            TObjectProperties ObjProps = CreateStandardObjProps();
            TOneCellRef CommentAddress = new TOneCellRef(-1, -1);
            TObjectType ObjType = TObjectType.Undefined;

            string ShapeName = DataStream.GetAttribute("id");
            if (!String.IsNullOrEmpty(ShapeName) && !ShapeName.StartsWith("\0")) ObjProps.ShapeName = ShapeName;

            ObjProps.AltText = DataStream.GetAttribute("alt");
            TFillStyle FillStyle = ReadLegacyFillColor(DataStream.GetAttribute("fillcolor"));
            string HasFill = DataStream.GetAttribute("filled");
            ObjProps.FShapeFill = new TShapeFill(GetLegacyBoolean(HasFill, true), FillStyle);

            TFillStyle BorderStyle = ReadLegacyFillColor(DataStream.GetAttribute("strokecolor"));
            string HasBorder = DataStream.GetAttribute("stroked");
            ObjProps.ShapeBorder = new TShapeLine(GetLegacyBoolean(HasBorder, true), new TLineStyle(BorderStyle));

            ObjProps.FHidden = ReadShapeStyle(DataStream.GetAttribute("style"));

            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case TOpenXmlManager.LegDrawExcelNamespace + ":ClientData":
                        ReadClientData(ref ObjProps, aSheet, ref ObjType, ref CommentAddress);
                        break;

                    case TOpenXmlManager.LegDrawMainNamespace + ":textbox":
                        ReadTextBox(ObjProps, aSheet);
                        break;

                    case TOpenXmlManager.LegDrawOfficeNamespace + ":lock":
                        ReadLock(ObjProps, false);
                        break;

                    case TOpenXmlManager.LegDrawMainNamespace + ":fill":
                        TFillStyle fs = ReadLegacyFill();
                        if (fs != null) ObjProps.ShapeFill = new TShapeFill(true, fs);
                        break;

                    default:
                        DataStream.GetXml(); break;
                }
            }


            if (ObjProps.Anchor != null)
            {
                switch (ObjType)
                {
                    case TObjectType.CheckBox:
                        int cbIndex = aSheet.Drawing.AddCheckbox(ObjProps.Anchor.Dec(), ObjProps.FText, ObjProps.FCheckboxState, xls, aSheet, ObjProps, null);
                        SetObjLink(WorkingSheet, aSheet, ObjProps, cbIndex);
                        SetMacro(WorkingSheet, aSheet, ObjProps, cbIndex);

                        break;

                    case TObjectType.OptionButton:
                        int rbIndex = aSheet.Drawing.AddRadioButton(ObjProps.Anchor.Dec(), ObjProps.FText, ObjProps.FCheckboxState, xls, aSheet, ObjProps, null);
                        SetObjLink(WorkingSheet, aSheet, ObjProps, rbIndex);
                        SetMacro(WorkingSheet, aSheet, ObjProps, rbIndex);
                        break;

                    case TObjectType.GroupBox:
                        int gbIndex = aSheet.Drawing.AddGroupBox(ObjProps.Anchor.Dec(), ObjProps.FText, xls, aSheet, ObjProps, null);
                        SetMacro(WorkingSheet, aSheet, ObjProps, gbIndex);
                        break;

                    case TObjectType.Button:
                        int btnIndex = aSheet.Drawing.AddButton(ObjProps.Anchor.Dec(), xls, aSheet, ObjProps);
                        SetMacro(WorkingSheet, aSheet, ObjProps, btnIndex);
                        break;

                    case TObjectType.ComboBox:
                        int cobIndex = aSheet.Drawing.AddComboBox(ObjProps.Anchor.Dec(), 0, xls, aSheet, ObjProps, null);
                        SetObjLink(WorkingSheet, aSheet, ObjProps, cobIndex);
                        SetObjectRange(WorkingSheet, aSheet, ObjProps, cobIndex);
                        SetMacro(WorkingSheet, aSheet, ObjProps, cobIndex);
                        break;

                    case TObjectType.ListBox:
                        int lobIndex = aSheet.Drawing.AddListBox(ObjProps.Anchor.Dec(), 0, xls, aSheet, ObjProps, null, ObjProps.FComboBoxProperties.SelectionType);
                        SetObjLink(WorkingSheet, aSheet, ObjProps, lobIndex);
                        SetObjectRange(WorkingSheet, aSheet, ObjProps, lobIndex);
                        SetMacro(WorkingSheet, aSheet, ObjProps, lobIndex);
                        break;

                    case TObjectType.Label:
                        int lIndex = aSheet.Drawing.AddLabel(ObjProps.Anchor.Dec(), ObjProps.FText, xls, aSheet, ObjProps, null);
                        SetMacro(WorkingSheet, aSheet, ObjProps, lIndex);
                        break;

                    case TObjectType.Spinner:
                        int sIndex = aSheet.Drawing.AddSpinner(ObjProps.Anchor.Dec(), xls, aSheet, ObjProps, null);
                        SetObjLink(WorkingSheet, aSheet, ObjProps, sIndex);
                        SetMacro(WorkingSheet, aSheet, ObjProps, sIndex);
                        break;
                    
                    case TObjectType.ScrollBar:
                        int sbIndex = aSheet.Drawing.AddScrollBar(ObjProps.Anchor.Dec(), xls, aSheet, ObjProps, null);
                        SetObjLink(WorkingSheet, aSheet, ObjProps, sbIndex);
                        SetMacro(WorkingSheet, aSheet, ObjProps, sbIndex);
                        break;

                    case TObjectType.Comment:
                        if (CommentAddress.Row >= 0 && CommentAddress.Col >= 0)
                        {
                            CommentProperties.Add(CommentAddress, ObjProps);
                        }
                        break;
                }
            }

        }

        private void SetObjLink(int WorkingSheet, TSheet aSheet, TObjectProperties ObjProps, int cbIndex)
        {
            if (ObjProps.FLinkedFmla != null)
            {
                {
                    TFormulaConvertTextToInternal Converter = new TFormulaConvertTextToInternal(xls, WorkingSheet, false, TFormulaMessages.TokenString(TFormulaToken.fmStartFormula) + ObjProps.FLinkedFmla, true, false, false, null, TFmReturnType.Ref, false);
                    Converter.SetReadingXlsx();
                    Converter.Parse();

                    aSheet.Drawing.SetObjectLink(cbIndex, null, null, Converter.GetTokens(), xls, true);
                }
            }
        }

        private void SetObjectRange(int WorkingSheet, TSheet aSheet, TObjectProperties ObjProps, int cbIndex)
        {
            if (ObjProps.FComboBoxProperties != null && ObjProps.FComboBoxProperties.FormulaRange != null)
            {
                {
                    TFormulaConvertTextToInternal Converter = new TFormulaConvertTextToInternal(xls, WorkingSheet, false, TFormulaMessages.TokenString(TFormulaToken.fmStartFormula) + ObjProps.FComboBoxProperties.FormulaRange, true, false, false, null, TFmReturnType.Ref, false);
                    Converter.SetReadingXlsx();
                    Converter.Parse();

                    aSheet.Drawing.SetFormulaRange(cbIndex, null, null, Converter.GetTokens(), xls, true);
                }
            }
        }

        private void SetMacro(int WorkingSheet, TSheet aSheet, TObjectProperties ObjProps, int btnIndex)
        {
            if (ObjProps.Macro != null)
            {
                {
                    TFormulaConvertTextToInternal Converter = new TFormulaConvertTextToInternal(xls, WorkingSheet, false, TFormulaMessages.TokenString(TFormulaToken.fmStartFormula) + ObjProps.Macro, true, false, false, null, TFmReturnType.Ref, false);
                    Converter.SetReadingXlsx();
                    Converter.Parse();

                    aSheet.Drawing.SetButtonMacro(btnIndex, null, Converter.GetTokens(), xls);
                }
            }
        }

        private void ReadLegacyShapeHF(TDrawing Drawing)
        {
            string ShapeName = DataStream.GetAttribute("id");
            if (string.IsNullOrEmpty(ShapeName) || ShapeName.Length < 2)
            {
                DataStream.FinishTag();
                return;
            }
            THeaderAndFooterPos HFPos;
            if (!GetHeaderAndFooterPos(ShapeName.Substring(0, 2), out HFPos))
            {
                DataStream.FinishTag();
                return;
            }

            THeaderAndFooterKind HFKind = THeaderAndFooterKind.Default;
            if (ShapeName.Length > 2)
            {
                if (!GetHeaderAndFooterKind(ShapeName.Substring(2), out HFKind))
                {
                    DataStream.FinishTag();
                    return;
                }
            }

            byte[] Data = null;
            TXlsImgType ImageType = TXlsImgType.Unknown;
            THeaderOrFooterImageProperties Props = new THeaderOrFooterImageProperties();

            long Width = 0;
            long Height = 0;
            string Style = DataStream.GetAttribute("style");
            string[] Styles = Style.Split(';');
            foreach (string s in Styles)
            {
                if (string.IsNullOrEmpty(s)) continue;
                string s1 = s.Trim();
                int ColonPos = s1.IndexOf(":");
                if (ColonPos < 1 || ColonPos >= s1.Length - 1) continue;

                string att = s1.Substring(0, ColonPos).Trim();
                string val = s1.Substring(ColonPos + 1).Trim();

                switch (att)
                {
                    case "width":
                        Width = GetLegacyMeasure(val);
                        break;
                    case "height":
                        Height = GetLegacyMeasure(val);
                        break;

                }
            }

            Props.Anchor = new THeaderOrFooterAnchor(Width, Height);
            Props.PreferRelativeSize = GetLegacyBoolean(DataStream.GetAttribute("preferrelative", TOpenXmlManager.LegDrawOfficeNamespace), true);

            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case TOpenXmlManager.LegDrawMainNamespace + ":imagedata":
                        string ContType;
                        string FileName;
                        Data = DataStream.GetRelationshipData("relid", 0, 0, TOpenXmlManager.LegDrawOfficeNamespace, out ContType, out FileName);
                        ImageType = TDrawing.GetImageType(ContType);
                        Props.FileName = DataStream.GetAttribute("title", TOpenXmlManager.LegDrawOfficeNamespace);
                        Props.CropArea = new TCropArea(ReadLegacyPercentFF("croptop", 0), ReadLegacyPercentFF("cropbottom", 0), ReadLegacyPercentFF("cropleft", 0), ReadLegacyPercentFF("cropright", 0));
                        Props.Brightness = ReadLegacyPercentFF("blacklevel", FlxConsts.DefaultBrightness);
                        Props.Contrast = ReadLegacyPercentFF("gain", FlxConsts.DefaultContrast);
                        Props.Gamma = ReadLegacyPercentFF("gamma", FlxConsts.DefaultGamma);
                        Props.Grayscale = GetLegacyBoolean(DataStream.GetAttribute("grayscale"));
                        Props.BiLevel = GetLegacyBoolean(DataStream.GetAttribute("bilevel"));
                        TDrawingColor Col;
                        if (ReadLegacyColor(DataStream.GetAttribute("chromakey"), out Col))
                        {
                            Props.TransparentColor = Col.ToColor(xls).ToArgb();
                        }

                        DataStream.FinishTag();
                        break;

                    case TOpenXmlManager.LegDrawOfficeNamespace + ":lock":
                        ReadLock(Props, true);
                        break;

                    default:
                        DataStream.GetXml(); break;
                }
            }

            if (Data == null) return;
            if (ImageType == TXlsImgType.Unknown) return;
            Drawing.AssignHeaderOrFooterDrawing(HFKind, HFPos, Data, ImageType, Props);

        }

        private long GetLegacyMeasure(string val)
        {
            if (string.IsNullOrEmpty(val)) return 0;
            val = val.Trim();
            string vunit = "";
            string sv = val;
            if (val.Length > 2)
            {
                vunit = val.Substring(val.Length - 2);
                sv = val.Substring(0, val.Length - 2).Trim();
            }

            double v;
            if (!TCompactFramework.ConvertToNumber(sv, CultureInfo.InvariantCulture, out v)) return 0;

            switch (vunit)
            {
                case "cm": return GetLong(v / 2.54 * 96.0);
                case "mm": return GetLong(v / 254 * 96.0);
                case "in": return GetLong(v * 96.0);
                case "pt": return GetLong(v * 96.0 / 72.0);
                case "pc": return GetLong(v * 12.0 * 96.0 / 72.0);
                case "px": return GetLong(v);
            }
            return 0;
        }

        private long GetLong(double v)
        {
            if (v < long.MinValue) return long.MinValue;
            if (v > long.MaxValue) return long.MaxValue;
            return (long)v;
        }

        private int ReadLegacyPercentFF(string AttName, int DefaultValue)
        {
            string p = DataStream.GetAttribute(AttName);

            if (string.IsNullOrEmpty(p)) return DefaultValue;
            p = p.Trim();
            if (string.IsNullOrEmpty(p)) return DefaultValue;

            if (p.EndsWith("f"))
            {
                int r;
                if (Int32.TryParse(p.Substring(0, p.Length - 1), NumberStyles.Any, CultureInfo.InvariantCulture, out r)) return r;
                return DefaultValue;
            }

            double d;
            if (!TCompactFramework.ConvertToNumber(p, CultureInfo.InvariantCulture, out d)) return DefaultValue;

            d = d * 65536;
            if (d > Int32.MaxValue) return Int32.MaxValue;
            if (d < Int32.MinValue) return Int32.MinValue;

            return (Int32)d;
        }

        private static bool GetHeaderAndFooterPos(string ShapeName, out THeaderAndFooterPos HFPos)
        {
            switch (ShapeName)
            {
                case "LH":
                    HFPos = THeaderAndFooterPos.HeaderLeft;
                    return true;
                case "CH":
                    HFPos = THeaderAndFooterPos.HeaderCenter;
                    return true;
                case "RH":
                    HFPos = THeaderAndFooterPos.HeaderRight;
                    return true;
                case "LF":
                    HFPos = THeaderAndFooterPos.FooterLeft;
                    return true;
                case "CF":
                    HFPos = THeaderAndFooterPos.FooterCenter;
                    return true;
                case "RF":
                    HFPos = THeaderAndFooterPos.FooterRight;
                    return true;
            }

            HFPos = THeaderAndFooterPos.HeaderLeft;
            return false;
        }

        private bool GetHeaderAndFooterKind(string ShapeName, out THeaderAndFooterKind HFKind)
        {
            switch (ShapeName)
            {
                case "FIRST": HFKind = THeaderAndFooterKind.FirstPage; return true;
                case "EVEN": HFKind = THeaderAndFooterKind.EvenPages; return true;
            }

            HFKind = THeaderAndFooterKind.Default;
            return false;
        }


        private TFillStyle ReadLegacyFill()
        {
            DataStream.FinishTagAndIgnoreChildren();
            return null;
        }

        private TFillStyle ReadLegacyFillColor(string FillColor)
        {
            TDrawingColor fc;
            if (!ReadLegacyColor(FillColor, out fc)) return null;
            return new TSolidFill(fc);
        }

        private bool ReadLegacyColor(string aColor, out TDrawingColor Col)
        {
            if (!string.IsNullOrEmpty(aColor) && aColor.IndexOf("[") > 0) return GetLegacyIndexedColor(aColor, out Col);

            Color c = THtmlColors.GetColor(aColor);
            Col = c;
            if (ColorUtil.Empty.Equals(c)) return false;
            return true;
        }

        private bool GetLegacyIndexedColor(string aColor, out TDrawingColor Col)
        {            
            int a1 = aColor.IndexOf("[");
            int a2 = aColor.IndexOf("]");
            if (a1 < 0 || a2 < a1)
            {
                Col = TDrawingColor.FromSystem(TSystemColor.Window); 
                return false;
            }

            if (aColor.StartsWith("#"))
            {
                Col = THtmlColors.GetColor(aColor.Substring(0, a1 - 1));
                return true;
            }

            string c = aColor.Substring(a1 + 1, a2 - a1 - 1);
            int ci;
            if (!int.TryParse(c, NumberStyles.Any, CultureInfo.InvariantCulture, out ci))
            {
                Col = TDrawingColor.FromSystem(TSystemColor.Window);
                return false;
            }


            if (ci < 56)
            {
                Col = TDrawingColor.FromColor(TExcelColor.FromBiff8ColorIndex(ci).ToColor(xls));
                return true;
            }

            TSystemColor syscol = ColorUtil.GetSystemColor(ci - 56);
            if (syscol == TSystemColor.None)
            {
                Col = TDrawingColor.FromSystem(TSystemColor.Window);
                return false;
            }


            Col = TDrawingColor.FromSystem(syscol);
            return true;
        }

        private void ReadTextBox(TObjectProperties TextProps, TSheet aSheet)
        {
            ReadTextStyle(TextProps);

            string s = DataStream.ReadLegacyValue();
            if (s != null)
            {
                TextProps.FText = new TRichString();
                TextProps.FText.SetFromHtml(s, xls.GetDefaultFormat, xls, true);
            }
        }

        private void ReadTextStyle(TObjectProperties TextProps)
        {
            string Style = DataStream.GetAttribute("style");
            if (String.IsNullOrEmpty(Style))
            {
                return;
            }

            bool Vertical = false;
            TTextRotation TextRot = TTextRotation.Normal;

            string[] Styles = Style.Split(';');
            foreach (string s in Styles)
            {
                if (string.IsNullOrEmpty(s)) continue;
                string s1 = s.Trim();
                int ColonPos = s1.IndexOf(":");
                if (ColonPos < 1 || ColonPos >= s1.Length -1) continue;

                string att = s1.Substring(0, ColonPos).Trim();
                string val = s1.Substring(ColonPos + 1).Trim();

                switch (att)
                {
                    case "layout-flow":
                        Vertical = val == "vertical";
                        break;

                    case "mso-layout-flow-alt":
                        switch (val)
                        {
                            case "top-to-bottom": TextRot = TTextRotation.Vertical; break;
                            case "bottom-to-top": TextRot = TTextRotation.Rotated90Degrees; break;
                        }
                        break;

                    case "mso-fit-shape-to-text":
                        TextProps.FAutoSize = GetLegacyBoolean(val);
                        break;
                }
            }

            if (Vertical)
            {
                switch (TextRot)
                {
                    case TTextRotation.Rotated90Degrees:
                        TextProps.FTextProperties.TextRotation = TTextRotation.Rotated90Degrees;
                        break;
                    case TTextRotation.Vertical:
                        TextProps.FTextProperties.TextRotation = TTextRotation.Vertical;
                        break;
                    default:
                        TextProps.FTextProperties.TextRotation = TTextRotation.RotatedMinus90Degrees;
                        break;
                }
            }

        }

        private bool ReadShapeStyle(string Style)
        {
            if (String.IsNullOrEmpty(Style))
            {
                return false;
            }

            string[] Styles = Style.Split(';');
            foreach (string s in Styles)
            {
                if (string.IsNullOrEmpty(s)) continue;
                string s1 = s.Trim();
                int ColonPos = s1.IndexOf(":");
                if (ColonPos < 1 || ColonPos >= s1.Length - 1) continue;

                string att = s1.Substring(0, ColonPos).Trim();
                string val = s1.Substring(ColonPos + 1).Trim();

                if (att == "visibility" && val == "hidden") return true;
            }
            return false;
        }

        private void ReadLock(TBaseImageProperties ImgProps, bool DefaultsToTrue)
        {
            string AspectRatio = DataStream.GetAttribute("aspectratio");
            ImgProps.LockAspectRatio = GetLegacyBoolean(AspectRatio, DefaultsToTrue);
            DataStream.FinishTagAndIgnoreChildren();
        }

        private void ReadClientData(ref TObjectProperties CommentProps, TSheet aSheet, ref TObjectType ObjType, ref TOneCellRef CommentAddress)
        {
            string sObjType = DataStream.GetAttribute("ObjectType");
            switch (sObjType)
            {
                case "Note":
                    {
                        ObjType = TObjectType.Comment;
                        ReadActualClientData(CommentProps, ref CommentAddress, aSheet);
                        return;
                    }

                case "Checkbox":
                    {
                        ObjType = TObjectType.CheckBox;
                        ReadActualClientData(CommentProps, ref CommentAddress, aSheet);
                        return;
                    }

                case "Radio":
                    {
                        ObjType = TObjectType.OptionButton;
                        ReadActualClientData(CommentProps, ref CommentAddress, aSheet);
                        return;
                    }

                case "GBox":
                    {
                        ObjType = TObjectType.GroupBox;
                        ReadActualClientData(CommentProps, ref CommentAddress, aSheet);
                        return;
                    }

                case "Button":
                    {
                        ObjType = TObjectType.Button;
                        ReadActualClientData(CommentProps, ref CommentAddress, aSheet);
                        return;
                    }

                case "Drop":
                    {
                        ObjType = TObjectType.ComboBox;
                        ReadActualClientData(CommentProps, ref CommentAddress, aSheet);
                        return;
                    }

                case "List":
                    {
                        ObjType = TObjectType.ListBox;
                        ReadActualClientData(CommentProps, ref CommentAddress, aSheet);
                        return;
                    }

                case "Label":
                    {
                        ObjType = TObjectType.Label;
                        ReadActualClientData(CommentProps, ref CommentAddress, aSheet);
                        return;
                    }

                case "Spin":
                    {
                        ObjType = TObjectType.Spinner;
                        ReadActualClientData(CommentProps, ref CommentAddress, aSheet);
                        return;
                    }

                case "Scroll":
                    {
                        ObjType = TObjectType.ScrollBar;
                        ReadActualClientData(CommentProps, ref CommentAddress, aSheet);
                        return;
                    }

            }
            DataStream.FinishTagAndIgnoreChildren();            
        }

        private void ReadActualClientData(TObjectProperties CommentProps, ref TOneCellRef CommentAddress, TSheet aSheet)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            TFlxAnchorType AnchorType = TFlxAnchorType.MoveAndResize;
            string CellAnchor = null;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case TOpenXmlManager.LegDrawExcelNamespace + ":MoveWithCells":
                        if (ReadLegacyBoolean()) //this is reversed
                        {
                            AnchorType = TFlxAnchorType.DontMoveAndDontResize;
                        }
                        break;

                    case TOpenXmlManager.LegDrawExcelNamespace + ":SizeWithCells":
                        if (ReadLegacyBoolean()) //this is reversed
                        {
                            if (AnchorType == TFlxAnchorType.MoveAndResize) AnchorType = TFlxAnchorType.MoveAndDontResize; else AnchorType = TFlxAnchorType.DontMoveAndDontResize;
                        }
                        break;

                    case TOpenXmlManager.LegDrawExcelNamespace + ":Anchor":
                        CellAnchor = DataStream.ReadValueAsString();
                        break;

                    case TOpenXmlManager.LegDrawExcelNamespace + ":Locked":
                        CommentProps.Lock = ReadLegacyBoolean();
                        break;

                    case TOpenXmlManager.LegDrawExcelNamespace + ":DefaultSize":
                        CommentProps.DefaultSize = ReadLegacyBoolean();
                        break;

                    case TOpenXmlManager.LegDrawExcelNamespace + ":Disabled":
                        CommentProps.Disabled = ReadLegacyBoolean();
                        break;

                    case TOpenXmlManager.LegDrawExcelNamespace + ":PrintObject":
                        CommentProps.Print = ReadLegacyBoolean();
                        break;

                    case TOpenXmlManager.LegDrawExcelNamespace + ":AutoFill":
                        CommentProps.AutoFill = ReadLegacyBoolean();
                        break;

                    case TOpenXmlManager.LegDrawExcelNamespace + ":AutoLine":
                        CommentProps.AutoLine = ReadLegacyBoolean();
                        break;

                   // case TOpenXmlManager.LegDrawExcelNamespace + ":Published":
                   //     ImgProps.Published = ReadLegacyBoolean();
                   //     break;


                    case TOpenXmlManager.LegDrawExcelNamespace + ":FmlaMacro":
                        CommentProps.Macro = DataStream.ReadValueAsString();
                        break;

                    case TOpenXmlManager.LegDrawExcelNamespace + ":TextHAlign":
                        CommentProps.FTextProperties.HAlignment = ReadLegacyHAlignment();
                        break;

                    case TOpenXmlManager.LegDrawExcelNamespace + ":TextVAlign":
                        CommentProps.FTextProperties.VAlignment = ReadLegacyVAlignment();
                        break;
                    
                    case TOpenXmlManager.LegDrawExcelNamespace + ":LockText":
                        CommentProps.FTextProperties.LockText = ReadLegacyBoolean();
                        break;

                    case TOpenXmlManager.LegDrawExcelNamespace + ":Row":
                        CommentAddress.Row = DataStream.ReadValueAsInt() + 1;
                        break;

                    case TOpenXmlManager.LegDrawExcelNamespace + ":Column":
                        CommentAddress.Col = DataStream.ReadValueAsInt() + 1;
                        break;

                    case TOpenXmlManager.LegDrawExcelNamespace + ":Checked":
                        switch (DataStream.ReadValueAsInt())
                        {
                            case 1:
                                CommentProps.FCheckboxState = TCheckboxState.Checked;
                                break;
                            case 2:
                                CommentProps.FCheckboxState = TCheckboxState.Indeterminate;
                                break;
                        }
                        break;

                    case TOpenXmlManager.LegDrawExcelNamespace + ":FmlaLink":
                        CommentProps.FLinkedFmla = DataStream.ReadValueAsString();
                        break;

                        /* This one is not really useful, as it doesn't have NextIds.
                    case TOpenXmlManager.LegDrawExcelNamespace + ":FirstButton":
                        CommentProps.FFirstButton = ReadLegacyBoolean();
                        break;*/

                    case TOpenXmlManager.LegDrawExcelNamespace + ":NoThreeD":
                    case TOpenXmlManager.LegDrawExcelNamespace + ":NoThreeD2":
                        CommentProps.FIs3D = !ReadLegacyBoolean();
                        break;

                    case TOpenXmlManager.LegDrawExcelNamespace + ":FmlaRange":
                        if (CommentProps.FComboBoxProperties == null) CommentProps.FComboBoxProperties = new TComboBoxProperties();
                        CommentProps.FComboBoxProperties.FormulaRange = DataStream.ReadValueAsString();
                        break;

                    case TOpenXmlManager.LegDrawExcelNamespace + ":Val":
                        if (CommentProps.FSpinProperties == null) CommentProps.FSpinProperties = new TSpinProperties();
                        CommentProps.FSpinProperties.FVal = DataStream.ReadValueAsInt();
                        break;

                    case TOpenXmlManager.LegDrawExcelNamespace + ":Min":
                        if (CommentProps.FSpinProperties == null) CommentProps.FSpinProperties = new TSpinProperties();
                        CommentProps.FSpinProperties.Min = DataStream.ReadValueAsInt();
                        break;

                    case TOpenXmlManager.LegDrawExcelNamespace + ":Max":
                        if (CommentProps.FSpinProperties == null) CommentProps.FSpinProperties = new TSpinProperties();
                        CommentProps.FSpinProperties.Max = DataStream.ReadValueAsInt();
                        break;

                    case TOpenXmlManager.LegDrawExcelNamespace + ":Inc":
                        if (CommentProps.FSpinProperties == null) CommentProps.FSpinProperties = new TSpinProperties();
                        CommentProps.FSpinProperties.Incr = DataStream.ReadValueAsInt();
                        break;

                    case TOpenXmlManager.LegDrawExcelNamespace + ":Page":
                        if (CommentProps.FSpinProperties == null) CommentProps.FSpinProperties = new TSpinProperties();
                        CommentProps.FSpinProperties.Page = DataStream.ReadValueAsInt();
                        break;

                    case TOpenXmlManager.LegDrawExcelNamespace + ":DropLines":
                        if (CommentProps.FComboBoxProperties == null) CommentProps.FComboBoxProperties = new TComboBoxProperties();
                        CommentProps.FComboBoxProperties.DropLines = DataStream.ReadValueAsInt();
                        break;

                    case TOpenXmlManager.LegDrawExcelNamespace + ":Dx":
                        if (CommentProps.FSpinProperties == null) CommentProps.FSpinProperties = new TSpinProperties();
                        CommentProps.FSpinProperties.Dx = DataStream.ReadValueAsInt();
                        break;

                    case TOpenXmlManager.LegDrawExcelNamespace + ":Sel":
                        if (CommentProps.FComboBoxProperties == null) CommentProps.FComboBoxProperties = new TComboBoxProperties();
                        CommentProps.FComboBoxProperties.Sel = DataStream.ReadValueAsInt();
                        break;

                    case TOpenXmlManager.LegDrawExcelNamespace + ":SelType":
                        if (CommentProps.FComboBoxProperties == null) CommentProps.FComboBoxProperties = new TComboBoxProperties();
                        CommentProps.FComboBoxProperties.SelectionType = GetSelType(DataStream.ReadValueAsString());
                        break;


                    default:
                        DataStream.GetXml(); break;
                }

                if (CellAnchor != null)
                {
                    string[] Coords = CellAnchor.Split(',');
                    if (Coords.Length == 8)
                    {
                        IRowColSize rc = new RowColSize(xls.HeightCorrection, xls.WidthCorrection, aSheet);
                        CommentProps.Anchor = new TClientAnchor(AnchorType, gs(Coords, 2) + 1, gs(Coords, 3), gs(Coords, 0) + 1, gs(Coords, 1),
                                                                        gs(Coords, 6) + 1, gs(Coords, 7), gs(Coords, 4) + 1, gs(Coords, 5), rc);
                    }
                }
            }                     
        }

        private TListBoxSelectionType GetSelType(string p)
        {
            switch (p)
            {
                case "Multi": return TListBoxSelectionType.Multi;
                case "Extend": return TListBoxSelectionType.Extend;
            }
            return TListBoxSelectionType.Single;
        }

        private TVFlxAlignment ReadLegacyVAlignment()
        {
            switch (DataStream.ReadValueAsString())
            {
                case "Center": return TVFlxAlignment.center;
                case "Bottom": return TVFlxAlignment.bottom;
                case "Justify": return TVFlxAlignment.justify;
                case "Distributed": return TVFlxAlignment.distributed;
            }
            return TVFlxAlignment.top;
        }

        private THFlxAlignment ReadLegacyHAlignment()
        {
            switch (DataStream.ReadValueAsString())
            {
                case "Center": return THFlxAlignment.center;
                case "Right": return THFlxAlignment.right;
                case "Justify": return THFlxAlignment.justify;
                case "Distributed": return THFlxAlignment.distributed;
            }
            return THFlxAlignment.left;
        }

        private int gs(string[] Coords, int p)
        {
            if (string.IsNullOrEmpty(Coords[p])) return 0;
            return Convert.ToInt32(Coords[p].Trim(), CultureInfo.InvariantCulture);
        }

        private bool ReadLegacyBoolean()
        {
            string b = DataStream.ReadValueAsString();
            return b != "false" && b != "f" && b != "False";
        }

        private static bool GetLegacyBoolean(string val)
        {
            return val == "t" || val == "T" || val == "true" || val == "True";
        }

        private static bool GetLegacyBoolean(string val, bool DefaultVal)
        {
            if (val == null) return DefaultVal;
            return GetLegacyBoolean(val);
        }


        private TObjectProperties CreateStandardObjProps()
        {
            return new TObjectProperties(null, String.Empty, null, new TCropArea(), FlxConsts.NoTransparentColor,
                FlxConsts.DefaultBrightness, FlxConsts.DefaultContrast, FlxConsts.DefaultGamma, 
                true, true, false, false, false, false, true, null, null, false, new TObjectTextProperties(),
                false, null, null, false, true, true);
        }

        #endregion

        #region Comments
        private void ReadComments(TSheet aSheet, Dictionary<TOneCellRef, TObjectProperties> CommentProperties)
        {
            if (aSheet == null) return;
            List<Uri> CommentUris = DataStream.GetUrisForCurrentPartRelationship(TOpenXmlManager.CommentsRelationshipType);
            foreach (Uri CommentUri in CommentUris)
            {
                ReadCommentPart(aSheet, CommentUri, CommentProperties);
            }
        }

        private void ReadCommentPart(TSheet aSheet, Uri CommentUri, Dictionary<TOneCellRef, TObjectProperties> CommentProperties)
        {
            DataStream.SelectMasterPart(CommentUri, TOpenXmlManager.MainNamespace);
            while (DataStream.NextTag())
            {
                switch (DataStream.RecordName())
                {
                    case "comments":
                        ReadActualComments(aSheet, CommentProperties);
                        break;

                    default: //shouldn't happen
                        DataStream.GetXml();
                        break;
                }
            }
        }

        private void ReadActualComments(TSheet aSheet, Dictionary<TOneCellRef, TObjectProperties> CommentProperties)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            List<string> AuthorList = new List<string>();
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "authors": ReadAuthors(AuthorList); break;
                    case "commentList": ReadCommentList(AuthorList, aSheet, CommentProperties); break;
                    
                    case "extLst":
                    default:
                        aSheet.Notes.AddFutureStorage(new TFutureStorageRecord(DataStream.GetXml())); break;
                }
            }
        }

        private void ReadAuthors(List<string> AuthorList)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "author": AuthorList.Add(DataStream.ReadValueAsString()); break;
                    default:
                        DataStream.GetXml(); break;
                }
            }            
        }

        private void ReadCommentList(List<string> AuthorList, TSheet aSheet, Dictionary<TOneCellRef, TObjectProperties> CommentProperties)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "comment": ReadComment(AuthorList, aSheet, CommentProperties); break;
                    default:
                        DataStream.GetXml(); break;
                }
            }
        }

        private void ReadComment(List<string> AuthorList, TSheet aSheet, Dictionary<TOneCellRef, TObjectProperties> CommentProperties)
        {
            TRichString CommentText = null;

            TCellAddress CellRef = DataStream.GetAttributeAsAddress("ref");
            int AuthorId = DataStream.GetAttributeAsInt("authorId", -1);
            string Author = null;
            if (AuthorId >= 0 && AuthorId < AuthorList.Count) Author = AuthorList[AuthorId];

            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "text": CommentText = DataStream.ReadValueAsRichString(xls); break;

                    case "commentPr": //doesn't seem to be used by Excel.
                    default:
                        DataStream.GetXml(); break;
                }
            }

            if (CommentText != null && CellRef != null)
            {
                TObjectProperties CommentProp;
                if (CommentProperties.TryGetValue(new TOneCellRef(CellRef.Row, CellRef.Col), out CommentProp)) //we could be more forgiving here, and load even comments that don't have an associated legacy drawing. But Excel doesn't do it, so we won't either.
                {
                    aSheet.Notes.AddNewComment(CellRef.Row - 1, CellRef.Col - 1, CommentText, Author, aSheet.Drawing, CommentProp.Dec(), xls, aSheet, true);
                }
            }
        }

        #endregion

        #region External Links
        internal void ReadExternalLink(TWorkbookGlobals Globals, string LinkId)
        {
            DataStream.SelectPart(LinkId, TOpenXmlManager.MainNamespace);
            while (DataStream.NextTag())
            {
                switch (DataStream.RecordName())
                {
                    case "externalLink":
                        ReadActualExternalLink(Globals);
                        break;

                    default: //shouldn't happen
                        DataStream.GetXml();
                        break;
                }
            }
        }

        private void ReadActualExternalLink(TWorkbookGlobals Globals)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            TSupBookRecord Sup = null;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "externalBook": Sup = ReadExternalBook(); Globals.References.AddSupBook(Sup); break;
                    case "oleLink": Sup = ReadOleLink(); Globals.References.AddSupBook(Sup); break;
                    case "ddeLink": Sup = ReadDdeLink(); Globals.References.AddSupBook(Sup); break;

                    case "extLst":
                    default:
                        if (Sup != null) Sup.AddFutureStorage(new TFutureStorageRecord(DataStream.GetXml()));
                        break;
                }
            }
        }

        private TSupBookRecord ReadExternalBook()
        {
            string rId = DataStream.GetRelationship("id");

            TSupBookRecord Result = TSupBookRecord.CreateExternalRef(GetExternalFilename(rId), null);
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return Result; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return Result;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "definedNames": ReadExternalDefinedNames(Result); break;
                    case "sheetDataSet": ReadExternalSheetDataSet(); break;
                    case "sheetNames": ReadExternalSheetNames(Result); break;
                 
                    default:
                        DataStream.GetXml();
                        break;
                }
            }
            return Result;
        }

        private string GetExternalFilename(string rId)
        {
            string s = DataStream.GetExternalLink(rId);
            if (s.StartsWith("file:///")) s = s.Substring("file:///".Length);
            return Uri.UnescapeDataString(s);
 /*           //Here we can have an uri, or a string starting with "/". We will know which it is by looking at IsAbsolute.
            Uri UFileName = DataStream.GetExternalLink(rId);
            if (UFileName.IsAbsoluteUri)
            {
                return UFileName.LocalPath; //don't use Uri.UnescapeDataString here as it is already unescaped.
            }
            else
            {
                return Uri.UnescapeDataString(UFileName.ToString()); //We need Uri.UnescapeDataString because when the Uri is relative, it is not unencoded.
            }*/
        }

        private void ReadExternalDefinedNames(TSupBookRecord SupBook)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "definedName":
                        SupBook.AddExternName(TExternNameRecord.CreateExternName(DataStream.GetAttributeAsInt("sheetId", -1) + 1, DataStream.GetAttribute("name")));
                        DataStream.FinishTag();
                        break;

                    default:
                        DataStream.GetXml();
                        break;
                }
            }
        }

        private void ReadExternalSheetDataSet()
        {
            DataStream.GetXml(); //we won't load or save cached data.
        }

        private void ReadExternalSheetNames(TSupBookRecord SupBook)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            List<String> SheetNames = new List<string>(); //The structure in SupBookRecord is not very nice right now, so we will use this to speed it up. If we make the sheets in supbook an array, this could be removed.
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "sheetName":
                        SheetNames.Add(DataStream.GetAttribute("val")); 
                        DataStream.FinishTag();
                        break;

                    default:
                        DataStream.GetXml();
                        break;
                }
            }

            SupBook.AddExternalSheets(SheetNames);
        }
                    

        private TSupBookRecord ReadOleLink()
        {
            string rId = DataStream.GetRelationship("id");
            string ProgId = DataStream.GetAttribute("progId");

            TSupBookRecord Result = TSupBookRecord.CreateOleOrDdeLink(ProgId, GetExternalFilename(rId));
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return Result; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return Result;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "oleItems": ReadExternalOleItems(Result); break;

                    default:
                        DataStream.GetXml();
                        break;
                }
            }
            return Result;
            
        }

        private void ReadExternalOleItems(TSupBookRecord SupBook)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "oleItem":
                        SupBook.AddExternName(TExternNameRecord.CreateOleLink(
                            DataStream.GetAttribute("name"),
                            DataStream.GetAttributeAsBool("icon", false),
                            DataStream.GetAttributeAsBool("advise", false),
                            DataStream.GetAttributeAsBool("preferPic", false)
                            ));
                        DataStream.FinishTag();
                        break;

                    default:
                        DataStream.GetXml();
                        break;
                }
            }
        }

        private TSupBookRecord ReadDdeLink()
        {
            string DdeService = DataStream.GetAttribute("ddeService");
            string DdeTopic = DataStream.GetAttribute("ddeTopic");

            TSupBookRecord Result = TSupBookRecord.CreateOleOrDdeLink(DdeService, DdeTopic);
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return Result; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return Result;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "ddeItems": ReadExternalDdeItems(Result); break;

                    default:
                        DataStream.GetXml();
                        break;
                }
            }
            return Result;

        }

        private void ReadExternalDdeItems(TSupBookRecord SupBook)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "ddeItem":
                        SupBook.AddExternName(TExternNameRecord.CreateDdeLink(
                            DataStream.GetAttribute("name"),
                            DataStream.GetAttributeAsBool("ole", false),
                            DataStream.GetAttributeAsBool("advise", false),
                            DataStream.GetAttributeAsBool("preferPic", false)
                            ));
                        DataStream.FinishTagAndIgnoreChildren();
                        break;

                    default:
                        DataStream.GetXml();
                        break;
                }
            }
        }

        #endregion

        #region Theme
        internal void ReadTheme(TThemeRecord ThemeRecord)
        {
            DrawingLoader.ReadTheme(ThemeRecord);
        }
        #endregion

        #region File Properties
        internal void ReadCustomFileProperties(TFileProps FileProps)
        {
            FileProps.Custom = null;
            DataStream.SelectCustomFileProps();
            if (DataStream.Eof) return;

            DataStream.NextTag();
            FileProps.Custom = DataStream.GetXml();
        }

        #endregion

        internal void ReadCustomXMLData(TCustomXMLDataStorageList CustomXMLData)
        {
            CustomXMLData.Clear();
            bool me;
            DataStream.SelectWorkbook(out me);
            List<Uri> Uris = DataStream.GetUrisForCurrentPartRelationship(TOpenXmlManager.CustomXmlDataRelationshipType);
            if (Uris == null) return;

            foreach (Uri u in Uris)
            {
                DataStream.SelectMasterPart(u, null);
                DataStream.NextTag();
                TCustomXMLDataStorage ds = CustomXMLData.Add(new TCustomXMLDataStorage(u, DataStream.GetXml()));

                AddCustomXMLDataProps(ds);
            }
        }

        private void AddCustomXMLDataProps(TCustomXMLDataStorage ds)
        {
            List<Uri> Uris = DataStream.GetUrisForCurrentPartRelationship(TOpenXmlManager.CustomXmlDataPropsRelationshipType);
            if (Uris == null) return;

            foreach (Uri u in Uris)
            {
                DataStream.SelectMasterPart(u, TOpenXmlManager.CustomXmlDataPropsNamespace);
                DataStream.NextTag();
                ds.CustomXMLDataStorageProps.Add(new TCustomXMLDataStorageProp(u, DataStream.GetXml()));
            }
        }

        #region Macros
        internal void ReadMacros()
        {
            byte[] MacroData = DataStream.GetPart(TOpenXmlManager.VBARelationshipType, 0, 0);
            xls.SetMacrodata(MacroData);
        }
        #endregion

        internal void SwitchSheets()
        {
            if (VirtualReader != null) VirtualReader.ClearSheet();
        }
    }
}
