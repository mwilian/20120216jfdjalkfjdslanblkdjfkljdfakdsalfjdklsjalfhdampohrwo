using System;
using FlexCel.Core;

using System.IO;

#if (MONOTOUCH)
using Color = MonoTouch.UIKit.UIColor;
#endif

#if (WPF)
using real = System.Double;
using System.Windows.Media;
#else
using System.Drawing;
using System.Collections.Generic;
#endif

namespace FlexCel.XlsAdapter
{
    #region Bound sheets
    /// <summary>
    /// List with BoundSheets.
    /// </summary>
    internal class TBoundSheetList
    {
        private TFutureStorage FutureStorage;
        private TSheetNameList FSheetNames;  //Cache with all the sheet names to speed up searching
        private TBoundSheetRecordList FBoundSheets;
        private UInt32List FTabList;
        private UInt32 MaxTabId;
                          
        internal TBoundSheetRecordList BoundSheets{get{return FBoundSheets;}}

        internal TBoundSheetList()
        {
            FSheetNames= new TSheetNameList();
            FBoundSheets= new TBoundSheetRecordList();
            FTabList = new UInt32List();
            MaxTabId = 0;
        }

        internal void Clear()
        {
            if (FSheetNames!=null) FSheetNames.Clear();
            if (FBoundSheets!=null) FBoundSheets.Clear();
            FTabList.Clear();
            MaxTabId = 0;
        }

        internal void AddFromFile(TBoundSheetRecord aRecord)
        {
            //No TabId here, it will be loaded at its time.
            FSheetNames.Add(null, 0, aRecord.SheetName, FBoundSheets.Count);
            FBoundSheets.Add(aRecord); //Last
        }

        internal void SaveToStream(IDataStream DataStream, TSaveData SaveData)
        {
            FBoundSheets.SaveToStream(DataStream, SaveData, 0);
        }

        internal void SaveToPxl(TPxlStream PxlStream, TPxlSaveData SaveData)
        {
            FBoundSheets.SaveToPxl(PxlStream, 0, SaveData);
        }

        internal void SaveRangeToStream(IDataStream DataStream, TSaveData SaveData, int SheetIndex)
        {
            if ((SheetIndex>=FBoundSheets.Count)|| (SheetIndex<0)) XlsMessages.ThrowException(XlsErr.ErrInvalidSheetNo, SheetIndex,0,FBoundSheets.Count-1);
            FBoundSheets[SheetIndex].SaveToStream(DataStream, SaveData,  0);
        }

        internal long TotalSize()
        {
            return FBoundSheets.TotalSize;
        }

        internal long TotalRangeSize(int SheetIndex)
        {
            if ((SheetIndex>=FBoundSheets.Count) || (SheetIndex<0)) XlsMessages.ThrowException(XlsErr.ErrInvalidSheetNo, SheetIndex,0,FBoundSheets.Count-1);
            return FBoundSheets[SheetIndex].TotalSize();
        }

        internal void InsertSheet(int BeforeSheet, int OptionFlags, string SheetName)
        {
            CheckTabId();
            if (FTabList.Count > 0)
            {
                MaxTabId++;
                FTabList.Insert(BeforeSheet, MaxTabId);
                if (FTabList.Count > XlsConsts.MaxTabIdCount || MaxTabId > XlsConsts.MaxTabIdValue)
                {
                    FTabList.Clear();
                    MaxTabId = 0;
                }
            }

            string NewName= FSheetNames.AddUniqueName(FBoundSheets, BeforeSheet, SheetName, BeforeSheet);
            FBoundSheets.Insert(BeforeSheet, new TBoundSheetRecord(OptionFlags, NewName)); //.CreateNew
        }

        internal void DeleteSheet(int SheetIndex)
        {
            CheckTabId();
            if (FTabList.Count > 0)
            {
                FTabList.RemoveAt(SheetIndex); //MaxTabId keeps the same, we don't want to reuse old ids.
            }
            FSheetNames.DeleteSheet(FBoundSheets.GetSheetName(SheetIndex), FBoundSheets, SheetIndex);
            FBoundSheets.Delete(SheetIndex);
        }

        internal TSheetNameList SheetNames { get{return FSheetNames;}}

        internal void AddFutureStorage(TFutureStorageRecord R)
        {
            TFutureStorage.Add(ref FutureStorage, R);
        }

        private void CheckTabId()
        {
            if (FTabList.Count != BoundSheets.Count)
            {
                FTabList.Clear();
            }
        }

        internal void AddTabIdFromFile(TTabIdRecord TabId)
        {
            MaxTabId = 0;
            for (int i = 0; i < TabId.Data.Length; i+=2)
            {
                UInt32 r = (UInt32)BitOps.GetWord(TabId.Data, i);
                if (r <= 0) //Invalid tab id. Shouldn't happen, but might in 3rd party files.
                {
                    MaxTabId = 0;
                    FTabList.Clear();
                    return;
                }
                FTabList.Add(r);
                if (r > MaxTabId) MaxTabId = r;
            }
        }

        internal void AddTabIdFromFile(uint r)
        {
            FTabList.Add(r);
            if (r > MaxTabId) MaxTabId = r;
        }


        internal long TabIdsTotalSize()
        {
            CheckTabId();
            if (FTabList.Count > XlsConsts.MaxTabIdCount || FTabList.Count != FBoundSheets.Count) return 0;
            return XlsConsts.SizeOfTRecordHeader + 2 * FTabList.Count;
        }

        internal int GetTabId(int i)
        {
            CheckTabId();
            if (i >= FTabList.Count) return i + 1;
            return (int)FTabList[i];
        }

        internal void SaveTabIds(IDataStream DataStream)
        {
            CheckTabId();
            if (FTabList.Count > XlsConsts.MaxTabIdCount || FTabList.Count != FBoundSheets.Count) return;
            DataStream.WriteHeader((UInt16)xlr.TABID,(UInt16) (2 * FTabList.Count));
            for (int i = 0; i < FTabList.Count; i++)
            {
                DataStream.Write16((UInt16)FTabList[i]);
            }
        }
    }
    #endregion

    struct TCopiedGen
    {
        internal long Generation;
        internal int Level;// a chart is a sheet that can be inside other sheet. We need to make sure when copying a chart inside a sheet we don't corrupt Generation.

        public TCopiedGen(int gen)
        {
            Generation = gen;
            Level = 0;
        }

        public override bool Equals(object obj)
        {
            if (!(obj is TCopiedGen)) return false;
            TCopiedGen g = (TCopiedGen)obj;

            return g == this;
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

        public static bool operator ==(TCopiedGen a1, TCopiedGen a2)
        {
            return a1.Generation == a2.Generation && a1.Level == a2.Level;
        }

        public static bool operator !=(TCopiedGen a1, TCopiedGen a2)
        {
            return !(a1 == a2);
        }

        public void IncGeneration()
        {
            Generation++;
        }

        public void Push()
        {
            Level++;
            Generation = 0;
        }
    }

    /// <summary>
    /// Global Section of the workbook.
    /// </summary>
    internal class TWorkbookGlobals: TBaseSection
    {
        #region Variables
        ExcelFile FWorkbook;

        TFileEncryption FFileEncryption;
        TMiscRecordList FLel;
        TBoundSheetList FBoundSheets;
        TMiscRecordList FFnGroups;

        TWorkbookProtection FWorkbookProtection;
        TFontRecordList FFonts;
        TFormatRecordList FFormats;
        TXFRecordList FStyleXF;
        TXFRecordList FCellXF;
        TDXFRecordList FDXF;
        TStyleRecordList FStyles;
        TTableStyleRecordList FTableStyles;

        TMiscRecordList FPivotCache;
#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
        TXlsxPivotCacheList FXlsxPivotCache;
        string FXlsxConnections;
#endif
        TMiscRecordList FDocRoute;
        TMiscRecordList FUserBView;

        TMiscRecordList FMetaData;

        TReferences FReferences;
        TNameRecordList FNames;
        TMiscRecordList FRealTimeData;

        TDrawingGroup FHeaderImages;
        TDrawingGroup FDrawingGroup;

        TSST FSST;
        TMiscRecordList FWebPub;

        TMiscRecordList FFeatHdr;
        TMiscRecordList FDConn;

        TBorderList FBorders;
        TPatternList FPatterns;
        
        TMiscRecordList FFutureRecords;
        internal TFutureStorage StylesFutureStorage;

        internal TCalcOptions CalcOptions; //doesn't go here in biff8, but it does in Excel 2007.
        bool FIsXltTemplate;
        internal TCodePageRecord CodePage;
        internal TExcel9FileRecord Excel9File;
        internal TObProjRecord ObjProj;
        internal TObNoMacrosRecord ObNoMacros;
        internal TCodeNameRecord CodeNameRecord;
#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
        internal string CodeName07;
#endif

        internal TOleObjectSizeRecord OleObjectSize;

        internal TWindow1Record[] Window1;        
        internal bool Backup;
        internal THideObj HideObj;
        internal bool Dates1904;
        internal bool PrecisionAsDisplayed;
        internal bool RefreshAll;
        internal TBookBoolRecord BookBool;
        

        internal TPaletteRecord Palette;
        internal TClrtClientRecord ClrtClient;

        internal bool UsesELFs;

        internal TMTRSettingsRecord MTRSettings;
        internal TForceFullCalculationRecord ForceFullCalculation;

        internal TCountryRecord Country;
        internal TRecalcIdRecord RecalcId;
        internal TWOptRecord WOpt;
        internal TBookExtRecord BookExt;

        internal TThemeRecord ThemeRecord;

        internal TCompressPicturesRecord CompressPictures;
        internal TCompat12Record Compat12;
        internal TGUIDTypeLibRecord GUIDTypeLib;

        internal TFutureStorage MruColors; //only xlsx

#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
        internal TCustomXMLDataStorageList CustomXMLData;
#endif

        internal bool LoadingInterfaceHdr;
        private TBiff8XFMap FBiff8XF;

        internal TCopiedGen DrawingGen; //used when copying to make sure we don't repeat.
        #endregion

        #region Constructor
        internal TWorkbookGlobals(ExcelFile aWorkbook)
        {
            FWorkbook = aWorkbook;

            FFileEncryption = new TFileEncryption();
            FLel = new TMiscRecordList();
            FBoundSheets = new TBoundSheetList();
            FFnGroups = new TMiscRecordList();

            FWorkbookProtection = new TWorkbookProtection();
            FFonts = new TFontRecordList();
            FFormats = TFormatRecordList.Create();
            FStyleXF = new TXFRecordList();
            FCellXF = new TXFRecordList();
            FDXF = new TDXFRecordList();
            FStyles = new TStyleRecordList();
            FTableStyles = new TTableStyleRecordList();

            FPivotCache = new TMiscRecordList();
#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
            FXlsxPivotCache = new TXlsxPivotCacheList();
            FXlsxConnections = null;
#endif
            FDocRoute = new TMiscRecordList();
            FUserBView = new TMiscRecordList();

            FMetaData = new TMiscRecordList();
            FNames = new TNameRecordList();
            FRealTimeData = new TMiscRecordList();

            FReferences = new TReferences();

            FHeaderImages = new TDrawingGroup(xlr.HEADERIMG, 14);
            FDrawingGroup = new TDrawingGroup(xlr.MSODRAWINGGROUP, 0);

            FSST = new TSST();
            FWebPub = new TMiscRecordList();

            FFeatHdr = new TMiscRecordList();
            FDConn = new TMiscRecordList();
            
            FBorders = new TBorderList();
            FPatterns = new TPatternList();
            FFutureRecords = new TMiscRecordList();
            StylesFutureStorage = null;

            CalcOptions = new TCalcOptions();
            ThemeRecord = new TThemeRecord();

#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
            CustomXMLData = new TCustomXMLDataStorageList();
#endif

        }
        #endregion

        #region Public collections
        internal ExcelFile Workbook { get { return FWorkbook; } }
        internal TFileEncryption FileEncryption { get { return FFileEncryption; } }
        internal TMiscRecordList Lel { get { return FLel; } }
        internal TBoundSheetList BoundSheets { get { return FBoundSheets; } }
        internal TMiscRecordList FnGroups { get { return FFnGroups; } }

        internal TWorkbookProtection WorkbookProtection { get { return FWorkbookProtection; } }
        internal TFontRecordList Fonts { get { return FFonts; } }
        internal TFormatRecordList Formats { get { return FFormats; } }
        internal TXFRecordList StyleXF { get { return FStyleXF; } }
        internal TXFRecordList CellXF { get { return FCellXF; } }
        internal TDXFRecordList DXF { get { return FDXF; } }
        internal TStyleRecordList Styles { get { return FStyles; } }
        internal TTableStyleRecordList TableStyles { get { return FTableStyles; } }

        internal TMiscRecordList PivotCache { get { return FPivotCache; } }
#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
        internal TXlsxPivotCacheList XlsxPivotCache { get { return FXlsxPivotCache; } }
        internal string XlsxConnections { get { return FXlsxConnections; } set { FXlsxConnections = value; } }
#endif
        internal TMiscRecordList DocRoute { get { return FDocRoute; } }
        internal TMiscRecordList UserBView { get { return FUserBView; } }

        internal TMiscRecordList MetaData { get { return FMetaData; } }

        internal TReferences References { get { return FReferences; } }
        internal TNameRecordList Names { get { return FNames; } }
        internal TMiscRecordList RealTimeData { get { return FRealTimeData; } }
        internal TDrawingGroup HeaderImages { get { return FHeaderImages; } }
        internal TDrawingGroup DrawingGroup { get { return FDrawingGroup; } }

        internal TSST SST { get { return FSST; } }

        internal TMiscRecordList WebPub { get { return FWebPub; } }

        internal TMiscRecordList FeatHdr { get { return FFeatHdr; } }
        internal TMiscRecordList DConn { get { return FDConn; } }

        internal TBorderList Borders { get { return FBorders; } }
        internal TPatternList Patterns { get { return FPatterns; } }

        internal TMiscRecordList FutureRecords { get { return FFutureRecords; } }

        internal TTheme Theme { get { return ThemeRecord.Theme; } }

        public TBiff8XFMap Biff8XF { get { return FBiff8XF; } set { FBiff8XF = value; } }

        #endregion

        #region Workbook Properties

        internal bool IsXltTemplate { get { return FIsXltTemplate; } set { FIsXltTemplate = value; } }

        internal int ActiveSheet { get { return Window1[0].ActiveSheet; } set { Window1[0].ActiveSheet = value; } }

        internal string GetSheetName(int SheetIndex)
        {
            return FBoundSheets.BoundSheets.GetSheetName(SheetIndex);
        }
        
        internal void SetSheetName(int SheetIndex, string Value)
        {
            string RealName=TSheetNameList.MakeValidSheetName(Value);
            FBoundSheets.SheetNames.Rename(FBoundSheets.BoundSheets.GetSheetName(SheetIndex), RealName);
            FBoundSheets.BoundSheets.SetSheetName(SheetIndex, RealName);
        }

        internal void SetFirstSheetVisible(int Index)
        {
            Window1[0].FirstSheetVisible=Index;
        }

        internal int GetFirstSheetVisible()
        {
            return Window1[0].FirstSheetVisible;
        }

        internal TXlsSheetVisible GetSheetVisible(int SheetIndex)
        {
            return FBoundSheets.BoundSheets.GetSheetVisible(SheetIndex);
        }

        internal void SetSheetVisible(int SheetIndex, TXlsSheetVisible Value)
        {
            FBoundSheets.BoundSheets.SetSheetVisible(SheetIndex, Value);
            //Window1[0].SetSheetVisible(Value);  //NO! It would hid the full workbook!
        }

        internal int SheetCount{get{ return FBoundSheets.BoundSheets.Count;}}
        
        internal int SheetOptionFlags(int SheetIndex) 
        {
            return FBoundSheets.BoundSheets[SheetIndex].OptionFlags;
        }

        internal void SheetSetOffset(int SheetIndex, long Offset)
        {
            FBoundSheets.BoundSheets[SheetIndex].SetOffset(Offset);
        }

        internal string SheetRelationshipId(int SheetIndex)
        {
            return FBoundSheets.BoundSheets[SheetIndex].XlsxRelationshipId;
        }

        private const TSheetWindowOptions AllOptions =
            TSheetWindowOptions.HideWindow |
            TSheetWindowOptions.MinimizeWindow |
            TSheetWindowOptions.ShowHorizontalScrollBar |
            TSheetWindowOptions.ShowSheetTabBar |
            TSheetWindowOptions.ShowVerticalScrollBar; 
        
        internal TSheetWindowOptions WindowOptions
        {
            get
            {
                return ((TSheetWindowOptions)Window1[0].Options) & AllOptions;
            }
            set
            {
                TSheetWindowOptions FilteredOptions = value & AllOptions;
                TSheetWindowOptions ExistingOptions = ((TSheetWindowOptions) Window1[0].Options) & (~AllOptions);
                Window1[0].Options = (int)(ExistingOptions | FilteredOptions); 
            }
        }

        internal bool HasMacro
        {
            get
            {
                return ObjProj != null;
            }
            set
            {
                if ((value==false) && (HasMacro))
                {
                    ObjProj = null;
                    ObNoMacros = null;
                }
                else if (value && !HasMacro)
                {
                    ObjProj = new TObProjRecord((int)xlr.OBPROJ, new byte[]{});
                }
            }
        }

        internal string CodeName
        {
            get
            {
                if (CodeNameRecord == null) return String.Empty;
                return CodeNameRecord.SheetName;
            }
            set
            {
                if (value == null || value.Length == 0) CodeNameRecord = null;
                else CodeNameRecord = new TCodeNameRecord(value);
            }
        }

        internal bool SaveExternalLinkValues
        {
            get
            {
                if (BookBool != null) return BookBool.SaveExternalLinkValues; else return true;
            }
            set
            {
                EnsureBookBool();
                BookBool.SaveExternalLinkValues = value;
            }
        }

        internal bool HasEnvelope
        {
            get
            {
                if (BookBool != null) return BookBool.GetFlag(TBookBoolOption.HasEnvelope); else return false;
            }
            set
            {
                EnsureBookBool();
                BookBool.SetFlag(TBookBoolOption.HasEnvelope, value);
            }
        }

        internal bool EnvelopeVisible
        {
            get
            {
                if (BookBool != null) return BookBool.GetFlag(TBookBoolOption.EnvelopeVisible); else return false;
            }
            set
            {
                EnsureBookBool();
                BookBool.SetFlag(TBookBoolOption.EnvelopeVisible, value);
            }
        }

        internal bool EnvelopeInitDone
        {
            get
            {
                if (BookBool != null) return BookBool.GetFlag(TBookBoolOption.EnvelopeInitDone); else return false;
            }
            set
            {
                EnsureBookBool();
                BookBool.SetFlag(TBookBoolOption.EnvelopeInitDone, value);
            }
        }

        internal bool HideBorderUnselLists
        {
            get
            {
                if (BookBool != null) return BookBool.GetFlag(TBookBoolOption.HideBorderUnselLists); else return false;
            }
            set
            {
                EnsureBookBool();
                BookBool.SetFlag(TBookBoolOption.HideBorderUnselLists, value);
            }
        }

        internal TUpdateLinkOption UpdateLinks
        {
            get
            {
                if (BookBool != null) return BookBool.UpdateLinks; else return TUpdateLinkOption.PromptUser;
            }
            set
            {
                EnsureBookBool();
                BookBool.UpdateLinks = value;
            }
        }

        private void EnsureBookBool()
        {
            if (BookBool == null) BookBool = new TBookBoolRecord();
        }

        internal void AddBoundSheetFromFile(int TabId, TBoundSheetRecord bs)
        {
            FBoundSheets.AddFromFile(bs);
            FBoundSheets.AddTabIdFromFile((uint)TabId);
            
        }

        internal void DeleteCountry()
        {
            Country = null;
        }

        internal int MultithreadRecalc
        {
            get
            {
                if (MTRSettings == null) return -1;
                return MTRSettings.NumberOfThreads;
            }
            set
            {
                if (value < 0)
                {
                    MTRSettings = null;
                    return;
                }

                if (MTRSettings == null) MTRSettings = new TMTRSettingsRecord();
                MTRSettings.NumberOfThreads = value;
            }
        }

        internal bool ForceFullRecalc
        {
            get
            {
                if (ForceFullCalculation == null) return false;
                return ForceFullCalculation.FullRecalc;
            }
            set
            {
                if (!value)
                {
                    ForceFullCalculation = null;
                    return;
                }

                if (ForceFullCalculation == null) ForceFullCalculation = new TForceFullCalculationRecord();
                ForceFullCalculation.FullRecalc = value;
            }
        }

        internal bool AutoCompressPictures
        {
            get
            {
                if (CompressPictures == null) return false;
                return CompressPictures.Compression;
            }
            set
            {
                if (!value)
                {
                    CompressPictures = null;
                    return;
                }

                if (CompressPictures == null) CompressPictures = new TCompressPicturesRecord();
                CompressPictures.Compression = value;
            }
        }

        internal bool CheckCompatibility
        {
            get
            {
                if (Compat12 == null) return true;
                return Compat12.CompatCheck;
            }
            set
            {
                if (value)
                {
                    Compat12 = null;
                    return;
                }

                if (Compat12 == null) Compat12 = new TCompat12Record();
                Compat12.CompatCheck = value;
            }
        }

        #endregion

        #region Palette
        internal Color GetRgbColorPalette(int Index)
        {
            TPaletteRecord p = Palette == null ? TPaletteRecord.StandardPalette : Palette;
            return p.GetRgbColor(Index);
        }

        internal TLabColor GetLabColorPalette(int Index)
        {
            TPaletteRecord p = Palette == null ? TPaletteRecord.StandardPalette : Palette;
            return p.GetLabColor(Index);
        }

        internal void SetColorPalette(int Index, Color Color)
        {
            if (Palette==null)
            {
                //We have to create a standard palette first.
                Palette = TPaletteRecord.CreateStandard();
            }

            Palette.SetColor(Index, Color);
        }

        internal bool PaletteContainsColor(Color value)
        {
            TPaletteRecord p = Palette == null ? TPaletteRecord.StandardPalette : Palette;
            return p.ContainsColor(value);
        }
        #endregion

        #region Biff8XF
        internal TBiff8XFGuard GetBiff8XFGuard()
        {
            return new TBiff8XFGuard(this);
        }

        internal void InitBiff8X()
        {
            FBiff8XF = new TBiff8XFMap(StyleXF, CellXF);
        }

        internal void DestroyBiff8X()
        {
            FBiff8XF = null;
        }

        #endregion

        #region Load

        internal void Clear()
        {
            sBOF = null;
            sEOF = null;

            FFileEncryption.Clear();
            FLel.Clear();
            FBoundSheets.Clear();
            FFnGroups.Clear();

            FWorkbookProtection.Clear();
            FFonts.Clear();
            FFormats = TFormatRecordList.Create();
            FStyleXF.Clear();
            FCellXF.Clear();
            FDXF.Clear();
            FStyles.Clear();
            FTableStyles.Clear();

            FPivotCache.Clear();
#if(FRAMEWORK30 && !COMPACTFRAMEWORK)
            FXlsxPivotCache.Clear();
            FXlsxConnections = null;
#endif
            FDocRoute.Clear();
            FUserBView.Clear();

            FMetaData.Clear();
            FNames.Clear();
            FRealTimeData.Clear();

            FReferences.Clear();

            FHeaderImages.Clear();
            FDrawingGroup.Clear();

            FSST.Clear();
            FWebPub.Clear();

            FFeatHdr.Clear();
            FDConn.Clear();

            FBorders.Clear();
            FPatterns.Clear();
            FFutureRecords.Clear();
            StylesFutureStorage = null;

            FIsXltTemplate = false;
            CodePage = null;
            Excel9File = null;
            ObjProj = null;
            ObNoMacros = null;
            CodeNameRecord = null;

            OleObjectSize = null;

            Window1 = null;
            Backup = false;
            HideObj = THideObj.ShowAll;
            Dates1904 = false;
            PrecisionAsDisplayed = false;
            RefreshAll = false;
            BookBool = null;


            Palette = null;
            ClrtClient = null;

            UsesELFs = false;

            MTRSettings = null;
            ForceFullCalculation = null;

            Country = null;
            RecalcId = null;
            WOpt = null;
            BookExt = null;

            ThemeRecord = new TThemeRecord();

#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
            CustomXMLData.Clear();
#endif

            CalcOptions = new TCalcOptions();

            CompressPictures = null;
            Compat12 = null;
            GUIDTypeLib = null;

            MruColors = null; //only xlsx

            LoadingInterfaceHdr = false;
        }

        internal override void LoadFromStream(TBaseRecordLoader RecordLoader, TBOFRecord First)
        {
            TWorkbookLoader WorkbookLoader = new TWorkbookLoader(RecordLoader);
        
            int RecordId;
            do
            {
                RecordId=RecordLoader.RecordHeader.Id;
                TBaseRecord R=RecordLoader.LoadRecord(true);

                if (R != null) R.LoadIntoWorkbook(this, WorkbookLoader);

            } while (RecordId != (int)xlr.EOF);

            ThemeRecord.LoadFromBiff8(); //done after continues are loaded.

            if (WorkbookLoader.XFCount == RecordLoader.XFCount && WorkbookLoader.XFCRC == RecordLoader.XFCRC && WorkbookLoader.XFExtList.Count > 0
                && RecordLoader.XlsBiffVersion != TXlsBiffVersion.Excel2003) //We won't check First.BiffVersion to see if this was saved with Excel 2007. CRC should be enough.
            {
                Biff8XF.AddExt(WorkbookLoader.XFExtList, this);
            }

            EnsureRequiredRecords();
            FStyles.AddBiff8Outlines();
            if (First != null) sBOF=First; //Last statement
        }

        public void AddNewWindow1()
        {
            if (Window1 == null)
            {
                Window1 = new TWindow1Record[1];
            }
            else
            {
                //I have never seen this happen, but it could according to docs. In any case, it will be slow but it doesn't matter.
                TWindow1Record[] OldWin1 = Window1;
                Window1 = new TWindow1Record[OldWin1.Length + 1];
                Array.Copy(OldWin1, 0, Window1, 0, OldWin1.Length);
            }

        }

        public void EnsureRequiredRecords()
        {
            if (sBOF == null) sBOF = TBOFRecord.CreateEmptyWorkbook(4);
            if (sEOF == null) sEOF = new TEOFRecord();
            if (Window1 == null)
            {
                Window1 = new TWindow1Record[1];
                Window1[0] = new TWindow1Record();
            }

            CellXF.EnsureMinimumCellXFs(this);
            StyleXF.EnsureMinimumStyles(Styles, this);
        }

        #endregion

        #region Total Size
        private int ts(TBaseRecord[] Records)
        {
            if (Records == null) return 0;
            int Result = 0;
            foreach (TBaseRecord r in Records)
            {
                Result += ts(r);
            }
            return Result;
        }

        private int ts(TBaseRecord r)
        {
            if (r == null) return 0;
            return r.TotalSize();
        }

        internal long TotalSize(TEncryptionData Encryption, bool Repeatable, int SheetIndex)
        {
            return
                FileEncryption.TotalSize() +
                Encryption.TotalSize() +

                TTemplateRecord.GetSize(FIsXltTemplate) +
                ts(CodePage) +

                FLel.TotalSize +
                TDSFRecord.StandardSize() +
                ts(Excel9File) +

                (SheetIndex < 0 ? FBoundSheets.TabIdsTotalSize() : 0) 
                +
                ts(GetMacroRec(ObjProj)) +
                ts(GetMacroRec(ObNoMacros)) +
                ts(CodeNameRecord) +

                FFnGroups.TotalSize +
                ts(OleObjectSize) +
                FWorkbookProtection.TotalSize() +
                ts(Window1) +
                TBackupRecord.StandardSize() +
                THideObjRecord.StandardSize() +
                T1904Record.StandardSize() +

                TPrecisionRecord.StandardSize() +
                TRefreshAllRecord.StandardSize() +

                ts(BookBool) +
                Fonts.TotalSize +
                Formats.TotalSize +
                StyleXF.SizeWithXFExt(this, CellXF) +
                DXF.TotalSize +
                Styles.TotalSize +
                TableStyles.TotalSize +
                ts(Palette) +
                ts(ClrtClient) +

                PivotCache.TotalSize +
                DocRoute.TotalSize +

                UserBView.TotalSize +
                TUsesELFsRecord.StandardSize()+

                (SheetIndex < 0 ?
                  FBoundSheets.TotalSize() +
                  MetaData.TotalSize
                :
                  FBoundSheets.TotalRangeSize(SheetIndex)
                )

                +

                ts(MTRSettings) +
                ts(ForceFullCalculation) +
                ts(Country) +

                FReferences.TotalSize() +
                FNames.TotalSize +
                RealTimeData.TotalSize +
                ts(RecalcId) +

                (SheetIndex < 0 ?
                  FHeaderImages.TotalSize() +
                  FDrawingGroup.TotalSize()
                :
                  0)
                  +

                FSST.TotalSize(Repeatable) +
                WebPub.TotalSize +
                ts(WOpt) +
                //CrErr is ignored
                ts(BookExt) +
                FeatHdr.TotalSize +
                DConn.TotalSize +
                ts(ThemeRecord) +
                ts(CompressPictures) +
                ts(Compat12) +
                ts(GUIDTypeLib) +
                FFutureRecords.TotalSize;
        }

        internal override long TotalSize(TEncryptionData Encryption, bool Repeatable)
        {
            return base.TotalSize(Encryption, Repeatable) +
                TotalSize(Encryption, Repeatable, -1);
        }

        internal override long TotalRangeSize(int SheetIndex, TXlsCellRange CellRange, TEncryptionData Encryption, bool Repeatable)
        {
            return base.TotalRangeSize(SheetIndex, CellRange, Encryption, Repeatable) +
                TotalSize(Encryption, Repeatable, SheetIndex);
        }
        #endregion

        #region Save
        private void SaveToStream(IDataStream DataStream, TSaveData SaveData, int SheetIndex)
        {
            if ((sBOF == null) || (sEOF == null)) XlsMessages.ThrowException(XlsErr.ErrSectionNotLoaded);

            sBOF.SaveToStream(DataStream, SaveData, 0);
            FileEncryption.SaveFirstPart(DataStream, SaveData); //WriteProt
            if (DataStream.Encryption.Engine != null) //FilePass
            {
                byte[] Fp = DataStream.Encryption.Engine.GetFilePassRecord();
                DataStream.WriteRaw(Fp, Fp.Length);
            }

            if (IsXltTemplate) TTemplateRecord.SaveNewRecord(DataStream);

            FileEncryption.SaveSecondPart(DataStream, SaveData);

            TGlobalRecordSaver g = new TGlobalRecordSaver(DataStream, SaveData);
            g.SaveRecord(CodePage);

            FLel.SaveToStream(DataStream, SaveData, 0);
            TDSFRecord.SaveDSF(DataStream);
            g.SaveRecord(Excel9File);

            if (SheetIndex < 0) FBoundSheets.SaveTabIds(DataStream);
            g.SaveRecord(GetMacroRec(ObjProj));
            g.SaveRecord(GetMacroRec(ObNoMacros));
            g.SaveRecord(CodeNameRecord);

            FFnGroups.SaveToStream(DataStream, SaveData, 0);
            g.SaveRecord(OleObjectSize);
            FWorkbookProtection.SaveToStream(DataStream, SaveData);
            foreach (TWindow1Record w1 in Window1)
            {
                g.SaveRecord(w1);
            }
            TBackupRecord.SaveRecord(DataStream, Backup);
            THideObjRecord.SaveRecord(DataStream, HideObj);
            T1904Record.SaveRecord(DataStream, Dates1904);

            TPrecisionRecord.SaveRecord(DataStream, PrecisionAsDisplayed);
            TRefreshAllRecord.SaveRecord(DataStream, RefreshAll);

            g.SaveRecord(BookBool);

            Fonts.SaveToStream(DataStream, SaveData, 0);
            Formats.SaveToStream(DataStream, SaveData, 0);
            StyleXF.SaveAllToStream(DataStream, ref SaveData, CellXF);

            DXF.SaveToStream(DataStream, SaveData, 0);
            Styles.SaveToStream(DataStream, SaveData, 0);
            TableStyles.SaveToStream(DataStream, SaveData, 0);
            g.SaveRecord(Palette);
            g.SaveRecord(ClrtClient);

            PivotCache.SaveToStream(DataStream, SaveData, 0);
            DocRoute.SaveToStream(DataStream, SaveData, 0);

            UserBView.SaveToStream(DataStream, SaveData, 0);
            TUsesELFsRecord.SaveRecord(DataStream, UsesELFs);

            if (SheetIndex < 0)
            {
                FBoundSheets.SaveToStream(DataStream, SaveData);
                MetaData.SaveToStream(DataStream, SaveData, 0);
            }
            else
            {
                FBoundSheets.SaveRangeToStream(DataStream, SaveData, SheetIndex);
            }

            g.SaveRecord(MTRSettings);
            g.SaveRecord(ForceFullCalculation);
            g.SaveRecord(Country);
            
            FReferences.SaveToStream(DataStream, SaveData);
            FNames.SaveToStream(DataStream, SaveData, 0); //Should be after FBoundSheets.SaveToStream
            RealTimeData.SaveToStream(DataStream, SaveData, 0);
            g.SaveRecord(RecalcId);

            if (SheetIndex < 0)
            {
                FHeaderImages.SaveToStream(DataStream, SaveData);
                FDrawingGroup.SaveToStream(DataStream, SaveData);
            }

            FSST.SaveToStream(DataStream, SaveData);
            WebPub.SaveToStream(DataStream, SaveData, 0);
            g.SaveRecord(WOpt);
            //CrErr is ignored
            g.SaveRecord(BookExt);
            FeatHdr.SaveToStream(DataStream, SaveData, 0);
            DConn.SaveToStream(DataStream, SaveData, 0);
            ThemeRecord.SaveToStream(DataStream, SaveData, 0);
            g.SaveRecord(CompressPictures);
            g.SaveRecord(Compat12);
            g.SaveRecord(GUIDTypeLib);
            FFutureRecords.SaveToStream(DataStream, SaveData, 0);
            sEOF.SaveToStream(DataStream, SaveData, 0);
        }

        private TBaseRecord GetMacroRec(TBaseRecord Rec)
        {
            if (FWorkbook.HasMacroXlsm()) return null;
            return Rec;
        }

        internal override void SaveToStream(IDataStream DataStream, TSaveData SaveData)
        {
            SaveToStream(DataStream, SaveData, -1);
        }

        internal override void SaveRangeToStream(IDataStream DataStream, TSaveData SaveData, int SheetIndex, TXlsCellRange CellRange)
        {
            //Someday this can be optimized to only save SST for labels used on the range
            //But even Excel does not do it...
            SaveToStream(DataStream, SaveData, SheetIndex);
        }

        internal override void SaveToPxl(TPxlStream PxlStream, TPxlSaveData SaveData)
        {
            if ((sBOF==null)||(sEOF==null)) XlsMessages.ThrowException(XlsErr.ErrSectionNotLoaded);

            sBOF.SaveToPxl(PxlStream, 0, SaveData);

            foreach (TWindow1Record w1 in Window1)
            {
                if (w1 != null) w1.SaveToPxl(PxlStream, 0, SaveData);
            }

            Fonts.SaveToPxl(PxlStream, 0, SaveData);
            Formats.SaveToPxl(PxlStream, 0, SaveData);
            CellXF.SaveToPxl(PxlStream, 0, SaveData);
            
            FBoundSheets.SaveToPxl(PxlStream, SaveData);
            FNames.SaveToPxl(PxlStream, 0, SaveData); //Should be after FBoundSheets.SaveToStream
            sEOF.SaveToPxl(PxlStream, 0, SaveData);
        }
        #endregion

        #region InsertAndCopy
        internal void InsertAndCopyRange(TXlsCellRange SourceRange, TFlxInsertMode InsertMode, int DestRow, int DestCol, int aRowCount, int aColCount, TSheetInfo SheetInfo)
        {
            FNames.ArrangeInsertRange(SourceRange.OffsetForIns(DestRow, DestCol, InsertMode), aRowCount, aColCount, SheetInfo);
#if(FRAMEWORK30 && !COMPACTFRAMEWORK)
            FXlsxPivotCache.ArrangeInsertRange(SourceRange.OffsetForIns(DestRow, DestCol, InsertMode), aRowCount, aColCount, SheetInfo);
#endif
        }
        
        internal void DeleteRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
        {
            FNames.ArrangeInsertRange(CellRange, -aRowCount, -aColCount, SheetInfo);
#if(FRAMEWORK30 && !COMPACTFRAMEWORK)
            FXlsxPivotCache.ArrangeInsertRange(CellRange, -aRowCount, -aColCount, SheetInfo);
#endif
        }

        internal void MoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
        {
            FNames.ArrangeMoveRange(CellRange, NewRow, NewCol, SheetInfo);
#if(FRAMEWORK30 && !COMPACTFRAMEWORK)
            FXlsxPivotCache.ArrangeMoveRange(CellRange, NewRow, NewCol, SheetInfo);
#endif
        }

        internal void DeleteSheets(int SheetIndex, int SheetCount, TWorkbook Workbook)
        {
            if (HasMacro) XlsMessages.ThrowException(XlsErr.ErrCantDeleteSheetWithMacros);  //If we delete a sheet that has a corresponding macro on the vba stream, Excel 2000 will crash when opening the file. Excel Xp seems to handle this ok.
            for (int i=0; i< SheetCount;i++)
                FBoundSheets.DeleteSheet(SheetIndex);
            FReferences.InsertSheets(SheetIndex, -SheetCount);
            FNames.DeleteSheets(SheetIndex, SheetCount, Workbook);
        }

        internal void UpdateDeletedRanges(int SheetIndex, int SheetCount, TDeletedRanges DeletedRanges)
        {
            FNames.UpdateDeletedRanges(SheetIndex, SheetCount, DeletedRanges);
#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
            XlsxPivotCache.UpdateDeletedRanges(SheetIndex, SheetCount, DeletedRanges);
#endif

        }

        internal void InsertSheets(int CopyFrom, int BeforeSheet, int OptionFlags, string Name, int SheetCount, TSheetList SheetList)
        {
            for (int i=0; i< SheetCount;i++)
                FBoundSheets.InsertSheet(BeforeSheet + i, OptionFlags, Name);

            FReferences.InsertSheets(BeforeSheet, SheetCount);
        }
        #endregion

        #region Misc

        internal void MergeFromPxlGlobals(TWorkbookGlobals SourceGlobals)
        {
            sBOF = SourceGlobals.sBOF;
            FSST = SourceGlobals.FSST;
            FReferences = SourceGlobals.FReferences;
            FBoundSheets = SourceGlobals.FBoundSheets;

            //Pxl doesn't have styles.
            CellXF.MergeFromPxlXF(SourceGlobals.CellXF, Fonts.Count - 1, this, SourceGlobals);  //-1 because fonts[0] will be merged
            FFonts.MergeFromPxlFont(SourceGlobals.Fonts);
            //Formats are added in FXF.Merge
            
            FNames = SourceGlobals.FNames;
            Window1 = SourceGlobals.Window1;

            CodePage = SourceGlobals.CodePage;
            Country = SourceGlobals.Country;
        }

        internal int AddStyleFormat(TFlxFormat format, string StyleName)
        {
            if (!format.IsStyle)
            {
                format = (TFlxFormat)format.Clone();
                format.IsStyle = true;
            }

            int OldFmt = Styles.GetStyle(StyleName);
            if (OldFmt >= 0 && OldFmt < StyleXF.Count) 
            {
                TXFRecord XF1 = new TXFRecord(format, OldFmt == 0, this, false);
                StyleXF[OldFmt] = XF1;
                CellXF.UpdateChangedStyleInCellXF(OldFmt, XF1, false);
                return OldFmt;
            }

            TXFRecord XFRec = new TXFRecord(format, false, this, false);
            StyleXF.Add(XFRec);
            return StyleXF.Count - 1;
        }

        internal TFlxFormat GetStyleFormat(int XFid)
        {
            if ((XFid < 0) || (XFid > StyleXF.Count)) XFid = 0;
            return StyleXF[XFid].FlxFormat(Styles, Fonts, Formats, Borders, Patterns);
        }

        internal TFlxFormat GetCellFormat(int XFid)
        {
            if ((XFid<0) || (XFid> CellXF.Count)) XFid = 0;
            return CellXF[XFid].FlxFormat(Styles, Fonts, Formats, Borders, Patterns);
        }

        internal void AddStylesFutureStorage(TFutureStorageRecord R)
        {
            TFutureStorage.Add(ref StylesFutureStorage, R);
        }

        #endregion

    }


    internal class TCalcOptions
    {
        public TSheetCalcMode CalcMode;
        public int CalcCount;
        public bool A1RefMode;
        public bool IterationEnabled;
        public double Delta;
        public bool SaveRecalc;

        internal TCalcOptions()
        {
            CalcMode = TSheetCalcMode.Automatic;
            CalcCount = 100;
            A1RefMode = true;
            Delta = 0.001;
            SaveRecalc = true;
        }

        internal void SaveToStream(IDataStream DataStream)
        {
            TCalcModeRecord.SaveRecord(DataStream, CalcMode);
            TCalcCountRecord.SaveRecord(DataStream, CalcCount);
            TRefModeRecord.SaveRecord(DataStream, A1RefMode);
            TIterationRecord.SaveRecord(DataStream, IterationEnabled);
            TDeltaRecord.SaveRecord(DataStream, Delta);
            TSaveRecalcRecord.SaveRecord(DataStream, SaveRecalc);
        }

        internal static int TotalSize()
        {
            return
                    TCalcModeRecord.StandardSize() +
                    TCalcCountRecord.StandardSize() +
                    TRefModeRecord.StandardSize() +
                    TIterationRecord.StandardSize() +
                    TDeltaRecord.StandardSize() +
                    TSaveRecalcRecord.StandardSize();
        }
    }

    internal class TWorkbookLoader
    {
        internal int XFCount;
        internal UInt32 XFCRC;
        internal TXFExtRecordList XFExtList;
        internal TBaseRecordLoader RecordLoader;

        internal TWorkbookLoader(TBaseRecordLoader aRecordLoader)
        {
            XFCount = -1;
            XFCRC = 0;
            XFExtList = new TXFExtRecordList();
            RecordLoader = aRecordLoader;
        }
    }

    internal class TGlobalRecordSaver
    {
        IDataStream DataStream;
        TSaveData SaveData;

        internal TGlobalRecordSaver(IDataStream aDataStream, TSaveData aSaveData)
        {
            DataStream = aDataStream;
            SaveData = aSaveData;
        }

        internal void SaveRecord(TBaseRecord r)
        {
            if (r != null) r.SaveToStream(DataStream, SaveData, 0);
        }
    }

    /// <summary>
    /// A small class to make sure we set the Biff8XF member to null once it has been used to load a file.
    /// </summary>
    internal sealed class TBiff8XFGuard : IDisposable
    {
        TWorkbookGlobals Globals;

        internal TBiff8XFGuard(TWorkbookGlobals aGlobals)
        {
            aGlobals.InitBiff8X();
            Globals = aGlobals;
        }

        #region IDisposable Members

        public void Dispose()
        {
            if (Globals != null) Globals.DestroyBiff8X();
            GC.SuppressFinalize(this);
        }

        #endregion
    }

}
