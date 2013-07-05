using System;
using System.Text;
using FlexCel.Core;

namespace FlexCel.XlsAdapter
{
    /// <summary>
    ///Record IDs.
    /// </summary>
    internal enum xlb
    {
        Globals = 0x0005,
        Worksheet = 0x0010,
        Chart = 0x0020,
        Macro = 0x0040
    }

    internal enum xlr : int
    {
        BofVersion = 0x0600,
        INTEGER = 0x0002,
        FORMULA = 0x0006,
        FORMULABiff4 = 0x0406,
        EOF = 0x000A,

        CALCCOUNT = 0x000C,
        CALCMODE = 0x000D,
        PRECISION = 0x000E,
        REFMODE = 0x000F,
        DELTA = 0x0010,
        ITERATION = 0x0011,
        PROTECT = 0x0012,
        PASSWORD = 0x0013,
        HEADER = 0x0014,
        FOOTER = 0x0015,

        HEADERFOOTER = 0x089C,

        EXTERNCOUNT = 0x0016,
        EXTERNSHEET = 0x0017,
        NAME = 0x0018,

        WINDOWPROTECT = 0x0019,
        VERTICALPAGEBREAKS = 0x001A,
        HORIZONTALPAGEBREAKS = 0x001B,
        NOTE = 0x001C,
        SELECTION = 0x001D,

        FORMATCOUNT = 0x001F,
        COLUMNDEFAULT = 0x0020,

        ARRAY2 = 0x0021,

        x1904 = 0x0022,

        COLWIDTH = 0x0024,

        LEFTMARGIN = 0x0026,
        RIGHTMARGIN = 0x0027,
        TOPMARGIN = 0x0028,
        BOTTOMMARGIN = 0x0029,
        PRINTHEADERS = 0x002A,
        PRINTGRIDLINES = 0x002B,
        FILEPASS = 0x002F,

        PRINTSIZE = 0x0033,

        CONTINUE = 0x003C,
        CONTINUEFRT = 0x0812,
        CONTINUEFRT11 = 0x875,
        CONTINUEFRT12 = 0x87F,
        CONTINUECRTMLFRT = 0x089F,

        WINDOW1 = 0x003D,

        BACKUP = 0x0040,
        PANE = 0x0041,
        CODEPAGE = 0x0042,

        IXFE = 0x0044,
        PLS = 0x004D,
        DCON = 0x0050,
        DCONREF = 0x0051,
        DCONNAME = 0x0052,
        DEFCOLWIDTH = 0x0055,
        BUILTINFMTCNT = 0x0056,
        XCT = 0x0059,
        CRN = 0x005A,
        FILESHARING = 0x005B,
        WRITEACCESS = 0x005C,
        OBJ = 0x005D,
        UNCALCED = 0x005E,
        SAVERECALC = 0x005F,
        TEMPLATE = 0x0060,
        OBJPROTECT = 0x0063,
        COLINFO = 0x007D,


        //MetaData
        MDB = 2186,
        MDTInfo = 2180,
        MDXKPI = 2185,
        MDXProp = 2184, 
        MDXSet = 2183,
        MDXStr = 2181,
        MDXTuple = 2182,

        RTD = 0x0813,
        DCONN = 0x876,

        IMDATA = 0x007F,
        GUTS = 0x0080,
        WSBOOL = 0x0081,
        GRIDSET = 0x0082,
        HCENTER = 0x0083,
        VCENTER = 0x0084,
        BOUNDSHEET = 0x0085,
        WRITEPROT = 0x0086,
        ADDIN = 0x0087,
        EDG = 0x0088,
        PUB = 0x0089,
        COUNTRY = 0x008C,
        HIDEOBJ = 0x008D,
        BUNDLESOFFSET = 0x008E,
        BUNDLEHEADER = 0x008F,
        
        SORT = 0x0090,
        SORTDATA = 0x0895,
        DROPDOWNOBJIDS = 0x0874,

        SUB = 0x0091,
        PALETTE = 0x0092,

        LHRECORD = 0x0094,
        LHNGRAPH = 0x0095,
        SOUND = 0x0096,
        LPR = 0x0098,
        PLV = 0x088B,
        STANDARDWIDTH = 0x0099,
        FNGROUPNAME = 0x009A,
        FNGRP12 = 0x0898,

        SYNC = 0x0097,

        FILTERMODE = 0x009B,
        FNGROUPCOUNT = 0x009C,
        AutoFilterINFO = 0x009D,
        AutoFilter = 0x009E,
        AutoFilter12 = 0x087E,
        SCL = 0x00A0,
        SETUP = 0x00A1,
        COORDLIST = 0x00A9,
        GCW = 0x00AB,
        SCENMAN = 0x00AE,
        SCENARIO = 0x00AF,
        SXVIEW = 0x00B0,
        SXVD = 0x00B1,
        SXVI = 0x00B2,
        SXIVD = 0x00B4,
        SXLI = 0x00B5,
        SXPI = 0x00B6,
        DOCROUTE = 0x00B8,
        RECIPNAME = 0x00B9,

        MULRK = 0x00BD,
        MULBLANK = 0x00BE,
        MMS = 0x00C1,
        ADDMENU = 0x00C2,
        DELMENU = 0x00C3,

        SXDI = 0x00C5,
        SXDB = 0x00C6,
        SXFIELD = 0x00C7,
        SXINDEXLIST = 0x00C8,
        SXDOUBLE = 0x00C9,
        SXSTRING = 0x00CD,
        SXDATETIME = 0x00CE,
        SXTBL = 0x00D0,
        SXTBRGITEM = 0x00D1,
        SXTBPG = 0x00D2,

        SXADDL = 0x0864,
        SXADDL12 = 0x0881,

        OBPROJ = 0x00D3,
        SXIDSTM = 0x00D5,
        RSTRING = 0x00D6,
        DBCELL = 0x00D7,
        BOOKBOOL = 0x00DA,
        SXEXTPARAMQRY = 0x00DC, //DBORPARAMQUERY in new docs
        SCENPROTECT = 0x00DD,
        OLESIZE = 0x00DE,
        UDDESC = 0x00DF,

        LRNG = 0x015F,
        RRSORT = 0x13F,

        INTERFACEHDR = 0x00E1,
        INTERFACEEND = 0x00E2,
        SXVS = 0x00E3,
        CELLMERGING = 0x00E5,
        BITMAP = 0x00E9,
        MSODRAWINGGROUP = 0x00EB,
        MSODRAWING = 0x00EC,
        MSODRAWINGSELECTION = 0x00ED,
        ENTEXU2 = 0x01C2,
        PHONETIC = 0x00EF,
        SXRULE = 0x00F0,
        SXEX = 0x00F1,
        SXFILT = 0x00F2,
        SXNAME = 0x00F6,
        SXSELECT = 0x00F7,
        SXPAIR = 0x00F8,
        SXFMLA = 0x00F9,
        SXFORMAT = 0x00FB,
        SST = 0x00FC,
        LABELSST = 0x00FD,
        EXTSST = 0x00FF,
        SXVDEX = 0x0100,
        SXFORMULA = 0x0103,
        SXDBEX = 0x0122,
        CHTRINSERT = 0x0137,
        CHTRINFO = 0x0138,
        CHTRCELLCONTENT = 0x013B,
        TABID = 0x013D,
        CHTRMOVERANGE = 0x0140,
        CHTRINSERTTAB = 0x014D,
        USESELFS = 0x0160,
        XL5MODIFY = 0x0162,
        CHTRHEADER = 0x0196,
        USERBVIEW = 0x01A9,
        USERSVIEWBEGIN = 0x01AA,
        USERSVIEWEND = 0x01AB,

        QSI = 0x01AD,
        QSIR = 0x0806,
        QSIF = 0x0807,

        SUPBOOK = 0x01AE,
        PROT4REV = 0x01AF,
        DSF = 0x0161,
        CONDFMT = 0x01B0,
        CF = 0x01B1,
        DVAL = 0x01B2,
        DCONBIN = 0x01B5,
        TXO = 0x01B6,
        REFRESHALL = 0x01B7,
        HLINK = 0x01B8,
        CODENAME = 0x01BA,
        SXFDBTYPE = 0x01BB,
        PROT4REVPASS = 0x01BC,
        DV = 0x01BE,
        XL9FILE = 0x01C0,
        RECALCID = 0x01C1,
        DIMENSIONS = 0x0200,
        BLANK = 0x0201,
        NUMBER = 0x0203,
        LABEL = 0x0204,
        BOOLERR = 0x0205,

        SXDXF = 0x00F4,
        SXITM = 0x00F5,
        SXVIEWEX = 0x080C,
        SXTH = 0x080D,
        SXPIEX = 0x080E,
        SXVDTEX = 0x080F,
        SXVIEWEX9 = 0x0810,

        FEAT = 0x0868,
        FEATHDR11= 0x0871,
        FEAT11 = 0x0872,
        FEAT12 = 0x0878,
        LIST12= 0x0877,

        QSISXTAG = 0x0802,
        DBQUERYEXT = 0x0803,
        EXTSTRING = 0x0804,
        TXTQRY = 0x0805,
        OLEDBCONN = 0x080A,

        STRING = 0x0207,
        ROW = 0x0208,

        BIGNAME = 0x0418,
        CONTINUEBIGNAME = 0x043C,

        INDEX = 0x020B,
        ARRAY = 0x0221,
        EXTERNNAME = 0x0223,
        EXTERNNAME2 = 0x0023,
        DEFAULTROWHEIGHT = 0x0225,
        FONT = 0x0031,
        TABLE = 0x0236,
        WINDOW2 = 0x023E,

        INTL = 0x0061,

        RK = 0x027E,
        STYLE = 0x0293,
        STYLEEX = 0x0892,

        xFORMAT = 0x041E,
        XF = 0x00E0,
        XFCRC = 0x087C,
        XFEXT = 0x087D,
        THEME = 0x0896,

        SHRFMLA = 0x04BC,
        SCREENTIP = 0x0800,
        WEBQRYSETTINGS = 0x0803,
        WEBQRYTABLES = 0x0804,
        BOF = 0x0809,

        SHEETEXT = 0x0862,
        BOOKEXT = 0x0863,
        HEADERIMG = 0x0866,

        FEATHDR = 0x0867,
        LEL = 0x01B9,

        DXF = 0x88D,
        TABLESTYLE = 0x88F,
        TABLESTYLES = 0x88E,
        TABLESTYLEELEMENT = 0x890,

        NAMECMT = 0x0894,

        OBNOMACROS = 0x01BD,
        CLRTCLIENT = 0x105C,
        FRTINFO = 0x0850,

        MTRSETTINGS = 0x089A,
        FORCEFULLCALCULATION = 0x08A3,
        
        WEBPUB = 0x0801,
        WOPT = 0x080B,

        CRERR = 0x0865,

        COMPAT12 = 0x88C,
        GUIDTYPELIB = 0x0897,

        COMPRESSPICTURES = 0x089B,

        CRTMLFRT = 0x089E,

        UNITS = 0x1001,
        ChartChart = 0x1002,
        ChartSeries = 0x1003,
        ChartDataformat = 0x1006,
        ChartLineformat = 0x1007,
        ChartMarkerformat = 0x1009,
        ChartAreaformat = 0x100A,
        ChartPieformat = 0x100B,
        ChartAttachedlabel = 0x100C,
        ChartSeriestext = 0x100D,
        ChartChartformat = 0x1014,
        ChartLegend = 0x1015,
        ChartSerieslist = 0x1016,
        ChartBar = 0x1017,
        ChartLine = 0x1018,
        ChartPie = 0x1019,
        ChartArea = 0x101A,
        ChartScatter = 0x101B,
        ChartChartline = 0x101C,
        ChartAxis = 0x101D,
        ChartTick = 0x101E,
        ChartValuerange = 0x101F,
        ChartCatserrange = 0x1020,
        ChartAxislineformat = 0x1021,
        ChartFormatlink = 0x1022,
        ChartDefaulttext = 0x1024,
        ChartText = 0x1025,
        ChartFontx = 0x1026,
        ChartObjectLink = 0x1027,
        ChartDataLabExtContent = 0x086B,
        ChartFrame = 0x1032,
        BEGIN = 0x1033,
        END = 0x1034,
        ChartPlotarea = 0x1035,
        Chart3D = 0x103A,
        ChartPicf = 0x103C,
        ChartDropbar = 0x103D,
        ChartRadar = 0x103E,
        ChartSurface = 0x103F,
        ChartRadararea = 0x1040,
        ChartAxisparent = 0x1041,
        ChartLegendxn = 0x1043,
        ChartShtprops = 0x1044,
        ChartSertocrt = 0x1045,
        ChartAxesused = 0x1046,
        ChartSbaseref = 0x1048,
        ChartSerparent = 0x104A,
        ChartSerauxtrend = 0x104B,
        ChartIfmt = 0x104E,
        ChartPos = 0x104F,
        ChartAlruns = 0x1050,
        ChartAI = 0x1051,
        ChartSerauxerrbar = 0x105B,
        ChartClrClient = 0x105C,
        ChartSerfmt = 0x105D,
        Chart3DDataFormat = 0x105F,
        ChartFbi = 0x1060,
        ChartBoppop = 0x1061,
        ChartAxcext = 0x1062,
        ChartDat = 0x1063,
        ChartPlotgrowth = 0x1064,
        ChartSiindex = 0x1065,
        ChartGelframe = 0x1066,
        ChartBoppcustom = 0x1067,
        ChartFbi2 = 0x1068,

        SXVIEWLINK = 0x0858,
        PIVOTCHARTBITS = 0x0859

    }

    ///////////////////////////////////////Object Types /////////////////////

    internal enum ft
    {
        End = 0x0000,
        Macro = 0x0004,
        Button = 0x0005,
        Gmo = 0x0006,
        Cf = 0x0007,
        PioGrbit = 0x0008,
        PictFmla = 0x0009,
        Cbls = 0x000A,
        Rbo = 0x000B,
        Sbs = 0x000C,
        Nts = 0x000D,
        SbsFmla = 0x000E,
        GboData = 0x000F,
        EdoData = 0x0010,
        RboData = 0x0011,
        CblsData = 0x0012,
        LbsData = 0x0013,
        CblsFmla = 0x0014,
        Cmo = 0x0015
    }

    //////////////////////////////////////Escher Records //////////////////

    internal enum Msofbt
    {
        DggContainer = 0xF000,
        Dgg = 0xF006,
        CLSID = 0xF016,
        OPT = 0xF00B,
        ColorMRU = 0xF11A,
        SplitMenuColors = 0xF11E,
        BstoreContainer = 0xF001,
        BSE = 0xF007,
        DgContainer = 0xF002,
        Dg = 0xF008,
        RegroupItem = 0xF118,
        ColorScheme = 0xF120,
        SpgrContainer = 0xF003,
        SpContainer = 0xF004,
        Spgr = 0xF009,
        Sp = 0xF00A,
        Textbox = 0xF00C,
        ClientTextbox = 0xF00D,
        Anchor = 0xF00E,
        ChildAnchor = 0xF00F,
        ClientAnchor = 0xF010,
        ClientData = 0xF011,
        OleObject = 0xF11F,
        DeletedPspl = 0xF11D,
        SolverContainer = 0xF005,
        ConnectorRule = 0xF012,
        AlignRule = 0xF013,
        ArcRule = 0xF014,
        ClientRule = 0xF015,
        CalloutRule = 0xF017,
        Selection = 0xF119
    }

    internal enum msoblip
    {
        /// <summary>
        /// An error occurred during loading
        /// </summary>
        ERROR = 0,
        /// <summary>
        /// An unknown blip type
        /// </summary>
        UNKNOWN = 1,
        /// <summary>
        /// Windows Enhanced Metafile
        /// </summary>
        EMF = 2,
        /// <summary>
        /// Windows Metafile
        /// </summary>
        WMF = 3,
        /// <summary>
        /// Macintosh PICT
        /// </summary>
        PICT = 4,
        /// <summary>
        /// JFIF
        /// </summary>
        JPEG = 5,
        /// <summary>
        /// PNG
        /// </summary>
        PNG = 6,
        /// <summary>
        /// Windows DIB
        /// </summary>
        DIB = 7
    }

    internal enum msobi
    {
        UNKNOWN = 0,
        WMF = 0x216,      // Metafile header then compressed WMF
        EMF = 0x3D4,      // Metafile header then compressed EMF
        PICT = 0x542,      // Metafile header then compressed PICT
        PNG = 0x6E0,      // One byte tag then PNG data
        JFIF = 0x46A,      // One byte tag then JFIF data
        //	JPEG = JFIF,
        DIB = 0x7A8,      // One byte tag then DIB data
        Client = 0x800      // Clients should set this bit
    }


    //////////////////////////////////////Tokens///////////////////////////
    /// <summary>
    /// Formula tokens
    /// </summary>
    internal sealed class XlsTokens
    {
        private XlsTokens() { }

        //Globals
        internal const int tk_Exp = 0x1;
        internal const int tk_Table = 0x2;
        internal static bool IsBinaryOp(byte b) { return (b >= 0x3) && (b <= 0x11); }
        internal static bool IsUnaryOp(byte b) { return (b >= 0x12) && (b <= 0x15); }

        //Constants
        internal const int tk_MissArg = 0x16;
        internal const int tk_Str = 0x17;
        internal const int tk_Attr = 0x19;
        internal const int tk_Err = 0x1C;
        internal const int tk_Bool = 0x1D;
        internal const int tk_Int = 0x1E;
        internal const int tk_Num = 0x1F;

        internal const int tk_MemArea = 0x26;
        internal const int tk_MemErr = 0x27;
        internal const int tk_MemNoMem = 0x28;
        internal const int tk_MemFunc = 0x29;

        //Func
        internal static bool Is_tk_Func(byte b) { return (b == 0x21) || (b == 0x41) || (b == 0x61); }
        internal static bool Is_tk_FuncVar(byte b) { return (b == 0x22) || (b == 0x42) || (b == 0x62); }

        //Operand
        internal static bool Is_tk_Array(byte b) { return (b == 0x20) || (b == 0x40) || (b == 0x60); }
        internal static bool Is_tk_Name(byte b) { return (b == 0x23) || (b == 0x43) || (b == 0x63); }
        internal static bool Is_tk_Ref(byte b) { return (b == 0x24) || (b == 0x44) || (b == 0x64); }
        internal static bool Is_tk_Area(byte b) { return (b == 0x25) || (b == 0x45) || (b == 0x65); }
        internal static bool Is_tk_RefErr(byte b) { return (b == 0x2A) || (b == 0x4A) || (b == 0x6A); }
        internal static bool Is_tk_AreaErr(byte b) { return (b == 0x2B) || (b == 0x4B) || (b == 0x6B); }
        internal static bool Is_tk_RefN(byte b) { return (b == 0x2C) || (b == 0x4C) || (b == 0x6C); }  //Reference relative to the current row. Can be < 0
        internal static bool Is_tk_AreaN(byte b) { return (b == 0x2D) || (b == 0x4D) || (b == 0x6D); }  //Area relative to the current row
        internal static bool Is_tk_NameX(byte b) { return (b == 0x39) || (b == 0x59) || (b == 0x79); }
        internal static bool Is_tk_Ref3D(byte b) { return (b == 0x3A) || (b == 0x5A) || (b == 0x7A); }
        internal static bool Is_tk_Area3D(byte b) { return (b == 0x3B) || (b == 0x5B) || (b == 0x7B); }
        internal static bool Is_tk_Ref3DErr(byte b) { return (b == 0x3C) || (b == 0x5C) || (b == 0x7C); }
        internal static bool Is_tk_Area3DErr(byte b) { return (b == 0x3D) || (b == 0x5D) || (b == 0x7D); }

        internal const int tk_RefToRefErr = 0x2A - 0x24;
        internal const int tk_AreaToAreaErr = 0x2B - 0x25;
        internal const int tk_Ref3DToRef3DErr = 0x3C - 0x3A;
        internal const int tk_Area3DToArea3DErr = 0x3D - 0x3B;
        internal const int tk_RefNToRefNErr = 0x2A - 0x2C; //there is no RefNErr. We will change it to RefErr
        internal const int tk_AreaNToAreaNErr = 0x2B - 0x2D; //there is no AreaNErr. We will change it to AreaErr

        internal static bool Is_tk_AnyRef(byte b)
        {
            return
                Is_tk_Name(b) ||
                Is_tk_Ref(b) ||
                Is_tk_Area(b) ||
                Is_tk_RefErr(b) ||
                Is_tk_AreaErr(b) ||
                Is_tk_RefN(b) ||
                Is_tk_AreaN(b) ||
                Is_tk_NameX(b) ||
                Is_tk_Ref3D(b) ||
                Is_tk_Area3D(b) ||
                Is_tk_Ref3DErr(b) ||
                Is_tk_Area3DErr(b);

        }

        internal static bool Is_tk_Operand(byte b)
        {
            /*return
                Is_tk_Array(b) ||
                Is_tk_Name(b) ||
                Is_tk_Ref(b) ||
                Is_tk_Area(b) ||
                Is_tk_RefErr(b) ||
                Is_tk_AreaErr(b) ||
                Is_tk_RefN(b) ||
                Is_tk_AreaN(b) ||
                Is_tk_NameX(b) ||
                Is_tk_Ref3D(b) ||
                Is_tk_Area3D(b) ||
                Is_tk_Ref3DErr(b) ||
                Is_tk_Area3DErr(b);
                */
            if (b >= 0x60 && b <= 0x7D) b -= 0x40;
            else
                if (b >= 0x40 && b <= 0x5D) b -= 0x20;
            return (b == 0x20 || (b >= 0x23 && b <= 0x25) || (b >= 0x2A && b <= 0x2D) || (b >= 0x39 && b <= 0x3D));
        }
    }


    internal sealed class XlsConsts
    {
        private XlsConsts() { }

        public const int SizeOfTRecordHeader = 4;

        public const int MaxCFRules2007 = int.MaxValue; //In Excel 2007
        public const int MaxCFRules97_2003 = 4; //In biff8

        public static int MaxCFRules { get { return FlxConsts.ExcelVersion == TExcelVersion.v97_2003 ? MaxCFRules97_2003 : MaxCFRules2007; } }

        public const int MaxXFDefs2007 = 0xFFFF; //In Excel 2007
        public const int MaxXFDefs97_2003 = 4000; //In biff8

        public const int MaxRowHeight = 8192; //In Excel 2003/2007

        public const int MaxTabIdCount = 4112;
        public const int MaxTabIdValue = 0xFFFE;

        public static int MaxXFDefs { get { return FlxConsts.ExcelVersion == TExcelVersion.v97_2003 ? MaxXFDefs97_2003 : MaxXFDefs2007; } }

        public const int MaxRecordDataSize = 8223;  //Real max is 8224... but I prefer to keep it safe.
        internal const int MaxExternSheetDataSize = 8220;  // 1370 records of 6 bytes each, and 2 bytes for the count

        internal const string WorkbookString = "Workbook"; //Do not localize
        internal const string DocumentPropertiesStringExtended = "\u0005DocumentSummaryInformation"; //Do not localize
        internal const string DocumentPropertiesString = "\u0005SummaryInformation"; //Do not localize
        internal const string ProjectString = "PROJECT"; //Do not localize

        internal static readonly string VBAMainStreamFullPath = (char)0 + "Root Entry" + (char)0 + "_VBA_PROJECT_CUR"; //Do not localize //STATIC*
        internal static readonly string[] VBAStreams = { "_VBA_PROJECT_CUR", "_VBA_PROJECT" }; //Do not localize //STATIC*
        internal static readonly string[] PropStreams = { DocumentPropertiesString, DocumentPropertiesStringExtended }; //Do not localize //STATIC*

        internal const int LowColorPaletteRange = 1;
        internal const int HighColorPaletteRange = 56;

        internal const int MaxHPageBreaks = 1026;
        internal const int MaxVPageBreaks = 1026;

        internal const int MinNumFormatId = 0x00A4;
        internal const int MaxNumFormatId = 0x017E;

        internal const string EmptyExcelPassword = "VelvetSweatshop";
    }

    internal static class XlsxConsts
    {
        internal const string ContentString = "EncryptedPackage"; //Do not localize
        internal const string EncryptionInfoString = "EncryptionInfo"; //Do not localize
    }

    internal sealed class XlsEscherConsts
    {
        private XlsEscherConsts() { }

        public const int SizeOfTEscherRecordHeader = 8;

        public static msoblip XlsImgConv(TXlsImgType img)
        {
            switch (img)
            {
                case TXlsImgType.Emf: return msoblip.EMF;
                case TXlsImgType.Wmf: return msoblip.WMF;
                case TXlsImgType.Jpeg: return msoblip.JPEG;
                case TXlsImgType.Png: return msoblip.PNG;
                case TXlsImgType.Bmp: return msoblip.DIB;
                default: return msoblip.UNKNOWN;
            }
        }

        public static int XlsBlipHeaderConv(TXlsImgType img)
        {
            switch (img)
            {
                case TXlsImgType.Emf: return 0xF01A;
                case TXlsImgType.Wmf: return 0xF01B;
                case TXlsImgType.Jpeg: return 0xF01D;
                case TXlsImgType.Png: return 0xF01E;
                case TXlsImgType.Bmp: return 0xF01F;
                default: return 0xF01A - 1;
            }
        }

        public static msobi XlsBlipSignConv(TXlsImgType img)
        {
            switch (img)
            {
                case TXlsImgType.Emf: return msobi.EMF;
                case TXlsImgType.Wmf: return msobi.WMF;
                case TXlsImgType.Jpeg: return msobi.JFIF;
                case TXlsImgType.Png: return msobi.PNG;
                case TXlsImgType.Bmp: return msobi.DIB;
                default: return msobi.UNKNOWN;
            }
        }

    }


    /// <summary>
    ///  Holds global data on all the sheets on the workbook.
    /// </summary>
    internal class TSheetInfo
    {
        /// <summary>
        /// Sheet where we inserted/deleted things.
        /// </summary>
        internal int InsSheet;

        /// <summary>
        /// Sheet where the formula is. For example, we can insert a row on Sheet1 (so InsSheet=1), and we are fixing a Formula on Sheet2 that references Sheet1 (So SourceFormulaSheet=2) 
        /// </summary>
        internal int SourceFormulaSheet;

        /// <summary>
        /// Sheet where the new formula will be. This only makes sense when copying from one file to another. For example, SourceFormulaSheet could be 1 in book1.xls and DestFormulaSheet 2 in book2.xls.
        /// </summary>
        internal int DestFormulaSheet;

        /// <summary>
        /// Sheet refered by SourceFormulaSheet.
        /// </summary>
        internal TSheet SourceSheet;

        /// <summary>
        /// Sheet referred by DestFormulaSheet.
        /// </summary>
        internal TSheet DestSheet;

        /// <summary>
        /// Globals for the source. Used when copying to different files.
        /// </summary>
        internal TWorkbookGlobals SourceGlobals;

        /// <summary>
        /// Globals for the destination. Used when copying to different files.
        /// </summary>
        internal TWorkbookGlobals DestGlobals;

        /// <summary>
        /// References for the source sheet.
        /// </summary>
        internal TReferences SourceReferences { get { if (SourceGlobals == null) return null; return SourceGlobals.References; } }

        /// <summary>
        /// References for the destination sheet. This will be different from References only if copying to other file.
        /// </summary>
        internal TReferences DestReferences { get { if (DestGlobals == null) return null; return DestGlobals.References; } }

        /// <summary>
        /// List of names.
        /// </summary>
        internal TNameRecordList SourceNames { get { if (SourceGlobals == null) return null; return SourceGlobals.Names; } }

        /// <summary>
        /// List of names.
        /// </summary>
        internal TNameRecordList DestNames { get { if (DestGlobals == null) return SourceNames; return DestGlobals.Names; } }


        /// <summary>
        /// If true, absolute references in a block will be updated when copied.
        /// </summary>
        internal bool SemiAbsoluteMode;

        /// <summary>
        /// A way to avoid having to recurse all records to clear CopiedTo. If CopiedGen is different, then CopiedTo is null
        /// </summary>
        internal TCopiedGen CopiedGen { get { return SourceGlobals.DrawingGen; } }

        /// <summary>
        /// A list of drawings that are in a range of cells.
        /// </summary>
        internal TExcelObjectList ObjectsInRange;

        internal TSheetInfo(int aInsSheet, int aFormulaSheet, int aFormulaSheetDest, TWorkbookGlobals aSourceGlobals,
            TWorkbookGlobals aDestGlobals, TSheet aSourceSheet, TSheet aDestSheet, bool aSemiAbsoluteMode)
        {
            InsSheet = aInsSheet;
            SourceFormulaSheet = aFormulaSheet;
            DestFormulaSheet = aFormulaSheetDest;
            SourceGlobals = aSourceGlobals;
            DestGlobals = aDestGlobals;
            SourceSheet = aSourceSheet;
            DestSheet = aDestSheet;
            SemiAbsoluteMode = aSemiAbsoluteMode;
            ObjectsInRange = null;
        }

        internal TSheetInfo(int aInsSheet, int aFormulaSheet, int aFormulaSheetDest, TWorkbookGlobals aSourceGlobals, TWorkbookGlobals aDestGlobals, TSheetList SourceSheetList, TSheetList DestSheetList, bool aSemiAbsoluteMode)
            : this(aInsSheet, aFormulaSheet, aFormulaSheetDest, aSourceGlobals, aDestGlobals, (TSheet)null, (TSheet)null, aSemiAbsoluteMode)
        {
            if (SourceSheetList != null && aFormulaSheet >= 0) SourceSheet = SourceSheetList[aFormulaSheet];
            if (DestSheetList != null && aFormulaSheetDest >= 0) DestSheet = DestSheetList[aFormulaSheetDest];
        }

        internal static TSheetInfo EmptyInstance = new TSheetInfo(-1, -1, -1, null, null, (TSheet)null, (TSheet)null, false);


        internal void IncCopiedGen()
        {
            SourceGlobals.DrawingGen.IncGeneration();
        }

        internal void PushCopiedGen()
        {
            SourceGlobals.DrawingGen.Push();
        }

        internal void PopCopiedGen(TCopiedGen SaveCopiedGen)
        {
            SourceGlobals.DrawingGen = SaveCopiedGen;
        }
    }


    /// <summary>
    /// Simple read/writes from/to array of bytes to primitive types.
    /// </summary>
    internal sealed class BitOps
    {
        private BitOps() { }

        internal static void IncWord(byte[] Data, int tPos, int Offset, int Max, XlsErr ErrWhenTooMany)
        {
            int w = GetWord(Data, tPos) + Offset;
            if ((w < 0) || (w > Max)) XlsMessages.ThrowException(ErrWhenTooMany, w + 1, Max + 1);
            SetWord(Data, tPos, (UInt16)w);
        }

        internal static void IncWord(ref int v, int Offset, int Max, XlsErr ErrWhenTooMany)
        {
            long w = v + Offset;
            if ((w < 0) || (w > Max)) XlsMessages.ThrowException(ErrWhenTooMany, w + 1, Max + 1);
            v = (int)w;
        }

        internal static void IncCardinal(byte[] Data, int tPos, long Offset)
        {
            long w = GetCardinal(Data, tPos) + Offset;
            SetCardinal(Data, tPos, w);
        }


        internal static int GetWord(byte[] Data, int tPos) //{return (int)BitConverter.ToUInt16(Data, tPos);}  
        {
            //Optimization for a routine called million times
            unchecked
            {
                return Data[tPos] + (Data[tPos + 1] << 8);
            }
        }

        internal static void SetWord(byte[] Data, int tPos, int number) // BitConverter.GetBytes((UInt16)number).CopyTo(Data, tPos);}
        {
            unchecked
            {
                Data[tPos] = (byte)number;
                Data[tPos + 1] = (byte)(number >> 8);
            }
        }

        internal static Int64 GetCardinal(byte[] Data, int tPos) //{return (UInt32)BitConverter.ToUInt32(Data, tPos);}  
        {
            //Optimization for a routine called million times
            unchecked
            {
                return (UInt32)(Data[tPos] + (Data[tPos + 1] << 8) + (Data[tPos + 2] << 16) + (Data[tPos + 3] << 24));
            }
        }

        internal static Int32 GetInt32(byte[] Data, int tPos) //{return (UInt32)BitConverter.ToUInt32(Data, tPos);}  
        {
            //Optimization for a routine called million times
            unchecked
            {
                return (Data[tPos] + (Data[tPos + 1] << 8) + (Data[tPos + 2] << 16) + (Data[tPos + 3] << 24));
            }
        }

        internal static void SetCardinal(byte[] Data, int tPos, long number) //{ BitConverter.GetBytes((UInt32)number).CopyTo(Data, tPos);}
        {
            unchecked
            {
                Data[tPos] = (byte)number;
                Data[tPos + 1] = (byte)(number >> 8);
                Data[tPos + 2] = (byte)(number >> 16);
                Data[tPos + 3] = (byte)(number >> 24);
            }
        }

        internal static int GetIncMaxMin(int X, int N, int Max, int Min)
        {
            if (N + X > Max) X = Max; else if (N + X < Min) X = Min; else X += N;
            return X;
        }

        internal static int BoolToBit(bool Value, int ofs)
        {
            if (Value) return (1 << ofs); else return 0;
        }



        //Read memory taking in count "Continue" Records
        internal static void ReadMem(ref TxBaseRecord aRecord, ref int aPos, byte[] Result, int StartResultPos, int aSize)
        {
            int ResultPos = StartResultPos;
            int lr = 0;
            do
            {
                lr = aRecord.DataSize - aPos;

                if (lr < 0) XlsMessages.ThrowException(XlsErr.ErrReadingRecord);
                if ((lr == 0) && (aSize > 0)) //Goto next continue
                {
                    aPos = 0;
                    aRecord = aRecord.Continue;
                    if (aRecord == null) XlsMessages.ThrowException(XlsErr.ErrReadingRecord);
                }

                lr = aRecord.DataSize - aPos;

                int RealLr = Math.Min(aSize, lr);
                if (Result != null) Array.Copy(aRecord.Data, aPos, Result, ResultPos, RealLr);
                aPos += RealLr;
                ResultPos += RealLr;
                aSize -= RealLr;
            } while (aSize > 0);
        }

        internal static void ReadMem(ref TxBaseRecord aRecord, ref int aPos, byte[] Result)
        {
            ReadMem(ref aRecord, ref aPos, Result, 0, Result.Length);
        }

        internal static bool CompareMem(byte[] a1, byte[] a2)
        {
            if (a1 == null)
            {
                if ((a2 == null) || (a2.Length == 0)) return true; else return false;
            }

            if (a2 == null)
            {
                if (a1.Length == 0) return true; else return false;
            }

            if (a1.Length != a2.Length) return false;
            for (int i = 0; i < a1.Length; i++)
                if (a1[i] != a2[i]) return false;
            return true;
        }

        internal static bool CompareMem(byte[] a1, byte[] a2, int a2Pos)
        {
            for (int i = 0; i < a1.Length; i++)
                if (a1[i] != a2[i + a2Pos]) return false;
            return true;
        }

        internal static int CompareMemOrdinal(byte[] a1, byte[] a2)
        {
            if (a1 == null)
            {
                if ((a2 == null) || (a2.Length == 0)) return 0; else return -1;
            }

            if (a2 == null)
            {
                if (a1.Length == 0) return 0; else return 1;
            }

            int Result = a1.Length.CompareTo(a2.Length);
            if (Result != 0) return Result;
            for (int i = 0; i < a1.Length; i++)
            {
                Result = a1[i].CompareTo(a2[i]);
                if (Result != 0) return Result;
            }
            return 0;
        }


        internal static int GetBool(params bool[] b)
        {
            int Result = 0;
            int Mask = 1;

            for (int i = 0; i < b.Length; i++)
            {
                if (b[i]) Result |= Mask;
                Mask <<= 1;
            }

            return Result;
        }
    }

    /// <summary>
    /// Simple string functions. They don't consider continue records.
    /// </summary>
    internal sealed class StrOps
    {
        private StrOps() { }

        /// <summary>
        /// Length of a simple string. Does not consider continue records, so use it with care.
        /// </summary>
        /// <param name="Length16Bit">When true, string can have up to 65535 chars. When false, max string length=255</param>
        /// <param name="PData">Data</param>
        /// <param name="tPos">Where it starts.</param>
        /// <param name="UseExtStrLen">True if length is included on ExStrLen</param>
        /// <param name="ExtStrLen">StringLength, if it is not on the Data array</param>
        /// <returns>Length of the string</returns>
        internal static long GetStrLen(bool Length16Bit, byte[] PData, int tPos, bool UseExtStrLen, long ExtStrLen)
        {
            int myPos = tPos;
            long l = 0;
            if (UseExtStrLen) l = ExtStrLen;
            else
            {
                if (Length16Bit)
                {
                    l = BitOps.GetWord(PData, myPos); myPos += 2;
                }
                else
                {
                    l = PData[myPos];
                    myPos++;
                }
            }

            byte oField = PData[myPos];
            myPos++;

            byte bsize = (byte)(oField & 0x1); // 8bit/16 bit string

            long rt = 0;
            if ((oField & 0x8) == 0x8)  //RTF Info
            {
                rt = BitOps.GetWord(PData, myPos);
                myPos += 2;
            }

            long sz = 0;
            if ((oField & 0x4) == 0x4) //Far East Info
            {
                sz = BitOps.GetCardinal(PData, myPos);
                myPos += 4;
            }

            return (long)myPos - tPos + (l << bsize) + rt * 4 + sz;

        }

        internal static bool CompressUnicode(string s, byte[] data, int startPos)
        {
            for (int i = 0; i < s.Length; i++)
                if (s[i] <= '\u00FF') data[i + startPos] = (byte)s[i]; else return false;
            return true;
        }

        internal static bool CompressBestUnicode(string s, byte[] data, int startPos)
        {
            bool Result = true;
            for (int i = 0; i < s.Length; i++)
                if (s[i] <= '\u00FF') data[i + startPos] = (byte)s[i]; else { data[i + startPos] = (byte)'?'; Result = false; }
            return Result;
        }

        internal static bool IsWide(string s)
        {
            for (int i = 0; i < s.Length; i++)
                if (s[i] > '\u00FF') return true;
            return false;
        }

        internal static string UnCompressUnicode(byte[] data, int start, int len)
        {
            StringBuilder sb = new StringBuilder(len);
            sb.Length = len;
            for (int i = 0; i < len; i++)
                sb[i] = (char)data[i + start];
            return sb.ToString();
        }

        internal static void GetSimpleString(bool Length16Bit, byte[] Pdata, int tPos, bool UseExtStrLen, long ExtStrLen, ref string St, ref long StSize)
        {
            int myPos = tPos;
            long l = 0;
            if (UseExtStrLen) l = ExtStrLen;
            else
            {
                if (Length16Bit)
                {
                    l = BitOps.GetWord(Pdata, myPos); myPos += 2;
                }
                else
                {
                    l = Pdata[myPos]; myPos++;
                }
            }
            
			byte oField = Pdata[myPos];
            myPos++;

            byte bsize = (byte)(oField & 0x1);  //8bit/16 bit string

            long rt = 0;
            if ((oField & 0x8) == 0x8)  //RTF Info
            {
                rt = BitOps.GetWord(Pdata, myPos);
                myPos += 2;
            }

            long sz = 0;
            if ((oField & 0x4) == 0x4) //Far East Info
            {
                sz = BitOps.GetCardinal(Pdata, myPos);
                myPos += 4;
            }

            StSize = (long)myPos - tPos + (l << bsize) + rt * 4 + sz;
            if (bsize == 0)
            {
                St = UnCompressUnicode(Pdata, myPos, (int)l);
            }

            else
            {
                St = Encoding.Unicode.GetString(Pdata, myPos, (int)l * 2);
            }
        }


        /// <summary>
        /// Read a string taking in count "Continue" Records. It returns the last position, so you can continue reading other things.
        /// </summary>
        /// <param name="aRecord">Record to read. If continued, the next record will be returned here.</param>
        /// <param name="aPos">Position of the string on the Data array</param>
        /// <param name="Data">Read string</param>
        /// <param name="OptionFlags">Original optionFlags of the string</param>
        /// <param name="ActualOptionFlags">Final Option flags. They can change if there are Continue records</param>
        /// <param name="StrLen">Length of the string to read</param>
        internal static void ReadStr(ref TxBaseRecord aRecord, ref int aPos, StringBuilder Data, byte OptionFlags, ref byte ActualOptionFlags, int StrLen)
        {
            byte CurrentOptionFlags = OptionFlags;
            int lr = aRecord.DataSize - aPos;

            if (lr < 0) XlsMessages.ThrowException(XlsErr.ErrReadingRecord);
            if ((lr == 0) && (StrLen > 0))  //Move to next continue
            {
                /* In a real Excel file, this is impossible, optionflags and the first character should go in the same
                 * record. But sometimes the JET engine can generate files split between optionflags and the first char.
                 * 
				if (Data.Length==0)  //we are beginning the record
				{
					aPos=0;
					if (aRecord.Continue==null) XlsMessages.ThrowException(XlsErr.ErrReadingRecord);
					aRecord=aRecord.Continue;
				}
				else*/
                {   //We are in the middle of a string. First byte are the new OptionFlags
                    aPos = 1;
                    if (aRecord.Continue == null) XlsMessages.ThrowException(XlsErr.ErrReadingRecord);
                    aRecord = aRecord.Continue;

                    CurrentOptionFlags = aRecord.Data[0];
                    if (((CurrentOptionFlags & 1) == 1) && ((ActualOptionFlags & 1) == 0))
                    {
                        //WideData=StringToWideStringNoCodePage(ShortData);  //Not needed as we always use widestrings.
                        ActualOptionFlags = (byte)(OptionFlags | 1);
                    }
                }
            }

            lr = aRecord.DataSize - aPos;
            int Remaining = 0;

            int CharSize = 0;
            if ((CurrentOptionFlags & 1) == 0) //Remember that we can have a wide string continuing on a compressed string.
            {
                CharSize = 1;
            }
            else
            {
                CharSize = 2;
            }

            Remaining = (StrLen - Data.Length) * CharSize;

            int LeftInRecord = 0;
            if (Remaining <= lr) //Record ends here
                LeftInRecord = Remaining;//Record continues
            else
                LeftInRecord = lr;

            if ((CurrentOptionFlags & 1) == 0)
                //Convert the compressed data to uncompressed result
                Data.Append(UnCompressUnicode(aRecord.Data, aPos, LeftInRecord));

                //uncompressed to uncompressed
            else Data.Append(Encoding.Unicode.GetString(aRecord.Data, aPos, LeftInRecord));

            aPos += LeftInRecord;

            if (Remaining > lr)
                ReadStr(ref aRecord, ref aPos, Data, CurrentOptionFlags, ref ActualOptionFlags, StrLen);
        }
    }

    /// <summary>
    /// Methods to handle biff8 format.
    /// </summary>
    internal sealed class Biff8Utils
    {
        private Biff8Utils() { }

        internal static int ExpandBiff8Row(int r)
        {
            if (r == FlxConsts.Max_Rows97_2003 && !FlxConsts.KeepMaxRowsAndColumsWhenUpdating) return FlxConsts.Max_Rows;
            return r;
        }

        internal static int ExpandBiff8Col(int c)
        {
            if (c == FlxConsts.Max_Columns97_2003 && !FlxConsts.KeepMaxRowsAndColumsWhenUpdating) return FlxConsts.Max_Columns;
            return c;
        }

		internal static void CheckChart(ref int Coord)
		{
			if (Coord < 0) Coord = 0;
            if (Coord > 4000) Coord = 4000;
		}


        internal static void CheckRow(int Row)
        {
            if (Row < 0 || Row > FlxConsts.Max_Rows97_2003) XlsMessages.ThrowException(XlsErr.ErrTooManyRows, Row + 1, FlxConsts.Max_Rows97_2003 + 1); //an xls file cannot have more than 65536 rows. File must be saved as xlsx
        }

        internal static void CheckCol(int Col)
        {
            if (Col < 0 || Col > FlxConsts.Max_Columns97_2003) XlsMessages.ThrowException(XlsErr.ErrTooManyColumns, Col + 1, FlxConsts.Max_Columns97_2003 + 1); //an xls file cannot have more than 255 columns. File must be saved as xlsx
        }

        internal static void CheckXF(int xf)
        {
            if (xf < 0 || xf > XlsConsts.MaxXFDefs97_2003) XlsMessages.ThrowException(XlsErr.ErrTooManyXFDefs); //saving to biff8
        }


        internal static int CheckAndContractBiff8Row(int r)
        {
            if (r == FlxConsts.Max_Rows2007 && !FlxConsts.KeepMaxRowsAndColumsWhenUpdating) return FlxConsts.Max_Rows97_2003;
            CheckRow(r);
            return r;
        }

        internal static int CheckAndContractBiff8Col(int c)
        {
            if (c == FlxConsts.Max_Columns2007 && !FlxConsts.KeepMaxRowsAndColumsWhenUpdating) return FlxConsts.Max_Columns97_2003;
            CheckCol(c);
            return c;
        }

        internal static int CheckAndContractRelativeBiff8Row(int r)
        {
            if (r >= Int16.MaxValue || r < Int16.MinValue)
            {
                XlsMessages.ThrowException(XlsErr.ErrTooManyRows, r, FlxConsts.Max_Rows97_2003 + 1);
            }

            unchecked
            {
                r = (UInt16)r;
            }
            return r;
        }

        internal static int CheckAndContractRelativeBiff8Col(int c)
        {
            if (c >= sbyte.MaxValue || c <sbyte.MinValue)
            {
                XlsMessages.ThrowException(XlsErr.ErrTooManyColumns, c, FlxConsts.Max_Columns97_2003 + 1);
            }
            unchecked
            {
                c = (byte)c;
            }
            return c;
        }

    }

}



