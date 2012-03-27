using System;

using System.Resources;
using System.Reflection;
using System.Collections.Generic;

namespace FlexCel.Core
{
    internal class TCellFunctionDataDictionary : Dictionary<string, TCellFunctionData> { }

    /// <summary>
    /// Functions that aren't defined in the biff8/xlsb spec (Excel 2010 or newer), but that we define as internal
    /// </summary>
    internal enum TFutureFunctions
    {
        CeilingPrecise = 0x200,
        FloorPrecise = 0x201,
        IsoCeiling = 0x202,
        Aggregate = 0x203,

        PercentileExc = 0x204,
        QuartileExc = 0x205,

        BetaDist = 0x206,
        BetaInv = 0x207,
        BinomDist = 0x208,
        BinomInv = 0x209,
        ChisqDistRt = 0x20A,
        ChisqInvRt = 0x20B,
        ChisqTest = 0x20C,
        ConfidenceNorm = 0x20D,
        CovarianceP = 0x20E,
        ExponDist = 0x20F,
        FDistRt = 0x210,
        FInvRt = 0x211,
        FTest = 0x212,
        GammaDist = 0x213,
        GammaInv = 0x214,
        HypGeomDist = 0x215,
        LogNormDist = 0x216,
        LogNormInv = 0x217,
        ModeSngl = 0x218,
        NegBinom = 0x219,
        NormDist = 0x21A,
        NormInv = 0x21B,
        NormSDist = 0x21C,
        NormSInv = 0x21D,
        PercentileInc = 0x21E,
        QuartileInc = 0x21F,
        PercentRankInc = 0x220,
        PoissonDist = 0x221,
        RankEq = 0x222,
        StDevP = 0x223,
        StDevS = 0x224,
        TDist2T = 0x225,
        TDistRT = 0x226,
        TInv2T = 0x227,
        TTest = 0x228,
        VarP = 0x229,
        VarS = 0x22A,
        WeibullDist = 0x22B,
        ZTest = 0x22C

    }

	/// <summary>
	/// A list with the functions on an excel sheet.
	/// </summary>
	internal class TXlsFunction
	{
        private static readonly TCellFunctionDataDictionary Ht = CreateHashTable();//STATIC* 
        
        private static readonly TCellFunctionData[] IndexFunc = CreateIndexFunc();//STATIC*

		internal const int MaxFunctions = 0x230; // Total of functions indexed.

		private TXlsFunction()
		{
		}

        private static TCellFunctionDataDictionary CreateHashTable()
        {
            ResourceManager rm = new ResourceManager("FlexCel.Core.FunctionNames", Assembly.GetExecutingAssembly());
            TCellFunctionDataDictionary Result = new TCellFunctionDataDictionary();

			Add(Result, new TCellFunctionData(0, rm.GetString("COUNT"), 0, 255, true, TFmReturnType.Value, "R"));
			Add(Result, new TCellFunctionData(1, rm.GetString("IF"), 2, 3, true, TFmReturnType.Ref, "VRR"));
			Add(Result, new TCellFunctionData(2, rm.GetString("ISNA"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(3, rm.GetString("ISERROR"), 1, 1, true, TFmReturnType.Value, "V"));
            Add(Result, new TCellFunctionData(4, rm.GetString("SUM"), 0, 255, true, TFmReturnType.Value, "R"));
            Add(Result, new TCellFunctionData(5, rm.GetString("AVERAGE"), 1, 255, true, TFmReturnType.Value, "R"));
            Add(Result, new TCellFunctionData(6, rm.GetString("MIN"), 1, 255, true, TFmReturnType.Value, "R"));
            Add(Result, new TCellFunctionData(7, rm.GetString("MAX"), 1, 255, true, TFmReturnType.Value, "R"));
			Add(Result, new TCellFunctionData(8, rm.GetString("ROW"), 0, 1, true, TFmReturnType.Value, "R"));
			Add(Result, new TCellFunctionData(9, rm.GetString("COLUMN"), 0, 1, true, TFmReturnType.Value, "R"));
			Add(Result, new TCellFunctionData(10, rm.GetString("NA"), 0, 0, true, TFmReturnType.Value, "-"));
            Add(Result, new TCellFunctionData(11, rm.GetString("NPV"), 2, 255, true, TFmReturnType.Value, "VR"));
            Add(Result, new TCellFunctionData(12, rm.GetString("STDEV"), 1, 255, true, TFmReturnType.Value, "R"));
			Add(Result, new TCellFunctionData(13, rm.GetString("DOLLAR"), 1, 2, true, TFmReturnType.Value, "VV"));
			Add(Result, new TCellFunctionData(14, rm.GetString("FIXED"), 1, 3, true, TFmReturnType.Value, "VVV"));
			Add(Result, new TCellFunctionData(15, rm.GetString("SIN"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(16, rm.GetString("COS"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(17, rm.GetString("TAN"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(18, rm.GetString("ATAN"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(19, rm.GetString("PI"), 0, 0, true, TFmReturnType.Value, "-"));
			Add(Result, new TCellFunctionData(20, rm.GetString("SQRT"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(21, rm.GetString("EXP"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(22, rm.GetString("LN"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(23, rm.GetString("LOG10"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(24, rm.GetString("ABS"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(25, rm.GetString("INT"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(26, rm.GetString("SIGN"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(27, rm.GetString("ROUND"), 2, 2, true, TFmReturnType.Value, "VV"));
			Add(Result, new TCellFunctionData(28, rm.GetString("LOOKUP"), 2, 3, true, TFmReturnType.Value, "VRR"));
			Add(Result, new TCellFunctionData(29, rm.GetString("INDEX"), 2, 4, true, TFmReturnType.Ref, "RVVV"));
			Add(Result, new TCellFunctionData(30, rm.GetString("REPT"), 2, 2, true, TFmReturnType.Value, "VV"));
			Add(Result, new TCellFunctionData(31, rm.GetString("MID"), 3, 3, true, TFmReturnType.Value, "VVV"));
			Add(Result, new TCellFunctionData(32, rm.GetString("LEN"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(33, rm.GetString("VALUE"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(34, rm.GetString("TRUE"), 0, 0, true, TFmReturnType.Value, "-"));
			Add(Result, new TCellFunctionData(35, rm.GetString("FALSE"), 0, 0, true, TFmReturnType.Value, "-"));
            Add(Result, new TCellFunctionData(36, rm.GetString("AND"), 1, 255, true, TFmReturnType.Value, "R"));
            Add(Result, new TCellFunctionData(37, rm.GetString("OR"), 1, 255, true, TFmReturnType.Value, "R"));
			Add(Result, new TCellFunctionData(38, rm.GetString("NOT"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(39, rm.GetString("MOD"), 2, 2, true, TFmReturnType.Value, "VV"));
			Add(Result, new TCellFunctionData(40, rm.GetString("DCOUNT"), 3, 3, true, TFmReturnType.Value, "RRR"));
			Add(Result, new TCellFunctionData(41, rm.GetString("DSUM"), 3, 3, true, TFmReturnType.Value, "RRR"));
			Add(Result, new TCellFunctionData(42, rm.GetString("DAVERAGE"), 3, 3, true, TFmReturnType.Value, "RRR"));
			Add(Result, new TCellFunctionData(43, rm.GetString("DMIN"), 3, 3, true, TFmReturnType.Value, "RRR"));
			Add(Result, new TCellFunctionData(44, rm.GetString("DMAX"), 3, 3, true, TFmReturnType.Value, "RRR"));
			Add(Result, new TCellFunctionData(45, rm.GetString("DSTDEV"), 3, 3, true, TFmReturnType.Value, "RRR"));
            Add(Result, new TCellFunctionData(46, rm.GetString("VAR"), 1, 255, true, TFmReturnType.Value, "R"));
			Add(Result, new TCellFunctionData(47, rm.GetString("DVAR"), 3, 3, true, TFmReturnType.Value, "RRR"));
			Add(Result, new TCellFunctionData(48, rm.GetString("TEXT"), 2, 2, true, TFmReturnType.Value, "VV"));
			Add(Result, new TCellFunctionData(49, rm.GetString("LINEST"), 1, 4, true, TFmReturnType.Array, "RRVV"));
			Add(Result, new TCellFunctionData(50, rm.GetString("TREND"), 1, 4, true, TFmReturnType.Array, "RRRV"));
			Add(Result, new TCellFunctionData(51, rm.GetString("LOGEST"), 1, 4, true, TFmReturnType.Array, "RRVV"));
			Add(Result, new TCellFunctionData(52, rm.GetString("GROWTH"), 1, 4, true, TFmReturnType.Array, "RRRV"));
			Add(Result, new TCellFunctionData(56, rm.GetString("PV"), 3, 5, true, TFmReturnType.Value, "VVVVV"));
			Add(Result, new TCellFunctionData(57, rm.GetString("FV"), 3, 5, true, TFmReturnType.Value, "VVVVV"));
			Add(Result, new TCellFunctionData(58, rm.GetString("NPER"), 3, 5, true, TFmReturnType.Value, "VVVVV"));
			Add(Result, new TCellFunctionData(59, rm.GetString("PMT"), 3, 5, true, TFmReturnType.Value, "VVVVV"));
			Add(Result, new TCellFunctionData(60, rm.GetString("RATE"), 3, 6, true, TFmReturnType.Value, "VVVVVV"));
			Add(Result, new TCellFunctionData(61, rm.GetString("MIRR"), 3, 3, true, TFmReturnType.Value, "RVV"));
			Add(Result, new TCellFunctionData(62, rm.GetString("IRR"), 1, 2, true, TFmReturnType.Value, "RV"));
			Add(Result, new TCellFunctionData(63, rm.GetString("RAND"), 0, 0, false, TFmReturnType.Value, "-"));
			Add(Result, new TCellFunctionData(64, rm.GetString("MATCH"), 2, 3, true, TFmReturnType.Value, "VRR"));
			Add(Result, new TCellFunctionData(65, rm.GetString("DATE"), 3, 3, true, TFmReturnType.Value, "VVV"));
			Add(Result, new TCellFunctionData(66, rm.GetString("TIME"), 3, 3, true, TFmReturnType.Value, "VVV"));
			Add(Result, new TCellFunctionData(67, rm.GetString("DAY"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(68, rm.GetString("MONTH"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(69, rm.GetString("YEAR"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(70, rm.GetString("WEEKDAY"), 1, 2, true, TFmReturnType.Value, "VV"));
			Add(Result, new TCellFunctionData(71, rm.GetString("HOUR"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(72, rm.GetString("MINUTE"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(73, rm.GetString("SECOND"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(74, rm.GetString("NOW"), 0, 0, false, TFmReturnType.Value, "-"));
			Add(Result, new TCellFunctionData(75, rm.GetString("AREAS"), 1, 1, true, TFmReturnType.Value, "R"));
			Add(Result, new TCellFunctionData(76, rm.GetString("ROWS"), 1, 1, true, TFmReturnType.Value, "R"));
			Add(Result, new TCellFunctionData(77, rm.GetString("COLUMNS"), 1, 1, true, TFmReturnType.Value, "R"));
			Add(Result, new TCellFunctionData(78, rm.GetString("OFFSET"), 3, 5, false, TFmReturnType.Ref, "RVVVV"));
			Add(Result, new TCellFunctionData(82, rm.GetString("SEARCH"), 2, 3, true, TFmReturnType.Value, "VVV"));
			Add(Result, new TCellFunctionData(83, rm.GetString("TRANSPOSE"), 1, 1, true, TFmReturnType.Array, "A"));
			Add(Result, new TCellFunctionData(86, rm.GetString("TYPE"), 1, 1, true, TFmReturnType.Value, "V", true));
			Add(Result, new TCellFunctionData(97, rm.GetString("ATAN2"), 2, 2, true, TFmReturnType.Value, "VV"));
			Add(Result, new TCellFunctionData(98, rm.GetString("ASIN"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(99, rm.GetString("ACOS"), 1, 1, true, TFmReturnType.Value, "V"));
            Add(Result, new TCellFunctionData(100, rm.GetString("CHOOSE"), 2, 255, true, TFmReturnType.Ref, "VR"));
			Add(Result, new TCellFunctionData(101, rm.GetString("HLOOKUP"), 3, 4, true, TFmReturnType.Value, "VRRV"));
			Add(Result, new TCellFunctionData(102, rm.GetString("VLOOKUP"), 3, 4, true, TFmReturnType.Value, "VRRV"));
			Add(Result, new TCellFunctionData(105, rm.GetString("ISREF"), 1, 1, true, TFmReturnType.Value, "R"));
			Add(Result, new TCellFunctionData(109, rm.GetString("LOG"), 1, 2, true, TFmReturnType.Value, "VV"));
			Add(Result, new TCellFunctionData(111, rm.GetString("CHAR"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(112, rm.GetString("LOWER"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(113, rm.GetString("UPPER"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(114, rm.GetString("PROPER"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(115, rm.GetString("LEFT"), 1, 2, true, TFmReturnType.Value, "VV"));
			Add(Result, new TCellFunctionData(116, rm.GetString("RIGHT"), 1, 2, true, TFmReturnType.Value, "VV"));
			Add(Result, new TCellFunctionData(117, rm.GetString("EXACT"), 2, 2, true, TFmReturnType.Value, "VV"));
			Add(Result, new TCellFunctionData(118, rm.GetString("TRIM"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(119, rm.GetString("REPLACE"), 4, 4, true, TFmReturnType.Value, "VVVV"));
			Add(Result, new TCellFunctionData(120, rm.GetString("SUBSTITUTE"), 3, 4, true, TFmReturnType.Value, "VVVV"));
			Add(Result, new TCellFunctionData(121, rm.GetString("CODE"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(124, rm.GetString("FIND"), 2, 3, true, TFmReturnType.Value, "VVV"));
			Add(Result, new TCellFunctionData(125, rm.GetString("CELL"), 1, 2, false, TFmReturnType.Value, "VR"));
			Add(Result, new TCellFunctionData(126, rm.GetString("ISERR"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(127, rm.GetString("ISTEXT"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(128, rm.GetString("ISNUMBER"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(129, rm.GetString("ISBLANK"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(130, rm.GetString("T"), 1, 1, true, TFmReturnType.Value, "R"));
			Add(Result, new TCellFunctionData(131, rm.GetString("N"), 1, 1, true, TFmReturnType.Value, "R"));
			Add(Result, new TCellFunctionData(140, rm.GetString("DATEVALUE"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(141, rm.GetString("TIMEVALUE"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(142, rm.GetString("SLN"), 3, 3, true, TFmReturnType.Value, "VVV"));
			Add(Result, new TCellFunctionData(143, rm.GetString("SYD"), 4, 4, true, TFmReturnType.Value, "VVVV"));
			Add(Result, new TCellFunctionData(144, rm.GetString("DDB"), 4, 5, true, TFmReturnType.Value, "VVVVV"));
			Add(Result, new TCellFunctionData(148, rm.GetString("INDIRECT"), 1, 2, false, TFmReturnType.Ref, "VV"));
			Add(Result, new TCellFunctionData(162, rm.GetString("CLEAN"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(163, rm.GetString("MDETERM"), 1, 1, true, TFmReturnType.Value, "A"));
			Add(Result, new TCellFunctionData(164, rm.GetString("MINVERSE"), 1, 1, true, TFmReturnType.Array, "A"));
			Add(Result, new TCellFunctionData(165, rm.GetString("MMULT"), 2, 2, true, TFmReturnType.Array, "AA"));
			Add(Result, new TCellFunctionData(167, rm.GetString("IPMT"), 4, 6, true, TFmReturnType.Value, "VVVVVV"));
			Add(Result, new TCellFunctionData(168, rm.GetString("PPMT"), 4, 6, true, TFmReturnType.Value, "VVVVVV"));
            Add(Result, new TCellFunctionData(169, rm.GetString("COUNTA"), 0, 255, true, TFmReturnType.Value, "R"));
            Add(Result, new TCellFunctionData(183, rm.GetString("PRODUCT"), 0, 255, true, TFmReturnType.Value, "R"));
			Add(Result, new TCellFunctionData(184, rm.GetString("FACT"), 1, 1, true, TFmReturnType.Value, "V"));

            //this are in the new definition.
            Add(Result, new TCellFunctionData(0xB9, rm.GetString("GET.CELL"), 1, 2, false, TFmReturnType.Value, "VR"));
            Add(Result, new TCellFunctionData(0xBA, rm.GetString("GET.WORKSPACE"), 1, 1, false, TFmReturnType.Value, "V"));
            Add(Result, new TCellFunctionData(0xBB, rm.GetString("GET.WINDOW"), 1, 2, false, TFmReturnType.Value, "VV"));
            Add(Result, new TCellFunctionData(0xBB, rm.GetString("GET.DOCUMENT"), 1, 2, false, TFmReturnType.Value, "VV"));


			Add(Result, new TCellFunctionData(189, rm.GetString("DPRODUCT"), 3, 3, true, TFmReturnType.Value, "RRR"));
			Add(Result, new TCellFunctionData(190, rm.GetString("ISNONTEXT"), 1, 1, true, TFmReturnType.Value, "V"));
            Add(Result, new TCellFunctionData(193, rm.GetString("STDEVP"), 1, 255, true, TFmReturnType.Value, "R"));
            Add(Result, new TCellFunctionData(194, rm.GetString("VARP"), 1, 255, true, TFmReturnType.Value, "R"));
			Add(Result, new TCellFunctionData(195, rm.GetString("DSTDEVP"), 3, 3, true, TFmReturnType.Value, "RRR"));
			Add(Result, new TCellFunctionData(196, rm.GetString("DVARP"), 3, 3, true, TFmReturnType.Value, "RRR"));
			Add(Result, new TCellFunctionData(197, rm.GetString("TRUNC"), 1, 2, true, TFmReturnType.Value, "VV"));
			Add(Result, new TCellFunctionData(198, rm.GetString("ISLOGICAL"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(199, rm.GetString("DCOUNTA"), 3, 3, true, TFmReturnType.Value, "RRR"));
			Add(Result, new TCellFunctionData(204, rm.GetString("USDOLLAR"), 1, 2, true, TFmReturnType.Value, "VV"));
			Add(Result, new TCellFunctionData(205, rm.GetString("FINDB"), 2, 3, true, TFmReturnType.Value, "VVV"));
			Add(Result, new TCellFunctionData(206, rm.GetString("SEARCHB"), 2, 3, true, TFmReturnType.Value, "VVV"));
			Add(Result, new TCellFunctionData(207, rm.GetString("REPLACEB"), 4, 4, true, TFmReturnType.Value, "VVVV"));
			Add(Result, new TCellFunctionData(208, rm.GetString("LEFTB"), 1, 2, true, TFmReturnType.Value, "VV"));
			Add(Result, new TCellFunctionData(209, rm.GetString("RIGHTB"), 1, 2, true, TFmReturnType.Value, "VV"));
			Add(Result, new TCellFunctionData(210, rm.GetString("MIDB"), 3, 3, true, TFmReturnType.Value, "VVV"));
			Add(Result, new TCellFunctionData(211, rm.GetString("LENB"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(212, rm.GetString("ROUNDUP"), 2, 2, true, TFmReturnType.Value, "VV"));
			Add(Result, new TCellFunctionData(213, rm.GetString("ROUNDDOWN"), 2, 2, true, TFmReturnType.Value, "VV"));
			Add(Result, new TCellFunctionData(214, rm.GetString("ASC"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(215, rm.GetString("DBSC"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(216, rm.GetString("RANK"), 2, 3, true, TFmReturnType.Value, "VRV"));
			Add(Result, new TCellFunctionData(219, rm.GetString("ADDRESS"), 2, 5, true, TFmReturnType.Value, "VVVVV"));
			Add(Result, new TCellFunctionData(220, rm.GetString("DAYS360"), 2, 3, true, TFmReturnType.Value, "VVV"));
			Add(Result, new TCellFunctionData(221, rm.GetString("TODAY"), 0, 0, false, TFmReturnType.Value, "-"));
			Add(Result, new TCellFunctionData(222, rm.GetString("VDB"), 5, 7, true, TFmReturnType.Value, "VVVVVVV"));
            Add(Result, new TCellFunctionData(227, rm.GetString("MEDIAN"), 1, 255, true, TFmReturnType.Value, "R"));
            Add(Result, new TCellFunctionData(228, rm.GetString("SUMPRODUCT"), 1, 255, true, TFmReturnType.Value, "A"));
			Add(Result, new TCellFunctionData(229, rm.GetString("SINH"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(230, rm.GetString("COSH"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(231, rm.GetString("TANH"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(232, rm.GetString("ASINH"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(233, rm.GetString("ACOSH"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(234, rm.GetString("ATANH"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(235, rm.GetString("DGET"), 3, 3, true, TFmReturnType.Value, "RRR"));
			Add(Result, new TCellFunctionData(244, rm.GetString("INFO"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(247, rm.GetString("DB"), 4, 5, true, TFmReturnType.Value, "VVVVV"));
			Add(Result, new TCellFunctionData(252, rm.GetString("FREQUENCY"), 2, 2, true, TFmReturnType.Array, "RR"));
			Add(Result, new TCellFunctionData(261, rm.GetString("ERROR.TYPE"), 1, 1, true, TFmReturnType.Value, "V"));
            Add(Result, new TCellFunctionData(269, rm.GetString("AVEDEV"), 1, 255, true, TFmReturnType.Value, "R"));
			Add(Result, new TCellFunctionData(270, rm.GetString("BETADIST"), 3, 5, true, TFmReturnType.Value, "VVVVV"));
			Add(Result, new TCellFunctionData(271, rm.GetString("GAMMALN"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(272, rm.GetString("BETAINV"), 3, 5, true, TFmReturnType.Value, "VVVVV"));
			Add(Result, new TCellFunctionData(273, rm.GetString("BINOMDIST"), 4, 4, true, TFmReturnType.Value, "VVVV"));
			Add(Result, new TCellFunctionData(274, rm.GetString("CHIDIST"), 2, 2, true, TFmReturnType.Value, "VV"));
			Add(Result, new TCellFunctionData(275, rm.GetString("CHIINV"), 2, 2, true, TFmReturnType.Value, "VV"));
			Add(Result, new TCellFunctionData(276, rm.GetString("COMBIN"), 2, 2, true, TFmReturnType.Value, "VV"));
			Add(Result, new TCellFunctionData(277, rm.GetString("CONFIDENCE"), 3, 3, true, TFmReturnType.Value, "VVV"));
			Add(Result, new TCellFunctionData(278, rm.GetString("CRITBINOM"), 3, 3, true, TFmReturnType.Value, "VVV"));
			Add(Result, new TCellFunctionData(279, rm.GetString("EVEN"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(280, rm.GetString("EXPONDIST"), 3, 3, true, TFmReturnType.Value, "VVV"));
			Add(Result, new TCellFunctionData(281, rm.GetString("FDIST"), 3, 3, true, TFmReturnType.Value, "VVV"));
			Add(Result, new TCellFunctionData(282, rm.GetString("FINV"), 3, 3, true, TFmReturnType.Value, "VVV"));
			Add(Result, new TCellFunctionData(283, rm.GetString("FISHER"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(284, rm.GetString("FISHERINV"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(285, rm.GetString("FLOOR"), 2, 2, true, TFmReturnType.Value, "VV"));
			Add(Result, new TCellFunctionData(286, rm.GetString("GAMMADIST"), 4, 4, true, TFmReturnType.Value, "VVVV"));
			Add(Result, new TCellFunctionData(287, rm.GetString("GAMMAINV"), 3, 3, true, TFmReturnType.Value, "VVV"));
			Add(Result, new TCellFunctionData(288, rm.GetString("CEILING"), 2, 2, true, TFmReturnType.Value, "VV"));
			Add(Result, new TCellFunctionData(289, rm.GetString("HYPGEOMDIST"), 4, 4, true, TFmReturnType.Value, "VVVV"));
			Add(Result, new TCellFunctionData(290, rm.GetString("LOGNORMDIST"), 3, 3, true, TFmReturnType.Value, "VVV"));
			Add(Result, new TCellFunctionData(291, rm.GetString("LOGINV"), 3, 3, true, TFmReturnType.Value, "VVV"));
			Add(Result, new TCellFunctionData(292, rm.GetString("NEGBINOMDIST"), 3, 3, true, TFmReturnType.Value, "VVV"));
			Add(Result, new TCellFunctionData(293, rm.GetString("NORMDIST"), 4, 4, true, TFmReturnType.Value, "VVVV"));
			Add(Result, new TCellFunctionData(294, rm.GetString("NORMSDIST"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(295, rm.GetString("NORMINV"), 3, 3, true, TFmReturnType.Value, "VVV"));
			Add(Result, new TCellFunctionData(296, rm.GetString("NORMSINV"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(297, rm.GetString("STANDARDIZE"), 3, 3, true, TFmReturnType.Value, "VVV"));
			Add(Result, new TCellFunctionData(298, rm.GetString("ODD"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(299, rm.GetString("PERMUT"), 2, 2, true, TFmReturnType.Value, "VV"));
			Add(Result, new TCellFunctionData(300, rm.GetString("POISSON"), 3, 3, true, TFmReturnType.Value, "VVV"));
			Add(Result, new TCellFunctionData(301, rm.GetString("TDIST"), 3, 3, true, TFmReturnType.Value, "VVV"));
			Add(Result, new TCellFunctionData(302, rm.GetString("WEIBULL"), 4, 4, true, TFmReturnType.Value, "VVVV"));
			Add(Result, new TCellFunctionData(303, rm.GetString("SUMXMY2"), 2, 2, true, TFmReturnType.Value, "AA"));
			Add(Result, new TCellFunctionData(304, rm.GetString("SUMX2MY2"), 2, 2, true, TFmReturnType.Value, "AA"));
			Add(Result, new TCellFunctionData(305, rm.GetString("SUMX2PY2"), 2, 2, true, TFmReturnType.Value, "AA"));
			Add(Result, new TCellFunctionData(306, rm.GetString("CHITEST"), 2, 2, true, TFmReturnType.Value, "AA"));
			Add(Result, new TCellFunctionData(307, rm.GetString("CORREL"), 2, 2, true, TFmReturnType.Value, "AA"));
			Add(Result, new TCellFunctionData(308, rm.GetString("COVAR"), 2, 2, true, TFmReturnType.Value, "AA"));
			Add(Result, new TCellFunctionData(309, rm.GetString("FORECAST"), 3, 3, true, TFmReturnType.Value, "VAA"));
			Add(Result, new TCellFunctionData(310, rm.GetString("FTEST"), 2, 2, true, TFmReturnType.Value, "AA"));
			Add(Result, new TCellFunctionData(311, rm.GetString("INTERCEPT"), 2, 2, true, TFmReturnType.Value, "AA"));
			Add(Result, new TCellFunctionData(312, rm.GetString("PEARSON"), 2, 2, true, TFmReturnType.Value, "AA"));
			Add(Result, new TCellFunctionData(313, rm.GetString("RSQ"), 2, 2, true, TFmReturnType.Value, "AA"));
			Add(Result, new TCellFunctionData(314, rm.GetString("STEYX"), 2, 2, true, TFmReturnType.Value, "AA"));
			Add(Result, new TCellFunctionData(315, rm.GetString("SLOPE"), 2, 2, true, TFmReturnType.Value, "AA"));
			Add(Result, new TCellFunctionData(316, rm.GetString("TTEST"), 4, 4, true, TFmReturnType.Value, "AAVV"));
			Add(Result, new TCellFunctionData(317, rm.GetString("PROB"), 3, 4, true, TFmReturnType.Value, "AAVV"));
            Add(Result, new TCellFunctionData(318, rm.GetString("DEVSQ"), 1, 255, true, TFmReturnType.Value, "R"));
            Add(Result, new TCellFunctionData(319, rm.GetString("GEOMEAN"), 1, 255, true, TFmReturnType.Value, "R"));
            Add(Result, new TCellFunctionData(320, rm.GetString("HARMEAN"), 1, 255, true, TFmReturnType.Value, "R"));
            Add(Result, new TCellFunctionData(321, rm.GetString("SUMSQ"), 0, 255, true, TFmReturnType.Value, "R"));
            Add(Result, new TCellFunctionData(322, rm.GetString("KURT"), 1, 255, true, TFmReturnType.Value, "R"));
            Add(Result, new TCellFunctionData(323, rm.GetString("SKEW"), 1, 255, true, TFmReturnType.Value, "R"));
			Add(Result, new TCellFunctionData(324, rm.GetString("ZTEST"), 2, 3, true, TFmReturnType.Value, "RVV"));
			Add(Result, new TCellFunctionData(325, rm.GetString("LARGE"), 2, 2, true, TFmReturnType.Value, "RV"));
			Add(Result, new TCellFunctionData(326, rm.GetString("SMALL"), 2, 2, true, TFmReturnType.Value, "RV"));
			Add(Result, new TCellFunctionData(327, rm.GetString("QUARTILE"), 2, 2, true, TFmReturnType.Value, "RV"));
			Add(Result, new TCellFunctionData(328, rm.GetString("PERCENTILE"), 2, 2, true, TFmReturnType.Value, "RV"));
			Add(Result, new TCellFunctionData(329, rm.GetString("PERCENTRANK"), 2, 3, true, TFmReturnType.Value, "RVV"));
            Add(Result, new TCellFunctionData(330, rm.GetString("MODE"), 1, 255, true, TFmReturnType.Value, "A"));
			Add(Result, new TCellFunctionData(331, rm.GetString("TRIMMEAN"), 2, 2, true, TFmReturnType.Value, "RV"));
			Add(Result, new TCellFunctionData(332, rm.GetString("TINV"), 2, 2, true, TFmReturnType.Value, "VV"));
            Add(Result, new TCellFunctionData(336, rm.GetString("CONCATENATE"), 0, 255, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(337, rm.GetString("POWER"), 2, 2, true, TFmReturnType.Value, "VV"));
			Add(Result, new TCellFunctionData(342, rm.GetString("RADIANS"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(343, rm.GetString("DEGREES"), 1, 1, true, TFmReturnType.Value, "V"));
            Add(Result, new TCellFunctionData(344, rm.GetString("SUBTOTAL"), 2, 255, true, TFmReturnType.Value, "VR"));
			Add(Result, new TCellFunctionData(345, rm.GetString("SUMIF"), 2, 3, true, TFmReturnType.Value, "RVR"));
			Add(Result, new TCellFunctionData(346, rm.GetString("COUNTIF"), 2, 2, true, TFmReturnType.Value, "RV"));
			Add(Result, new TCellFunctionData(347, rm.GetString("COUNTBLANK"), 1, 1, true, TFmReturnType.Value, "R"));
			Add(Result, new TCellFunctionData(350, rm.GetString("ISPMT"), 4, 4, true, TFmReturnType.Value, "VVVV"));
			Add(Result, new TCellFunctionData(351, rm.GetString("DATEDIF"), 3, 3, true, TFmReturnType.Value, "VVV"));
			Add(Result, new TCellFunctionData(352, rm.GetString("DATESTRING"), 1, 1, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(353, rm.GetString("NUMBERSTRING"), 2, 2, true, TFmReturnType.Value, "VV"));
			Add(Result, new TCellFunctionData(354, rm.GetString("ROMAN"), 1, 2, true, TFmReturnType.Value, "VV"));
            Add(Result, new TCellFunctionData(358, rm.GetString("GETPIVOTDATA"), 2, 255, true, TFmReturnType.Value, "V"));
			Add(Result, new TCellFunctionData(359, rm.GetString("HYPERLINK"), 1, 2, true, TFmReturnType.Value, "VV"));
			Add(Result, new TCellFunctionData(360, rm.GetString("PHONETIC"), 1, 1, true, TFmReturnType.Value, "R"));
            Add(Result, new TCellFunctionData(361, rm.GetString("AVERAGEA"), 1, 255, true, TFmReturnType.Value, "R"));
            Add(Result, new TCellFunctionData(362, rm.GetString("MAXA"), 1, 255, true, TFmReturnType.Value, "R"));
            Add(Result, new TCellFunctionData(363, rm.GetString("MINA"), 1, 255, true, TFmReturnType.Value, "R"));
            Add(Result, new TCellFunctionData(364, rm.GetString("STDEVPA"), 1, 255, true, TFmReturnType.Value, "R"));
            Add(Result, new TCellFunctionData(365, rm.GetString("VARPA"), 1, 255, true, TFmReturnType.Value, "R"));
            Add(Result, new TCellFunctionData(366, rm.GetString("STDEVA"), 1, 255, true, TFmReturnType.Value, "R"));
            Add(Result, new TCellFunctionData(367, rm.GetString("VARA"), 1, 255, true, TFmReturnType.Value, "R"));

            //Excel 2007
            Add(Result, new TCellFunctionData(0x1E0, rm.GetString("IFERROR"), 2, 2, true, TFmReturnType.Value, "VV", true, false));
            Add(Result, new TCellFunctionData(0x1E1, rm.GetString("COUNTIFS"), 2, 255, true, TFmReturnType.Value, "(RV)", true, false));
            Add(Result, new TCellFunctionData(0x1E2, rm.GetString("SUMIFS"), 3, 255, true, TFmReturnType.Value, "RRV(RV)", true, false));
            Add(Result, new TCellFunctionData(0x1E3, rm.GetString("AVERAGEIF"), 2, 3, true, TFmReturnType.Value, "RVR", true, false));
            Add(Result, new TCellFunctionData(0x1E4, rm.GetString("AVERAGEIFS"), 3, 255, true, TFmReturnType.Value, "RRV(RV)", true, false));

            //Excel 2010
            Add(Result, new TCellFunctionData((int)TFutureFunctions.CeilingPrecise, rm.GetString("CEILING.PRECISE"), 1, 2, true, TFmReturnType.Value, "VV", true, true));
            Add(Result, new TCellFunctionData((int)TFutureFunctions.FloorPrecise, rm.GetString("FLOOR.PRECISE"), 1, 2, true, TFmReturnType.Value, "VV", true, true));
            Add(Result, new TCellFunctionData((int)TFutureFunctions.IsoCeiling, rm.GetString("ISO.CEILING"), 1, 2, true, TFmReturnType.Value, "VV", true, true));
            Add(Result, new TCellFunctionData((int)TFutureFunctions.Aggregate, rm.GetString("AGGREGATE"), 3, 255, true, TFmReturnType.Value, "VVR", true, true));
            Add(Result, new TCellFunctionData((int)TFutureFunctions.PercentileExc, rm.GetString("PERCENTILE.EXC"), 2, 2, true, TFmReturnType.Value, "RV", true, true));
            Add(Result, new TCellFunctionData((int)TFutureFunctions.QuartileExc, rm.GetString("QUARTILE.EXC"), 2, 2, true, TFmReturnType.Value, "RV", true, true));

            //Renamed in Excel 2010
            Add(Result, new TCellFunctionData((int)TFutureFunctions.BetaDist, rm.GetString("BETA.DIST"), 4, 6, true, TFmReturnType.Value, "VVVVVV", true, true));  //changed params
            Add(Result, new TCellFunctionData((int)TFutureFunctions.BetaInv, rm.GetString("BETA.INV"), 3, 5, true, TFmReturnType.Value, "VVVVV", true, true));
            Add(Result, new TCellFunctionData((int)TFutureFunctions.BinomDist, rm.GetString("BINOM.DIST"), 4, 4, true, TFmReturnType.Value, "VVVV", true, true));
            Add(Result, new TCellFunctionData((int)TFutureFunctions.BinomInv, rm.GetString("BINOM.INV"), 3, 3, true, TFmReturnType.Value, "VVV", true, true));
            Add(Result, new TCellFunctionData((int)TFutureFunctions.ChisqDistRt, rm.GetString("CHISQ.DIST.RT"), 2, 2, true, TFmReturnType.Value, "VV", true, true));
            
            Add(Result, new TCellFunctionData((int)TFutureFunctions.ChisqInvRt, rm.GetString("CHISQ.INV.RT"), 2, 2, true, TFmReturnType.Value, "VV", true, true));
            Add(Result, new TCellFunctionData((int)TFutureFunctions.ChisqTest, rm.GetString("CHISQ.TEST"), 2, 2, true, TFmReturnType.Value, "AA", true, true));
            Add(Result, new TCellFunctionData((int)TFutureFunctions.ConfidenceNorm, rm.GetString("CONFIDENCE.NORM"), 3, 3, true, TFmReturnType.Value, "VVV", true, true));
            Add(Result, new TCellFunctionData((int)TFutureFunctions.CovarianceP, rm.GetString("COVARIANCE.P"), 2, 2, true, TFmReturnType.Value, "AA", true, true));
            Add(Result, new TCellFunctionData((int)TFutureFunctions.ExponDist, rm.GetString("EXPON.DIST"), 3, 3, true, TFmReturnType.Value, "VVV", true, true));
            
            Add(Result, new TCellFunctionData((int)TFutureFunctions.FDistRt, rm.GetString("F.DIST.RT"), 3, 3, true, TFmReturnType.Value, "VVV", true, true));
            Add(Result, new TCellFunctionData((int)TFutureFunctions.FInvRt, rm.GetString("F.INV.RT"), 3, 3, true, TFmReturnType.Value, "VVV", true, true));
            Add(Result, new TCellFunctionData((int)TFutureFunctions.FTest, rm.GetString("F.TEST"), 2, 2, true, TFmReturnType.Value, "AA", true, true));
            Add(Result, new TCellFunctionData((int)TFutureFunctions.GammaDist, rm.GetString("GAMMA.DIST"), 4, 4, true, TFmReturnType.Value, "VVVV", true, true));
            Add(Result, new TCellFunctionData((int)TFutureFunctions.GammaInv, rm.GetString("GAMMA.INV"), 3, 3, true, TFmReturnType.Value, "VVV", true, true));

            Add(Result, new TCellFunctionData((int)TFutureFunctions.HypGeomDist, rm.GetString("HYPGEOM.DIST"), 5, 5, true, TFmReturnType.Value, "VVVVV", true, true)); //changed params
            Add(Result, new TCellFunctionData((int)TFutureFunctions.LogNormDist, rm.GetString("LOGNORM.DIST"), 4, 4, true, TFmReturnType.Value, "VVVV", true, true)); //changed params
            Add(Result, new TCellFunctionData((int)TFutureFunctions.LogNormInv, rm.GetString("LOGNORM.INV"), 3, 3, true, TFmReturnType.Value, "VVV", true, true));
            Add(Result, new TCellFunctionData((int)TFutureFunctions.ModeSngl, rm.GetString("MODE.SNGL"), 1, 255, true, TFmReturnType.Value, "A", true, true));
            Add(Result, new TCellFunctionData((int)TFutureFunctions.NegBinom, rm.GetString("NEGBINOM.DIST"), 4, 4, true, TFmReturnType.Value, "VVVV", true, true)); //changed params
            
            Add(Result, new TCellFunctionData((int)TFutureFunctions.NormDist, rm.GetString("NORM.DIST"), 4, 4, true, TFmReturnType.Value, "VVVV", true, true));
            Add(Result, new TCellFunctionData((int)TFutureFunctions.NormInv, rm.GetString("NORM.INV"), 3, 3, true, TFmReturnType.Value, "VVV", true, true));
            Add(Result, new TCellFunctionData((int)TFutureFunctions.NormSDist, rm.GetString("NORM.S.DIST"), 2, 2, true, TFmReturnType.Value, "VV", true, true)); //changed params
            Add(Result, new TCellFunctionData((int)TFutureFunctions.NormSInv, rm.GetString("NORM.S.INV"), 1, 1, true, TFmReturnType.Value, "V", true, true));
            Add(Result, new TCellFunctionData((int)TFutureFunctions.PercentileInc, rm.GetString("PERCENTILE.INC"), 2, 2, true, TFmReturnType.Value, "RV", true, true));

            Add(Result, new TCellFunctionData((int)TFutureFunctions.QuartileInc, rm.GetString("QUARTILE.INC"), 2, 2, true, TFmReturnType.Value, "RV", true, true));
            Add(Result, new TCellFunctionData((int)TFutureFunctions.PercentRankInc, rm.GetString("PERCENTRANK.INC"), 2, 3, true, TFmReturnType.Value, "RVV", true, true));
            Add(Result, new TCellFunctionData((int)TFutureFunctions.PoissonDist, rm.GetString("POISSON.DIST"), 3, 3, true, TFmReturnType.Value, "VVV", true, true));
            Add(Result, new TCellFunctionData((int)TFutureFunctions.RankEq, rm.GetString("RANK.EQ"), 2, 3, true, TFmReturnType.Value, "VRV", true, true));
            Add(Result, new TCellFunctionData((int)TFutureFunctions.StDevP, rm.GetString("STDEV.P"), 1, 255, true, TFmReturnType.Value, "R", true, true));

            Add(Result, new TCellFunctionData((int)TFutureFunctions.StDevS, rm.GetString("STDEV.S"), 1, 255, true, TFmReturnType.Value, "R", true, true));
            Add(Result, new TCellFunctionData((int)TFutureFunctions.TDist2T, rm.GetString("T.DIST.2T"), 2, 2, true, TFmReturnType.Value, "VV", true, true)); //changed params
            Add(Result, new TCellFunctionData((int)TFutureFunctions.TDistRT, rm.GetString("T.DIST.RT"), 2, 2, true, TFmReturnType.Value, "VV", true, true)); //changed params
            Add(Result, new TCellFunctionData((int)TFutureFunctions.TInv2T, rm.GetString("T.INV.2T"), 2, 2, true, TFmReturnType.Value, "VV", true, true));
            Add(Result, new TCellFunctionData((int)TFutureFunctions.TTest, rm.GetString("T.TEST"), 4, 4, true, TFmReturnType.Value, "AAVV", true, true));

            Add(Result, new TCellFunctionData((int)TFutureFunctions.VarP, rm.GetString("VAR.P"), 1, 255, true, TFmReturnType.Value, "R", true, true));
            Add(Result, new TCellFunctionData((int)TFutureFunctions.VarS, rm.GetString("VAR.S"), 1, 255, true, TFmReturnType.Value, "R", true, true));
            Add(Result, new TCellFunctionData((int)TFutureFunctions.WeibullDist, rm.GetString("WEIBULL.DIST"), 4, 4, true, TFmReturnType.Value, "VVVV", true, true));
            Add(Result, new TCellFunctionData((int)TFutureFunctions.ZTest, rm.GetString("Z.TEST"), 2, 3, true, TFmReturnType.Value, "RVV", true, true));
            
            Add(Result, new TCellFunctionData(255, rm.GetString("USER.DEFINED"), 0, 255, true, TFmReturnType.Value, "R"));

			foreach (TCellFunctionData fd in Result.Values)
            {
                if (fd.Name.Length==0)
                    FlxMessages.ThrowException(FlxErr.ErrUndefinedFunction, fd.Index);
            }

            return Result;
        }

        private static void Add(TCellFunctionDataDictionary Result, TCellFunctionData Func)
        {
            Result.Add(Func.Name, Func);
        }

        internal static TCellFunctionData GetData(string Name)
        {
#if(FRAMEWORK20)
            TCellFunctionData Value;
            if (Ht.TryGetValue(Name, out Value)) return Value;
            return null;
#else
            return (TCellFunctionData)Ht[Name];
#endif
        }

		private static TCellFunctionData[] CreateIndexFunc()
		{
			TCellFunctionData[] Result = new TCellFunctionData[MaxFunctions];
			foreach (TCellFunctionData fd in Ht.Values)
			{
				Result[fd.Index]=fd;
			}
			return Result;
		}

        internal static TCellFunctionData GetData(int Index, out bool Found)
        {
			if ((Index<0) || (Index>=MaxFunctions))
			{
				Found = false;
				return null;
			}
			
            TCellFunctionData Result = IndexFunc[Index];
			if (Result == null)
			{
				Found = false;
				return null;
			}

			Found = true;
			return Result;
        }

		internal static TCellFunctionData GetData(int Index)
		{
			bool Found;
			TCellFunctionData Result = GetData(Index, out Found);
			if (!Found)	FlxMessages.ThrowException(FlxErr.ErrInvalidValue, "Function id", Index, 0, MaxFunctions-1);
			return Result;		
		}
	}
}
