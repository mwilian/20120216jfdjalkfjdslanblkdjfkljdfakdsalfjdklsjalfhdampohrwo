using System;
using FlexCel.Core;
using System.Collections.Generic;

#if (MONOTOUCH)
using Color = MonoTouch.UIKit.UIColor;
#endif

#if (WPF)
using System.Windows.Media;
#else
using System.Drawing;
#endif

namespace FlexCel.XlsAdapter
{
	/// <summary>
	/// An Excel color Palette.
	/// </summary>
	internal class TPaletteRecord: TxBaseRecord
	{
        #region Standard palette
        private static readonly long[] StandardPaletteData ={       //STATIC*
                                                     0,
                                                     16777215,
                                                     255,
                                                     65280,
                                                     16711680,
                                                     65535,
                                                     16711935,
                                                     16776960,
                                                     128,
                                                     32768,
                                                     8388608,
                                                     32896,
                                                     8388736,
                                                     8421376,
                                                     12632256,
                                                     8421504,
                                                     16751001,
                                                     6697881,
                                                     13434879,
                                                     16777164,
                                                     6684774,
                                                     8421631,
                                                     13395456,
                                                     16764108,
                                                     8388608,
                                                     16711935,
                                                     65535,
                                                     16776960,
                                                     8388736,
                                                     128,
                                                     8421376,
                                                     16711680,
                                                     16763904,
                                                     16777164,
                                                     13434828,
                                                     10092543,
                                                     16764057,
                                                     13408767,
                                                     16751052,
                                                     10079487,
                                                     16737843,
                                                     13421619,
                                                     52377,
                                                     52479,
                                                     39423,
                                                     26367,
                                                     10053222,
                                                     9868950,
                                                     6697728,
                                                     6723891,
                                                     13056,
                                                     13107,
                                                     13209,
                                                     6697881,
                                                     10040115,
                                                     3355443
                                                 };
        #endregion

        //Must be created after the standard palette.
        internal static readonly TPaletteRecord StandardPalette = CreateStandard(); //STATIC*

        Dictionary<int, string> ColorLocator;
        Color[] RgbColorCache;
        TLabColor[] LabColorCache;

		internal TPaletteRecord(int aId, byte[] aData): base(aId, aData)
        {
            if (Data.Length < 2 + XlsConsts.HighColorPaletteRange * 4 || GetWord(0) < XlsConsts.HighColorPaletteRange)
            {
                byte[] NewData = new byte[2 + XlsConsts.HighColorPaletteRange * 4];
                SetWord(0, XlsConsts.HighColorPaletteRange);
                Array.Copy(Data, 0, NewData, 0, Data.Length);
                Array.Copy(StandardPalette.Data, Data.Length, NewData, Data.Length, NewData.Length - Data.Length);
                Data = NewData;
            }
            ColorLocator = new Dictionary<int,string>();
            LabColorCache = new TLabColor[XlsConsts.HighColorPaletteRange - XlsConsts.LowColorPaletteRange + 1];
            RgbColorCache = new Color [XlsConsts.HighColorPaletteRange - XlsConsts.LowColorPaletteRange + 1];

            for (int i = XlsConsts.LowColorPaletteRange - 1; i < XlsConsts.HighColorPaletteRange; i++)
            {
                unchecked
                {
                    Color aColor = ColorUtil.FromArgb((int)((uint)0xFF000000 | (uint)ColorUtil.BgrToRgb(GetColor(i))));
                    LabColorCache[i] = aColor;
                    RgbColorCache[i] = aColor;
                    ColorLocator[aColor.ToArgb()] = string.Empty;
                }
            }
        }

        internal static TPaletteRecord CreateStandard()
        {
            byte[] aData= new byte[2+XlsConsts.HighColorPaletteRange*4];
            BitOps.SetWord(aData,0,XlsConsts.HighColorPaletteRange);
            for (int i=XlsConsts.LowColorPaletteRange-1; i< XlsConsts.HighColorPaletteRange;i++)
                BitOps.SetCardinal(aData,2+i*4, StandardPaletteData[i]);
            return new TPaletteRecord((int)xlr.PALETTE, aData);
        }

        private long GetColor(int Index)
        {
            if ((Index>=Count) || (Index<0)) XlsMessages.ThrowException(XlsErr.ErrXlsIndexOutBounds,Index,"Palette Color Index",0, Count-1);
            return GetCardinal(2+Index*4);
        }
         
        public void SetColor(int Index, Color aColor)
        {
            long Value = ColorUtil.BgrToRgb(aColor.ToArgb());
            if ((Index>=Count) || (Index<0)) XlsMessages.ThrowException(XlsErr.ErrTooManyEntries,Index, Count-1);
            SetCardinal(2+Index*4, Value);

            ColorLocator.Remove(RgbColorCache[Index].ToArgb());
            RgbColorCache[Index] = aColor;
            LabColorCache[Index] = aColor;
            ColorLocator[RgbColorCache[Index].ToArgb()] = string.Empty;
        }

        public Color GetRgbColor(int Index)
        {
            if ((Index >= Count) || (Index < 0)) XlsMessages.ThrowException(XlsErr.ErrXlsIndexOutBounds, Index, "Palette Color Index", 0, Count - 1);
            return RgbColorCache[Index];
        }

        public TLabColor GetLabColor(int Index)
        {
            if ((Index >= Count) || (Index < 0)) XlsMessages.ThrowException(XlsErr.ErrXlsIndexOutBounds, Index, "Palette Color Index", 0, Count - 1);
            return LabColorCache[Index];
        }

        public bool ContainsColor(Color Value)
        {
            return ColorLocator.ContainsKey(Value.ToArgb());
        }

        internal static int Count
        {
            get
            {
                return XlsConsts.HighColorPaletteRange;
            }
        }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.Palette = this;
        }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            TFlxChart chart = ws as TFlxChart;
            if (chart != null)
            {
                chart.Palette = this;
                return;
            }
            base.LoadIntoSheet(ws, rRow, RecordLoader, ref Loader);
        }

        internal bool IsStandard()
        {
            for (int i = 0; i < Count; i++)
            {
                if (StandardPalette.GetColor(i) != GetColor(i)) return false;
            }
            return true;
        }
    }

    internal class TClrtClientRecord : TxBaseRecord
    {
		internal TClrtClientRecord(int aId, byte[] aData): base(aId, aData)
        {
        }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.ClrtClient = this;
        }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            TFlxChart chart = ws as TFlxChart;
            if (chart != null)
            {
                chart.ClrtClient = this;
                return;
            }
            base.LoadIntoSheet(ws, rRow, RecordLoader, ref Loader);
        }
    }
}
