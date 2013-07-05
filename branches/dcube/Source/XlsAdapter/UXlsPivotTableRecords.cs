using System;
using System.IO;
using System.Diagnostics;
using System.Text;
using FlexCel.Core;

/* 
 * Records inside a pivot table. Some of them go into the Workbook stream, some others go into a dedicated cache stream.
 */ 
namespace FlexCel.XlsAdapter
{
    internal class TPivotTableRecord : TxBaseRecord
    {
        internal TPivotTableRecord(int aId, byte[] aData): base(aId, aData){}

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.PivotView.PivotItems.Add(this);
        }
    }

    /// <summary>
    /// View definition.
    /// </summary>
    internal class TSxViewRecord: TPivotTableRecord
    {
        internal TSxViewRecord(int aId, byte[] aData): base(aId, aData){}

        internal int FirstRow{get {return GetWord(0);} set {SetWord(0, value);}}

        internal int LastRow{get {return GetWord(2);} set {SetWord(2, value);}}
        internal int FirstCol{get {return GetWord(4);} set {SetWord(4, value);}}
        internal int LastCol{get {return GetWord(6);} set {SetWord(6, value);}}
    }

    #region Global records
    internal class TPivotCacheRecord : TPivotTableRecord
    {
        internal TPivotCacheRecord(int aId, byte[] aData) : base(aId, aData) { }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.PivotCache.Add(this);
        }
    }
    #endregion

    #region Sheet records
    internal class TPivotSheetRecord : TPivotTableRecord
    {
        internal TPivotSheetRecord(int aId, byte[] aData) : base(aId, aData) { }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.PivotCache.Add(this);
        }
    }
    #endregion

    #region ChartRecords
    internal class TSxViewLinkRecord : TPivotTableRecord
    {
        internal TSxViewLinkRecord(int aId, byte[] aData) : base(aId, aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            TFlxChart chart = ws as TFlxChart;
            if (chart != null)
            {
                chart.ViewLink = this;
                return;
            }

            base.LoadIntoSheet(ws, rRow, RecordLoader, ref Loader);
        }
    }

    internal class TPivotChartBitsRecord : TPivotTableRecord
    {
        internal TPivotChartBitsRecord(int aId, byte[] aData) : base(aId, aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            TFlxChart chart = ws as TFlxChart;
            if (chart != null)
            {
                chart.PivotChartBits = this;
                return;
            }

            base.LoadIntoSheet(ws, rRow, RecordLoader, ref Loader);
        }
    }

    internal class TChartSBaseRefRecord : TPivotTableRecord
    {
        internal TChartSBaseRefRecord(int aId, byte[] aData) : base(aId, aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            TFlxChart chart = ws as TFlxChart;
            if (chart != null)
            {
                chart.SBaseRef = this;
                return;
            }

            base.LoadIntoSheet(ws, rRow, RecordLoader, ref Loader);
        }
    }

    #endregion
}
