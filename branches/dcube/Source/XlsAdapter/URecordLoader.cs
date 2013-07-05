using System;
using FlexCel.Core;
using System.Diagnostics;
using System.Collections.Generic;

namespace FlexCel.XlsAdapter
{
    internal class TVirtualCellList: ICellList
    {
        ExcelFile FXls;
        TWorkbook FWorkbook;
        Dictionary<TOneCellRef, TFormulaRecord> ArrayFmlas;
        Dictionary<TOneCellRef, TFormulaRecord> TableFmlas;

        internal TVirtualCellList(ExcelFile aXls, TWorkbook aWorkbook)
        {
            FWorkbook = aWorkbook;
            FXls = aXls;
            ArrayFmlas = new Dictionary<TOneCellRef, TFormulaRecord>();
            TableFmlas = new Dictionary<TOneCellRef, TFormulaRecord>();
        }

        public bool FoundArrayFormula(int RowArr, int ColArr, out TArrayRecord ArrData)
        {
            ArrData = null;
            TFormulaRecord f;
            if (!ArrayFmlas.TryGetValue(new TOneCellRef(RowArr, ColArr), out f)) return false;
            ArrData = f.ArrayRecord;
            return ArrData != null;
        }

        public bool FoundTableFormula(int TopRow, int LeftCol, out TTableRecord TableData)
        {
            TableData = null;
            TFormulaRecord f;
            if (!TableFmlas.TryGetValue(new TOneCellRef(TopRow, LeftCol), out f)) return false;
            TableData = f.TableRecord;
            return TableData != null;
        }

        public TParsedTokenList ArrayFormula(int Row, int Col)
        {
            TFormulaRecord f;
            if (!ArrayFmlas.TryGetValue(new TOneCellRef(Row, Col), out f)) return null;
            TArrayRecord ArrData = f.ArrayRecord;
            if (ArrData == null) return null;
            return ArrData.Data;
        }

        public TTableRecord TableFormula(int Row, int Col)
        {
            TFormulaRecord f;
            if (!TableFmlas.TryGetValue(new TOneCellRef(Row, Col), out f)) return null;
            TTableRecord TableData = f.TableRecord;
            if (TableData == null) return null;
            return TableData;
        }

        public ExcelFile Workbook
        {
            get { return FXls; }
        }

        public TWorkbookGlobals Globals
        {
            get { return FWorkbook.Globals; }
        }

        public void AddArray(int aRow, int aCol, TFormulaRecord fm)
        {
            ArrayFmlas.Add(new TOneCellRef(aRow, aCol), fm);
        }

        public void AddTable(int aRow, int aCol, TFormulaRecord fm)
        {
            TableFmlas.Add(new TOneCellRef(aRow, aCol), fm);
        }

        public void ClearSheet()
        {
            ArrayFmlas.Clear();
            TableFmlas.Clear();
        }
    }

    internal class T2dCellList
    {
        Dictionary<TOneCellRef, TFormulaRecord> ArrayFmlas;

        public T2dCellList()
        {
            ArrayFmlas = new Dictionary<TOneCellRef, TFormulaRecord>();
        }

        public void Add(TXlsCellRange Range, TFormulaRecord Fmla)
        {
            for (int r = Range.Top; r <= Range.Bottom; r++)
            {
                for (int c = Range.Left; c <= Range.Right; c++)
                {
                    if (r != Range.Top || c != Range.Left)
                    {
                        ArrayFmlas.Add(new TOneCellRef(r, c), Fmla);
                    }
                }
            }
        }

        public TFormulaRecord Get(int aRow, int aCol)
        {
            TFormulaRecord f;
            if (!ArrayFmlas.TryGetValue(new TOneCellRef(aRow, aCol), out f)) return null;
            return f;
        }

        internal void ClearSheet()
        {
            ArrayFmlas.Clear();
        }
    }

    internal class TVirtualReader
    {
        int Sheet;
        int Row;
        int Col;
        TCellRecord CellRecord;

        ExcelFile Xls;
        public TVirtualCellList CellList;
        T2dCellList ArrayFormulas;

        public TVirtualReader(ExcelFile aXls, TWorkbook aWorkbook)
        {
            CellList = new TVirtualCellList(aXls, aWorkbook);
            ArrayFormulas = new T2dCellList();
            Xls = aXls;
        }

        public void ClearSheet()
        {
            CellList.ClearSheet();
            ArrayFormulas.ClearSheet();
        }

        internal void Read(int aSheet, int aRow, int aCol, TCellRecord aRecord)
        {
            Flush();

            Sheet = aSheet;
            Row = aRow;
            Col = aCol;
            CellRecord = aRecord;
        }

        internal void Flush()
        {
            if (CellRecord != null) //We delay by one so we give time for formula records to read their string records.
            {
                Xls.OnVirtualCellRead(Xls, new VirtualCellReadEventArgs(new CellValue(Sheet, Row, Col, CellRecord.GetValue(CellList), CellRecord.XF)));
                CellRecord = null;
            }
        }

        internal void AddArray(TXlsCellRange Range, TFormulaRecord Fmla)
        {
            ArrayFormulas.Add(Range, Fmla);
        }

        internal TFormulaRecord GetArray(int aRow, int aCol)
        {
            TFormulaRecord f = ArrayFormulas.Get(aRow, aCol);

            if (f == null) return null;
            TFormulaRecord f1 = new TFormulaRecord(f.Id, aRow, aCol, null, f.XF, f.CloneData(), null, null, (int)f.OptionFlags, false, 0, f.bx); //ArrayOptionFlags doesn't matter here, this doesn't have an array.
            return f1;
            
        }

        internal void StartReading()
        {
            Xls.OnVirtualCellStartReading(Xls, new VirtualCellStartReadingEventArgs(Xls));
        }
    }

    /// <summary>
    /// An Abstract loader that can be used to load different file formats on FlexCel.
    /// </summary>
    internal abstract class TBaseRecordLoader
    {
        #region Record State
        internal TRecordHeader RecordHeader;
        internal TSST SST;
        internal IFlexCelFontList FontList;
        internal TEncryptionData Encryption;
        internal UInt32 XFCRC;
        internal int XFCount;
        internal TXlsBiffVersion XlsBiffVersion;
        protected TBiff8XFMap XFMap;
        internal TNameRecordList Names;
        internal TVirtualReader VirtualReader;
        #endregion

        protected TBaseRecordLoader(TSST aSST, IFlexCelFontList aFontList, TEncryptionData aEncryption, TXlsBiffVersion aXlsBiffVersion, 
            TBiff8XFMap aXFMap, TNameRecordList aNames, TVirtualReader aVirtualReader)
        {
            RecordHeader = new TRecordHeader();
            SST = aSST;
            FontList = aFontList;
            Encryption = aEncryption;
            XlsBiffVersion = aXlsBiffVersion;
            XFMap = aXFMap;
            Names = aNames;
            VirtualReader = aVirtualReader;
        }

        internal TBaseRecord LoadRecord(bool InGlobalSection)
        {
            int rRow = 0;
            return LoadRecord(out rRow, InGlobalSection);
        }

        internal abstract TBaseRecord LoadRecord(out int rRow, bool InGlobalSection);

        internal abstract TBaseRecord LoadUnsupportdedRecord();

        internal abstract void ReadHeader();
    }

    internal abstract class TBinRecordLoader : TBaseRecordLoader
    {
        protected TBinRecordLoader(TSST aSST, IFlexCelFontList aFontList, TEncryptionData aEncryption, 
            TXlsBiffVersion aXlsBiffVersion, TBiff8XFMap aXFMap, TNameRecordList aNames, TVirtualReader aVirtualReader) 
            : base(aSST, aFontList, aEncryption, aXlsBiffVersion, aXFMap, aNames, aVirtualReader) { }

        internal abstract bool Eof { get; }

        internal void SwitchSheet()
        {
            if (VirtualReader != null) VirtualReader.ClearSheet();
        }
    }
}
