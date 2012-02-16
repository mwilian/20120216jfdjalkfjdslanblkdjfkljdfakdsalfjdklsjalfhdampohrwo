using System;
using System.IO;
using System.Diagnostics;
using System.Text;
using FlexCel.Core;

/* 
 * Basic records that we can find on an Excel file, including abstract ones.
 */
namespace FlexCel.XlsAdapter
{
    internal struct TSaveData
    {
        internal IFlexCelPalette Palette; //Might not be in the correct activesheet, so it is not really valis as XlsFile.
        internal TSheet SavingSheet;
        internal TBorderList BorderList;
        internal TWorkbookGlobals Globals;
        internal bool Repeatable;
        internal TExcludedRecords ExcludedRecords;
        internal TXlsBiffVersion BiffVersion;
        private TExcelFileErrorActions ErrorActions;
        internal double ObjectHeight; //Currently used by listboxes to calculate page.
        internal bool FixPage;//Currently used by listboxes to calculate page.

        internal int[] StylesSavedAtPos; //Will be set when saving stylexf, and used when saving the styles.
        internal int[] AddedRecords; //This should be an int, but as this is an strucure and passed by value, its value won't be preserved.

        internal TSaveData(IFlexCelPalette aPalette, TBorderList aBorderList, TWorkbookGlobals aGlobals, TXlsBiffVersion aBiffVersion, TExcelFileErrorActions aErrorActions)
        {
            Palette = aPalette;
            BorderList = aBorderList;
            Globals = aGlobals;
            ExcludedRecords = TExcludedRecords.None;
            Repeatable = false;
            BiffVersion = aBiffVersion;
            ErrorActions = aErrorActions;
            StylesSavedAtPos = null;
            AddedRecords = new int[1];
            ObjectHeight = 0;
            FixPage = false;
            SavingSheet = null;
        }

        internal TSaveData(IFlexCelPalette aPalette, TBorderList aBorderList, TWorkbookGlobals aGlobals, TXlsBiffVersion aBiffVersion, TExcelFileErrorActions aErrorActions, TExcludedRecords aExcludedRecords, bool aRepeatable) :
            this(aPalette, aBorderList, aGlobals, aBiffVersion, aErrorActions)
        {
            ExcludedRecords = aExcludedRecords;
            Repeatable = aRepeatable;
        }

        internal bool ThrowExceptionOnTooManyPageBreaks
        {
            get
            {
                return (ErrorActions & TExcelFileErrorActions.ErrorOnTooManyPageBreaks) != 0;
            }
        }

        internal UInt16 GetBiff8FromCellXF(int CellXF)
        {
            if (CellXF < 0) return 0; //not defined.
            if (CellXF == 0) return FlxConsts.DefaultFormatIdBiff8; // normal
            return (UInt16)(CellXF + Globals.StyleXF.Count + AddedRecords[0]);
        }
    }

    /// <summary>
    /// An Excel Record Header. It allows us to access it as an array or as a record
    /// </summary>
    internal class TRecordHeader
    {
        private byte[] FData;


        internal TRecordHeader()
        {
            FData = new byte[XlsConsts.SizeOfTRecordHeader];
        }

        internal TRecordHeader(int aId, int aSize)
        {
            FData = new byte[XlsConsts.SizeOfTRecordHeader];
            Id = aId;
            Size = aSize;
        }

        internal byte[] Data { get { return FData; } }

        internal int Id { get { return (Int32)(FData[0] + (FData[1] << 8)); } set { BitConverter.GetBytes((UInt16)value).CopyTo(FData, 0); } }
        internal long Size
        {
            get
            {
                int Result = (Int32)(FData[2] + (FData[3] << 8));
                if (Result > XlsConsts.MaxRecordDataSize + 10) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid); //Guard against too large record sizes.
                return Result;
            }
            set { BitConverter.GetBytes((UInt16)value).CopyTo(FData, 2); }
        }


        internal int Length { get { return FData.Length; } }
    }

    /// <summary>
    /// The base for all Excel record hierachy.
    /// All Excel file records inherit from here.
    /// </summary>
    /// 
    internal abstract class TBaseRecord
    {
        protected abstract TBaseRecord DoCopyTo(TSheetInfo SheetInfo);  //Actual method doing the copy
        internal abstract void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row);
        internal virtual void SaveToPxl(TPxlStream PxlStream, int Row, TPxlSaveData SaveData)
        {
        }

        internal virtual bool PxlRecordIsValid(int Row)
        {
            return true;
        }

        internal abstract int TotalSize();
        internal abstract int TotalSizeNoHeaders();
        internal static TBaseRecord Clone(TBaseRecord Self, TSheetInfo SheetInfo) //this should be non-virtual. It allows you to obtain a clone, even if the object is null
        {
            if (Self == null) return null;   //for this to work, this can't be a virtual method
            else return Self.DoCopyTo(SheetInfo);
        }

        internal virtual void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.FutureRecords.Add(this);
        }

        internal virtual void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            if (Globals.LoadingInterfaceHdr)
            {
                Globals.FileEncryption.InterfaceHdr.Add(this);
            }
            else
            {
                Globals.FutureRecords.Add(this);
            }
        }

        internal abstract int GetId { get; }

        internal virtual void AddContinue(TContinueRecord aContinue)
        {
            XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
        }

    }

    internal class TxBaseRecord : TBaseRecord
    {
        internal int Id;
        internal byte[] Data;
        internal TContinueRecord Continue;

        internal TxBaseRecord(int aId, byte[] aData)
        {
            Id = aId;
            Data = aData;
        }

        #region Bit Manipulation
        protected byte GetByte(int tPos) { return Data[tPos]; }

        protected int GetWord(int tPos) //{return BitOps.GetWord(Data, tPos);}  
        {
            //Optimization for a routine called million times
            unchecked
            {
                return (UInt16)(Data[tPos] + (Data[tPos + 1] << 8));
            }
        }

        protected int GetInt16(int tPos) //{return BitOps.GetWord(Data, tPos);}  
        {
            //Optimization for a routine called million times
            unchecked
            {
                return (Int16)(Data[tPos] + (Data[tPos + 1] << 8));
            }
        }

        protected void SetWord(int tPos, int number) { BitOps.SetWord(Data, tPos, number); }

        protected Int64 GetCardinal(int tPos) { return BitOps.GetCardinal(Data, tPos); }
        protected void SetCardinal(int tPos, Int64 Value) { BitOps.SetCardinal(Data, tPos, Value); }

        protected byte[] GetArray(int StartPos, int Length)
        {
            byte[] Result = new byte[Length];
            Array.Copy(Data, StartPos, Result, 0, Length);
            return Result;
        }
        #endregion

        internal int DataSize { get { return Data.Length; } }

        internal override int GetId { get { return Id; } }

        internal void SaveDataToStream(IDataStream Workbook, byte[] aData)
        {
            Workbook.WriteHeader((UInt16)Id, (UInt16)Data.Length);
            Workbook.Write(aData, aData.Length);
        }

        protected override TBaseRecord DoCopyTo(TSheetInfo SheetInfo)
        {
            TxBaseRecord b = (TxBaseRecord)MemberwiseClone();
            if (Data != null) b.Data = (byte[])Data.Clone();
            b.Id = Id;
            if (Continue != null) b.Continue = (TContinueRecord)TContinueRecord.Clone(Continue, SheetInfo);
            return b;
        }

        internal override void AddContinue(TContinueRecord aContinue)
        {
            if (Continue != null) XlsMessages.ThrowException(XlsErr.ErrInvalidContinue);
            Continue = aContinue;
        }

        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
        {
            SaveDataToStream(Workbook, Data);
            if (Continue != null) Continue.SaveToStream(Workbook, SaveData, Row);
        }

        internal override int TotalSize()
        {
            int Result = Data.Length + XlsConsts.SizeOfTRecordHeader;
            if (Continue != null) Result += Continue.TotalSize();
            return Result;
        }

        internal override int TotalSizeNoHeaders()
        {
            int Result = Data.Length;
            if (Continue != null) Result += Continue.TotalSizeNoHeaders();
            return Result;
        }
    }

    /// <summary>
    /// This record continues another normal record
    /// A continue record can contain another continue record, that can contain another...
    /// </summary>
    internal class TContinueRecord : TxBaseRecord
    {
        internal TContinueRecord(int aId, byte[] aData) : base(aId, aData) { }
    }


    /// <summary>
    /// Base for all records including row on first word and column on second.
    /// </summary>
    internal class TBaseRowColRecord : TBaseRecord
    {
        internal UInt16 Id;
        internal int Col;
        internal TBaseRowColRecord(int aId, int aCol)
            : base()
        {
            Id = (UInt16)aId;
            Col = aCol;
        }

        internal override int GetId
        {
            get { return Id; }
        }

        /// <summary>
        /// Implements deterministic destruction, for decreasing references.
        /// </summary>
        internal virtual void Destroy()
        {
        }

        internal static void IncRef(ref int v, int Offset, int Max, XlsErr ErrorWhenTooMany)
        {
            v += Offset;
            if ((v < 0) || (v > Max)) XlsMessages.ThrowException(ErrorWhenTooMany, v + 1, Max + 1);
        }

        internal virtual bool AllowCopyOnOnlyFormula
        {
            get
            {
                return false;
            }
        }

        internal virtual void ArrangeInsertRange(int Row, TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
        {
            if ((SheetInfo.InsSheet < 0) || (SheetInfo.SourceFormulaSheet != SheetInfo.InsSheet)) return;
            if ((aColCount != 0) && (Col >= CellRange.Left) && (Row >= CellRange.Top) && (Row <= CellRange.Bottom)) IncRef(ref Col, aColCount * CellRange.ColCount, FlxConsts.Max_Columns, XlsErr.ErrTooManyColumns);  //col;
        }
        internal virtual void ArrangeCopyRange(TXlsCellRange SourceRange, int Row, int RowOffset, int ColOffset, TSheetInfo SheetInfo)
        {
            Col += ColOffset;  //col;
        }

        protected override TBaseRecord DoCopyTo(TSheetInfo SheetInfo)
        {
            TBaseRowColRecord b = (TBaseRowColRecord)MemberwiseClone();
            return b;
        }

        internal override int TotalSize()
        {
            return TotalSizeNoHeaders() + XlsConsts.SizeOfTRecordHeader;
        }

        internal override int TotalSizeNoHeaders()
        {
            return 4;
        }

        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
        {
            Biff8Utils.CheckRow(Row);
            Biff8Utils.CheckCol(Col);
            unchecked
            {
                Workbook.WriteHeader(Id, (UInt16)TotalSizeNoHeaders());
                Workbook.Write32((UInt32)((UInt16)Row + (((UInt16)Col) << 16)));
            }
        }

        internal override bool PxlRecordIsValid(int Row)
        {
            return Row <= FlxConsts.Max_PxlRows && Col <= FlxConsts.Max_PxlColumns;
        }

    }

    /// <summary>
    /// Record to hold a value on a cell.
    /// </summary>
    internal abstract class TCellRecord : TBaseRowColRecord
    {
        internal int XF;
        internal TFutureStorage FutureStorage;


        internal TCellRecord(int aId, byte[] aData, TBiff8XFMap XFMap)
            : base(aId, BitOps.GetWord(aData, 2))
        {
            XF = BitOps.GetWord(aData, 4);
            if (XFMap != null)
            {
                XF = XFMap.GetCellXF2007(XF);
            }
        }

        internal TCellRecord(int aId, int aCol, int aXF)
            : base(aId, aCol)
        {
            XF = aXF;
        }

        internal virtual object GetValue(ICellList Cells)
        {
            return null;
        }

        protected override TBaseRecord DoCopyTo(TSheetInfo SheetInfo)
        {
            TCellRecord Result = (TCellRecord)base.DoCopyTo(SheetInfo);
            Result.XF = XF;
            return Result;
        }

        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
        {
            base.SaveToStream(Workbook, SaveData, Row);
            Workbook.Write16(SaveData.GetBiff8FromCellXF(XF));
        }

        internal override int TotalSizeNoHeaders()
        {
            return base.TotalSizeNoHeaders() + 2;
        }

        internal virtual bool CanJoinNext(TCellRecord NextRecord, int MaxCol)
        {
            return false;
        }

        internal virtual void SaveFirstMul(IDataStream Workbook, TSaveData SaveData, int Row, int JoinedRecordSize)
        {
            SaveToStream(Workbook, SaveData, Row);
        }

        internal virtual void SaveMidMul(IDataStream Workbook, TSaveData SaveData)
        {
        }

        internal virtual void SaveLastMul(IDataStream Workbook, TSaveData SaveData)
        {
        }

        internal virtual int TotalSizeFirst()
        {
            return TotalSize();
        }

        internal virtual int TotalSizeMid()
        {
            return TotalSize();
        }

        internal virtual int TotalSizeLast()
        {
            return TotalSize();
        }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.Cells.AddCell(this, rRow, RecordLoader.VirtualReader);
        }

        internal void AddFutureStorage(TFutureStorageRecord R)
        {
            TFutureStorage.Add(ref FutureStorage, R);
        }

#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
        internal abstract void SaveToXlsx(TOpenXmlWriter DataStream, int Row, TCellList CellList, bool Dates1904);
#endif

    }

    /// <summary>
    /// Record with ROW attributes (like Height, XF of the row, etc)
    /// </summary>
    internal class TRowRecord : TBaseRowColRecord
    {
        internal UInt16 rMinCol;
        internal UInt16 rMaxCol;
        internal UInt16 FHeight;
        //IRWMAC=0;
        internal UInt16 Reserved;
        private byte OptionFlags1;
        private byte OptionFlags2;
        int FXF;
        internal bool ThickTop, ThickBottom;
        internal bool Phonetic;

        internal TFutureStorage FutureStorage;

        internal bool MarkForAutofit; //internal use.
        internal bool HasMergedCell;//internal use
        internal float AutofitAdjustment; //internal use.
        internal int AutofitAdjustmentFixed; //internal use.
        internal int MinHeight;//internal use
        internal int MaxHeight;//internal use
        internal int KeepTogether; //Internal use.
        internal TCellCondFmt[] CondFmt;//internal use

        internal TRowRecord(int aId, byte[] aData, TBiff8XFMap aXFMap)
            : base(aId, 0) //Column is not needed here.
        {
            rMinCol = BitConverter.ToUInt16(aData, 2);
            rMaxCol = BitConverter.ToUInt16(aData, 4);
            Height = BitConverter.ToUInt16(aData, 6);
            //IRWMAC=0;
            Reserved = BitConverter.ToUInt16(aData, 10);
            OptionFlags1 = aData[12];
            OptionFlags2 = aData[13];
            int o14 = BitConverter.ToUInt16(aData, 14);
            ThickTop = (o14 & 0x1000) != 0;
            ThickBottom = (o14 & 0x2000) != 0;
            Phonetic = (o14 & 0x4000) != 0;

            if (aXFMap == null)
            {
                FXF = o14 & 0xFFF;
            }
            else
            {
                if (IsFormatted())
                    FXF = aXFMap.GetCellXF2007(o14 & 0xFFF);
                else
                    FXF = FlxConsts.DefaultFormatId;
            }
        }

        internal TRowRecord(int aXF, bool CustomXF, int aHeight, bool CustomHeight, bool Collapsed, bool Hidden, int OutlineLevel, bool aPhonetic, bool aThickTop, bool aThickBot)
            : base((int)xlr.ROW, 0)
        {
            if (aXF < 0 || aXF > XlsConsts.MaxXFDefs || !CustomXF) FXF = (UInt16)FlxConsts.DefaultFormatId; else FXF = aXF;
            if (aHeight < 0) Height = 0; else if (aHeight > XlsConsts.MaxRowHeight) Height = XlsConsts.MaxRowHeight; else Height = (UInt16)aHeight;

            OptionFlags1 =
                (byte)
                (
                (Collapsed ? 0x10 : 0) |
                (Hidden ? 0x20 : 0) |
                (CustomHeight ? 0x40 : 0) |
                (CustomXF ? 0x80 : 0)
                );

            SetRowOutlineLevel(OutlineLevel);
            OptionFlags2 = 1;


            ThickTop = aThickTop;
            ThickBottom = aThickBot;
            Phonetic = aPhonetic;

        }

        /// <summary>
        /// Create a Standard Row
        /// </summary>
        internal TRowRecord(int DefaultHeight)
            : base((int)xlr.ROW, 0)
        {
            rMinCol = 0;
            rMaxCol = 0;
            Height = (UInt16) DefaultHeight;
            OptionFlags1 = 0;
            OptionFlags2 = 1;
            FXF = FlxConsts.DefaultFormatId;
        }

        internal int XF
        {
            get
            {
                if (IsFormatted()) return FXF; else return FlxConsts.DefaultFormatId;
            }
            set
            {
                OptionFlags1 = (byte)(OptionFlags1 | 0x80);  //Row has been formatted
                OptionFlags2 = (byte)(OptionFlags2 | 0x01);
                FXF = value;
            }
        }

        internal UInt16 Height
        {
            get
            {
                return FHeight;
            }
            set
            {
                unchecked
                {
                    if (value > XlsConsts.MaxRowHeight) value = XlsConsts.MaxRowHeight;
                    if (value < 0) value = 0;
                    FHeight = value;
                }
            }
        }

        internal void ManualHeight()
        {
            OptionFlags1 = (byte)(OptionFlags1 | 0x40);
        }

        internal void AutoHeight()
        {
            OptionFlags1 = (byte)(OptionFlags1 & (~0x40));
        }

        internal void Collapse(int level, TCollapseChildrenMode CollapseChildren, bool IsNode)
        {
            const byte mask = 0x20;  // 0x10 notes if children must be collapsed when parent is collapsed. 0x20 means row hidden and also collapse.
            const byte CMask = 0x10;

            int rowlevel = GetRowOutlineLevel();
            if (rowlevel == 0) return;

            bool NeedsCollapse = rowlevel >= level;
            if (NeedsCollapse)
            {
                OptionFlags1 = (byte)(OptionFlags1 | mask);

                if (IsNode)
                {
                    if (CollapseChildren == TCollapseChildrenMode.Collapsed) OptionFlags1 = (byte)(OptionFlags1 | CMask);
                    if (CollapseChildren == TCollapseChildrenMode.Expanded) OptionFlags1 = (byte)(OptionFlags1 & ~CMask);
                }
            }
            else
            {
                if (IsNode && rowlevel == level - 1)
                    OptionFlags1 = (byte)(OptionFlags1 & (~mask) | CMask);  //The parent node has one level less than the hidden children.
                else
                    OptionFlags1 = (byte)(OptionFlags1 & (~mask) & ~CMask);
            }


        }

        internal void Hide(bool Value)
        {
            if (Value) OptionFlags1 = (byte)(OptionFlags1 | (0x20)); else OptionFlags1 = (byte)(OptionFlags1 & (~0x20));
        }

        internal bool IsAutoHeight()
        {
            return !((OptionFlags1 & 0x40) == 0x40);
        }

        internal bool IsCollapsed()
        {
            return ((OptionFlags1 & 0x10) == 0x10);
        }

        internal bool IsHidden()
        {
            return ((OptionFlags1 & 0x20) == 0x20);
        }

        internal bool IsFormatted()
        {
            return ((OptionFlags1 & 0x80) == 0x80);
        }

        internal bool IsModified(int DefaultRowHeight, bool DefaultIsZero)
        {
            bool IsZero = (OptionFlags1 & 0x20) != 0;
            bool DifferentZero = DefaultIsZero ^ IsZero;
            bool IsMod = (((OptionFlags1 & ~0x20) != 0) || DifferentZero || OptionFlags2 != 1);
            return IsMod || Height != DefaultRowHeight;  //this will make autosized rows not to be discarded.
        }

        internal int GetOptions()
        {
            unchecked
            {
                return (UInt16)(OptionFlags1 + (OptionFlags2 << 8));
            }
        }

        internal void SetOptions(int value)
        {
            OptionFlags1 = (byte)(value & 0xFF);
            OptionFlags2 = (byte)((value >> 8) & 0xFF);
        }

        internal void SetRowOutlineLevel(int Level)
        {
            OptionFlags1 = (byte)((OptionFlags1 & ~0x7) | (Level & 0x7));
        }

        internal int GetRowOutlineLevel()
        {
            return OptionFlags1 & 0x7;
        }

        internal void SaveRangeToStream(IDataStream DataStream, TSaveData SaveData, int aMinCol, int aMaxCol, int Row)
        {
            int sMinCol = rMinCol;
            int sMaxCol = rMaxCol;
            try
            {
                if (sMinCol < aMinCol) rMinCol = (UInt16)aMinCol;
                if (sMaxCol > aMaxCol + 1) rMaxCol = (UInt16)(aMaxCol + 1);
                if (rMinCol > rMaxCol) rMinCol = rMaxCol;
                SaveToStream(DataStream, SaveData, Row);
            }
            finally
            {
                rMinCol = (UInt16)sMinCol;
                rMaxCol = (UInt16)sMaxCol;
            } //Finally

        }

        protected override TBaseRecord DoCopyTo(TSheetInfo SheetInfo)
        {
            TRowRecord Result = (TRowRecord)base.DoCopyTo(SheetInfo);
            //No need to set up all. This has been memberwisecloned.
            return Result;
        }

        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
        {
            Col = rMinCol;
            base.SaveToStream(Workbook, SaveData, Row);
            //Workbook.Write16(MinCol);
            Workbook.Write16(rMaxCol);

            UInt16 mrh = Height;
            if (mrh > XlsConsts.MaxRowHeight)
            {
                mrh = XlsConsts.MaxRowHeight;
            }

            Workbook.Write16(mrh);

            Workbook.Write16(0);//IRWMAC=0;
            Workbook.Write16(Reserved);
            Workbook.Write16((UInt16)GetOptions());
            int Biff8XF = IsFormatted() ? (int)SaveData.GetBiff8FromCellXF(FXF) : FlxConsts.DefaultFormatIdBiff8;
            Biff8Utils.CheckXF(Biff8XF);
            Workbook.Write16(XFPlusFlags(Biff8XF));
        }

        internal UInt16 XFPlusFlags(int Biff8XF)
        {
            return
                (UInt16)
                (
                    Biff8XF +
                    (ThickTop ? 1 << 12 : 0) +
                    (ThickBottom ? 1 << 13 : 0) +
                    (Phonetic ? 1 << 14 : 0)
                );
        }

        internal override void SaveToPxl(TPxlStream PxlStream, int Row, TPxlSaveData SaveData)
        {
            base.SaveToPxl(PxlStream, Row, SaveData);
            if (!PxlRecordIsValid(Row)) return;

            PxlStream.WriteByte((byte)pxl.ROW);
            PxlStream.Write16((UInt16)Row);

            UInt16 mrh = Height;
            if (mrh > XlsConsts.MaxRowHeight)
            {
                mrh = XlsConsts.MaxRowHeight;
            }

            PxlStream.Write16(mrh);
            UInt16 Options = 0;
            if (IsFormatted() && IsHidden())
            {
                Options |= 0x1; //Row hidden.
            }

            PxlStream.Write16(Options);
            PxlStream.Write16(SaveData.GetBiff8FromCellXF(XF));  //the flags are not included here.
        }


        internal override int TotalSizeNoHeaders()
        {
            return base.TotalSizeNoHeaders() + 12;
        }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            if (RecordLoader.VirtualReader != null)
            {
                return;
            }

            ws.Cells.AddRow(rRow, this);
        }

        internal void AddFutureStorage(TFutureStorageRecord R)
        {
            TFutureStorage.Add(ref FutureStorage, R);
        }
    }

    /// <summary>
    /// Dimensions of the spreadsheet. This record has to be recreated when saving.
    /// Note that the biff8 specification for this record already alows more than 65536 rows, so there is no need to change it.
    /// </summary>
    internal class TDimensionsRecord : TxBaseRecord
    {
        internal TDimensionsRecord(int aId, byte[] aData) : base(aId, aData) { }
        internal long FirstRow() { return GetCardinal(0); }
        /// <summary>
        /// Last defined row
        /// </summary>
        /// <returns>Last defined row+1</returns>
        internal long LastRow() { return GetCardinal(4); }
        internal int FirstCol() { return GetWord(8); }
        /// <summary>
        /// Last defined column
        /// </summary>
        /// <returns>Last defined column+1</returns>
        internal int LastCol() { return GetWord(10); }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.OriginalDimensions = this;
        }

    }

    /// <summary>
    /// Formula result, when it is a string.
    /// </summary>
    internal class TStringRecord : TxBaseRecord
    {
        internal TStringRecord(int aId, byte[] aData) : base(aId, aData) { }
        internal TStringRecord(string Value)
            : base((int)xlr.STRING, null)
        {
            TExcelString Xs = new TExcelString(TStrLenLength.is16bits, Value, null, false);
            Data = new byte[Xs.TotalSize()];
            Xs.CopyToPtr(Data, 0, true);
        }

        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
        {
            //We are not saving out this record...
        }

        protected override TBaseRecord DoCopyTo(TSheetInfo SheetInfo)
        {
            Debug.Assert(true, "String record can't be copied"); //
            return null;
        }


        internal override int TotalSize()
        {
            return 0;
        }

        internal override int TotalSizeNoHeaders()
        {
            return 0;
        }

        internal string Value()
        {
            TxBaseRecord Myself = this; int Ofs = 0;
            TExcelString Xs = new TExcelString(TStrLenLength.is16bits, ref Myself, ref Ofs);
            return Xs.Data;
        }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            if (Loader.LastFormula == null) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
            Loader.LastFormula.SetFormulaValue(this);
        }
    }

    internal class TStringRecordData
    {
        byte[][] Values;

        internal TStringRecordData(TStringRecord aValue)
        {
            TxBaseRecord xBase = aValue;
            int Len = 0;
            while (xBase != null)
            {
                Len++;
                xBase = xBase.Continue;
            }

            Values = new byte[Len][];

            xBase = aValue;
            int i = 0;
            while (xBase != null)
            {
                Values[i] = FillData(xBase.Id, xBase.Data);
                i++;
                xBase = xBase.Continue;
            }
        }

        private static byte[] FillData(int Id, byte[] aValue)
        {
            byte[] Result = new byte[aValue.Length + 4];
            BitOps.SetWord(Result, 0, Id);
            BitOps.SetWord(Result, 2, Result.Length - 4);
            Array.Copy(aValue, 0, Result, 4, aValue.Length);
            return Result;
        }

        internal TStringRecordData(string aValue)
        {
            TExcelString Xs = new TExcelString(TStrLenLength.is16bits, aValue.ToString(), null, false);
            byte[] StringValue = new byte[Xs.TotalSize() + 4];
            Xs.CopyToPtr(StringValue, 4);


            int Pos = XlsConsts.MaxRecordDataSize;

            int Count = 1;
            while (Pos < StringValue.Length)
            {
                Pos += XlsConsts.MaxRecordDataSize - 4 - 1; //4 is for the record header, 1 for the optionflags.
                Count++;
            }

            Values = new byte[Count][];

            if (Count == 1) //Most Common case
            {
                Values[0] = StringValue;
                BitOps.SetWord(Values[0], 0, (int)xlr.STRING);
                BitOps.SetWord(Values[0], 2, Values[0].Length - 4);
                return;
            }

            Values[0] = new byte[XlsConsts.MaxRecordDataSize];
            BitOps.SetWord(Values[0], 0, (int)xlr.STRING);
            BitOps.SetWord(Values[0], 2, Values[0].Length - 4);
            Array.Copy(StringValue, 4, Values[0], 4, XlsConsts.MaxRecordDataSize - 4);

            Pos = XlsConsts.MaxRecordDataSize;
            int i = 1;
            while (Pos < StringValue.Length)
            {
                Values[i] = new byte[Math.Min(XlsConsts.MaxRecordDataSize, StringValue.Length - Pos + 4 + 1)];

                BitOps.SetWord(Values[i], 0, (int)xlr.CONTINUE);
                BitOps.SetWord(Values[i], 2, Values[i].Length - 4);
                Values[i][4] = StringValue[6];
                Array.Copy(StringValue, Pos, Values[i], 5, Values[i].Length - 5);

                Pos += XlsConsts.MaxRecordDataSize - 4 - 1;
                i++;
            }

            Debug.Assert(i == Count, "Both calculated length and real length must be the same");
        }

        internal void SaveToStream(IDataStream Workbook)
        {
            foreach (byte[] StringRecord in Values)
            {
                Workbook.WriteHeader((UInt16)(StringRecord[0] + (StringRecord[1] << 8)), (UInt16)(StringRecord[2] + (StringRecord[3] << 8)));
                Workbook.Write(StringRecord, 4, StringRecord.Length - 4);
            }
        }

        internal void SaveToPxl(TPxlStream PxlStream)
        {
            if (Values.Length != 1) return; //Continue records are not supported.
            PxlStream.WriteByte((byte)pxl.STRING);
            byte[] StringRecord = Values[0];
            string StrValue = string.Empty;
            long StSize = 0;
            StrOps.GetSimpleString(true, StringRecord, 4, false, 0, ref StrValue, ref StSize);
            PxlStream.WriteString16(StrValue); //Wrong Docs!
        }


        internal int Length
        {
            get
            {
                int Result = 0;
                foreach (byte[] StringRecord in Values)
                {
                    Result += StringRecord.Length;
                }
                return Result;
            }
        }
    }


    /// <summary>
    /// Global window settings
    /// </summary>
    internal class TWindow1Record : TxBaseRecord
    {
        internal TFutureStorage FutureStorage;

        internal TWindow1Record(int aId, byte[] aData) : base(aId, aData) { }

        internal TWindow1Record()
            : base((int)xlr.WINDOW1,
                new byte[] { 0x00, 0x00, 0x00, 0x00, 0xF9, 0x51, 0x47, 0x31, 0x38, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x58, 0x02 })
        {
        }

        internal int xWin { get { unchecked { return (Int16)GetWord(0); } } set { unchecked { SetWord(0, (UInt16)value); } } }
        internal int yWin { get { unchecked { return (Int16)GetWord(2); } } set { unchecked { SetWord(2, (UInt16)value); } } }
        internal int dxWin { get { return MakePositiveNotZero(4); } set { SetWord(4, value); } }
        internal int dyWin { get { return MakePositiveNotZero(6); } set { SetWord(6, value); } }

        private int MakePositiveNotZero(int p)
        {
            int Result = GetWord(p);
            if (Result > 0) return Result;
            return 1;
        }

        internal int Options { get { return GetWord(8); } set { SetWord(8, value); } }

        internal int ActiveSheet
        {
            get
            {
                return GetWord(10);
            }
            set
            {
                SetWord(10, value);
                SetWord(12, 0);
                SetWord(14, 1);
            }
        }


        internal int FirstSheetVisible
        {
            get
            {
                return GetWord(12);
            }
            set
            {
                SetWord(12, value);
            }
        }

        internal int TabsSelected { get { return GetWord(14); } set { SetWord(14, value); } }
        internal int TabRatio { get { return GetWord(16); } set { SetWord(16, value); } }

        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
        {
            if ((SaveData.ExcludedRecords & TExcludedRecords.SheetSelected) != 0) return; //Note that this will invalidate the size, but it doesnt matter as this is not saved for real use. We could write blanks here if we wanted to keep the offsets right.
            dxWin = dxWin; dyWin = dyWin; //ensure values are in bounds
            base.SaveToStream(Workbook, SaveData, Row);
        }


        internal override void SaveToPxl(TPxlStream PxlStream, int Row, TPxlSaveData SaveData)
        {
            dxWin = dxWin; dyWin = dyWin; //ensure values are in bounds
            base.SaveToPxl(PxlStream, Row, SaveData);
            PxlStream.WriteByte((byte)pxl.WINDOW1);
            PxlStream.Write16(0); //option flags
            PxlStream.Write(Data, 10, 2);  //Worksheet index
        }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.AddNewWindow1();
            Globals.Window1[Globals.Window1.Length - 1] = this;
        }

        internal void AddFutureStorage(TFutureStorageRecord R)
        {
            TFutureStorage.Add(ref FutureStorage, R);
        }


        internal void SetSheetVisible(TXlsSheetVisible Value)
        {
            switch (Value)
            {
                case TXlsSheetVisible.Hidden:
                    Options = (Options & ~5) | 1;
                    break;

                case TXlsSheetVisible.VeryHidden:
                    Options = (Options & ~5) | 4;
                    break;

                case TXlsSheetVisible.Visible:
                    Options = (Options & ~5);
                    break;
            }
        }
    }

    /// <summary>
    /// Per-Sheet window settings
    /// </summary>
    internal class TWindow2Record : TBaseRecord
    {
        internal int Id;
        private int FFlags;
        internal int FirstRow;
        internal int FirstCol;
        private TExcelColor FGridLinesColor;
        internal int ScaleInPageBreakPreview;
        internal int ScaleInNormalView;
        internal int Reserved;
        internal int Len;
        private bool InChart;

        internal TWindow2Record(bool ShortVersion)
        {
            Id = (int)xlr.WINDOW2;
            Flags = 0x06B6;
            FGridLinesColor = TExcelColor.FromBiff8ColorIndex(0x40);
            if (ShortVersion) Len = 10; else Len = 18;
        }


        internal TWindow2Record(int aId, byte[] aData)
            : base()
        {
            Id = aId;
            Flags = BitOps.GetWord(aData, 0);
            FirstRow = BitOps.GetWord(aData, 2);
            FirstCol = BitOps.GetWord(aData, 4);
            FGridLinesColor = TExcelColor.FromBiff8ColorIndex(BitOps.GetWord(aData, 6));
            if (aData.Length > 10)
            {
                ScaleInPageBreakPreview = BitOps.GetWord(aData, 10);
                ScaleInNormalView = BitOps.GetWord(aData, 12);
                Reserved = BitOps.GetWord(aData, 14);
            }

            Len = aData.Length;
            if (Len != 10 && Len != 18) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
        }

        internal override int GetId { get { return Id; } }

        protected override TBaseRecord DoCopyTo(TSheetInfo SheetInfo)
        {
            TWindow2Record Result = MemberwiseClone() as TWindow2Record;
            Result.Selected = false;
            return Result;
        }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.Window.Window2 = this;
            InChart = ws is TFlxChart;
        }


        private const TSheetOptions AllOptions =
            TSheetOptions.ShowFormulaText |
            TSheetOptions.ShowGridLines |
            TSheetOptions.ShowRowAndColumnHeaders |
            TSheetOptions.ZeroValues |
            TSheetOptions.RightToLeft |
            TSheetOptions.AutomaticGridLineColors |
            TSheetOptions.OutlineSymbols |
            TSheetOptions.PageBreakView;

        private int Flags
        {
            get
            {
                if (InChart) return (int)TSheetOptions.ZeroValues | (FFlags & (1 << 9));
                return FFlags;
            }
            set
            {
                FFlags = value;
            }
        }

        internal TSheetOptions Options
        {
            get
            {
                return (TSheetOptions)Flags & AllOptions;
            }
            set
            {
                TSheetOptions FilteredOptions = value & AllOptions;
                TSheetOptions ExistingOptions = ((TSheetOptions)Flags) & (~AllOptions);
                Flags = (int)(ExistingOptions | FilteredOptions);
            }
        }

        internal int RawOptions
        {
            get
            {
                return Flags;
            }
            set
            {
                Flags = value;
            }
        }

        internal bool IsFrozenButNoSplit
        {
            get
            {
                return (Flags & 0x100) != 0;
            }
            set
            {
                if (value) Flags |= 0x100;
                else Flags &= ~0x100;
            }
        }

        internal bool IsFrozen
        {
            get
            {
                return (Flags & 0x8) != 0;
            }
            set
            {
                if (value) Flags |= 0x8;
                else Flags &= ~0x8;
            }
        }

        internal bool Selected
        {
            get
            {
                return (Flags & (1 << 9)) != 0;
            }
            set
            {
                if (value) FFlags |= (3 << 9); //Selected=true, showing on window=true
                else FFlags &= ~(3 << 9); //Selected=false, showing on window=false
            }
        }

        internal int SheetZoom
        {
            get
            {
                return ScaleInNormalView;
            }
            set
            {
                if (value < 10) ScaleInNormalView = 10;
                else
                    if (value > 400) ScaleInNormalView = 400;
                    else
                        ScaleInNormalView = value;
            }
        }
        internal bool ShowGridLines
        {
            get
            {
                return (Flags & 0x2) != 0;
            }
            set
            {
                if (value) Flags |= 0x2; //GridLines=true
                else Flags &= ~0x2; //GridLines=false
            }
        }

        internal bool ShowFormulaText
        {
            get
            {
                return (Flags & 0x1) != 0;
            }
            set
            {
                if (value) Flags |= 0x1; //ShowFormulaText=true
                else Flags &= ~0x1; //ShowFormulaText=false
            }
        }

        internal bool HideZeroValues
        {
            get
            {
                return (Flags & 0x10) == 0; //this is reversed
            }
            set
            {
                if (!value) Flags |= 0x10; //HideZeroValues=false
                else Flags &= ~0x10; //HideZeroValues=true
            }
        }

        internal TExcelColor GetGridLinesColor(ExcelFile xls)
        {
            TExcelColor Result = FGridLinesColor;
            return Result;
        }

        internal void SetGridLinesColor(TExcelColor aColor)
        {
            FGridLinesColor = aColor;
            Flags &= ~0x20; //Automatic gridline color.
        }

        internal override void SaveToPxl(TPxlStream PxlStream, int Row, TPxlSaveData SaveData)
        {
            base.SaveToPxl(PxlStream, Row, SaveData);
            PxlStream.WriteByte((byte)pxl.WINDOW2);
            PxlStream.Write16((UInt16)FirstRow);  //row and col
            PxlStream.WriteByte((byte)FirstCol);  //row and col

            UInt16 OptionFlags = 0;
            if (IsFrozen) OptionFlags |= 0x08;  //panes frozen
            if (IsFrozenButNoSplit) OptionFlags |= 0x0100;  //frozen no split.
            PxlStream.Write16(OptionFlags);
        }

        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
        {
            Workbook.WriteHeader((UInt16)Id, (UInt16)Len);
            Workbook.Write16((UInt16)FFlags);

            //Biff8Utils.CheckRow(FirstRow); We might have larger number here.
            //Biff8Utils.CheckCol(FirstCol);
            Workbook.Write16((UInt16)FirstRow);
            Workbook.Write16((UInt16)FirstCol);

            Workbook.Write16((UInt16)FGridLinesColor.GetBiff8ColorIndex(SaveData.Palette, TAutomaticColor.DefaultForeground));
            Workbook.Write16(0);

            if (Len == 18)
            {
                Workbook.Write16((UInt16)ScaleInPageBreakPreview);
                Workbook.Write16((UInt16)ScaleInNormalView);
                Workbook.Write16((UInt16)Reserved);
                Workbook.Write16(0);
            }
        }

        internal override int TotalSize()
        {
            return TotalSizeNoHeaders() + XlsConsts.SizeOfTRecordHeader;
        }

        internal override int TotalSizeNoHeaders()
        {
            return Len;
        }



    }


    /// <summary>
    /// Window Zoom Magnification 
    /// </summary>
    internal class TSCLRecord : TxBaseRecord
    {
        internal TSCLRecord(int aId, byte[] aData) : base(aId, aData) { }
        internal TSCLRecord(int aZoom)
            : base((int)xlr.SCL, new byte[4])
        {
            Zoom = aZoom;
        }

        internal int Zoom
        {
            get
            {
                if (GetWord(2) == 0) return 100;
                else
                    return (int)Math.Round((100.0 * GetWord(0)) / (double)GetWord(2));
            }
            set
            {
                int v = 0;
                if (value < 10) v = 10; else if (value > 400) v = 400; else v = value;
                SetWord(0, v);
                SetWord(2, 100);
            }
        }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.Window.Scl = this;
        }
    }

    /// <summary>
    /// Default Width for Columns  
    /// </summary>
    internal class TDefColWidthRecord : TWordRecord
    {
        internal TDefColWidthRecord(int aId, byte[] aData) : base(aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.Columns.DefColWidthChars = Value; 
        }

        internal static void SaveRecord(IDataStream DataStream, int DefColWidth)
        {
            if (DefColWidth > 255) DefColWidth = 255; if (DefColWidth < 0) DefColWidth = 0;
            DataStream.WriteHeader((UInt16)xlr.DEFCOLWIDTH, 2);
            DataStream.Write16((UInt16)DefColWidth);
        }

        internal static void SaveToPxl(TPxlStream PxlStream, int DefColWidth)
        {
            PxlStream.WriteByte((byte)pxl.DEFCOLWIDTH);
            PxlStream.Write16(0); //options not available on biff8.
            if (DefColWidth > 255) DefColWidth = 255; if (DefColWidth < 0) DefColWidth = 0;
            PxlStream.Write16((UInt16)(DefColWidth << 8));
            PxlStream.Write16(0); //XF not available on biff8.
        }
    }

    /// <summary>
    /// Standard column width, in increments of 1/256th of a character width 
    /// The STANDARDWIDTH record records the measurement from the Standard Width dialog box.  
    /// </summary>
    internal class TStandardWidthRecord : TWordRecord
    {
        internal TStandardWidthRecord(int aId, byte[] aData) : base(aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.Columns.DefColWidthChars256 = Value;
        }

        internal static void SaveRecord(IDataStream DataStream, int StdWidth)
        {
            DataStream.WriteHeader((UInt16)xlr.STANDARDWIDTH, 2);
            DataStream.Write16((UInt16)StdWidth);
        }

    }

    /// <summary>
    /// The DEFAULTROWHEIGHT record specifies the height of all undefined rows on the sheet. The miyRw field contains the row height in units of 1/20th of a point. This record does not affect the row height of any rows that are explicitly defined.  
    /// </summary>
    internal class TDefaultRowHeightRecord : TxBaseRecord
    {
        internal TDefaultRowHeightRecord(int aId, byte[] aData) : base(aId, aData) { }

        internal TDefaultRowHeightRecord() : base((int)xlr.DEFAULTROWHEIGHT, new byte[4]) { Height = 0xFF; }

        internal int Height
        {
            get
            {
                return GetWord(2);
            }
            set
            {
                SetWord(2, value);
            }
        }

        internal int Flags
        {
            get
            {
                return GetWord(0);
            }
            set
            {
                SetWord(0, value);
            }
        }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.SheetGlobals.DefRowHeight = this;
        }

        internal override void SaveToPxl(TPxlStream PxlStream, int Row, TPxlSaveData SaveData)
        {
            base.SaveToPxl(PxlStream, Row, SaveData);
            PxlStream.WriteByte((byte)pxl.DEFAULTROWHEIGHT);
            if ((Data[0] & 0x2) != 0) PxlStream.Write16(1); else PxlStream.Write16(0); //Row hidden
            PxlStream.Write(Data, 2, 2); //Row Height
        }

    }

    /// <summary>
    /// Print Header or Footer on Each Page
    /// </summary>
    internal class TPageHeaderFooterRecord : TNotStorableRecord
    {
        protected string Text;

        internal TPageHeaderFooterRecord(int aId, byte[] aData)
        {
            if (aData == null || aData.Length == 0) Text = String.Empty;
            else
            {
                long TextSize = 0;
                StrOps.GetSimpleString(true, aData, 0, false, 0, ref Text, ref TextSize);
            }
        }

        internal static int StandardSize(string Text)
        {
            if (Text == null || Text.Length == 0) return XlsConsts.SizeOfTRecordHeader;
            if (Text.Length > 255) XlsMessages.ThrowException(XlsErr.ErrHeaderFooterStringTooLong, Text, 255);
            TExcelString Xs = new TExcelString(TStrLenLength.is16bits, Text, null, false);
            return XlsConsts.SizeOfTRecordHeader + Xs.TotalSize();
        }

        internal static void SaveRecord(IDataStream DataStream, xlr Id, string Text)
        {
            int Len = StandardSize(Text);
            if (Len > 0) Len -= XlsConsts.SizeOfTRecordHeader;
            DataStream.WriteHeader((UInt16)Id, (UInt16)Len);
            if (Text == null || Text.Length == 0) return;
            if (Text.Length > 255) XlsMessages.ThrowException(XlsErr.ErrHeaderFooterStringTooLong, Text, 255);
            TExcelString Xs = new TExcelString(TStrLenLength.is16bits, Text, null, false);
            byte[] Data = new byte[Xs.TotalSize()];
            Xs.CopyToPtr(Data, 0);
            DataStream.Write(Data, Data.Length);
        }


        #region AsHtml
        private static void AddAttributes(StringBuilder Text, TPageHeaderFooterState State, ref string LastAtt)
        {
            StringBuilder css = new StringBuilder();
            string Decoration = String.Empty;

            if (State.U || State.dU) Decoration += "underline;";
            if (State.Striked) Decoration += " line-through";

            if (Decoration.Length > 0) css.Append("text-decoration:" + Decoration + ";");

            if (State.Bold) css.Append("font-weight:bold;");
            if (State.It) css.Append("font-style: italic;");
            if (State.SubScript) css.Append("vertical-align:sub;");
            if (State.SuperScript) css.Append("vertical-align:super;");

            if (State.FontSize != TPageHeaderFooterState.DefaultFontSize)
            {
                css.Append("font-size:"); css.Append(State.FontSize.ToString()); css.Append("pt;");
            }

            if (State.FontName.Length > 0)
            {
                css.Append("font-family:"); css.Append(State.FontName); css.Append(";");
            }

            string NewAtt = css.ToString();
            if (NewAtt != LastAtt)
            {
                if (LastAtt.Length > 0) Text.Append("</span>");
                if (NewAtt.Length > 0)
                {
                    Text.Append("<span style = '");
                    Text.Append(NewAtt);
                    Text.Append("'>");
                }
                LastAtt = NewAtt;
            }

        }

        private static void AddText(StringBuilder Res, string s, THtmlVersion htmlVersion, Encoding encoding, TPageHeaderFooterState State, ref string LastAtt)
        {
            if (s.Length != 0) AddAttributes(Res, State, ref LastAtt);
            Res.Append(THtmlEntities.EncodeAsHtml(s, htmlVersion, encoding));
        }

        private static string FinishText(StringBuilder Res, string LastAtt)
        {
            if (LastAtt.Length > 0) Res.Append("</span>");
            return Res.ToString();
        }

        internal static string AsHtml(ExcelFile Workbook, string Text, string ImageTag, int CurrentPage, int TotalPages, THtmlVersion htmlVersion, Encoding encoding, IHtmlFontEvent onFont)
        {
            const string ItalicStr = "Italic";  //Do not localize
            const string BoldStr = "Bold";  //Do not localize

            StringBuilder Result = new StringBuilder();

            string aText = Text + "&";
            TPageHeaderFooterState State = TPageHeaderFooterState.Init();
            string LastAtt = String.Empty;
            int p = 0;
            int o = 0;

            do
            {
                int q = aText.IndexOfAny(new char[] { '&', '\n' }, p + o);
                if (q < 0) return FinishText(Result, LastAtt); //might happen on an unterminated string, f.i.
                o = 0;

                AddText(Result, aText.Substring(p, q - p), htmlVersion, encoding, State, ref LastAtt);

                p = q + 1;

                if (aText[p - 1] == '\n') //new line.
                {
                    Result.Append("<br");
                    Result.Append(THtmlEntities.EndOfTag(htmlVersion));
                    continue;
                }


                if (p < aText.Length)
                {
                    switch (aText[p])
                    {
                        case 'U':
                            State.dU = false;
                            State.U = !State.U;
                            p++;
                            break;

                        case 'E':
                            State.U = false;
                            State.dU = !State.dU;
                            p++;
                            break;

                        case 'S':
                            State.Striked = !State.Striked;
                            p++;
                            break;

                        case 'B':
                        case 'b':
                            State.Bold = !State.Bold;
                            p++;
                            break;

                        case 'I':
                        case 'i':
                            State.It = !State.It;
                            p++;
                            break;

                        case 'Y':
                            State.SubScript = !State.SubScript;
                            State.SuperScript = false;
                            p++;
                            break;

                        case 'X':
                            State.SuperScript = !State.SuperScript;
                            State.SubScript = false;
                            p++;
                            break;

                        case 'G':
                            if (ImageTag != null)
                            {
                                Result.Append(ImageTag);
                            }
                            p++;
                            break;

                        case '"':
                            p++;
                            q = aText.IndexOf('"', p);
                            if (q >= 0)
                            {
                                string FontName = String.Empty;
                                string FontSt = String.Empty;
                                int r = aText.IndexOf(',', p, q - p);
                                if (r >= 0)
                                {
                                    FontName = aText.Substring(p, r - p);
                                    FontSt = aText.Substring(r + 1, q - r - 1);
                                }
                                else
                                    FontName = aText.Substring(p, q - p);

                                HtmlFontEventArgs e = new HtmlFontEventArgs(Workbook, new TFlxFont(), FontName);
                                if (e.FontFamily != null && e.FontFamily.IndexOf(" ") >= 0) e.FontFamily = "\"" + e.FontFamily + "\"";  //font names with spaces must be quoted, and they must be double quotes so they do not clash with the style tag.
                                if (onFont != null) onFont.DoHtmlFont(e);

                                State.FontName = e.FontFamily;

                                if (FontSt.IndexOf(BoldStr) >= 0)
                                {
                                    State.Bold = true;
                                }
                                else
                                {
                                    State.Bold = false;
                                }

                                if (FontSt.IndexOf(ItalicStr) >= 0)
                                {
                                    State.It = true;
                                }
                                else
                                {
                                    State.It = false;
                                }

                                p = q + 1;
                            }
                            break;

                        case '0':
                        case '1':
                        case '2':
                        case '3':
                        case '4':
                        case '5':
                        case '6':
                        case '7':
                        case '8':
                        case '9':
                            {
                                int FSize = 0;
                                do
                                {
                                    FSize = FSize * 10 + (int)aText[p] - (int)'0';
                                    p++;
                                }
                                while (p < aText.Length && aText[p] >= '0' && aText[p] <= '9');

                                State.FontSize = FSize;
                                break;
                            }

                        case '&': //double & means a simple one.
                            o = 1; //to skip next & from search
                            break;

                        case 'A': //SheetName
                            p++; //to skip next from search
                            AddText(Result, Workbook.SheetName, htmlVersion, encoding, State, ref LastAtt);
                            break;

                        case 'D': //Date
                            p++; //to skip next from search
                                AddText(Result, DateTime.Now.Date.ToShortDateString(), htmlVersion, encoding, State, ref LastAtt);
                            break;

                        case 'T': //Time
                            p++; //to skip next from search
                                AddText(Result, DateTime.Now.ToShortTimeString(), htmlVersion, encoding, State, ref LastAtt);
                            break;

                        case 'P': //Page Number
                            p++; //to skip next from search
                            AddText(Result, CurrentPage.ToString(), htmlVersion, encoding, State, ref LastAtt);
                            break;

                        case 'N': //PageCount
                            p++; //to skip next from search
                            AddText(Result, TotalPages.ToString(), htmlVersion, encoding, State, ref LastAtt);
                            break;

                        case 'F': //FileName
                            p++; //to skip next from search
                            AddText(Result, Path.GetFileName(Workbook.ActiveFileName), htmlVersion, encoding, State, ref LastAtt);
                            break;

                        case 'Z': //FullFileName
                            p++; //to skip next from search
                            AddText(Result, Path.GetFullPath(Workbook.ActiveFileName), htmlVersion, encoding, State, ref LastAtt);
                            break;

                        default: //unknown code
                            p++;
                            break;
                    }
                }
            }
            while (p < aText.Length);

            return FinishText(Result, LastAtt);
        }
        #endregion
    }

    internal struct TPageHeaderFooterState
    {
        internal const int DefaultFontSize = 10;
        internal bool U;
        internal bool dU;
        internal bool SubScript;
        internal bool SuperScript;
        internal bool It;
        internal bool Bold;
        internal bool Striked;
        internal int FontSize;

        internal string FontName;

        internal static TPageHeaderFooterState Init()
        {
            TPageHeaderFooterState Result;
            Result.U = false;
            Result.dU = false;
            Result.SubScript = false;
            Result.SuperScript = false;
            Result.It = false;
            Result.Bold = false;
            Result.Striked = false;
            Result.FontSize = DefaultFontSize;

            Result.FontName = string.Empty;

            return Result;
        }
    }

    /// <summary>
    /// Page Header
    /// </summary>
    internal class TPageHeaderRecord : TPageHeaderFooterRecord
    {
        internal TPageHeaderRecord(int aId, byte[] aData) : base(aId, aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            if (Loader.CustomView != null) Loader.CustomView.Setup.HeaderAndFooter.SetAllHeaders(Text);
            else ws.PageSetup.HeaderAndFooter.SetAllHeaders(Text);
        }

    }

    /// <summary>
    /// Page Footer
    /// </summary>
    internal class TPageFooterRecord : TPageHeaderFooterRecord
    {
        internal TPageFooterRecord(int aId, byte[] aData) : base(aId, aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            if (Loader.CustomView != null) Loader.CustomView.Setup.HeaderAndFooter.SetAllFooters(Text);
            else ws.PageSetup.HeaderAndFooter.SetAllFooters(Text);
        }
    }

    internal class THeaderFooterExtRecord : TNotStorableRecord
    {
        byte[] Data;

        internal THeaderFooterExtRecord(int aId, byte[] aData)
        {
            Data = aData;
        }

        private void Parse(TSheet ws, out THeaderAndFooter HeaderAndFooter, out TPageSetup Target)
        {
            HeaderAndFooter = new THeaderAndFooter();
            Target = null;

            byte[] bCustomView = new byte[16];
            if (Data.Length < 12 + 16)//might happen in files saved with Excel 2007 beta.
            {
                return;
            }
            
            Array.Copy(Data, 12, bCustomView, 0, 16);
            Guid CustomView = new Guid(bCustomView);

            //Now, we need to be sure this is a valid guid, and not a "beta" header that didn't had  guid.
            Target = ws.GetCustomViewSetup(CustomView);
            if (Target == null) return;
            

            byte oflags = Data[16 + 12];
            HeaderAndFooter.DiffEvenPages = (oflags & 0x01) != 0;
            HeaderAndFooter.DiffFirstPage = (oflags & 0x02) != 0;
            HeaderAndFooter.ScaleWithDoc = (oflags & 0x04) != 0;
            HeaderAndFooter.AlignMargins = (oflags & 0x08) != 0;

            long EvenHeaderSize = 0;
            long EvenFooterSize = 0;
            string hf = null;

            int HLen = BitOps.GetWord(Data, 30);
            if (HLen > 0)
            {
                StrOps.GetSimpleString(true, Data, 38, false, HLen, ref hf, ref EvenHeaderSize);
                HeaderAndFooter.EvenHeader = hf;
            }

            HLen = BitOps.GetWord(Data, 32);
            if (HLen > 0)
            {
                StrOps.GetSimpleString(true, Data, 38 + (int)EvenHeaderSize, false, HLen, ref hf, ref EvenFooterSize);
                HeaderAndFooter.EvenFooter = hf;
            }

            long FirstHeaderSize = 0;
            HLen = BitOps.GetWord(Data, 34);
            if (HLen > 0)
            {
                StrOps.GetSimpleString(true, Data, 38 + (int)(EvenHeaderSize + EvenFooterSize), false, HLen, ref hf, ref FirstHeaderSize);
                HeaderAndFooter.FirstHeader = hf;
            }

            long FirstFooterSize = 0;
            HLen = BitOps.GetWord(Data, 36);
            if (HLen > 0)
            {
                StrOps.GetSimpleString(true, Data, 38 + (int)(EvenHeaderSize + EvenFooterSize + FirstHeaderSize), false, HLen, ref hf, ref FirstFooterSize);
                HeaderAndFooter.FirstFooter = hf;
            }
        }
    
        static int SizeOneString(string Text)
        {
            if (Text == null || Text.Length == 0) return 0;
            if (Text.Length > 255) XlsMessages.ThrowException(XlsErr.ErrHeaderFooterStringTooLong, Text, 255);
            TExcelString Xs = new TExcelString(TStrLenLength.is16bits, Text, null, false);
            return Xs.TotalSize();
        }

        internal static bool NeedsToSave(THeaderAndFooter HeaderAndFooter)
        {
            return HeaderAndFooter.DiffEvenPages || HeaderAndFooter.DiffFirstPage
                    || HeaderAndFooter.AlignMargins || !HeaderAndFooter.ScaleWithDoc;
        }

        internal static int StandardSize(THeaderAndFooter HeaderAndFooter)
        {
            if (!NeedsToSave(HeaderAndFooter)) return 0;

            return XlsConsts.SizeOfTRecordHeader + 38 +
                SizeOneString(HeaderAndFooter.EvenHeader) +
                SizeOneString(HeaderAndFooter.EvenFooter) +
                SizeOneString(HeaderAndFooter.FirstHeader) +
                SizeOneString(HeaderAndFooter.FirstFooter);
        }

        static byte[] GetOneString(string Text, out int CharCount)
        {
            CharCount = 0;
            if (Text == null || Text.Length == 0) return new byte[0];
            if (Text.Length > 255) XlsMessages.ThrowException(XlsErr.ErrHeaderFooterStringTooLong, Text, 255);
            TExcelString Xs = new TExcelString(TStrLenLength.is16bits, Text, null, false);
            CharCount = Xs.Data.Length;
            byte[] Data = new byte[Xs.TotalSize()];
            Xs.CopyToPtr(Data, 0);
            return Data;
        }

        internal static void SaveRecord(IDataStream DataStream, Guid CustomView, THeaderAndFooter HeaderAndFooter)
        {
            if (!NeedsToSave(HeaderAndFooter)) return;

            int Len = StandardSize(HeaderAndFooter);
            if (Len > 0) Len -= XlsConsts.SizeOfTRecordHeader;
            DataStream.WriteHeader((UInt16)xlr.HEADERFOOTER, (UInt16)Len);
            DataStream.Write16((UInt16)xlr.HEADERFOOTER);
            DataStream.Write(new byte[10], 10);

            DataStream.Write(CustomView.ToByteArray(), 16);

            int oFlags = BitOps.GetBool(HeaderAndFooter.DiffEvenPages, HeaderAndFooter.DiffFirstPage, HeaderAndFooter.ScaleWithDoc, HeaderAndFooter.AlignMargins);
            DataStream.Write16((UInt16) oFlags);

            int ceh, cef, cfh, cff;
            byte[] eh = GetOneString(HeaderAndFooter.EvenHeader, out ceh);
            byte[] ef = GetOneString(HeaderAndFooter.EvenFooter, out cef);
            byte[] fh = GetOneString(HeaderAndFooter.FirstHeader, out cfh);
            byte[] ff = GetOneString(HeaderAndFooter.FirstFooter, out cff);

            DataStream.Write16((UInt16)ceh);
            DataStream.Write16((UInt16)cef);
            DataStream.Write16((UInt16)cfh);
            DataStream.Write16((UInt16)cff);

            DataStream.Write(eh, eh.Length);
            DataStream.Write(ef, ef.Length);
            DataStream.Write(fh, fh.Length);
            DataStream.Write(ff, ff.Length);
        }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            THeaderAndFooter HeaderAndFooter;
            TPageSetup Target;
            Parse(ws, out HeaderAndFooter, out Target);

            if (Target == null) return; //page doesn't have setup...

            //We need to load everything except the default header and default footer.
            Target.HeaderAndFooter.EvenHeader = HeaderAndFooter.EvenHeader;
            Target.HeaderAndFooter.EvenFooter = HeaderAndFooter.EvenFooter;
            Target.HeaderAndFooter.FirstHeader = HeaderAndFooter.FirstHeader;
            Target.HeaderAndFooter.FirstFooter = HeaderAndFooter.FirstFooter;

            Target.HeaderAndFooter.DiffEvenPages = HeaderAndFooter.DiffEvenPages;
            Target.HeaderAndFooter.DiffFirstPage = HeaderAndFooter.DiffFirstPage;

            Target.HeaderAndFooter.AlignMargins = HeaderAndFooter.AlignMargins;
            Target.HeaderAndFooter.ScaleWithDoc = HeaderAndFooter.ScaleWithDoc;
        }

    }

    /// <summary>
    /// Print Grid Lines
    /// </summary>
    internal class TPrintGridLinesRecord : TBoolRecord
    {
        internal TPrintGridLinesRecord(int aId, byte[] aData) : base(aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.SheetGlobals.PrintGridLines = Value;
        }

        internal static void SaveRecord(IDataStream DataStream, bool PrintGridLines)
        {
            DataStream.WriteHeader((UInt16)xlr.PRINTGRIDLINES, 2);
            if (PrintGridLines) DataStream.Write16(1); else DataStream.Write16(0);
        }
    }

    internal abstract class TMarginRecord: TDoubleRecord
    {
        internal TMarginRecord(byte[] aData) : base(aData) { }

        protected static void SaveMargins(xlr Id, IDataStream DataStream, double Margin)
        {
            if (Margin < 0) return; //won't be saved.
            if (Margin < 0) Margin = 0;
            if (Margin > 49) Margin = 49;
            DataStream.WriteHeader((UInt16)Id, 8);
            DataStream.Write(BitConverter.GetBytes(Margin), 8);
        }

        internal static int StandardSize(double Margin)
        {
            if (Margin < 0) return 0;
            return XlsConsts.SizeOfTRecordHeader + 8;
        }

    }

    internal class TLeftMarginRecord : TMarginRecord
    {
        internal TLeftMarginRecord(int aId, byte[] aData) : base(aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            if (Loader.CustomView != null) Loader.CustomView.Setup.LeftMargin = Value;
            else ws.PageSetup.LeftMargin = Value;
        }

        internal static void SaveRecord(IDataStream DataStream, double Margin)
        {
            SaveMargins(xlr.LEFTMARGIN, DataStream, Margin);
        }

    }

    internal class TRightMarginRecord : TMarginRecord
    {
        internal TRightMarginRecord(int aId, byte[] aData) : base(aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            if (Loader.CustomView != null) Loader.CustomView.Setup.RightMargin = Value;
            else ws.PageSetup.RightMargin = Value;
        }

        internal static void SaveRecord(IDataStream DataStream, double Margin)
        {
            SaveMargins(xlr.RIGHTMARGIN, DataStream, Margin);
        }

    }

    internal class TTopMarginRecord : TMarginRecord
    {
        internal TTopMarginRecord(int aId, byte[] aData) : base(aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            if (Loader.CustomView != null) Loader.CustomView.Setup.TopMargin = Value;
            else ws.PageSetup.TopMargin = Value;
        }

        internal static void SaveRecord(IDataStream DataStream, double Margin)
        {
            SaveMargins(xlr.TOPMARGIN, DataStream, Margin);
        }

    }

    internal class TBottomMarginRecord : TMarginRecord
    {
        internal TBottomMarginRecord(int aId, byte[] aData) : base(aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            if (Loader.CustomView != null) Loader.CustomView.Setup.BottomMargin = Value;
            else ws.PageSetup.BottomMargin = Value;
        }

        internal static void SaveRecord(IDataStream DataStream, double Margin)
        {
            SaveMargins(xlr.BOTTOMMARGIN, DataStream, Margin);
        }

    }

    /// <summary>
    /// Printer Setup
    /// </summary>
    internal class TSetupRecord : TxBaseRecord
    {
        internal TSetupRecord(): 
            base ((int)xlr.SETUP, new byte[]
            {0x00, 0x00, 0xFF, 0x00, 0x01, 0x00, 0x01, 0x00, 0x01, 0x00, 0x04, 0x01, 0x00, 0x01, 0x01, 
             0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xE0, 0x3F, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xE0, 0x3F, 0x01, 0x00})
        {
        }

        internal TSetupRecord(int aId, byte[] aData) : base(aId, aData) { }
        internal int PaperSize
        {
            get
            {
                if (PrintOptionsNotInit) return 0;
                else
                    return GetWord(0);
            }
            set
            {
                InitPrintOptions();
                SetWord(0, value);
            }
        }
        internal int Scale
        {
            get
            {
                if (PrintOptionsNotInit) return 100;
                else
                    return GetWord(2);
            }
            set
            {
                if ((value < 0) || (value > 0xFFFF))
                    XlsMessages.ThrowException(XlsErr.ErrXlsIndexOutBounds, value, "PrintScale", 0, 0xFFFF); 
                InitPrintOptions();
                SetWord(2, value);
            }
        }

        internal int? PageStart
        {
            get
            {
                if (!PageStartInitialized) return null;
                unchecked
                {
                    return (Int16)GetWord(4);
                }
            }
            set
            {
                if (value == null)
                {
                    PageStartInitialized = false;
                    return;
                }

                if ((value < Int16.MinValue) || (value > Int16.MaxValue))
                    XlsMessages.ThrowException(XlsErr.ErrXlsIndexOutBounds, value, "PageStart", Int16.MinValue, Int16.MaxValue);
                PageStartInitialized = true;
                SetWord(4, value.Value);
            }
        }

        private bool PageStartInitialized
        {
            get
            {
                return ((GetWord(10) & 0x80) != 0);
            }
            set
            {
                if (value) SetWord(10, GetWord(10) | 0x80);
                else SetWord(10, GetWord(10) & ~0x80); 
            }
        }

        internal int FitWidth
        {
            get { return GetWord(6); }
            set
            {
                if ((value < 0) || (value > 0xFFFF))
                    XlsMessages.ThrowException(XlsErr.ErrXlsIndexOutBounds, value, "PrintNumberHorizontalPages", 0, 0xFFFF);
                SetWord(6, value);
            }
        }
        internal int FitHeight
        {
            get { return GetWord(8); }
            set
            {
                if ((value < 0) || (value > 0xFFFF))
                    XlsMessages.ThrowException(XlsErr.ErrXlsIndexOutBounds, value, "PrintNumberVerticalPages", 0, 0xFFFF);
                SetWord(8, value);
            }
        }

        private bool PrintOptionsNotInit
        {
            get
            {
                return ((GetWord(10) & 0x4) == 0x4);
            }
        }

        private void InitPrintOptions()
        {
            if (PrintOptionsNotInit) SetWord(2, 100); //if not the zoom will be at 255.
            SetWord(10, GetWord(10) & ~0x4);

        }

        internal int GetPrintOptions(bool DefaultPortrait)
        {
            int op = GetWord(10);
            if (DefaultPortrait)
            {
                if ((op & 0x4) == 0x4) op |= 2; //default printing mode is portrait.
                //The default really depends on the default printer, when there is nothing
                //specified on the Excel file. we decided for portrait because it seems more useful,
                //but it could be anything and we can't know without asking for the default printer.
            }
            else
            {
                if ((op & 0x4) == 0x4) op &= ~2; //default printing mode is landscape
            }
            return op;
        }

        internal void SetPrintOptions(int value)
        {
            SetWord(10, value & 0xFFFF);
        }

        internal int HPrintRes
        {
            get
            {
                if (PrintOptionsNotInit) return 100;
                else
                    return GetWord(12);
            }
            set
            {
                if ((value < 0) || (value > 0xFFFF))
                    XlsMessages.ThrowException(XlsErr.ErrXlsIndexOutBounds, value, "PrintXResolution", 0, 0xFFFF);

                InitPrintOptions();
                SetWord(12, value);
            }
        }

        internal int VPrintRes
        {
            get
            {
                if (PrintOptionsNotInit) return 100;
                else
                    return GetWord(14);
            }
            set
            {
                if ((value < 0) || (value > 0xFFFF))
                    XlsMessages.ThrowException(XlsErr.ErrXlsIndexOutBounds, value, "PrintYResolution", 0, 0xFFFF);
                InitPrintOptions();
                SetWord(14, value);
            }
        }

        internal double HeaderMargin { get { return BitConverter.ToDouble(Data, 16); } set { BitConverter.GetBytes(value).CopyTo(Data, 16); } }
        internal double FooterMargin { get { return BitConverter.ToDouble(Data, 24); } set { BitConverter.GetBytes(value).CopyTo(Data, 24); } }
        internal int Copies
        {
            get
            {
                if (PrintOptionsNotInit) return 1;
                else
                    return GetWord(32);
            }
            set
            {
                if ((value < 0) || (value > 0xFFFF))
                    XlsMessages.ThrowException(XlsErr.ErrXlsIndexOutBounds, value, "PrintCopies", 0, 0xFFFF);

                InitPrintOptions();
                SetWord(32, value);
            }
        }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            if (Loader.CustomView != null) Loader.CustomView.Setup.Setup = this;
            else ws.PageSetup.Setup = this;
        }
    }

    /// <summary>
    /// Printer driver information.
    /// </summary>
    internal class TPlsRecord : TxBaseRecord
    {
        internal TPlsRecord(int aId, byte[] aData) : base(aId, aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            if (Loader.CustomView != null) Loader.CustomView.Setup.Pls = this;
            else ws.PageSetup.Pls = this;
        }

        /// <summary>
        /// Will create a Pls record including continues if data is too long.
        /// </summary>
        /// <param name="Pls"></param>
        /// <returns></returns>
        internal static TPlsRecord FromLongData(byte[] Pls)
        {
            TPlsRecord Result = null;
            using (MemoryStream ms = new MemoryStream(Pls))
            {
                byte[] Block = GetNextBlock(ms);
                Result = new TPlsRecord((int)xlr.PLS, Block);
                TxBaseRecord cont = Result;
                while (ms.Position < ms.Length)
                {
                    Block = GetNextBlock(ms);
                    cont.AddContinue(new TContinueRecord((int)xlr.CONTINUE, Block));
                    cont = cont.Continue;
                }

            }
            return Result;
        }

        private static byte[] GetNextBlock(MemoryStream ms)
        {
            byte[] Block = new byte[Math.Min(XlsConsts.MaxRecordDataSize + 1, ms.Length - ms.Position)]; //this needs to be exactly 8224 or excel will complain... weird.
            ms.Read(Block, 0, Block.Length);
            return Block;
        }
    }

    /// <summary>
    /// Set this record to true to print row and column Headings.
    /// </summary>
    internal class TPrintHeadersRecord : TBoolRecord
    {
        internal TPrintHeadersRecord(int aId, byte[] aData) : base(aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.SheetGlobals.PrintHeaders = Value;
        }

        internal static void SaveRecord(IDataStream DataStream, bool PrintHeaders)
        {
            DataStream.WriteHeader((UInt16)xlr.PRINTHEADERS, 2);
            if (PrintHeaders) DataStream.Write16(1); else DataStream.Write16(0);
        }
    }

    /// <summary>
    /// Set this record to true to center the sheet when printing.
    /// </summary>
    internal class THCenterRecord : TBoolRecord
    {
        internal THCenterRecord(int aId, byte[] aData) : base(aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            if (Loader.CustomView != null) Loader.CustomView.Setup.HCenter = Value;
 	        else ws.PageSetup.HCenter = Value;
        }

        internal static void SaveRecord(IDataStream DataStream, bool Centered)
        {
            DataStream.WriteHeader((UInt16)xlr.HCENTER, 2);
            if (Centered) DataStream.Write16(1); else DataStream.Write16(0);
        }
    
     }

    internal class TVCenterRecord : TBoolRecord
    {
        internal TVCenterRecord(int aId, byte[] aData) : base(aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            if (Loader.CustomView != null) Loader.CustomView.Setup.VCenter = Value;
            else ws.PageSetup.VCenter = Value;
        }

        internal static void SaveRecord(IDataStream DataStream, bool Centered)
        {
            DataStream.WriteHeader((UInt16)xlr.VCENTER, 2);
            if (Centered) DataStream.Write16(1); else DataStream.Write16(0);
        }
    }

    internal struct TWsBool
    {
        internal bool ShowAutoBreaks;

        internal bool Dialog;
        internal bool ApplyStyles;
        internal bool RowSumsBelow;
        internal bool ColSumsRight;

        internal bool FitToPage;
        internal bool DspGuts;
        internal bool SyncHoriz;
        internal bool SyncVert;
        internal bool AltExprEval;
        internal bool AltFormulaEntry;

        internal void Init()
        {
            ShowAutoBreaks = true;
            RowSumsBelow = true;
            ColSumsRight = true;
            DspGuts = true;
        }
    }

    /// <summary>
    /// Additional Workspace Information 
    /// </summary>
    internal class TWsBoolRecord : TNotStorableRecord
    {
        TWsBool WsBool;

        internal TWsBoolRecord(int aId, byte[] Data)
            : base()
        {
            WsBool.ShowAutoBreaks = (Data[0] & 0x1) != 0;

            WsBool.Dialog = (Data[0] & 0x10) != 0;
            WsBool.ApplyStyles = (Data[0] & 0x20) != 0;
            WsBool.RowSumsBelow = (Data[0] & 0x40) != 0;
            WsBool.ColSumsRight = (Data[0] & 0x80) != 0;

            WsBool.FitToPage = (Data[1] & 0x01) != 0;
            WsBool.DspGuts = (Data[1] & 0x00C) != 0;
            WsBool.SyncHoriz = (Data[1] & 0x10) != 0;
            WsBool.SyncVert = (Data[1] & 0x20) != 0;
            WsBool.AltExprEval = (Data[1] & 0x40) != 0;
            WsBool.AltFormulaEntry = (Data[1] & 0x80) != 0;
        }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.SheetGlobals.WsBool = WsBool;
            if (WsBool.Dialog)
            {
                ws.Columns.AllowStandardWidth = false;
                ws.Columns.IsDialog = false;
            }
        }

        internal static long StandardSize()
        {
            return XlsConsts.SizeOfTRecordHeader + 2;
        }

        internal static void SaveWsBool(IDataStream DataStream, TWsBool WsBool)
        {
            DataStream.WriteHeader((UInt16)xlr.WSBOOL, 2);
            DataStream.Write16((UInt16)BitOps.GetBool
                (
                    WsBool.ShowAutoBreaks,
                    false, false, false,
                    WsBool.Dialog,
                    WsBool.ApplyStyles,
                    WsBool.RowSumsBelow,
                    WsBool.ColSumsRight,

                    WsBool.FitToPage,
                    false,
                    WsBool.DspGuts,
                    false,
                    WsBool.SyncHoriz,
                    WsBool.SyncVert,
                    WsBool.AltExprEval,
                    WsBool.AltFormulaEntry
                ));
        }

    }

    /// <summary>
    /// AutoFilter Information 
    /// </summary>
    internal class TAutoFilterInfoRecord : TWordRecord
    {
        internal TAutoFilterInfoRecord(int aId, byte[] aData) : base(aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            if (Loader.CustomView != null) Loader.CustomView.AutoFilter.AutoFilterInfo = Value;
            else ws.SortAndFilter.AutoFilter.AutoFilterInfo = Value;
        }

        internal static void SaveRecord(IDataStream DataStream, int DropDownCount)
        {
            DataStream.WriteHeader((UInt16)xlr.AutoFilterINFO, 2);
            DataStream.Write16((UInt16)DropDownCount);
        }
    }

    internal class TAutoFilterRecord : TxBaseRecord
    {
        internal TAutoFilterRecord(int aId, byte[] aData)
            : base(aId, aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            if (Loader.CustomView != null) Loader.CustomView.AutoFilter.Filters.Add(this);
            else ws.SortAndFilter.AutoFilter.Filters.Add(this);
        }
    }

    /// <summary>
    /// Country for the Excel version and Windows version. 
    /// </summary>
    internal class TCountryRecord : TxBaseRecord
    {
        internal TCountryRecord(int aId, byte[] aData)
            : base(aId, aData)
        {
        }

        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
        {
            base.SaveToStream(Workbook, SaveData, Row);
        }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.Country = this;
        }
    }

    /// <summary>
    /// Coodepage Record. 
    /// </summary>
    internal class TCodePageRecord : TxBaseRecord
    {
        internal TCodePageRecord(int aId, byte[] aData)
            : base(aId, aData)
        {
        }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.CodePage = this;
        }
    }

    internal abstract class TNotStorableRecord : TBaseRecord
    {
        protected override TBaseRecord DoCopyTo(TSheetInfo SheetInfo)
        {
            FlxMessages.ThrowException(FlxErr.ErrInternal);
            return null;
        }

        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
        {
            FlxMessages.ThrowException(FlxErr.ErrInternal); //we won't store this record, so we don't need to save it.
        }

        internal override int TotalSize()
        {
            FlxMessages.ThrowException(FlxErr.ErrInternal);
            return 0;
        }

        internal override int TotalSizeNoHeaders()
        {
            FlxMessages.ThrowException(FlxErr.ErrInternal);
            return 0;
        }

        internal override int GetId
        {
            get { FlxMessages.ThrowException(FlxErr.ErrInternal); return 0; }
        }
    }

    internal abstract class TBoolRecord : TNotStorableRecord
    {
        protected bool Value;

        internal TBoolRecord(byte[] aData)
        {
            if (aData != null && aData.Length >= 2)
            {
                Value = BitOps.GetWord(aData, 0) == 1;
            }
        }

        internal static int StandardSize()
        {
            return XlsConsts.SizeOfTRecordHeader + 2;
        }
    }

    internal abstract class TDoubleRecord : TNotStorableRecord
    {
        protected double Value;

        internal TDoubleRecord(byte[] aData)
        {
            if (aData != null && aData.Length >= 8)
            {
                Value = BitConverter.ToDouble(aData, 0);
            }
        }
    }

    internal abstract class TWordRecord : TNotStorableRecord
    {
        protected int Value;

        internal TWordRecord(byte[] aData)
        {
            if (aData != null && aData.Length >= 2)
            {
                Value = BitOps.GetWord(aData, 0);
            }
        }

        internal static int StandardSize()
        {
            return XlsConsts.SizeOfTRecordHeader + 2;
        }
    }

    /// <summary>
    /// 1904 dates Record. 
    /// </summary>
    internal class T1904Record : TBoolRecord
    {
        internal T1904Record(int aId, byte[] aData) : base(aData) { }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.Dates1904 = Value;
        }

        internal static void SaveRecord(IDataStream DataStream, bool Dates1904)
        {
            DataStream.WriteHeader((UInt16)xlr.x1904, 2);
            if (Dates1904) DataStream.Write16(1); else DataStream.Write16(0);
        }
    }

    /// <summary>
    /// Backup. 
    /// </summary>
    internal class TBackupRecord : TBoolRecord
    {
        internal TBackupRecord(int aId, byte[] aData) : base(aData) { }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.Backup = Value;
        }

        internal static void SaveRecord(IDataStream DataStream, bool IsBackup)
        {
            DataStream.WriteHeader((UInt16)xlr.BACKUP, 2);
            if (IsBackup) DataStream.Write16(1); else DataStream.Write16(0);
        }
    }

    /// <summary>
    /// Precision Record. 
    /// </summary>
    internal class TPrecisionRecord : TBoolRecord
    {
        internal TPrecisionRecord(int aId, byte[] aData) : base(aData) { }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.PrecisionAsDisplayed = !Value;
        }

        internal static void SaveRecord(IDataStream DataStream, bool aPrecision)
        {
            DataStream.WriteHeader((UInt16)xlr.PRECISION, 2);
            if (aPrecision) DataStream.Write16(0); else DataStream.Write16(1);
        }
    }

    /// <summary>
    /// Refresh All. 
    /// </summary>
    internal class TRefreshAllRecord : TBoolRecord
    {
        internal TRefreshAllRecord(int aId, byte[] aData) : base(aData) { }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.RefreshAll = Value;
        }

        internal static void SaveRecord(IDataStream DataStream, bool aRefreshAll)
        {
            DataStream.WriteHeader((UInt16)xlr.REFRESHALL, 2);
            if (aRefreshAll) DataStream.Write16(1); else DataStream.Write16(0);
        }
    }

    /// <summary>
    /// Refresh All. 
    /// </summary>
    internal class TUsesELFsRecord : TBoolRecord
    {
        internal TUsesELFsRecord(int aId, byte[] aData) : base(aData) { }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.UsesELFs = Value;
        }

        internal static void SaveRecord(IDataStream DataStream, bool aUsesElfs)
        {
            DataStream.WriteHeader((UInt16)xlr.USESELFS, 2);
            if (aUsesElfs) DataStream.Write16(1); else DataStream.Write16(0);
        }

    }

    [Flags]
    internal enum TBookBoolOption
    {
        None = 0x00,
        NoSaveExtValues = 0x01,
        HasEnvelope = 0x04,
        EnvelopeVisible = 0x08,
        EnvelopeInitDone = 0x10,
        HideBorderUnselLists = 0x100
    }

    internal enum TUpdateLinkOption
    {
        PromptUser = 0x00,
        DontUpdate = 0x01,
        SilentlyUpdate = 0x02
    }

    /// <summary>
    /// Bookbool Record. 
    /// </summary>
    internal class TBookBoolRecord : TxBaseRecord
    {
        internal TBookBoolRecord(int aId, byte[] aData)
            : base(aId, aData)
        {
        }

        internal TBookBoolRecord()
            : base((int)xlr.BOOKBOOL, new byte[2])
        {
        }

        internal bool SaveExternalLinkValues
        {
            get
            {
                return (GetWord(0) & 0x1) == 0;
            }
            set
            {
                if (value) SetWord(0, GetWord(0) & ~1); else SetWord(0, GetWord(0) | 1);
            }
        }

        internal bool GetFlag(TBookBoolOption bo)
        {
            return (GetWord(0) & (int)bo) == 1;
        }

        internal void SetFlag(TBookBoolOption bo, bool value)
        {
            if (value) SetWord(0, GetWord(0) & ~(int)bo); else SetWord(0, GetWord(0) | (int)bo);
        }

        internal TUpdateLinkOption UpdateLinks
        {
            get
            {
                return (TUpdateLinkOption)((GetWord(0) >> 5) & 0x3);
            }
            set
            {
                if (!Enum.IsDefined(typeof(TUpdateLinkOption), value)) return;
                Data[0] = (byte)((Data[0] & ~0x60) | (((byte)value) << 5));
            }
        }



        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.BookBool = this;
        }
    }

    /// <summary>
    /// Multithread recalculation Record. 
    /// </summary>
    internal class TMTRSettingsRecord : TxBaseRecord
    {
        internal TMTRSettingsRecord(int aId, byte[] aData)
            : base(aId, aData)
        {
        }

        internal TMTRSettingsRecord()
            : base
            ((int)xlr.MTRSETTINGS,
            new byte[]{0x9A, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 
                0x01, 0x00, 0x00, 0x00, 
                0x00, 0x00, 0x00, 0x00, 
                0x00, 0x00, 0x00, 0x00, })
        {
        }

        /// <summary>
        /// returns -1 for auto, 0 for disabled or a positive number with the number of threads allowed.
        /// </summary>
        internal int NumberOfThreads
        {
            get
            {
                bool Enabled = GetCardinal(12) == 1;
                if (!Enabled) return 0;
                bool UserSet = GetCardinal(16) == 1;
                if (!UserSet) return -1;
                return GetWord(20);
            }
            set
            {
                if (value == 0)
                {
                    SetCardinal(12, 0);
                    SetCardinal(16, 0);
                    SetCardinal(20, 0);
                    return;
                }

                if (value == -1)
                {
                    SetCardinal(12, 1);
                    SetCardinal(16, 0);
                    SetCardinal(20, 0);
                    return;
                }

                SetCardinal(12, 0);
                SetCardinal(16, 1);
                if (value > 0x400) SetCardinal(20, 0x400); else SetCardinal(20, value);
            }
        }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.MTRSettings = this;
        }
    }

    /// <summary>
    /// Force Full recalculation Record. 
    /// </summary>
    internal class TForceFullCalculationRecord : TxBaseRecord
    {
        internal TForceFullCalculationRecord(int aId, byte[] aData)
            : base(aId, aData)
        {
        }

        internal TForceFullCalculationRecord()
            : base
            ((int)xlr.FORCEFULLCALCULATION,
            new byte[]{0xA3, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 
                0x00, 0x00, 0x00, 0x00
            })
        {
        }


        internal bool FullRecalc
        {
            get
            {
                return GetCardinal(12) == 1;
            }
            set
            {
                if (value) SetCardinal(12, 1); else SetCardinal(12, 0);
            }
        }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.ForceFullCalculation = this;
        }
    }

    /// <summary>
    /// CompressPictures. 
    /// </summary>
    internal class TCompressPicturesRecord : TxBaseRecord
    {
        internal TCompressPicturesRecord(int aId, byte[] aData)
            : base(aId, aData)
        {
        }

        internal TCompressPicturesRecord()
            : base
            ((int)xlr.COMPRESSPICTURES,
            new byte[]{0x9B, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 
                0x00, 0x00, 0x00, 0x01
            })
        {
        }


        internal bool Compression
        {
            get
            {
                return GetCardinal(12) == 1;
            }
            set
            {
                if (value) SetCardinal(12, 1); else SetCardinal(12, 0);
            }
        }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.CompressPictures = this;
        }
    }

    /// <summary>
    /// Compatibility checks. 
    /// </summary>
    internal class TCompat12Record : TxBaseRecord
    {
        internal TCompat12Record(int aId, byte[] aData)
            : base(aId, aData)
        {
        }

        internal TCompat12Record()
            : base
            ((int)xlr.COMPAT12,
            new byte[]{0x8C, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 
                0x00, 0x00, 0x00, 0x00
            })
        {
        }


        internal bool CompatCheck
        {
            get
            {
                return GetCardinal(12) == 0;
            }
            set
            {
                if (value) SetCardinal(12, 0); else SetCardinal(12, 1);
            }
        }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.Compat12 = this;
        }
    }

    /// <summary>
    /// GUID for the VBA project. 
    /// </summary>
    internal class TGUIDTypeLibRecord : TxBaseRecord
    {
        internal TGUIDTypeLibRecord(int aId, byte[] aData)
            : base(aId, aData)
        {
        }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.GUIDTypeLib = this;
        }
    }

    /// <summary>
    /// Excel tried to recover this file. 
    /// </summary>
    internal class TCRErrRecord : TxBaseRecord
    {
        internal TCRErrRecord(int aId, byte[] aData)
            : base(aId, aData)
        {
        }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            //We will ignore this record.
        }
    }

    internal enum THideObj
    {
        ShowAll = 0,
        ShowPlaceholder = 1,
        HideAll = 2
    }

    /// <summary>
    /// HideObj Record. 
    /// </summary>
    internal class THideObjRecord : TWordRecord
    {
        internal THideObjRecord(int aId, byte[] aData): base (aData) {}

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.HideObj = (THideObj)Value;
        }

        internal static void SaveRecord(IDataStream DataStream, THideObj aHideObj)
        {
            DataStream.WriteHeader((UInt16)xlr.HIDEOBJ, 2);
            DataStream.Write16((UInt16)aHideObj);
        }
    }

    /// <summary>
    /// RefMode Record. 
    /// </summary>
    internal class TRefModeRecord : TBoolRecord
    {
        internal TRefModeRecord(int aId, byte[] aData)
            : base(aData){ }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.FWorkbookGlobals.CalcOptions.A1RefMode = Value;
        }

        internal static void SaveRecord(IDataStream DataStream, bool A1RefMode)
        {
            DataStream.WriteHeader((UInt16)xlr.REFMODE, 2);
            if (A1RefMode) DataStream.Write16(1); else DataStream.Write16(0);
        }
    }

    internal class TIterationRecord : TBoolRecord
    {
        internal TIterationRecord(int aId, byte[] aData)
            : base(aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.FWorkbookGlobals.CalcOptions.IterationEnabled = Value;
        }

        internal static void SaveRecord(IDataStream DataStream, bool IterationEnabled)
        {
            DataStream.WriteHeader((UInt16)xlr.ITERATION, 2);
            if (IterationEnabled) DataStream.Write16(1); else DataStream.Write16(0);
        }
    }

    internal class TCalcModeRecord : TWordRecord
    {
        internal TCalcModeRecord(int aId, byte[] aData): base(aData){}

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.FWorkbookGlobals.CalcOptions.CalcMode = (TSheetCalcMode)Value;
        }

        internal static void SaveRecord(IDataStream DataStream, TSheetCalcMode CalcMode)
        {
            DataStream.WriteHeader((UInt16)xlr.CALCMODE, 2);
            DataStream.Write16((UInt16)CalcMode);
        }
    }

    internal class TCalcCountRecord : TWordRecord
    {
        internal TCalcCountRecord(int aId, byte[] aData)
            : base(aData){ }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.FWorkbookGlobals.CalcOptions.CalcCount = Value;
        }

        internal static void SaveRecord(IDataStream DataStream, int CalcCount)
        {
            DataStream.WriteHeader((UInt16)xlr.CALCCOUNT, 2);
            DataStream.Write16((UInt16)CalcCount);
        }
    }

    internal class TDeltaRecord : TDoubleRecord
    {
        internal TDeltaRecord(int aId, byte[] aData): base(aData) {}

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.FWorkbookGlobals.CalcOptions.Delta = Value;
        }

        internal static int StandardSize()
        {
            return XlsConsts.SizeOfTRecordHeader + 8;
        }

        internal static void SaveRecord(IDataStream DataStream, double Delta)
        {
            DataStream.WriteHeader((UInt16)xlr.DELTA, 8);
            DataStream.Write(BitConverter.GetBytes(Delta), 8);
        }
    }

    internal class TSaveRecalcRecord : TBoolRecord
    {
        internal TSaveRecalcRecord(int aId, byte[] aData)
            : base(aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.FWorkbookGlobals.CalcOptions.SaveRecalc = Value;
        }

        internal static void SaveRecord(IDataStream DataStream, bool SaveRecalc)
        {
            DataStream.WriteHeader((UInt16)xlr.SAVERECALC, 2);
            if (SaveRecalc) DataStream.Write16(1); else DataStream.Write16(0);
        }
    }

    internal class TGridSetRecord : TBoolRecord
    {
        internal TGridSetRecord(int aId, byte[] aData)
            : base(aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.SheetGlobals.GridSet = Value;
        }

        internal static void SaveRecord(IDataStream DataStream, bool GridSet)
        {
            DataStream.WriteHeader((UInt16)xlr.GRIDSET, 2);
            if (GridSet) DataStream.Write16(1); else DataStream.Write16(0);
        }
    }

    internal class TSyncRecord : TBaseRecord
    {
        internal int Row;
        internal int Col;

        internal TSyncRecord(int aId, byte[] aData)
        {
            Row = BitOps.GetWord(aData, 0);
            Col = BitOps.GetWord(aData, 2);
        }

        internal TSyncRecord(int aRow, int aCol)
        {
            Row = aRow;
            Col = aCol;
        }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.SheetGlobals.Sync = this;
        }

        protected override TBaseRecord DoCopyTo(TSheetInfo SheetInfo)
        {
            return (TBaseRecord)MemberwiseClone();
        }

        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
        {
            Workbook.WriteHeader((UInt16)xlr.SYNC, (UInt16)TotalSizeNoHeaders());
            Workbook.WriteRow(Row);
            Workbook.WriteCol(Col);
        }

        internal override int TotalSize()
        {
            return XlsConsts.SizeOfTRecordHeader + TotalSizeNoHeaders();
        }

        internal override int TotalSizeNoHeaders()
        {
            return 4;
        }

        internal override int GetId
        {
            get { return (int)xlr.SYNC; }
        }
    }

    internal class TLprRecord : TxBaseRecord
    {
        internal TLprRecord(int aId, byte[] aData)
            : base(aId, aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.SheetGlobals.Lpr = this;
        }
    }

    internal class TPlvRecord : TxBaseRecord
    {
        internal TPlvRecord(int aId, byte[] aData)
            : base(aId, aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.Window.Plv = this;
        }

        internal int Zoom
        {
            get
            {
                return GetWord(12);
            }
        }

        internal bool PageLayoutPreview
        {
            get
            {
                return (Data[14] & 0x01) != 0;
            }
        }

        internal bool ShowRuler
        {
            get
            {
                return (Data[14] & 0x02) != 0;
            }
        }

        internal bool ShowWhiteSpace
        {
            get
            {
                return (Data[14] & 0x04) == 0; //this is reversed.
            }
        }           
    }

    internal class TBgPicRecord : TxBaseRecord
    {
        internal TBgPicRecord(int aId, byte[] aData)
            : base(aId, aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.BgPic = this;
        }
    }

    internal class TBigNameRecord : TxBaseRecord
    {
        internal TBigNameRecord(int aId, byte[] aData)
            : base(aId, aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.BigNames.Add(this);
        }
    }

    internal class TScenManRecord : TxBaseRecord
    {
        internal TScenManRecord(int aId, byte[] aData)
            : base(aId, aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.Scenarios.ScenMan = this;
        }
    }

    internal class TScenarioRecord : TxBaseRecord
    {
        internal TScenarioRecord(int aId, byte[] aData)
            : base(aId, aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.Scenarios.Scenarios.Add(this);
        }
    }

    internal class TSortDataRecord : TxBaseRecord
    {
        internal TSortDataRecord(int aId, byte[] aData)
            : base(aId, aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            int spf = GetWord(12);
            spf = (spf >> 3) & 0x07;
            switch (spf)
            {
                case 0: ws.SortAndFilter.SortData = this; break;
                case 1: ws.Feat11.Add(this); break;


                    //This one is a little hack. For reasons only known to Excel developers,
                    //Excel doesn't include a CustomView GUID in SortData, so we can't know where it goes when
                    //saved by Excel 2003 which will put all of them at the end. But as order is preserved,
                    //and AutoFilter12 normally have a GUID, and this record should go after AutoFilter12, we can guess something.
                    //By the way, Excel just ignores the SortData in custom views. But it saves out the records... so we have to ignore the ones that don't apply.
                    //Custom views go after the main thing, so the first record should be the good one. Not 100% fool proof (a custom view could have a sort12 record and the
                    //main view not, and have been saved with Excel 2003, in which case we don't know (and neither does Excel).
                case 2:
                    if (ws.SortAndFilter.AutoFilter.Sort12.Count == 0 && Loader.CustomView == null)
                    {
                        ws.SortAndFilter.AutoFilter.Sort12.Add(this); 
                    }
                    break;
                case 3: ws.QueryTable.QueryItems.Add(this); break;

                default: ws.FutureRecords.Add(this); break;
                    

            }
            
        }
    }

    internal class TSortRecord : TxBaseRecord
    {
        internal TSortRecord(int aId, byte[] aData)
            : base(aId, aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.SortAndFilter.Sort = this;
        }
    }

    internal class TDropDownObjIdsRecord : TxBaseRecord
    {
        internal TDropDownObjIdsRecord(int aId, byte[] aData)
            : base(aId, aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.SortAndFilter.DropDownObjIds = this;
        }
    }

    internal class TRRSortRecord : TxBaseRecord
    {
        internal TRRSortRecord(int aId, byte[] aData)
            : base(aId, aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.RRSort.Add(this);
        }
    }

    internal class TLRngRecord : TxBaseRecord
    {
        internal TLRngRecord(int aId, byte[] aData)
            : base(aId, aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.LRng = this;
        }
    }

    internal class TPhoneticRecord: TxBaseRecord
    {
        internal TPhoneticRecord(int aId, byte[] aData):base(aId, aData){}

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.Phonetic = this;
        }
    }

    internal class TFilterModeRecord : TNotStorableRecord
    {
        internal TFilterModeRecord(int aId, byte[] aData) : base() { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.SortAndFilter.FilterMode = true;
        }

        internal static int StandardSize()
        {
            return XlsConsts.SizeOfTRecordHeader;
        }

        internal static void SaveRecord(IDataStream DataStream)
        {
            DataStream.WriteHeader((UInt16)xlr.FILTERMODE, 0);
        }
    }

    internal enum TChartPrintSize
    {
        NotDefined = -1,
        CustomView = 0,
        Strecth = 1,
        KeepAspectRatio = 2,
        DefinedInChart = 3
    }

    internal class TPrintSizeRecord : TWordRecord
    {
        internal TPrintSizeRecord(int aId, byte[] aData) : base(aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            if (Loader.CustomView != null) Loader.CustomView.Setup.ChartPrintSize = (TChartPrintSize)Value;
            else ws.PageSetup.ChartPrintSize = (TChartPrintSize)Value;
        }

        internal static int StandardSize(TChartPrintSize PrintSize)
        {
            if (PrintSize == TChartPrintSize.NotDefined) return 0;
            return StandardSize();
        }

        internal static void SaveRecord(IDataStream DataStream, TChartPrintSize PrintSize)
        {
            DataStream.WriteHeader((UInt16)xlr.PRINTSIZE, 2);
            DataStream.Write16((UInt16)PrintSize);
        }
    }

    internal class TDConRecord : TxBaseRecord
    {
        internal TDConRecord(int aId, byte[] aData) : base(aId, aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.Connections.DCon = this;
        }
    }

    internal class TDConNameRecord : TxBaseRecord
    {
        internal TDConNameRecord(int aId, byte[] aData) : base(aId, aData) { }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.PivotCache.Add(this);
        }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.Connections.DConList.Add(this);
        }
    }

    internal class TDConBinRecord : TxBaseRecord
    {
        internal TDConBinRecord(int aId, byte[] aData) : base(aId, aData) { }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.PivotCache.Add(this);
        }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.Connections.DConList.Add(this);
        }
    }

    internal class TDConRefRecord : TxBaseRecord
    {
        internal TDConRefRecord(int aId, byte[] aData) : base(aId, aData) { }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.PivotCache.Add(this);
        }
        
        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.Connections.DConList.Add(this);
        }
    }

    internal class TQSIRecord : TxBaseRecord
    {
        internal TQSIRecord(int aId, byte[] aData) : base(aId, aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.QueryTable.QueryItems.Add(this);
            TBaseRecord R = null;
            bool IsQueryRecord = false;
            do
            {
                R = RecordLoader.LoadRecord(out rRow, false);
                IsQueryRecord = IsQuery(R);
                if (IsQueryRecord) ws.QueryTable.QueryItems.Add(R);
            }
            while (IsQueryRecord && !Loader.Eof);

            if (R != null && ! IsQueryRecord)  //Null will be ignored
                R.LoadIntoSheet(ws, 0, RecordLoader, ref Loader);
        }

        private static bool IsQuery(TBaseRecord R)
        {
            xlr rid = (xlr)R.GetId;
            switch (rid)
            {
                case xlr.QSI: return true;
                case xlr.SXEXTPARAMQRY: return true;
                case xlr.SXSTRING: return true;
                case xlr.QSISXTAG: return true;
                case xlr.DBQUERYEXT: return true;
                case xlr.EXTSTRING: return true;
                case xlr.OLEDBCONN: return true;
                case xlr.TXTQRY: return true;
                case xlr.SXADDL: return true;
                case xlr.QSIR: return true;
                case xlr.QSIF: return true;
                case xlr.SORTDATA: return true;
            }
            return false;
        }
    }

    internal class TFeatRecord : TxBaseRecord
    {
        internal TFeatRecord(int aId, byte[] aData) : base(aId, aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.Feat.Add(this);
        }

    }

    internal class TFeat1112Record : TxBaseRecord
    {
        internal TFeat1112Record(int aId, byte[] aData) : base(aId, aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.Feat11.Add(this);
        }

    }

    internal class TFeatHdr11Record : TxBaseRecord
    {
        internal TFeatHdr11Record(int aId, byte[] aData) : base(aId, aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.Feat11.Add(this);
        }

    }

    internal class TList12Record : TxBaseRecord
    {
        internal TList12Record(int aId, byte[] aData) : base(aId, aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.Feat11.Add(this);
        }

    }

    internal class TAutoFilter12Record : TxBaseRecord
    {
        internal TAutoFilter12Record(int aId, byte[] aData) : base(aId, aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            if (IsSheetAutoFilter)
            {
                TAutoFilter Target = ws.GetCustomViewAutoFilter(CustomView);
                if (Target == null) return;
                Target.Filters.Add(this);
            }
            else
            {
                ws.Feat11.Add(this);
            }
        }

        private bool IsSheetAutoFilter
        {
            get
            {
                return GetCardinal(40) == 0xFFFFFFFF;
            }
        }

        private Guid CustomView
        {
            get
            {
                byte[] bCustomView = new byte[16];
                Array.Copy(Data, 44, bCustomView, 0, 16);
                return new Guid(bCustomView);
            }
        }

    }


    internal class TUserSViewBeginRecord : TxBaseRecord
    {
        internal TUserSViewBeginRecord(int aId, byte[] aData)
            : base(aId, aData)
        {
        }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            TCustomView cv = ws.CustomViews.Add(this);
            Loader.CustomView = cv;
        }

        internal Guid CustomView
        {
            get
            {
                byte[] bCustomView = new byte[16];
                Array.Copy(Data, 0, bCustomView, 0, 16);
                return new Guid(bCustomView);
            }
        }

    }

    internal class TUserSViewEndRecord : TNotStorableRecord
    {
        internal TUserSViewEndRecord(int aId, byte[] aData)
        {
        }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            Loader.CustomView = null;
        }

        internal static void SaveRecord(IDataStream DataStream)
        {
            DataStream.WriteHeader((UInt16)xlr.USERSVIEWEND, 2);
            DataStream.Write16(1);
        }

        internal static int StandardSize()
        {
            return XlsConsts.SizeOfTRecordHeader + 2;
        }
    }

    internal class TUnitsRecord : TWordRecord
    {
        internal TUnitsRecord(int aId, byte[] aData) : base(aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            TFlxChart chart = ws as TFlxChart;
            if (chart != null)
            {
                chart.Units = Value;
                return;
            }

            base.LoadIntoSheet(ws, rRow, RecordLoader, ref Loader);
        }

        internal static void SaveRecord(IDataStream DataStream, int aUnit)
        {
            DataStream.WriteHeader((UInt16)xlr.UNITS, 2);
            DataStream.Write16((UInt16)aUnit);
        }
    }

    internal class TCrtMlFrtRecord : TxBaseRecord
    {
        internal TCrtMlFrtRecord(int aId, byte[] aData)
            : base(aId, aData)
        {
        }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            TFlxChart chart = ws as TFlxChart;
            if (chart != null)
            {
                chart.CrtMlFrt.Add(this);
                return;
            }

            base.LoadIntoSheet(ws, rRow, RecordLoader, ref Loader);
        }
    }

    internal class TFrtInfoRecord : TxBaseRecord
    {
        internal TFrtInfoRecord(int aId, byte[] aData)
            : base(aId, aData)
        {
        }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            TFlxChart chart = ws as TFlxChart;
            if (chart != null)
            {
                chart.FrtInfo = this;
                return;
            }

            base.LoadIntoSheet(ws, rRow, RecordLoader, ref Loader);
        }
    }

    internal enum TChartSIIndexType
    {
        None = 0,
        SeriesValues = 1,
        CategoryLabels = 2,
        BubbleSizes = 3
    }

    internal class TChartSIIndexRecord : TBaseRecord
    {
        internal TChartSIIndexType NumIndex;
        internal TChartCellList Cells;

        internal TChartSIIndexRecord(TChartSIIndexType aNumIndex)
        {
            NumIndex = aNumIndex;
            Cells = new TChartCellList();
        }

        internal TChartSIIndexRecord(int aId, byte[] aData): this((TChartSIIndexType)BitOps.GetWord(aData, 0))
        {
        }

        protected override TBaseRecord DoCopyTo(TSheetInfo SheetInfo)
        {
            TChartSIIndexRecord Result = new TChartSIIndexRecord(NumIndex);
            Result.Cells.CopyFrom(Cells, SheetInfo);

            return Result;
        }

        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
        {
            Workbook.WriteHeader((UInt16)xlr.ChartSiindex, 2);
            Workbook.Write16((UInt16)NumIndex);
            Cells.SaveToStream(Workbook, SaveData, null);
        }

        internal override int TotalSize()
        {
            return XlsConsts.SizeOfTRecordHeader + 2 + (int)Cells.TotalSize(null);
        }

        internal override int TotalSizeNoHeaders()
        {
            return 2 + (int)Cells.TotalSizeNoHeaders();
        }

        internal override int GetId
        {
            get { return (int)xlr.ChartSiindex; }
        }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            TFlxChart chart = ws as TFlxChart;
            if (chart != null)
            {
                chart.SeriesData.Add(this);
                ReadCells(ws, rRow, RecordLoader, ref Loader);
                return;
            }
            
            
            base.LoadIntoSheet(ws, rRow, RecordLoader, ref Loader);
        }

        private void ReadCells(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            TBaseRecord R = null;
            bool IsCellRecord = false;
            do
            {
                R = RecordLoader.LoadRecord(out rRow, false);
                IsCellRecord = IsCell(R);
                if (IsCellRecord) Cells.AddRecord((TCellRecord)R, rRow);
            }
            while (IsCellRecord && !Loader.Eof);

            if (R != null && !IsCellRecord)  //Null will be ignored
                R.LoadIntoSheet(ws, rRow, RecordLoader, ref Loader);

        }

        private static bool IsCell(TBaseRecord R)
        {
            if (!(R is TCellRecord)) return false;
            xlr rid = (xlr)R.GetId;
            switch (rid)
            {
                case xlr.NUMBER: return true;
                case xlr.BOOLERR: return true;
                case xlr.BLANK: return true;
                case xlr.LABEL: return true;
            }
            return false;
        }

        internal void DeleteSeries(int index)
        {
            Cells.DeleteSeries(index);
        }
    }

    internal class TInternationalRecord : TWordRecord
    {
        internal TInternationalRecord(int aId, byte[] aData) : base(aData) { }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.International = true;
        }

        internal static void SaveRecord(IDataStream DataStream)
        {
            DataStream.WriteHeader((UInt16)xlr.INTL, 2);
            DataStream.Write16(0);
        }

    }

}
