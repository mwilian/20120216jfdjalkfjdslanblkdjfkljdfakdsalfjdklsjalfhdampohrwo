using System;
using System.IO;
using System.Text;
using FlexCel.Core;
using System.Collections.Generic;

namespace FlexCel.XlsAdapter
{
    /// <summary>
    /// a Biff8 HyperLink.  (Hyperlinks might be obtained with formulas too, but those are a different thing)
    /// </summary>
    internal class THLinkRecord: TBaseRecord, IComparable
    {
        #region Variables
        int Id;
        internal int FirstRow, FirstCol, LastRow, LastCol;
        private THyperLinkType FLinkType;

        private string FDescription;
        private string FTargetFrame;
        private string FTextMark;
        private string FText;

        internal TScreenTipRecord Hint;

        private static readonly byte[] URLGUID = { 0xE0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B, 0xA9, 0x0B };
        private static readonly byte[] FILEGUID = { 0x03, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 };
        #endregion

        #region Constructors
        private THLinkRecord(int aId) 
        {
            Id = aId;
            Hint=null;
        }

        private THLinkRecord(int aId, byte[] aData) : this(aId)
        {
            LoadFromBiff8(aData);
        }

        internal static THLinkRecord CreateFromBiff8(int aId, byte[] aData)
        {
            return new THLinkRecord(aId, aData);
        }

		internal override int GetId	{ get {	return Id; }}

        #endregion

        #region Export to Biff8
        internal void SaveBiff8Data(IDataStream DataStream, bool SaveCoords, out int Len)
        {
            if (DataStream != null)
            {
                if (SaveCoords)
                {
                    DataStream.WriteRow(FirstRow);
                    DataStream.WriteRow(LastRow);
                    DataStream.WriteCol(FirstCol);
                    DataStream.WriteCol(LastCol);
                }

                //CLSID
                DataStream.Write32(0x79EAC9D0);
                DataStream.Write32(0x11CEBAF9);
                DataStream.Write32(0xAA00828C);
                DataStream.Write32(0x0BA94B00);

                //Starts HyperLink Object - MS-OSHARED
                DataStream.Write32(0x02);
                DataStream.Write32(OptionFlags);
            }

            Len = 32;

            WriteString(DataStream, FDescription, ref Len);
            WriteString(DataStream, FTargetFrame, ref Len);

            //This part of the structure is different depending on the type of link.
            switch (FLinkType)
            {
                case THyperLinkType.URL:
                    if (DataStream != null) DataStream.Write(URLGUID, URLGUID.Length);
                    Len += URLGUID.Length;
                    WriteString(DataStream, FText, ref Len, 1); 
                    break;
                case THyperLinkType.LocalFile:
                    if (DataStream != null) DataStream.Write(FILEGUID, FILEGUID.Length);
                    Len += FILEGUID.Length;
                    WriteLocalFile(DataStream, FText, ref Len);
                    break;
                case THyperLinkType.UNC:
                    //String Moniker
                    WriteString(DataStream, FText, ref Len);
                    break;
                case THyperLinkType.CurrentWorkbook:
                    //CurrentWorkbook doesn't have monikers.
                    break;
                default:
                    XlsMessages.ThrowException(XlsErr.ErrInvalidHyperLinkType, (int)FLinkType);
                    break;
            }

            
            WriteString(DataStream, FTextMark, ref Len);
        }

        internal byte[] GetHLinkStream()
        {
            using (MemOle2 MemFile = new MemOle2())
            {
                SaveBiff8Data(MemFile, false);
                return MemFile.GetBytes();
            }
        }

        private void SaveBiff8Data(IDataStream DataStream, bool SaveCoords)
        {
            int Len;
            SaveBiff8Data(DataStream, SaveCoords, out Len);
        }

        private int CalcBiff8Len()
        {
            int Len;
            SaveBiff8Data(null, true, out Len);
            return Len;
        }

        private UInt32 OptionFlags
        {
            get
            {
                UInt32 Result = 0;
                if (FLinkType != THyperLinkType.CurrentWorkbook) Result |= 0x01;  //Has moniker. In Excel, only currentworkbook links don't have monikers.  - bit 0
                if (IsAbsolute(FText)) Result |= 0x02;  // Relative file path or url - bit 1.
                if (FTextMark != null && FTextMark.Length > 0) Result |= 0x08; //Has text mark   - bit 3
                if (FDescription != null && FDescription.Length > 0) Result |= 0x14; //Has description  - bits 2 and 4
                if (FTargetFrame != null && FTargetFrame.Length > 0) Result |= 0x80; //Has target frame  -bit 7
                if (FLinkType == THyperLinkType.UNC) Result |= 0x100;  //Has string moniker. In Excel, only unc links have string monikers.  - bit 8
                
                return Result;
            }
        }

        private bool IsAbsolute(string p)
        {
            if (FLinkType == THyperLinkType.URL || FLinkType == THyperLinkType.UNC) return true;
            if (FLinkType == THyperLinkType.CurrentWorkbook) return false;
            return Path.IsPathRooted(p);
        }

        private void WriteString(IDataStream DataStream, string value, ref int NewPos)
        {
            WriteString(DataStream, value, ref NewPos, 2);
        }

        private void WriteString(IDataStream DataStream, string value, ref int NewPos, int ByteSize)
        {
            if (value == null || value.Length == 0) return;
            byte[] ByteStr = Encoding.Unicode.GetBytes(value);

            if (DataStream != null) DataStream.Write32((UInt32)((ByteStr.Length  + 2)/ ByteSize));
            NewPos += 4;

            if (DataStream != null) DataStream.Write(ByteStr, ByteStr.Length);
            NewPos += ByteStr.Length;

            if (DataStream != null) DataStream.Write16(0);
            NewPos += 2;
        }

        private void WriteLocalFile(IDataStream DataStream, string value, ref int NewPos)
        {
            int i = 0;
            while (i + 3 <= value.Length && value.Substring(i, 3) == ".." + Path.DirectorySeparatorChar) i += 3;
            value = value.Substring(i);

            bool IsCompressed = !StrOps.IsWide(value);
            int WideDataLen = 0;
            byte[] ByteStr = null;
            if (!IsCompressed)
            {
                ByteStr = Encoding.Unicode.GetBytes(value);
                WideDataLen = 4 + 2 + ByteStr.Length;
            }

            NewPos += 2 + 4 + value.Length + 1 + 24 + 4 + WideDataLen;
            if (DataStream == null) return;


            DataStream.Write16((UInt16) (i / 3));

            DataStream.Write32((UInt32)(value.Length + 1));

            byte[] NewData = new byte[value.Length + 1];
            StrOps.CompressBestUnicode(value, NewData, 0);

            DataStream.Write(NewData, NewData.Length);

            
            DataStream.Write32(0xDEADFFFF);
            NewData = new byte[20];
            DataStream.Write(NewData, NewData.Length);

			if (IsCompressed)
			{
				DataStream.Write32(0);
				return;
			}
			else
			{
				DataStream.Write32((UInt32)( 4 + 2 + ByteStr.Length));
			}

            DataStream.Write32((UInt32) ByteStr.Length);
            DataStream.Write16(0x0003);

            DataStream.Write(ByteStr, ByteStr.Length);
        }

        #endregion

        #region Import From Biff8
        private void LoadFromBiff8(byte[] Data)
        {
            FirstRow = BitOps.GetWord(Data, 0);
            LastRow = BitOps.GetWord(Data, 2);
            FirstCol = BitOps.GetWord(Data, 4);
            LastCol = BitOps.GetWord(Data, 6);

            int pos = 32;
            int Flags = BitOps.GetWord(Data, 28);
            FDescription = ReadString(Data, Flags, ref pos, 0x14, 2);
            FTargetFrame = ReadString(Data, Flags, ref pos, 0x80, 2);

            FText = GetMoniker(Data, Flags, ref pos, ref FLinkType);

            FTextMark = ReadString(Data, Flags, ref pos, 0x08, 2);


        }

        private  static string ReadString(byte[] Data, int Flags, ref int Pos, int OptMask, int ByteSize)
        {
            if ((Flags & OptMask) != OptMask)
                return String.Empty;
            else
            {
                int OldPos = Pos;
                Pos += 4 + (int)BitOps.GetCardinal(Data, Pos) * ByteSize;
                string Result = Encoding.Unicode.GetString(Data, OldPos + 4, Pos - (OldPos + 4) - 2); //00 terminated.	
                int P = Result.IndexOf((char)0); //string might have a 0 inside. In this case we need to cut it.
                if (P >= 0) return Result.Substring(0, P);
                return Result;
            }
        }

        private string ReadLocalFile(byte[] Data, ref int Pos)
        {
            StringBuilder Result = new StringBuilder();
            int DirUp = BitOps.GetWord(Data, Pos);
            for (int i = 0; i < DirUp; i++)
                Result.Append(".." + Path.DirectorySeparatorChar);

            Pos += 2;
            int StrLen = (int)BitOps.GetCardinal(Data, Pos);
            if (StrLen > 1) StrLen--;
            string s8 = StrOps.UnCompressUnicode(Data, Pos + 4, StrLen);

            Pos += 4 + StrLen + 1 + 24;

            int RLen = (int)BitOps.GetCardinal(Data, Pos);
            Pos += 4;
            if (RLen == 0)
            {
                Result.Append(s8);
                return Result.ToString();
            }

            int XLen = (int)BitOps.GetCardinal(Data, Pos);
            Pos += 4 + 2;

            Result.Append(Encoding.Unicode.GetString(Data, Pos, XLen));
            Pos += XLen;

            return Result.ToString();
        }


        private static bool IsUrl(byte[] Data, int Flags, int pos)
        {
            if ((Flags & 0x01) == 0x01 && (Flags & 0x100) == 0)
                return BitOps.CompareMem(URLGUID, Data, pos);
            return false;
        }

        private static bool IsFile(byte[] Data, int Flags, int pos)
        {
            if ((Flags & 0x01) == 0x01 && (Flags & 0x100) == 0)
                return BitOps.CompareMem(FILEGUID, Data, pos);
            return false;
        }

        private static bool IsUNC(byte[] Data, int Flags, int pos)
        {
            return ((Flags & 0x01) == 0x01 && (Flags & 0x100) != 0);
        }

        private string GetMoniker(byte[] Data, int Flags, ref int pos, ref THyperLinkType HType)
        {
            if (IsUrl(Data, Flags, pos))
            {
                HType = THyperLinkType.URL;
                pos += 16;
                return ReadString(Data, Flags, ref pos, 0, 1);
            }
            if (IsFile(Data, Flags, pos))
            {
                HType = THyperLinkType.LocalFile;
                pos += 16;
                return ReadLocalFile(Data, ref pos);
            }
            if (IsUNC(Data, Flags, pos))
            {
                HType = THyperLinkType.UNC;
                return ReadString(Data, Flags, ref pos, 0, 2);
            }

            HType = THyperLinkType.CurrentWorkbook;
            return string.Empty;
        }


        #endregion

        #region Implementation
        internal static THLinkRecord CreateNew(TXlsCellRange CellRange, THyperLink HLink)
        {
            THLinkRecord Result= new THLinkRecord((int)xlr.HLINK);
            Result.FirstRow = CellRange.Top;
            Result.FirstCol = CellRange.Left;
            Result.LastRow = CellRange.Bottom;
            Result.LastCol = CellRange.Right;

            Result.SetProperties(HLink);
            return Result;
        }

        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
        {
			Workbook.WriteHeader((UInt16) Id, (UInt16)CalcBiff8Len());
            SaveBiff8Data(Workbook, true);
            if (Hint!=null) 
            {
                Hint.FirstRow=FirstRow;
                Hint.FirstCol=FirstCol;
                Hint.LastRow=LastRow;
                Hint.LastCol=LastCol;
                Hint.SaveToStream(Workbook, SaveData, Row);
            }
        }

        internal void SaveRangeToStream(IDataStream Workbook, TSaveData SaveData, TXlsCellRange CellRange)
        {
            if (FirstRow>CellRange.Bottom || LastRow<CellRange.Top || FirstCol>CellRange.Right || LastCol<CellRange.Left) return;
            SaveToStream(Workbook, SaveData, -1);
        }

        internal int TotalRangeSize(TXlsCellRange CellRange)
        {
            if (FirstRow>CellRange.Bottom || LastRow<CellRange.Top || FirstCol>CellRange.Right || LastCol<CellRange.Left) return 0;
            return TotalSize();
        }

        protected override TBaseRecord DoCopyTo(TSheetInfo SheetInfo)
        {
            THLinkRecord b = (THLinkRecord)MemberwiseClone();
            if (Hint != null) b.Hint= (TScreenTipRecord) TScreenTipRecord.Clone(Hint, SheetInfo);
            return b;
        }

        internal override int TotalSize()
        {
            int Result = CalcBiff8Len() + XlsConsts.SizeOfTRecordHeader;
            if (Hint!=null) Result+=Hint.TotalSize();
            return Result;
        }

        internal override int TotalSizeNoHeaders()
        {
            int Result=CalcBiff8Len(); 
            if (Hint!=null) Result+=Hint.TotalSizeNoHeaders();
            return Result;
        }

        internal THyperLink GetProperties()
        {
            THyperLink Result= new THyperLink();
            Result.Description= MakeNotNull(FDescription);
            Result.TargetFrame= MakeNotNull(FTargetFrame);

            Result.Text= MakeNotNull(FText);
            Result.LinkType=FLinkType;
            
            Result.TextMark= MakeNotNull(FTextMark);

            if (Hint==null) Result.Hint=String.Empty; else Result.Hint=Hint.Text;

            return Result;
        }

        private string MakeNotNull(string s)
        {
            if (s == null) return string.Empty;
            return s;
        }

        internal void SetProperties(THyperLink value)
        {
            if (value==null) value= new THyperLink();

            FLinkType = value.LinkType;
            FDescription = value.Description;
            FTargetFrame = value.TargetFrame;
            FText = value.Text;
            FTextMark =value.TextMark;

            if (value.Hint==null || value.Hint.Length==0) Hint=null;
            else 
            {
                if (Hint==null) Hint = TScreenTipRecord.CreateNew(value.Hint);
                else Hint.Text=value.Hint;
            }
        }

        internal TXlsCellRange GetCellRange()
        {
            return new TXlsCellRange(FirstRow, FirstCol, LastRow, LastCol);
        }

        internal void SetCellRange(TXlsCellRange CellRange)
        {
            FirstRow=CellRange.Top;
            FirstCol=CellRange.Left;
            LastRow=CellRange.Bottom;
            LastCol=CellRange.Right;
        }

		internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
		{
			ws.HLinks.Add(this);
		}

        internal void ArrangeInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
        {
            //Hyperlink data doesn't move when you insert/copy cells or sheets. It is a static text.

            if ((SheetInfo.InsSheet<0) || (SheetInfo.SourceFormulaSheet != SheetInfo.InsSheet)) return;
            
            if (aRowCount!=0 && FirstCol>=CellRange.Left && LastCol<=CellRange.Right)
            {
                if (FirstRow>=CellRange.Top) 
                    BitOps.IncWord(ref FirstRow, aRowCount*CellRange.RowCount, FlxConsts.Max_Rows, XlsErr.ErrTooManyRows);  //firstrow;
                if (LastRow>=CellRange.Top) 
                    BitOps.IncWord(ref LastRow, aRowCount*CellRange.RowCount, FlxConsts.Max_Rows, XlsErr.ErrTooManyRows);  // lastrow;
            }

            if (aColCount!=0 && FirstRow>=CellRange.Top && LastRow<=CellRange.Bottom)
            {
                if (FirstCol>=CellRange.Left) 
                    BitOps.IncWord(ref FirstCol, aColCount*CellRange.ColCount, FlxConsts.Max_Columns, XlsErr.ErrTooManyColumns);  //firstcol;
                if (LastCol>=CellRange.Left) 
                    BitOps.IncWord(ref LastCol, aColCount*CellRange.ColCount, FlxConsts.Max_Columns, XlsErr.ErrTooManyColumns);  //lastcol;
            }

        }

        internal THLinkRecord Offset(int DeltaRow, int DeltaCol)
        {
            FirstRow+=DeltaRow;
            LastRow+=DeltaRow;
            FirstCol+=DeltaCol;
            LastCol+=DeltaCol;
            return this;
        }
        #endregion

        #region Properties
        /// <summary>
        /// Can be null, different from THLink
        /// </summary>
        public string Description { get { return FDescription; } }

        public string TextMark { get { return FTextMark; } }

        public string Text { get { return FText; } }

        public THyperLinkType LinkType { get { return FLinkType; } }

        #endregion

        #region IComparable Members

        public int CompareTo(object obj)
        {
            THLinkRecord o2= (THLinkRecord)obj;
            int Result= FirstRow.CompareTo(o2.FirstRow);
            if (Result==0) return FirstCol.CompareTo(o2.FirstCol);
            return Result;
        }

        #endregion

    }

    internal class TScreenTipRecord : TBaseRecord
    {
        int Id;
        internal int FirstRow;
        internal int LastRow;
        internal int FirstCol;
        internal int LastCol;
        private string FText;

        private TScreenTipRecord(int aId)
        {
            Id = aId;
        }

        private TScreenTipRecord(int aId, byte[] aData)
            : this(aId)
        {
            FirstRow = BitOps.GetWord(aData, 2);
            LastRow = BitOps.GetWord(aData, 4);
            FirstCol = BitOps.GetWord(aData, 6);
            LastCol = BitOps.GetWord(aData, 8);
            FText = Encoding.Unicode.GetString(aData, 10, aData.Length - 10 - 2);
        }

        internal static TBaseRecord CreateFromBiff8(int Id, byte[] Data)
        {
            return new TScreenTipRecord(Id, Data);
        }

        internal static TScreenTipRecord CreateNew(string aDescription)
        {
            TScreenTipRecord Result = new TScreenTipRecord((int)xlr.SCREENTIP);
            Result.FText = aDescription;
            return Result;
        }

		internal override int GetId	{ get {	return Id; }}

        internal string Text
        {
            get
            {
                return FText;
            }
            set
            {
                FText = value;
            }
        }

        private string TrimmedText
        {
            get
            {
                if (FText == null || FText.Length <=255) return FText;
                return FText.Substring(0, 255);
            }
        }
        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
        {
            if (FText == null || FText.Length == 0) return;

			byte[] b = Encoding.Unicode.GetBytes(TrimmedText);

			Workbook.WriteHeader((UInt16)Id, (UInt16) (b.Length + 12));
            Workbook.Write16(0x0800); //record id
            Workbook.WriteRow(FirstRow);
            Workbook.WriteRow(LastRow);
            Workbook.WriteCol(FirstCol);
            Workbook.WriteCol(LastCol);

            Workbook.Write(b, b.Length);
            Workbook.Write16(0); //null terminated.
        }

        internal override int TotalSize()
        {
            if (FText == null || FText.Length == 0) return 0;
            return TotalSizeNoHeaders() + XlsConsts.SizeOfTRecordHeader;
        }

        internal override int TotalSizeNoHeaders()
        {
            if (FText == null || FText.Length == 0) return 0;
            return 12 + Encoding.Unicode.GetByteCount(TrimmedText);
        }

        protected override TBaseRecord DoCopyTo(TSheetInfo SheetInfo)
        {
            return (TBaseRecord) MemberwiseClone();
        }
    }


    /// <summary>
    /// A list of HLinks, Records are THLinkRecord
    /// </summary>
    internal class THLinkList
    {
        protected List<THLinkRecord> FList;

        internal THLinkList()
        {
            FList=new List<THLinkRecord>();
        }
            
        #region Generics
        internal void Add (THLinkRecord a)
        {
            FList.Add(a);
        }

        internal void Insert (int index, THLinkRecord a)
        {
            FList.Insert(index, a);
        }

        protected void SetThis(THLinkRecord value, int index)
        {
            FList[index]=value;
        }

        internal THLinkRecord this[int index] 
        {
            get {return (THLinkRecord) FList[index];} 
            set {SetThis(value, index);}
        }

        internal int Count
        {
            get {return FList.Count;}
        }

        internal void Sort()
        {
            FList.Sort();
        }

        internal void Clear()
        {
            FList.Clear();
        }

        internal void Delete(int index)
        {
            FList.RemoveAt(index);
        }
        #endregion

        internal void CopyFrom(THLinkList aHLinkList, TSheetInfo SheetInfo)
        {
            if (aHLinkList.FList == FList) XlsMessages.ThrowException(XlsErr.ErrInternal); //Should be different objects
            for (int i = 0; i < aHLinkList.Count; i++)
                Add((THLinkRecord)THLinkRecord.Clone(aHLinkList[i], SheetInfo));
        }

        internal void CopyObjectsFrom(THLinkList aHLinkList, TXlsCellRange CopyRange, int RowOfs, int ColOfs, TSheetInfo SheetInfo)
        {
            if (aHLinkList==null) return;
            
            int aCount=aHLinkList.Count;
            for (int i=0; i<aCount;i++)
            {
                THLinkRecord r=aHLinkList[i];
                if (r.FirstCol>=CopyRange.Left && r.LastCol<=CopyRange.Right &&
                    r.FirstRow >= CopyRange.Top && r.LastRow <= CopyRange.Bottom)
                {
                    Add(((THLinkRecord) THLinkRecord.Clone(r, SheetInfo)).Offset(RowOfs, ColOfs));
                }

            }
        }

        internal void SaveToStream(IDataStream DataStream, TSaveData SaveData, TXlsCellRange CellRange)
        {
            if (CellRange == null) SaveAllToStream(DataStream, SaveData); else SaveRangeToStream(DataStream, SaveData, CellRange);
        }

        private void SaveAllToStream(IDataStream DataStream, TSaveData SaveData)
        {
            Sort();
            int aCount=Count;
            for (int i=0; i< aCount;i++)
                this[i].SaveToStream(DataStream, SaveData, -1);
        }

        private void SaveRangeToStream(IDataStream DataStream, TSaveData SaveData, TXlsCellRange CellRange)
        {
            Sort();
            int aCount=Count;
            for (int i=0; i< aCount;i++)
                this[i].SaveRangeToStream(DataStream, SaveData, CellRange);
        }

        internal long TotalSize(TXlsCellRange CellRange)
        {
            if (CellRange == null) return TotalSizeAll();
            return TotalRangeSize(CellRange);
        }

        private long TotalSizeAll()
        {
            long Result=0;
            int aCount=Count;
            for (int i=0; i< aCount;i++)
                Result+=this[i].TotalSize();
            return Result;
        }

        private long TotalRangeSize(TXlsCellRange CellRange)
        {
            long Result=0;
            int aCount=Count;
            for (int i=0; i< aCount;i++)
                Result+=this[i].TotalRangeSize(CellRange);
            return Result;
        }

        internal void InsertAndCopyRange(TXlsCellRange SourceRange, TFlxInsertMode InsertMode, int DestRow, int DestCol, int aRowCount, int aColCount, TRangeCopyMode CopyMode, TSheetInfo SheetInfo)
        {
            int aCount=Count;
            for (int i=0; i< aCount;i++)
            {
                this[i].ArrangeInsertRange(SourceRange.OffsetForIns(DestRow, DestCol, InsertMode), aRowCount, aColCount, SheetInfo);
            }
         
            if (CopyMode==TRangeCopyMode.None) return;

            int RTop=SourceRange.Top;
            int RLeft=SourceRange.Left;
            if (DestRow<=SourceRange.Top) RTop+=aRowCount*SourceRange.RowCount;
            if (DestCol<=SourceRange.Left) RLeft+=aColCount*SourceRange.ColCount;
            int RRight=RLeft+SourceRange.ColCount-1;
            int RBottom=RTop+SourceRange.RowCount-1;
            //Copy the cells.

            if (aRowCount>0 || aColCount>0)
            {
                for (int i=0; i< aCount;i++)
                {
                    THLinkRecord r=this[i];
                    if (r.FirstCol>=RLeft && r.LastCol<=RRight &&
                        r.FirstRow >= RTop && r.LastRow <= RBottom)
                    {
                        for (int rc=0; rc<aRowCount;rc++)
                        {
                            Add(((THLinkRecord) TBaseRecord.Clone(r, SheetInfo)).Offset(DestRow-RTop+rc*SourceRange.RowCount, DestCol-RLeft));
                        }
                        for (int cc=0; cc<aColCount;cc++)
                        {
                            Add(((THLinkRecord) TBaseRecord.Clone(r, SheetInfo)).Offset(DestRow-RTop, DestCol-RLeft+cc*SourceRange.ColCount));
                        }
                    }
                }
            }            
        }

		internal void DeleteRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
		{
			int aCount=Count;
			for (int i=aCount-1; i>=0;i--)
			{
				THLinkRecord r=this[i];
				int bRowCount=aRowCount-1;if (bRowCount<0) bRowCount=0;
				int bColCount=aColCount-1;if (bColCount<0) bColCount=0;
				if (r.FirstRow >= CellRange.Top && r.LastRow <= CellRange.Bottom+CellRange.RowCount*bRowCount &&
					r.FirstCol >= CellRange.Left && r.LastCol <= CellRange.Right+CellRange.ColCount*bColCount)
					FList.RemoveAt(i);
				else
				{
					r.ArrangeInsertRange(CellRange, -aRowCount, -aColCount, SheetInfo);
					if (r.LastRow<r.FirstRow) FList.RemoveAt(i);
				}

			}
		}

		internal void MoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
		{
			if ((SheetInfo.InsSheet<0) || (SheetInfo.SourceFormulaSheet != SheetInfo.InsSheet)) return;

			int aCount=Count;
			for (int i=aCount-1; i>=0;i--)
			{
				THLinkRecord r=this[i];
				if (r.FirstRow >= CellRange.Top && r.LastRow <= CellRange.Bottom &&
					r.FirstCol >= CellRange.Left && r.LastCol <= CellRange.Right)
				{
					//Hyperlink data doesn't move when you insert/copy cells or sheets. It is a static text.
					r.Offset(NewRow - CellRange.Top, NewCol - CellRange.Left);
				}
				else
				{
					if (r.FirstRow >= NewRow && r.LastRow <= NewRow + CellRange.RowCount - 1 &&
						r.FirstCol >= NewCol && r.LastCol <= NewCol + CellRange.ColCount - 1)
					{
						//Hyperlink data doesn't move when you insert/copy cells or sheets. It is a static text.
						FList.RemoveAt(i);
					}
				}
			}
		}

        internal void ClearRange(TXlsCellRange CellRange)
        {
            int aCount=Count;
            for (int i=aCount-1; i>=0;i--)
            {
                THLinkRecord r=this[i];
                if (r.FirstRow >= CellRange.Top && r.LastRow <= CellRange.Bottom&&
                    r.FirstCol >= CellRange.Left && r.LastCol <= CellRange.Right)
                    FList.RemoveAt(i);
            }
        }

    }


}
