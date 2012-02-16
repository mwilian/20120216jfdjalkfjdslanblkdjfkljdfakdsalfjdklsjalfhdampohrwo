using System;
using System.Diagnostics;
using FlexCel.Core;
using System.Text;
using System.Collections.Generic;

namespace FlexCel.XlsAdapter
{
	/// <summary>
	/// Holds one String entry. Implements the hash, taking in count the RTF and other data.
	/// </summary>
	internal class TSSTEntry: TExcelString, IComparable
	{
		private int FRefs;

		internal long AbsStreamPos;
		internal UInt16 RecordStreamPos;
		internal UInt32 PosInTable;

		internal void AddRef(){FRefs++;}
		internal void ReleaseRef(){FRefs--;}
		internal int Refs {get {return FRefs;}}

		internal TSSTEntry(ref TxBaseRecord aRecord, ref int ofs): base( TStrLenLength.is16bits, ref aRecord, ref ofs){}
		internal TSSTEntry(string s, TRTFRun[] RTFRuns, bool ForceWide): base(TStrLenLength.is16bits, s, RTFRuns, ForceWide){}

		private static void AddContinue(IDataStream DataStream, byte[] Buffer, ref int BufferPos, ref long BeginRecordPos, ref long TotalSize)
		{
			if (DataStream!=null)
			{
				Array.Copy(BitConverter.GetBytes((UInt16) (BufferPos-4)), 0, Buffer, 2, 2);  //Adapt the record size before writing it. 
                DataStream.WriteHeader((UInt16)(Buffer[0]+(Buffer[1]<<8)), (UInt16)(Buffer[2]+(Buffer[3]<<8)));
                DataStream.Write(Buffer, 4, BufferPos-4);
				BeginRecordPos=DataStream.Position;
				Array.Copy(BitConverter.GetBytes((UInt16) xlr.CONTINUE), 0, Buffer, 0, 2);
				Buffer[4]=0; Buffer[5]=0; //Clear the OptionFlags. 
			}
			TotalSize+=BufferPos;
			BufferPos=4;
		}

		private static void WriteArray(byte[]DataToWrite, IDataStream DataStream, byte[] Buffer, ref int BufferPos, ref long BeginRecordPos, ref long TotalSize)
		{
			int StPos=0;
			while (StPos<DataToWrite.Length)
			{
				int BytesLeft=((Buffer.Length-BufferPos) / 4) * 4; //RTF runs can only be split on full runs.
				int Chars=Math.Min(DataToWrite.Length-StPos, BytesLeft);
				Array.Copy(DataToWrite,StPos, Buffer, BufferPos, Chars);
				BufferPos+=Chars;
				StPos+=Chars;
				if (StPos<DataToWrite.Length)
				{
					AddContinue(DataStream, Buffer, ref BufferPos, ref BeginRecordPos, ref TotalSize);
				}

			}
		}
		internal void SaveToStream(IDataStream DataStream, byte[] Buffer, ref int BufferPos, ref long BeginRecordPos, ref long TotalSize)
		{
			//First, see if we can write the header of this string on the current record, or if we have to create a new one 
			
			int BytesLeft=Buffer.Length-BufferPos;
			if (BytesLeft<32) //12 is really the required, but we play it safe. Anyway, starting a new continue does no harm.
			{
				AddContinue(DataStream, Buffer,ref  BufferPos, ref BeginRecordPos, ref TotalSize);
				BytesLeft=Buffer.Length-BufferPos;
			}

			if (DataStream!=null)
			{
				AbsStreamPos=DataStream.Position+BufferPos;
				RecordStreamPos= (UInt16)(AbsStreamPos- BeginRecordPos);
			}

			Debug.Assert(BytesLeft>=32);
            if (Data.Length < 0 || Data.Length > FlxConsts.Max_StringLenInCell) XlsMessages.ThrowException(XlsErr.ErrStringTooLong, Data, FlxConsts.Max_StringLenInCell);
			if (DataStream!=null) Array.Copy(BitConverter.GetBytes((UInt16)(Data.Length)),0, Buffer, BufferPos, 2);
			BufferPos+=2;

			int OpFlagsPos=BufferPos;
			Buffer[BufferPos]=OptionFlags;
			BufferPos++;

			if (HasRichText)
			{  
				if (DataStream!=null) Array.Copy(BitConverter.GetBytes((UInt16)(RichTextFormats.Length>>2)),0, Buffer, BufferPos, 2);
				BufferPos+=2;
			}

			if (HasFarInfo)
			{
				if (DataStream!=null) Array.Copy(BitConverter.GetBytes((UInt32)(FarEastData.Length)),0, Buffer, BufferPos, 4);
				BufferPos+=4;
			}

			// Write the actual string. It might span multiple continues
			int StPos=0;
			while (StPos<Data.Length)  //If Data.Length==0, we won't write this string.
			{
				BytesLeft=Buffer.Length-BufferPos;
				//Find if we can compress the unicode on this part. 
				//If the number of chars we can write using compressed is bigger or equal than using uncompressed, we compress...
				int CharsCompressed=Math.Min(Data.Length-StPos, BytesLeft);
				int CharsUncompressed=Math.Min(Data.Length-StPos, BytesLeft/2);
				if (CharSize!=1) //if charsize=1, string is already compressed.
				{
					for (int i=0; i < CharsCompressed; i++)
					{
						if (Data[StPos+i]>'\u00FF')
						{
							CharsCompressed=i;
							break;
						}
					}
				}

				bool CanCompress= CharsCompressed>=CharsUncompressed;
				if (CanCompress)
				{
					byte b=0xFE;
					Buffer[OpFlagsPos]=(byte)(Buffer[OpFlagsPos] & b);
					if (DataStream!=null)
					{
						for (int i=0; i< CharsCompressed; i++)
						{
							Buffer[BufferPos]=((byte)Data[StPos+i]);
							BufferPos++;
						}
					}
					else BufferPos+=CharsCompressed;

					StPos+=CharsCompressed;
					if (StPos<Data.Length)
					{
						AddContinue(DataStream, Buffer, ref BufferPos, ref BeginRecordPos, ref TotalSize);
						OpFlagsPos=BufferPos;
						BufferPos++;
					}
				}
				else
				{
					byte b=1;
					Buffer[OpFlagsPos]=(byte)(Buffer[OpFlagsPos] | b);
					if (DataStream!=null)
					{
						for (int i=0; i< CharsUncompressed; i++)
						{
							Array.Copy(BitConverter.GetBytes(Data[StPos+i]),0, Buffer, BufferPos, 2);
							BufferPos+=2;
						}
					} 
					else BufferPos+=CharsUncompressed*2;

					StPos+=CharsUncompressed;
					if (StPos<Data.Length)
					{
						AddContinue(DataStream, Buffer, ref BufferPos, ref BeginRecordPos, ref TotalSize);
						OpFlagsPos=BufferPos;
						BufferPos++;
					}				
				}
			}

			if (HasRichText)
				WriteArray(RichTextFormats, DataStream, Buffer, ref BufferPos, ref BeginRecordPos, ref TotalSize);

			if (HasFarInfo) 
				WriteArray(FarEastData, DataStream, Buffer, ref BufferPos, ref BeginRecordPos, ref TotalSize);
		}

        #region IComparable Members

        public int CompareTo(object obj)
        {
            //Only of use in Testing.
            TSSTEntry x = obj as TSSTEntry;
            if (x == null) return -1;
            int Result = OptionFlags.CompareTo(x.OptionFlags);
            if (Result != 0) return Result;
            Result = Data.CompareTo(x.Data);
            if (Result != 0) return Result;
            Result = BitOps.CompareMemOrdinal(RichTextFormats, x.RichTextFormats);
            if (Result != 0) return Result;
            Result = BitOps.CompareMemOrdinal(x.FarEastData, x.FarEastData);
            if (Result != 0) return Result;
            return 0;           
        }

        #endregion
    }

	/// <summary>
	/// Shared string table implementation
	/// </summary>
	internal class TSST
	{
        internal TFutureStorage FutureStorage;

#if (FRAMEWORK20)
        private Dictionary<TSSTEntry, TSSTEntry> Data;

        private static Dictionary<TSSTEntry, TSSTEntry> CreateData()
        {
            return new Dictionary<TSSTEntry, TSSTEntry>();
        }
#else
		private Hashtable Data;

        private static Hashtable CreateData()
        {
            return new Hashtable();
        }
#endif
        /// <summary>
        /// Holds a tmp list with indices until all data has been loaded and we can sort.
        /// </summary>
		internal List<TSSTEntry> IndexData;
		internal TSST()
		{
            Data = CreateData();  //Created here, just in case File doesn't have SST.
        }

        internal void Clear()
        {
            Data.Clear();
        }

		internal void Load(TSSTRecord aSSTRecord)
		{
			int Ofs=8;
			TxBaseRecord TmpSSTRecord= aSSTRecord;
			Data=CreateData();
			IndexData=new List<TSSTEntry>((int)aSSTRecord.Count);
            
            //Excel will always have the right count here, but silly third poarties might not.
            //long aCount = aSSTRecord.Count;
            //for (int i=0; i<aCount;i++)
            while (TmpSSTRecord.Continue != null || Ofs < TmpSSTRecord.DataSize)
            {
                TSSTEntry Es= new TSSTEntry(ref TmpSSTRecord, ref Ofs);
                if (Data.ContainsKey(Es))  //An excel file could have a repeated string.
                {
                    IndexData.Add(Data[Es]); 
                }
                else
                {
                    Data.Add(Es,Es);
                    IndexData.Add(Es);
                }
            }
			//We need IndexData until all the LABELSST records have been loaded
		}

        internal void LoadXml(TxSSTRecord aSSTRecord)
        {
            if (IndexData == null) IndexData = new List<TSSTEntry>();
            TSSTEntry es = new TSSTEntry(aSSTRecord.Text, aSSTRecord.RTFRuns, false);

            if (Data.ContainsKey(es))
            {
                IndexData.Add(Data[es]);
            }
            else
            {
                IndexData.Add(es);
                Data.Add(es, es);
            }
            //We need IndexData until all the LABELSST records have been loaded
        }

		internal void ClearIndexData()
		{
			IndexData=null;
		}

        internal void SaveToStream(IDataStream DataStream, TSaveData SaveData)
        {
            long BeginRecordPos = DataStream.Position;
            byte[] Buffer = new byte[XlsConsts.MaxRecordDataSize + 4];
            Array.Copy(BitConverter.GetBytes((UInt16)xlr.SST), 0, Buffer, 0, 2);

            bool Repeatable = SaveData.Repeatable;
            UInt32 TotalRefs;
            IEnumerator<KeyValuePair<TSSTEntry, TSSTEntry>> myEnumerator;
            TSSTEntry[] SortedEntries;
            PrepareToSave(Repeatable, out TotalRefs, out myEnumerator, out SortedEntries);

            Array.Copy(BitConverter.GetBytes(TotalRefs), 0, Buffer, 4, 4);
            Array.Copy(BitConverter.GetBytes((UInt32)Data.Count), 0, Buffer, 8, 4);

            int BufferPos = 4 + 8;
            long TotalSize = 0;
            if (Repeatable)
            {
                //REPEATABLEWRITES
                foreach (TSSTEntry Se in SortedEntries)
                {
                    Se.SaveToStream(DataStream, Buffer, ref BufferPos, ref BeginRecordPos, ref TotalSize);
                }
            }
            else
            {
                myEnumerator.Reset();
                while (myEnumerator.MoveNext())
                {
                    myEnumerator.Current.Key.SaveToStream(DataStream, Buffer, ref BufferPos, ref BeginRecordPos, ref TotalSize);
                }
            }

            //Flush the buffer.
            Array.Copy(BitConverter.GetBytes((UInt16)(BufferPos - 4)), 0, Buffer, 2, 2);  //Adapt the record size before writing it. 
            DataStream.WriteHeader((UInt16)(Buffer[0] + (Buffer[1] << 8)), (UInt16)(Buffer[2] + (Buffer[3] << 8)));
            DataStream.Write(Buffer, 4, BufferPos - 4);
            TotalSize += BufferPos;


            WriteExtSST(DataStream, Repeatable);

        }

        internal void PrepareToSave(bool Repeatable, out UInt32 TotalRefs, out IEnumerator<KeyValuePair<TSSTEntry, TSSTEntry>> myEnumerator, out TSSTEntry[] SortedEntries)
        {
            //Renum the items.  We need the order to serialize to disk
            UInt32 i = 0; TotalRefs = 0;

            myEnumerator = null;
            SortedEntries = null;
            if (Repeatable)
            {
                //REPEATABLEWRITES

                //Allow to create always the same file. It takes more time because it needs to sort the entries, and this is not needed, except for testing.
                SortedEntries = new TSSTEntry[Data.Count];
                Data.Keys.CopyTo(SortedEntries, 0);
                Array.Sort(SortedEntries);

                foreach (TSSTEntry Se in SortedEntries)
                {
                    Debug.Assert(Se.Refs > 0, "Refs should be >0");
                    Se.PosInTable = i;
                    TotalRefs += (UInt32)Se.Refs;
                    i++;
                }
            }
            else
            {
                myEnumerator = Data.GetEnumerator();
                while (myEnumerator.MoveNext())
                {
                    TSSTEntry Se = myEnumerator.Current.Key;
                    Debug.Assert(Se.Refs > 0, "Refs should be >0");
                    Se.PosInTable = i;
                    TotalRefs += (UInt32)Se.Refs;
                    i++;
                }
            }
        }

		internal void WriteExtSST(IDataStream DataStream, bool Repeatable)
		{
			// Calc number of strings per hash bucket
			UInt16 n=(UInt16) ((Data.Count / 128)+1);
			if (n<8) n=8;

			int nBuckets=0;
			if (Data.Count==0)nBuckets=0; else nBuckets= (Data.Count-1) / n + 1;

			DataStream.WriteHeader((UInt16)xlr.EXTSST, (UInt16)(2+8*nBuckets));
			DataStream.Write16(n);


            if (Repeatable)
            {
                //REPEATABLEWRITES
                TSSTEntry[] SortedEntries = new TSSTEntry[Data.Count];
                Data.Keys.CopyTo(SortedEntries, 0);
                Array.Sort(SortedEntries);
                int i = 0;
                while (i < SortedEntries.Length)
                {
                    TSSTEntry e = SortedEntries[i];
                    DataStream.Write32((UInt32)e.AbsStreamPos);
                    DataStream.Write32(e.RecordStreamPos);
                    i += n;
                }
            }
            else
            {
                Dictionary<TSSTEntry, TSSTEntry>.Enumerator myEnumerator = Data.GetEnumerator();
                while (myEnumerator.MoveNext())
                {
                    TSSTEntry e = myEnumerator.Current.Key;
                    DataStream.Write32((UInt32)e.AbsStreamPos);
                    DataStream.Write32(e.RecordStreamPos);
                    for (int i = 0; i < n - 1; i++)
                    {
                        if (!myEnumerator.MoveNext()) return;  //the if is needed to fix a bug in mono.
                    }
                }
            }
		}


		internal TSSTEntry AddString(string s, TRTFRun[] RTFRuns)
		{
			if (Data==null) Data= CreateData();  //this is for testing
			TSSTEntry es= new TSSTEntry(s , RTFRuns, false);
  
			if (Data.ContainsKey(es)) 
			{   
				es=((TSSTEntry)Data[es]);
				es.AddRef();
			}
			else 
			{
				es.AddRef();
				Data.Add(es, es);
			}
			return es;
		}

		internal void FixRefs()
		{
            List<TSSTEntry> KeysToRemove = new List<TSSTEntry>();

            Dictionary<TSSTEntry, TSSTEntry>.Enumerator myEnumerator = Data.GetEnumerator();
			while (myEnumerator.MoveNext() )
			{
                TSSTEntry en = myEnumerator.Current.Key;
                if ((en).Refs<=0) KeysToRemove.Add(en);
			}

			//Now do the actual remove
			for (int i=0;i<KeysToRemove.Count;i++)
				Data.Remove(KeysToRemove[i]);

		}

		private long ExtSSTRecordSize()  //This one is never continued
		{
			int n=Data.Count / 128+1;
			if (n<8) n=8;

			int nBuckets=0;
			if (Data.Count==0)nBuckets=0; else nBuckets= (Data.Count-1) / n + 1;
			return 2+8*nBuckets+ XlsConsts.SizeOfTRecordHeader;
		}

		//Simulates a write to know how much it takes.
		private long SSTRecordSize(bool Repeatable)
		{
			long BeginRecordPos=0;
			byte[] Buffer = new byte[XlsConsts.MaxRecordDataSize+4];
			int BufferPos=4+8;
			long TotalSize=0;


            if (Repeatable)
            {
                //REPEATABLEWRITES
                TSSTEntry[] SortedEntries = new TSSTEntry[Data.Count];
                Data.Keys.CopyTo(SortedEntries, 0);
                Array.Sort(SortedEntries);

                foreach (TSSTEntry Se in SortedEntries)
                {
                    Se.SaveToStream(null, Buffer, ref BufferPos, ref BeginRecordPos, ref TotalSize);
                }
            }
            else
            {
                Dictionary<TSSTEntry, TSSTEntry>.Enumerator myEnumerator = Data.GetEnumerator();
                while (myEnumerator.MoveNext())
                    myEnumerator.Current.Key.SaveToStream(null, Buffer, ref BufferPos, ref BeginRecordPos, ref TotalSize);
            }
	
			TotalSize+=BufferPos;
			return TotalSize;
		}

		internal long TotalSize(bool Repeatable)
		{
			return SSTRecordSize(Repeatable)+ ExtSSTRecordSize();
		}
		internal int Count{get {return Data.Count;}}


		internal TSSTEntry this[TSSTEntry es] {get{return (TSSTEntry)Data[es];}}

        internal void AddFutureStorage(TFutureStorageRecord R)
        {
            TFutureStorage.Add(ref FutureStorage, R);
        }
    }

	/// <summary>
	/// SST record. We are only interested in reading it. The SST class will write the new one.
	/// </summary>
	internal class TSSTRecord: TxBaseRecord
	{
		internal TSSTRecord(int aId, byte[] aData): base(aId, aData){}
		internal long Count {get{ return BitOps.GetCardinal(Data,4);}}

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            WorkbookLoader.RecordLoader.SST.Load(this);
        }
	}

	/// <summary>
	/// LABELSST. 
	/// </summary>
	internal class TLabelSSTRecord: TCellRecord
	{
		private TSSTEntry pSSTEntry;
		private TSST SST;
        private IFlexCelFontList FontList;

        internal override object GetValue(ICellList Cells)
        {
            if ((pSSTEntry != null) && (pSSTEntry.RichTextFormats != null) && (pSSTEntry.RichTextFormats.Length > 0))
                return AsRichString;
            else
                return AsString;
        }

		protected override TBaseRecord DoCopyTo(TSheetInfo SheetInfo) 
		{
			TLabelSSTRecord Result= (TLabelSSTRecord)base.DoCopyTo(SheetInfo);
			Result.SST = SST;
            Result.FontList = FontList;
			Result.pSSTEntry = pSSTEntry;
			pSSTEntry.AddRef();
			return Result;
		}


        internal TLabelSSTRecord(int aId, byte[] aData, TSST aSST, IFlexCelFontList aFontList, TBiff8XFMap XFMap)
            : base(aId, aData, XFMap)
        {
            AttachToSST(BitOps.GetCardinal(aData, 6), aSST, aFontList);
        }

        internal TLabelSSTRecord(int aCol, int aXF, long aPos, TSST aSST, IFlexCelFontList aFontList)
            : base((int)xlr.LABELSST, aCol, aXF)
        {
            AttachToSST(aPos, aSST, aFontList);
        }

        internal TLabelSSTRecord(int aCol, int aXF, TSST aSST, IFlexCelFontList aFontList, object Value)
            : base((int)xlr.LABELSST, aCol, aXF)
        {
            SST = aSST;
            FontList = aFontList;
            pSSTEntry = null;
            AsString = FlxConvert.ToStringWithArrays(Value);
        }

        internal TLabelSSTRecord(int aCol, int aXF, TSST aSST, IFlexCelFontList aFontList, TRichString Value)
            : base((int)xlr.LABELSST, aCol, aXF)
        {
            SST = aSST;
            FontList = aFontList;
            pSSTEntry = null;
            AsRichString = Value;
        }

        internal static TLabelSSTRecord CreateFromOtherString(int aCol, int aXF, TSST aSST, IFlexCelFontList aFontList, object Value)
        {
            TRichString rs = Value as TRichString;
            if (rs != null && rs.RTFRunCount > 0)
            {
                return new TLabelSSTRecord(aCol, aXF, aSST, aFontList, rs);
            }
            else
                return new TLabelSSTRecord(aCol, aXF, aSST, aFontList, FlxConvert.ToStringWithArrays(Value));
        }

        internal override void Destroy()
        {
            pSSTEntry.ReleaseRef();
            base.Destroy ();
        }

		/// <summary>
		/// Should be called before we release SST.IndexData
		/// </summary>
		/// <param name="aSST"></param>
		/// <param name="index"></param>
		/// <param name="aFontList"></param>
		private void AttachToSST(long index, TSST aSST, IFlexCelFontList aFontList)
		{
            if (aSST == null || aFontList == null || index >= aSST.IndexData.Count) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
            SST=aSST;
			pSSTEntry= SST.IndexData[(int)index];
			pSSTEntry.AddRef();
            FontList=aFontList;
		}

        private uint PosInTable
        {
            get
            {
                return pSSTEntry.PosInTable;
            }
        }

		internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
		{
			base.SaveToStream(Workbook, SaveData, Row);
            Workbook.Write32(pSSTEntry.PosInTable);
		}

        internal override int TotalSizeNoHeaders()
        {
            return base.TotalSizeNoHeaders()+4;
        }


		internal string AsString
		{ 
			get
			{
				return pSSTEntry.Data;
			} 
			set 
			{   
				TSSTEntry OldEntry=pSSTEntry;
				pSSTEntry=SST.AddString(value, null);
				if (OldEntry!=null) OldEntry.ReleaseRef();
			}
		}

        internal TRTFRun[] AdaptFontList(TRichString Source)
        {
            if (Source==null || Source.RTFRunCount==0) return new TRTFRun[0];
            TRTFRun[] Result = new TRTFRun[Source.RTFRunCount];
            for (int i=0;i<Source.RTFRunCount;i++)
            {
                Result[i].FirstChar=Source.RTFRun(i).FirstChar;
                Result[i].FontIndex=FontList.AddFont(Source.GetFont(Source.RTFRun(i).FontIndex));
            }
            return Result;
        }

		internal TRichString AsRichString
		{ 
			get
			{
				return new TRichString(pSSTEntry.Data, pSSTEntry.RichTextFormats, FontList);
			} 
			set 
			{   
				TSSTEntry OldEntry=pSSTEntry;
				pSSTEntry=SST.AddString(value.Value, AdaptFontList(value));
				if (OldEntry!=null) OldEntry.ReleaseRef();
			}
		}

        internal static void SaveToPxl(TPxlStream PxlStream, int Row, int Col, int XF, string LabelValue, TPxlSaveData SaveData)
        {
            PxlStream.WriteByte((byte)pxl.LABEL);
            PxlStream.Write16((UInt16) Row);
            PxlStream.WriteByte((byte) Col);
            PxlStream.Write16(SaveData.GetBiff8FromCellXF(XF));
            if (LabelValue == null) LabelValue = String.Empty;
            if (LabelValue.Length > 255) LabelValue = LabelValue.Substring(0, 255);
            PxlStream.WriteString16(LabelValue);
        }
       
        internal override void SaveToPxl(TPxlStream PxlStream, int Row, TPxlSaveData SaveData)
        {
            base.SaveToPxl (PxlStream, Row, SaveData);
            if (!PxlRecordIsValid(Row)) return;
            SaveToPxl(PxlStream, Row, Col, XF, AsString, SaveData);
        }

#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
        internal override void SaveToXlsx(TOpenXmlWriter DataStream, int Row, TCellList CellList, bool Dates1904)
        {
            DataStream.WriteAtt("t", "s");
            DataStream.WriteElement("v", PosInTable);
        }
#endif

    }

	/// <summary>
	/// LABEL. This record is deprecated in BIFF8 and should not be written.
	/// </summary>
	internal class TLabelRecord: TCellRecord
	{
        private byte[] Data;

        internal TLabelRecord(int aId, byte[] aData, TBiff8XFMap XFMap) : base(aId, aData, XFMap) { Data = aData; }
		internal override object GetValue(ICellList Cells)
		{
			TxBaseRecord MySelf=new TxBaseRecord(Id, Data);
			int MyOfs=6;
			TExcelString XS=new TExcelString(TStrLenLength.is16bits, ref MySelf, ref MyOfs);
			return XS.Data;
		}
		//We don't implement writing value to a label, as it is deprecated. All writing should go to a LabelSST

        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
        {
            base.SaveToStream(Workbook, SaveData, Row);
            Workbook.Write(Data, 6, Data.Length-6);
        }

        internal override int TotalSizeNoHeaders()
        {
            return base.TotalSizeNoHeaders()+Data.Length-6;
        }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            //Excel 2010 will complain if this record exists.
            TLabelSSTRecord lsst = new TLabelSSTRecord(Col, XF, RecordLoader.SST, RecordLoader.FontList, GetValue(ws.Cells.CellList));

            ws.Cells.AddCell(lsst, rRow, RecordLoader.VirtualReader);
        }

#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
        internal override void SaveToXlsx(TOpenXmlWriter DataStream, int Row, TCellList CellList, bool Dates1904)
        {
            TxBaseRecord MySelf = new TxBaseRecord(Id, Data);
            int MyOfs = 6;
            TExcelString XS = new TExcelString(TStrLenLength.is16bits, ref MySelf, ref MyOfs);

            DataStream.WriteAtt("t", "inlineStr");
            DataStream.WriteStartElement("is");
            DataStream.WriteRichText(XS, CellList.Workbook);
            DataStream.WriteEndElement();
        }
#endif
    }

    /// <summary>
    /// Used to load from xlsx
    /// </summary>
    internal class TxLabelRecord : TCellRecord
    {
        string Value;

        internal TxLabelRecord(string aValue, int aCol, int aXF)
            : base((int)xlr.LABEL, aCol, aXF)
        {
            Value = aValue;
        }

        internal override object GetValue(ICellList Cells)
        {
            return Value;
        }

#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
        internal override void SaveToXlsx(TOpenXmlWriter DataStream, int Row, TCellList CellList, bool Dates1904)
        {
            FlxMessages.ThrowException(FlxErr.ErrInternal);
        }
#endif
    }

	/// <summary>
	/// RSTRING. This record is deprecated in BIFF8 and should not be written. However, it is used when pasting from the clipboard.
	/// </summary>
	internal class TRStringRecord: TCellRecord
	{
		private byte[] Data;

        internal TRStringRecord(int aId, byte[] aData, TBiff8XFMap XFMap) : base(aId, aData, XFMap) { Data = aData; }

		internal override object GetValue(ICellList Cells)
		{
			TxBaseRecord MySelf=new TxBaseRecord(Id, Data);
			int MyOfs=6;
			TExcelString XS=new TExcelString(TStrLenLength.is16bits, ref MySelf, ref MyOfs);

			byte[] b = new byte[2];
			BitOps.ReadMem(ref MySelf, ref MyOfs, b); 
			TRTFRun[] RTFRuns = new TRTFRun[BitOps.GetWord(b, 0)];

			for (int i = 0; i< RTFRuns.Length; i++)
			{
				BitOps.ReadMem(ref MySelf, ref MyOfs, b); 
				RTFRuns[i].FirstChar = BitOps.GetWord(b, 0);
				BitOps.ReadMem(ref MySelf, ref MyOfs, b); 
				RTFRuns[i].FontIndex = BitOps.GetWord(b, 0);
			}

			return new TRichString(XS.Data, RTFRuns, Cells.Workbook);
		}
		//We don't implement writing value to a label, as it is deprecated. All writing should go to a LabelSST

		internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
		{
			base.SaveToStream(Workbook, SaveData, Row);
			Workbook.Write(Data, 6, Data.Length-6);
		}

		internal override int TotalSizeNoHeaders()
		{
			return base.TotalSizeNoHeaders()+Data.Length-6;
        }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            //Excel 2010 will complain if this record exists.
            TLabelSSTRecord lsst = new TLabelSSTRecord(Col, XF, RecordLoader.SST, RecordLoader.FontList, (TRichString)GetValue(ws.Cells.CellList));

            ws.Cells.AddCell(lsst, rRow, RecordLoader.VirtualReader);

        }

#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
        internal override void SaveToXlsx(TOpenXmlWriter DataStream, int Row, TCellList CellList, bool Dates1904)
        {
            TxBaseRecord MySelf = new TxBaseRecord(Id, Data);
            int MyOfs = 6;
            TExcelString XS = new TExcelString(TStrLenLength.is16bits, ref MySelf, ref MyOfs);

            DataStream.WriteAtt("t", "inlineStr");
            DataStream.WriteStartElement("is");
            DataStream.WriteRichText(XS, CellList.Workbook);
            DataStream.WriteEndElement();
        }
#endif

    }

}

