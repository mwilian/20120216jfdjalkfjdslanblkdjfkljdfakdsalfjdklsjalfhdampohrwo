using System;
using System.Text;
using FlexCel.Core;
using System.Globalization;
using System.Diagnostics;
using System.IO;

using System.Collections.Generic;

namespace FlexCel.XlsAdapter
{

	internal struct TSheetRange
	{
		internal int FirstSheet;
		internal int LastSheet;

		internal TSheetRange(int aFirstSheet, int aLastSheeet)
		{
			FirstSheet = aFirstSheet;
			LastSheet = aLastSheeet;
		}
	}

    /// <summary>
    /// Extern Name.
    /// </summary>
    internal class TExternNameRecord: TxBaseRecord
    {
        internal TExternNameRecord(int aId, byte[] aData): base(aId, aData){}

		internal static TExternNameRecord CreateAddin(string Name)
		{
			TExcelString xe = new TExcelString(TStrLenLength.is8bits, Name, null, false);
			byte[] pData = new byte[10 + xe.TotalSize()];
			xe.CopyToPtr(pData, 6);
			pData[pData.Length - 1- 3] = 0x02;
			pData[pData.Length - 1- 2] = 0x00;
			pData[pData.Length - 1- 1] = 0x1C;
			pData[pData.Length - 1- 0] = 0x17;

			return new TExternNameRecord((int)xlr.EXTERNNAME2, pData);
		}

		internal static TExternNameRecord CreateExternName(int SheetIndex, string Name)
		{
			TExcelString xe = new TExcelString(TStrLenLength.is8bits, Name, null, false);
			byte[] pData = new byte[8 + xe.TotalSize()];
			
			BitOps.SetWord(pData, 2, SheetIndex);
			
			xe.CopyToPtr(pData, 6);
			pData[pData.Length - 1] = 0x00; //We will not add any name definition, since it is external.
			pData[pData.Length - 1- 1] = 0x00;

			return new TExternNameRecord((int)xlr.EXTERNNAME2, pData);
		}

        internal static TExternNameRecord CreateOleLink(string Name, bool Icon, bool Advise, bool PreferPic)
        {
            UInt16 OptionFlags = (UInt16)
                (0 |
                (Advise ? 0x2 : 0) |
                (PreferPic ? 0x4 : 0) |
                0x0 |
                0x10 |
                (0x3FF << 5) |
                (Icon ? 0x8000 : 0)
                );

            TExcelString xe = new TExcelString(TStrLenLength.is8bits, Name, null, false);
            byte[] pData = new byte[6 + xe.TotalSize()];

            BitOps.SetWord(pData, 0, OptionFlags);

            xe.CopyToPtr(pData, 6);
            return new TExternNameRecord((int)xlr.EXTERNNAME2, pData);
        }

        internal static TExternNameRecord CreateDdeLink(string Name, bool Ole, bool Advise, bool PreferPic)
        {
            UInt16 OptionFlags = (UInt16)
                (0 |
                (Advise ? 0x2 : 0) |
                (PreferPic ? 0x4 : 0) |
                (Ole ? 0x8 : 0x0) |
                0x0 |
                (0x3FF << 5) |
                (0)
                );

            TExcelString xe = new TExcelString(TStrLenLength.is8bits, Name, null, false);
            byte[] pData = new byte[6 + xe.TotalSize()];

            BitOps.SetWord(pData, 0, OptionFlags);

            xe.CopyToPtr(pData, 6);
            return new TExternNameRecord((int)xlr.EXTERNNAME2, pData);
        }

        internal int OptionFlags
        {
            get
            {
                return GetWord(0);
            }
        }

        internal bool IsDdeLink
        {
            get
            {
                return (OptionFlags & 0x0010) == 0;
            }
        }

		internal int SheetIndexInOtherFile
		{
			get 
			{
				return GetWord(2);
			}
		}

		internal string Name
		{
			get
			{
				string St=null;
				long StSize=0;
				StrOps.GetSimpleString(false, Data, 7, true, NameLength, ref St, ref StSize);
				return St;
			}
		}

		internal int NameLength
		{
			get
			{
				return Data[6];
			}
		}

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.References.AddExternName(this);
        }
	}

    /// <summary>
    /// List of Extern Name records.
    /// </summary>
    internal class TExternNameRecordList: TBaseRecordList<TExternNameRecord>
    {
    }


    /// <summary>
    /// A Supbook with all the ExternNames linked to it.
    /// </summary>
    internal class TSupBookRecord  :TxBaseRecord
    {
        internal TFutureStorage FutureStorage;
        internal TExternNameRecordList FExternNameList;

        internal TSupBookRecord(int aId, byte[] aData): base(aId, aData)
        {
            FExternNameList = new TExternNameRecordList();
        }
                   
        /// <summary>
        /// Creates an empty local supbook.
        /// </summary>
        internal TSupBookRecord(int SheetCount): this((int)xlr.SUPBOOK, null)
        {
            byte[] sc= BitConverter.GetBytes((UInt16)SheetCount);
            byte[] bData={sc[0],sc[1], 0x01, 0x04};
            Data=bData;
        }

		/// <summary>
		/// Creates an empty addin supbook.
		/// </summary>
		internal static TSupBookRecord CreateAddin()
		{
			return new TSupBookRecord((int)xlr.SUPBOOK, new byte[] {0x01, 0x00, 0x01, 0x3A});
		}

		/// <summary>
		/// Creates an empty external file supbook.
		/// </summary>
		internal static TSupBookRecord CreateExternalRef(string FileName, string FirstSheetName)
		{
			TExcelString FileData = new TExcelString(TStrLenLength.is16bits, EncodeFileName(FileName), null, false);

			bool SheetIsNull = FirstSheetName == null || FirstSheetName.Length == 0;
			TExcelString FirstSheetData = null;
			int FirstSheetDataSize = 0;
			if (!SheetIsNull)
			{
				FirstSheetName = TSheetNameList.MakeValidSheetName(FirstSheetName);
				FirstSheetData = new TExcelString(TStrLenLength.is16bits, FirstSheetName, null, false);
				FirstSheetDataSize = FirstSheetData.TotalSize();
			}

			byte[] ResultData = new byte[2 + FileData.TotalSize() + FirstSheetDataSize];
			if (SheetIsNull) BitOps.SetWord(ResultData, 0, 0); else BitOps.SetWord(ResultData, 0, 1);
            FileData.CopyToPtr(ResultData, 2);
			if (FirstSheetData != null) FirstSheetData.CopyToPtr(ResultData, 2 + FileData.TotalSize());
			return new TSupBookRecord((int)xlr.SUPBOOK, ResultData);
		}

        internal static TSupBookRecord CreateOleOrDdeLink(string OleApp, string FileName)
        {
            TExcelString OleLink = new TExcelString(TStrLenLength.is16bits, OleApp + (char)3 + EncodeFileName(FileName), null, false);

            byte[] ResultData = new byte[2 + OleLink.TotalSize()];
            OleLink.CopyToPtr(ResultData, 2);
            return new TSupBookRecord((int)xlr.SUPBOOK, ResultData);
        }

        internal bool IsLocal
        {
            get
            {
                return (Data.Length == 4) && (GetWord (2)== 0x0401);
            }
        }

        internal bool IsAddin
        {
            get
            {
                return (Data.Length == 4) && (GetWord(2) == 0x3A01);
            }
        }

        internal string OleOrDdeLink
        {
            get
            {
                int cch = GetWord(2);
                if (cch < 1 || cch > 0xff) return null;
                if (GetWord(0) != 0) return null;

                string Result = null; long ResultSize = 0;
                StrOps.GetSimpleString(true, Data, 4, true, cch, ref Result, ref ResultSize);

                if (Result == null || Result.Length < 1 || Result[0] < (char)3) return null; //DDE links are not encoded. If it is encoded (starts with 1) is a filename.
                int pos3 = Result.IndexOf((char)3);
                if (pos3 < 0) return Result;

                string ProgId = Result.Substring(0, pos3);
                string FileName = string.Empty;
                if (pos3 + 1 < Result.Length)
                {
                    FileName = Result.Substring(pos3 + 1);
                    if (FileName.Length > 1)
                    {
                        switch ((int)FileName[0])
                        {
                            case 0:
                            case 2: FileName = String.Empty; break;

                            case 1: FileName = DecodeFileName(FileName.Substring(1)); break;
                        }
                    }
                }
                return ProgId + (char)3 + FileName;
            }
        }

        internal void InsertSheets(int SheetCount)
        {
            if (!IsLocal) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
            BitOps.IncWord(Data, 0, SheetCount, FlxConsts.Max_Sheets, XlsErr.ErrTooManySheets);
        }


		private static string EncodeFileName(string s)
		{
            if (string.Equals(s, TFormulaMessages.ErrString(TFlxFormulaErrorValue.ErrRef), StringComparison.Ordinal)) return s;
			StringBuilder Result = new StringBuilder();
			Result.Append((char)1);
            if (s == null || s.Trim().Length == 0) return Convert.ToString((char)0); //chEmpty

			string filename = Path.GetFileName(s);
			string filepath = Path.GetDirectoryName(s);

			if (filepath == null || filepath.Trim().Length == 0) return (char)1 + filename;

			filepath = Path.Combine(filepath, filename);
			int SourcePos = 0;
			if (filepath.StartsWith("\\\\"))
			{
				Result.Append((char) 1);
				Result.Append('@');
				SourcePos += 2;
			} 
			else
				if (filepath.Length >= 2 && filepath[1] == ':')
			{
				Result.Append((char) 1);
				Result.Append(filepath[0]);
				SourcePos += 2;

				if (filepath.Length >= 3 && filepath[2] == '\\') SourcePos++;
			}
			else
				if (filepath.Length >= 1 && filepath[0] == '\\')
			{
				Result.Append((char) 2);
				SourcePos += 1;

			}

			filepath = filepath.Substring(SourcePos);
			filepath = filepath.Replace("..\\", Convert.ToString((char)4));
			filepath = filepath.Replace("\\", Convert.ToString((char)3));

			Result.Append(filepath);
			return Result.ToString();
		}

        private static string DecodeFileName(string s)
        {
            StringBuilder Result = new StringBuilder(s.Length * 2); 
			int i=0;
            while (i< s.Length)
            {
				if (((int)s[i])==1)
				{
					i++;
					if (s[i]=='@')
					{
						Result.Append('\\');
						Result.Append('\\');
					}
					else
					{
						Result.Append(s[i]);
						Result.Append(':');
						Result.Append('\\');
					} 
				}
				else
					if (((int)s[i])==2)
				{
					Result.Append('\\');
				} 
				else
					if (((int)s[i])==3)
				{
					Result.Append('\\');
				} 
				else
					if (((int)s[i])==4)
				{
					Result.Append('.');
					Result.Append('.');
					Result.Append('\\');
				}
				else
					Result.Append(s[i]);

				i++;
			}

            return Result.ToString();
        }

        internal string BookName()
        {
			if (IsLocal || IsAddin) return String.Empty;
            TxBaseRecord MySelf= this;
            int MyPos=2;
            TExcelString Xs=new TExcelString(TStrLenLength.is16bits, ref MySelf, ref MyPos);
            string Result=Xs.Data;
            if (Result.Length > 0)
            {
                if ((int)Result[0] == 0) Result = String.Empty;
                else   //chEmpty
                    if ((int)Result[0] == 1) Result = DecodeFileName(Result.Substring(1));
                    else  //chEncode
                        if ((int)Result[0] == 2) Result = String.Empty;  //chSelf
            }
            return Result;
        }

        internal void SetBookName(string FileName)
        {
            if (string.IsNullOrEmpty(FileName)) XlsMessages.ThrowException(XlsErr.ErrStringEmpty, "SetLink");
            
            FileName = EncodeFileName(FileName);
            if (FileName.Length > 0xFF) XlsMessages.ThrowException(XlsErr.ErrStringTooLong, Data, 0xFF);

            if (IsLocal || IsAddin) return;
            if (Data.Length <= 4) return;
            int OldFileNameLen = GetWord(2);
            if (OldFileNameLen < 0x01 || OldFileNameLen > 0xFF) return;

            int OldByteLen = (int)StrOps.GetStrLen(true, Data, 2, false, 0);

            TExcelString FileData = new TExcelString(TStrLenLength.is16bits, FileName, null, false);
            int NewByteLen = FileData.TotalSize();

            byte[] ResultData = new byte[Data.Length - OldByteLen + NewByteLen];
            Array.Copy(Data, 0, ResultData, 0, 2);
            FileData.CopyToPtr(ResultData, 2);
            Array.Copy(Data, 2 + OldByteLen, ResultData, 2 + NewByteLen, Data.Length - OldByteLen - 2);

            Data = ResultData;
        }


        internal int SheetCount()
        {
            return GetWord(0);
        }

        internal string SheetName(int SheetIndex, TWorkbookGlobals Globals)
        {
            int n=GetWord(0);
            if ((SheetIndex<0) || (SheetIndex>=n)) //this might happen... on range references to another workbook
            {
                return TFormulaMessages.ErrString(TFlxFormulaErrorValue.ErrRef);
            }

            if (GetWord(2) == 0x0401) //current sheet
            {
                return Globals.GetSheetName(SheetIndex);
            }

            //A little slow... but it shouldn't be called much.
            //I don't think it justifies a cache.
            TxBaseRecord MySelf=this;
            int tPos=2;
            for (int i=0; i<= SheetIndex ;i++) //0 stands for the first unicode string, the book name.
            { 
                new TExcelString(TStrLenLength.is16bits, ref MySelf, ref tPos);
            }

            TExcelString Xs2=new TExcelString(TStrLenLength.is16bits, ref MySelf, ref tPos);
            return Xs2.Data;
        }

        #if (FRAMEWORK20)
        internal void AddExternalSheets(List<string> SheetNames)
        {
            if (SheetNames == null || SheetNames.Count == 0) return;
            int SheetDataSize = 0;
            TExcelString[] SheetData = null;

            SheetData = new TExcelString[SheetNames.Count];
            for (int i = 0; i < SheetNames.Count; i++)
            {
                string SheetName = TSheetNameList.MakeValidSheetName(SheetNames[i]);
                SheetData[i] = new TExcelString(TStrLenLength.is16bits, SheetName, null, false);
                SheetDataSize += SheetData[i].TotalSize();
            }

            byte[] ResultData = new byte[Data.Length + SheetDataSize];
            Array.Copy(Data, 0, ResultData, 0, Data.Length);
            BitOps.IncWord(ResultData, 0, SheetNames.Count, FlxConsts.Max_Sheets, XlsErr.ErrTooManySheets);
            int Pos = Data.Length;
            foreach (TExcelString xs in SheetData)
            {
                xs.CopyToPtr(ResultData, Pos);
                Pos += xs.TotalSize();
            }

            Data = ResultData;
        }
#endif

		internal int EnsureExternalSheet(string SheetName)
		{
			if (SheetName == null || SheetName.Length == 0) return 0xFFFE; //used by external names.

			SheetName = TSheetNameList.MakeValidSheetName(SheetName);
			Debug.Assert(GetWord(2) != 0x0401); //book can't be localsupbook.

			int n=GetWord(0);

			TxBaseRecord MySelf=this;
			int tPos=2;
			for (int i=0; i<= n ;i++) //0 stands for the first unicode string, the book name.
			{
				TExcelString CurrentSheetName = new TExcelString(TStrLenLength.is16bits, ref MySelf, ref tPos);
				if (i == 0) continue; //This is the book name.

				if (String.Equals(SheetName, CurrentSheetName.Data, StringComparison.InvariantCultureIgnoreCase))
				{
					return i - 1;
				}
			}

			//Couldn't find the name, add it.
			TExcelString NewSheet = new TExcelString(TStrLenLength.is16bits, SheetName, null, false);
			byte[] NewData = new byte[Data.Length + NewSheet.TotalSize()];
			Array.Copy(Data, 0, NewData, 0, Data.Length);
            NewSheet.CopyToPtr(NewData, Data.Length);			
			Data = NewData;
            BitOps.IncWord(Data, 0, 1, int.MaxValue, XlsErr.ErrTooManyEntries);

			return n;

		}

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.References.AddSupBook(this);
        }

        internal void AddExternName(TExternNameRecord ExternNameRecord)
        {
            FExternNameList.Add(ExternNameRecord);
        }

        internal void AddFutureStorage(TFutureStorageRecord R)
        {
            TFutureStorage.Add(ref FutureStorage, R);
        }


        #region TBaseRecord functionality
        protected override TBaseRecord DoCopyTo(TSheetInfo SheetInfo)
        {
            TSupBookRecord Result=(TSupBookRecord) base.DoCopyTo(SheetInfo);
            Result.FExternNameList= new TExternNameRecordList();
            Result.FExternNameList.CopyFrom(FExternNameList, SheetInfo);
            return Result;
        }

        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
        {
            base.SaveToStream(Workbook, SaveData, Row);
            FExternNameList.SaveToStream(Workbook, SaveData, Row);
        }
        internal override int TotalSize()
        {
            return base.TotalSize() + (int)FExternNameList.TotalSize;
        }
        internal override int TotalSizeNoHeaders ()
        {
            int Result = base.TotalSizeNoHeaders();
            for (int i=0; i< FExternNameList.Count;i++)
                Result=Result+ FExternNameList[i].TotalSizeNoHeaders();
            return Result;
        }

        #endregion
    }

    /// <summary>
    /// Extern Sheet.
    /// </summary>
    internal class TExternSheetRecord: TxBaseRecord
    {
        internal TExternSheetRecord(int aId, byte[] aData): base(aId, aData){}

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.References.AddExternRef(this);
        }
    }

    /// <summary>
    /// One external reference, parsed.
    /// </summary>
    internal class TExternRef
    {
        internal UInt16 SupBookRecord;
        internal int FirstSheet; 
        internal int LastSheet;
                    
        internal TExternRef(int aSupBookRecord, int aFirstSheet, int aLastSheet)
        {
            SupBookRecord=(UInt16)aSupBookRecord;
            FirstSheet=aFirstSheet;
            LastSheet=aLastSheet;
        }
    
        internal void SaveToStream(IDataStream DataStream, TSaveData SaveData)
        {
            DataStream.Write16(SupBookRecord);
            DataStream.Write16((UInt16)FirstSheet);
            DataStream.Write16((UInt16)LastSheet);
        }
    }

    /// <summary>
    /// List of SupBook records.
    /// </summary>
    internal class TSupBookRecordList : TBaseRecordList<TSupBookRecord>
    {
        internal override long TotalSize
        {
            get
            {
                long Result = 0;
                for (int i = 0; i < Count; i++)
                    Result += this[i].TotalSize();
                return Result;
            }
        }
    }
    

    internal class TExternRefList
    {
        protected List<TExternRef> FList;
        internal TExternRefList()
        {
            FList = new List<TExternRef>();
        }

        #region Generics
        internal int Add (TExternRef a)
        {
            FList.Add(a);
            return FList.Count - 1;
        }
        internal void Insert (int index, TExternRef a)
        {
            FList.Insert(index, a);
        }

        internal TExternRef this[int index] 
        {
            get {return FList[index];} 
            set {FList[index] = value;}
        }

        internal int Count
        {
            get {return FList.Count;}
        }

        internal void Clear()
        {
            FList.Clear();
        }
        #endregion

        internal void Load(TExternSheetRecord aRecord)
        {
            int n=BitOps.GetWord(aRecord.Data, 0);
            int aPos=2; TxBaseRecord MyRecord= aRecord;
            byte[] Index= new byte[2];
            byte[] Fs= new byte[2];
            byte[] Ls= new byte[2];
            for (int i=0; i< n;i++)
            {
                BitOps.ReadMem(ref MyRecord, ref aPos, Index);
                BitOps.ReadMem(ref MyRecord, ref aPos, Fs);
                BitOps.ReadMem(ref MyRecord, ref aPos, Ls);
                Add(new TExternRef(BitConverter.ToUInt16(Index,0),BitConverter.ToUInt16(Fs,0),BitConverter.ToUInt16(Ls,0)));
            }
        }

        internal void SaveToStream(IDataStream DataStream, TSaveData SaveData)
        {
            if (Count ==0)
            {
                //This will crash Excel 2010!!
                //DataStream.WriteHeader((UInt16)xlr.EXTERNSHEET, 2);
                //DataStream.Write16((UInt16)Count);
                return;
            }

            int Lines=(6* Count-1) / XlsConsts.MaxExternSheetDataSize;
            for (int i= 0; i<= Lines;i++)
            {
                int CountRecords=0;
                if (i<Lines) CountRecords = XlsConsts.MaxExternSheetDataSize / 6;
                else CountRecords=(((6*Count-1) % XlsConsts.MaxExternSheetDataSize) + 1) / 6 ;
                
                int RecordHeaderSize= CountRecords*6;
                int RecordHeaderId= (int)xlr.CONTINUE;
                if (i == 0)
                {
                    RecordHeaderId= (int)xlr.EXTERNSHEET;
                    RecordHeaderSize+=2;
                }

                DataStream.WriteHeader((UInt16)RecordHeaderId, (UInt16)RecordHeaderSize);
                if (i==0) DataStream.Write16((UInt16)Count);

                for (int k = i*(XlsConsts.MaxExternSheetDataSize / 6);k< i*(XlsConsts.MaxExternSheetDataSize / 6)+CountRecords;k++)
                    FList[k].SaveToStream(DataStream, SaveData);
            }
        }

        internal long TotalSize()
        {
            //Take in count Continues...
            if (Count==0) return 0; //2+ XlsConsts.SizeOfTRecordHeader; Would crash Excel 2010
            else
                return 2+ (((6* Count-1) / XlsConsts.MaxExternSheetDataSize)+1)* XlsConsts.SizeOfTRecordHeader  //header + continues
                    + 6*Count;
        }

        internal void InsertSheets(int BeforeSheet, int SheetCount, TSupBookRecordList SupBooks)
        {
			for (int i=0; i< Count;i++)
			{
                int SupBook = FList[i].SupBookRecord;
				if (SupBook >= 0 && SupBook < SupBooks.Count && SupBooks[SupBook].IsLocal)
				{
					//Handling of deleted references for Sheetcount<0
                    if ((FList[i].FirstSheet >= BeforeSheet) && (FList[i].FirstSheet < BeforeSheet - SheetCount))  // we will delete the reference
					{
                        FList[i].FirstSheet = 0xFFFF;
					}
                    if ((FList[i].LastSheet >= BeforeSheet) && (FList[i].LastSheet < BeforeSheet - SheetCount))  // we will delete the reference
					{
                        FList[i].LastSheet = 0xFFFF;
					}

                    if ((FList[i].FirstSheet < 0xFFFE) && (FList[i].FirstSheet >= BeforeSheet)) BitOps.IncWord(ref FList[i].FirstSheet, SheetCount, FlxConsts.Max_Sheets, XlsErr.ErrTooManySheets);
                    if ((FList[i].LastSheet < 0xFFFE) && (FList[i].LastSheet >= BeforeSheet)) BitOps.IncWord(ref FList[i].LastSheet, SheetCount, FlxConsts.Max_Sheets, XlsErr.ErrTooManySheets);
				}
			}
        }
    }

    internal class TReferences
    {
        private TSupBookRecordList FSupBooks;
        private TExternRefList FExternRefs;
        private int FLocalSupBook;
                      
        internal TReferences()
        {
            FSupBooks=new TSupBookRecordList();
            FExternRefs= new TExternRefList();
            FLocalSupBook=-1;
        }

        internal int ExternRefsCount {get { return FExternRefs.Count;}}

        internal long TotalSize()
        {
            long Result = FSupBooks.TotalSize;
            if (FSupBooks.Count > 0) Result += FExternRefs.TotalSize();
            return Result;
        }

        internal void Clear()
        {
            if (FSupBooks!=null) FSupBooks.Clear();
            if (FExternRefs!=null) FExternRefs.Clear();
            FLocalSupBook=-1;
        }

        internal TSupBookRecordList Supbooks { get { return FSupBooks; } }

        internal void SaveToStream(IDataStream DataStream, TSaveData SaveData)
        {
            FSupBooks.SaveToStream(DataStream, SaveData, 0);
            if (FSupBooks.Count > 0)
            {
                FExternRefs.SaveToStream(DataStream, SaveData);
            }
        }

        internal void AddSupBook(TSupBookRecord aRecord)
        {
            FSupBooks.Add(aRecord);
            if (aRecord.IsLocal) FLocalSupBook= FSupBooks.Count-1;
        }

        internal void AddExternRef(TExternSheetRecord aRecord)
        {
            FExternRefs.Load(aRecord);
        }

        internal void AddExternName(TExternNameRecord aRecord)
        {
            if (FSupBooks.Count <= 0) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
            FSupBooks[FSupBooks.Count - 1].AddExternName(aRecord);
        }

        internal void AddAddinExternalName(string Name, out int ExternSheetIndex, out int ExternNameIndex)
		{
            int AddinSupBook;
            AddinSupBook = EnsureAddinSupBook();
            ExternNameIndex = EnsureExternAddin(AddinSupBook, Name);
			ExternSheetIndex = FindExternSheet(AddinSupBook, 0xFFFE, 0XFFFE);
			if (ExternSheetIndex < 0)
			{
				ExternSheetIndex = FExternRefs.Add(new TExternRef(AddinSupBook, 0xFFFE, 0XFFFE));
			}
		}

        internal void InsertSheets(int BeforeSheet, int SheetCount)
        {
            FExternRefs.InsertSheets(BeforeSheet, SheetCount, FSupBooks);
            if (FLocalSupBook>=0) FSupBooks[FLocalSupBook].InsertSheets(SheetCount);
        }

		private bool IsLocal(int SupBook)
		{
			if (SupBook < 0) return true; //-1 is the local supbook when there is no supbook
			if (SupBook >= FSupBooks.Count) return false;
			return FSupBooks[SupBook].IsLocal;
		}

		private string BookName(int SupBook)
		{
			if (SupBook < 0) return null; //-1 is the local supbook when there is no supbook
			if (SupBook >= FSupBooks.Count) return null;
			return FSupBooks[SupBook].BookName();
		}

		internal bool IsLocalSheet(int SheetRef)
		{
            if (SheetRef >= FExternRefs.Count) XlsMessages.ThrowException(XlsErr.ErrXlsIndexOutBounds, SheetRef, "Sheet Reference", 0, FExternRefs.Count - 1);
			return IsLocal(FExternRefs[SheetRef].SupBookRecord);
		}

		private bool IsAddin(int SupBook)
		{
			if (SupBook < 0 || SupBook >= FSupBooks.Count) return false;
			return FSupBooks[SupBook].IsAddin;
		}

		internal bool IsAddinSheet(int SheetRef)
		{
			if (SheetRef>=FExternRefs.Count) XlsMessages.ThrowException(XlsErr.ErrXlsIndexOutBounds, SheetRef,"Sheet Reference",0,FExternRefs.Count-1);
			return IsAddin(FExternRefs[SheetRef].SupBookRecord);
		}

        internal int GetJustOneSheet(int SheetRef)
        {
            if (SheetRef>=FExternRefs.Count) XlsMessages.ThrowException(XlsErr.ErrXlsIndexOutBounds, SheetRef,"Sheet Reference",0,FExternRefs.Count-1);
            if ((IsLocal(FExternRefs[SheetRef].SupBookRecord)) &&
                (FExternRefs[SheetRef].FirstSheet == FExternRefs[SheetRef].LastSheet))

                return FExternRefs[SheetRef].FirstSheet; else return -1;
        }

		internal TSheetRange GetAllSheets(int SheetRef)
		{
			if (SheetRef>=FExternRefs.Count) XlsMessages.ThrowException(XlsErr.ErrXlsIndexOutBounds, SheetRef,"Sheet Reference",0,FExternRefs.Count-1);
			if (IsLocal(FExternRefs[SheetRef].SupBookRecord)) 
				return new TSheetRange(FExternRefs[SheetRef].FirstSheet, FExternRefs[SheetRef].LastSheet);

			return new TSheetRange(0, -1);
		}

		internal void EnsureLocalSupBook(int SheetCount)
		{
			if (FLocalSupBook<0)
			{
				AddSupBook(new TSupBookRecord(SheetCount));
			}
		}

		private int EnsureAddinSupBook()
		{
			for (int i = 0; i < FSupBooks.Count; i++)
			{
				if (IsAddin(i))
				{
					return i;
				}
			}

			AddSupBook(TSupBookRecord.CreateAddin());
			return FSupBooks.Count - 1;

		}

		private int EnsureSupBook(string FileName, string FirstSheetName, string LastSheetName, out int FirstSheet, out int LastSheet)
		{
			for (int i = 0; i < FSupBooks.Count; i++)
			{
				TSupBookRecord Book = FSupBooks[i];
				if (!Book.IsAddin && String.Equals(Book.BookName(), FileName, StringComparison.InvariantCultureIgnoreCase))
				{
					FirstSheet = Book.EnsureExternalSheet(FirstSheetName);
					LastSheet = Book.EnsureExternalSheet(LastSheetName);
					return i;
				}
			}

			TSupBookRecord NewBook = TSupBookRecord.CreateExternalRef(FileName, FirstSheetName);
			AddSupBook(NewBook);
			FirstSheet = NewBook.EnsureExternalSheet(LastSheetName); //might be 0xfffe if sheet is null, so we can't just set it to 0.
			LastSheet = NewBook.EnsureExternalSheet(LastSheetName); //can be 0 if FirstSheet = LastSheet, or 1 if they are different, or 0xfffe if sheetname is null.
			return FSupBooks.Count - 1;
		}

		private int EnsureExternAddin(int SupBook, string Name)
		{
			Debug.Assert(FSupBooks[SupBook].IsAddin);
			for (int i = 0; i < FSupBooks[SupBook].FExternNameList.Count; i++)
			{
				TExternNameRecord N = FSupBooks[SupBook].FExternNameList[i];
				if (String.Equals(N.Name, Name, StringComparison.InvariantCultureIgnoreCase)) return i;
			}

			FSupBooks[SupBook].FExternNameList.Add(TExternNameRecord.CreateAddin(Name.ToUpper(CultureInfo.InvariantCulture)));
            return FSupBooks[SupBook].FExternNameList.Count - 1;

		}

		internal int EnsureExternName(int ExternSheet, string Name) //SheetIndex = 0 is global.
		{
			int SheetIndex= FExternRefs[ExternSheet].FirstSheet + 1;
			int Sheet2= FExternRefs[ExternSheet].LastSheet + 1;
			Debug.Assert(SheetIndex == Sheet2);
			if (SheetIndex == 0xFFFF) SheetIndex = 0; //global name.

			int SupBook = FExternRefs[ExternSheet].SupBookRecord;

			Debug.Assert(!FSupBooks[SupBook].IsAddin);
			for (int i = 0; i < FSupBooks[SupBook].FExternNameList.Count; i++)
			{
				TExternNameRecord N = FSupBooks[SupBook].FExternNameList[i];
				if (N.SheetIndexInOtherFile == SheetIndex && String.Equals(N.Name, Name, StringComparison.InvariantCultureIgnoreCase)) return i;
			}

			FSupBooks[SupBook].FExternNameList.Add(TExternNameRecord.CreateExternName(SheetIndex, Name.ToUpper(CultureInfo.InvariantCulture)));
            return FSupBooks[SupBook].FExternNameList.Count - 1;

		}


		private int FindExternSheet(int SupBook, int Sheet1, int Sheet2)
		{
			for (int i=0; i< FExternRefs.Count;i++)
				if ((FExternRefs[i].SupBookRecord == SupBook) &&
					(FExternRefs[i].FirstSheet == Sheet1) &&
					(FExternRefs[i].LastSheet == Sheet2))
				{
					return i;
				}
			return -1;
		}

        private int FindExternName(int SupBook, TExternNameRecord ExtName)
        {
            TExternNameRecordList ExNameList = FSupBooks[SupBook].FExternNameList;

            for (int i = 0; i < ExNameList.Count; i++)
                if (BitOps.CompareMem(ExNameList[i].Data, ExtName.Data))
                {
                    return i;
                }
            return -1;
        }

        internal int AddSheet(int SheetCount, int Sheet)
        {
            return AddSheet(SheetCount, Sheet, Sheet);
        }

		internal int AddSheet(int SheetCount, int FirstSheet, int LastSheet)
		{      
			int exs = FindExternSheet(FLocalSupBook, FirstSheet, LastSheet);
			if (exs >= 0) return exs;

			//Ref doesn't exits...
			EnsureLocalSupBook(SheetCount);
			FExternRefs.Add(new TExternRef(FLocalSupBook, FirstSheet, LastSheet));
			return FExternRefs.Count-1;
		}

		internal int AddSheetFromExternalFile(string FileName, string FirstSheetName, string LastSheetName)
		{
			int FirstSheet; int LastSheet;

			int SupBook = EnsureSupBook(FileName, FirstSheetName, LastSheetName, out FirstSheet, out LastSheet);
			int exs = FindExternSheet(SupBook, FirstSheet, LastSheet);
			if (exs >= 0) return exs;

			//Ref doesn't exits...
			FExternRefs.Add(new TExternRef(SupBook, FirstSheet, LastSheet));
			return FExternRefs.Count-1;

		}

        internal int AddSheetFromXlsxFile(int SupBook, string FirstSheetName, string LastSheetName)
        {
            int FirstSheet = FSupBooks[SupBook].EnsureExternalSheet(FirstSheetName);
            int LastSheet = FSupBooks[SupBook].EnsureExternalSheet(LastSheetName);
            int exs = FindExternSheet(SupBook, FirstSheet, LastSheet);
            if (exs >= 0) return exs;

            //Ref doesn't exits...
            FExternRefs.Add(new TExternRef(SupBook, FirstSheet, LastSheet));
            return FExternRefs.Count - 1;

        }

		internal int CopySheet(int SourcePos, TSheetInfo SheetInfo)
		{
			TExternRef reference = SheetInfo.SourceReferences.FExternRefs[SourcePos];
			
			if (SheetInfo.SourceReferences.IsLocal(reference.SupBookRecord))  //convert this reference to the new file.
			{
                if (reference.FirstSheet == SheetInfo.SourceFormulaSheet && reference.LastSheet == SheetInfo.SourceFormulaSheet)
                {
                    return AddSheet(SheetInfo.DestGlobals.SheetCount, SheetInfo.DestFormulaSheet); //The reference is to the same sheet where the formula was. Convert it into a ref to the new sheet.
                }
				
				//Find sheetnames and see if it is possible to convert the reference.
				int FirstSheet = Math.Min(reference.FirstSheet, reference.LastSheet);
				int LastSheet = Math.Max(reference.FirstSheet, reference.LastSheet);
				if (LastSheet >= SheetInfo.SourceGlobals.SheetCount) return -1;
				string StartSheetName = SheetInfo.SourceGlobals.GetSheetName(FirstSheet);
				int found = -1;
				for (int i = 0; i < SheetInfo.DestGlobals.SheetCount; i++)
				{
					if (SheetInfo.DestGlobals.GetSheetName(i) == StartSheetName)
					{
						found = i;
						break;
					}
				}
				if (found < 0) return -1;
				int k = found + 1;
				for (int i = FirstSheet + 1; i <= LastSheet; i++)
				{
					if (k >= SheetInfo.DestGlobals.SheetCount || i >= SheetInfo.SourceGlobals.SheetCount ||  SheetInfo.DestGlobals.GetSheetName(k) != SheetInfo.SourceGlobals.GetSheetName(i)) return -1;
					k++;
				}
				//All sheets are named the same and in the same order, we can create an externsheet in the destination file.
                
				return AddSheet(SheetInfo.DestGlobals.SheetCount, found, found + LastSheet - FirstSheet);		
			}

			//Ref is already external...
            int supbook = CopySupBook(SheetInfo, reference);

			int exs = FindExternSheet(supbook, reference.FirstSheet, reference.LastSheet);
			if (exs >= 0) return exs;
			return FExternRefs.Add(new TExternRef(supbook, reference.FirstSheet, reference.LastSheet));  //reference is an external third book (not the source o the dest), so we do not have to chan+ge FirstSheet or LastSheet.
		}

        private int CopySupBook(TSheetInfo SheetInfo, TExternRef reference)
        {
            int supbook = -1;
            TSupBookRecord SourceSupbook = SheetInfo.SourceReferences.FSupBooks[reference.SupBookRecord];
            for (int i = 0; i < FSupBooks.Count; i++)
            {
                if (BitOps.CompareMem(FSupBooks[i].Data, SourceSupbook.Data))
                {
                    supbook = i;
                    break;
                }
            }

            if (supbook < 0)
            {
                TSupBookRecord sb = (TSupBookRecord)TSupBookRecord.Clone(SourceSupbook, SheetInfo);
                supbook = FSupBooks.Count;
                FSupBooks.Add(sb);
            }
            return supbook;
        }

		internal int GetSupBook(int ExternSheet)
		{
			return FExternRefs[ExternSheet].SupBookRecord;
		}

        internal int LocalSupBook { get { return FLocalSupBook; } }

        internal void GetSheetsFromExternSheet(int externSheet, out int Sheet1, out int Sheet2, out bool ExternalSheets, out string ExternBookName)
        {
			ExternalSheets = !IsLocal(FExternRefs[externSheet].SupBookRecord);
			ExternBookName = BookName(FExternRefs[externSheet].SupBookRecord);
            Sheet1= FExternRefs[externSheet].FirstSheet;
            Sheet2= FExternRefs[externSheet].LastSheet;
        }

		internal string GetName(int SheetRef, int NameIndex, TWorkbookGlobals Globals, bool WritingXlsx)
		{
			int idx = FLocalSupBook;
			if (SheetRef >=0)
			{
                if (SheetRef >= FExternRefs.Count) XlsMessages.ThrowException(XlsErr.ErrXlsIndexOutBounds, SheetRef, "Sheet Reference", 0, FExternRefs.Count - 1);
				idx=FExternRefs[SheetRef].SupBookRecord;
			}

			if (IsLocal(idx))
			{
				if (NameIndex < 0 || NameIndex >= Globals.Names.Count) XlsMessages.ThrowException(XlsErr.ErrXlsIndexOutBounds, NameIndex,"NameIndex",0,Globals.Names.Count - 1);
				string n = Globals.Names[NameIndex].Name;
                if (WritingXlsx)
                {
                    n = TXlsNamedRange.GetXlsxInternal(n);
                }

                return n;
			}

			if (idx < 0 || idx >= FSupBooks.Count) XlsMessages.ThrowException(XlsErr.ErrXlsIndexOutBounds, idx,"idx",0,FSupBooks.Count - 1);
			if (NameIndex < 0 || NameIndex >= FSupBooks[idx].FExternNameList.Count) XlsMessages.ThrowException(XlsErr.ErrXlsIndexOutBounds, NameIndex,"NameIndex",0,FSupBooks[idx].FExternNameList.Count - 1);

            string Name = FSupBooks[idx].FExternNameList[NameIndex].Name;
            bool IsInternal;
            if (!TXlsNamedRange.IsValidRangeName(Name, out IsInternal)) return TCellAddress.QuoteSheet(Name); else return Name; //a quoted name will raise an error, and so will an unquoted dde link
		}

		internal string GetSheetFromName(int ExternSheet, int NameIndex, TWorkbookGlobals Globals)
		{
			int idx = FLocalSupBook;
			if (ExternSheet >=0)
			{
				if (ExternSheet>=FExternRefs.Count) XlsMessages.ThrowException(XlsErr.ErrXlsIndexOutBounds, ExternSheet,"Sheet Reference",0,FExternRefs.Count - 1);
				idx=FExternRefs[ExternSheet].SupBookRecord;
			}

			if (IsLocal(idx))
			{
				if (NameIndex < 0 || NameIndex >= Globals.Names.Count) XlsMessages.ThrowException(XlsErr.ErrXlsIndexOutBounds, NameIndex,"NameIndex",0,Globals.Names.Count - 1);
				int SheetIndex = Globals.Names[NameIndex].RangeSheet; //already substracts 1
				if (SheetIndex < 0) return String.Empty; //Global name.
				return Globals.GetSheetName(SheetIndex);
			}

			if (idx < 0 || idx >= FSupBooks.Count) XlsMessages.ThrowException(XlsErr.ErrXlsIndexOutBounds, idx,"idx",0,FSupBooks.Count - 1);
			if (NameIndex < 0 || NameIndex >= FSupBooks[idx].FExternNameList.Count) XlsMessages.ThrowException(XlsErr.ErrXlsIndexOutBounds, NameIndex,"NameIndex",0,FSupBooks[idx].FExternNameList.Count - 1);

			int ShIndex = FSupBooks[idx].FExternNameList[NameIndex].SheetIndexInOtherFile - 1;
			if (ShIndex < 0) return String.Empty; //Global name. Not really needed, since SheetName below covers the case too.
			return FSupBooks[idx].SheetName(ShIndex, Globals);
		}


		internal string GetSheetName(int SheetRef, TWorkbookGlobals Globals, bool Writingxlsx)
		{
			if (SheetRef>=FExternRefs.Count) XlsMessages.ThrowException(XlsErr.ErrXlsIndexOutBounds, SheetRef,"Sheet Reference",0,FExternRefs.Count);
			int idx=FExternRefs[SheetRef].SupBookRecord;

			StringBuilder Sh1= new StringBuilder(FSupBooks[idx].SheetName(FExternRefs[SheetRef].FirstSheet, Globals));
			if (FExternRefs[SheetRef].FirstSheet!=FExternRefs[SheetRef].LastSheet)
				Sh1.Append(TFormulaMessages.TokenString(TFormulaToken.fmRangeSep)+FSupBooks[idx].SheetName(FExternRefs[SheetRef].LastSheet, Globals));

			return GetSheetName(idx, Sh1.ToString(), Globals, Writingxlsx);

		}

        internal string GetSheetName(int SupBook, String Sheets, TWorkbookGlobals Globals, bool WritingXlsx)
        {
            StringBuilder Result = new StringBuilder();
            string Quote = TFormulaMessages.TokenString(TFormulaToken.fmSingleQuote);
            if (!IsLocal(SupBook))
            {
                if (WritingXlsx)
                {
                    int sb = SupBook < FLocalSupBook ? SupBook + 1 : SupBook;  //sb is 1-based
                    if (!IsAddin(SupBook))
                    {
                        Result.Append(TFormulaMessages.TokenString(TFormulaToken.fmWorkbookOpen));
                        Result.Append(sb.ToString(CultureInfo.InvariantCulture));
                        Result.Append(TFormulaMessages.TokenString(TFormulaToken.fmWorkbookClose));
                    }
                    //Quote = String.Empty;
                }
                else
                {
                    Result.Append(FSupBooks[SupBook].BookName());
                    if (Sheets.Length > 0)
                    {
                        int Ld = Result.ToString().LastIndexOf(@"\");
                        if (Ld >= 0) Result.Insert(Ld + 1, TFormulaMessages.TokenString(TFormulaToken.fmWorkbookOpen)); else Result.Insert(0, TFormulaMessages.TokenString(TFormulaToken.fmWorkbookOpen));
                        Result.Append(TFormulaMessages.TokenString(TFormulaToken.fmWorkbookClose));
                    }
                }
            }
            else
            {
                if (WritingXlsx && string.IsNullOrEmpty(Sheets)) //When there is no sheet, we need to tel it is an external ref. Xlsx does it by adding a [0] ref here, like [0]Name. Others are like Sheet1!a1.
                {
                    Result.Append(TFormulaMessages.TokenString(TFormulaToken.fmWorkbookOpen));
                    Result.Append(0.ToString(CultureInfo.InvariantCulture));
                    Result.Append(TFormulaMessages.TokenString(TFormulaToken.fmWorkbookClose));
                }
            }
            Result.Append(Sheets);

            string ResultStr = Result.ToString();
            if (ResultStr.Length > 0 && Quote.Length > 0) ResultStr = ResultStr.Replace(Quote, "''");
            if (Result.Length > 0) return Quote + ResultStr + Quote + TFormulaMessages.TokenString(TFormulaToken.fmExternalRef); 
            else return String.Empty;
        }


		internal int CopyExternName(int ExternSheetIndex, ref int ExternNameIndex, TSheetInfo SheetInfo)
		{
			TExternRef reference = SheetInfo.SourceReferences.FExternRefs[ExternSheetIndex];
            
			TSupBookRecord SourceSupbook = SheetInfo.SourceReferences.FSupBooks[reference.SupBookRecord];
			TExternNameRecord SourceExternName = SourceSupbook.FExternNameList[ExternNameIndex - 1];
            
			int supbook = CopySupBook(SheetInfo, reference);

			int exn = FindExternName(supbook, SourceExternName);
            if (exn >= 0) ExternNameIndex = exn + 1;
            else
            {
                FSupBooks[supbook].FExternNameList.Add((TExternNameRecord)TExternNameRecord.Clone(SourceExternName, SheetInfo));
                ExternNameIndex = FSupBooks[supbook].FExternNameList.Count; //not -1
            }

			int exs = FindExternSheet(supbook, reference.FirstSheet, reference.LastSheet);
			if (exs >= 0) return exs;
			return FExternRefs.Add(new TExternRef(supbook, reference.FirstSheet, reference.LastSheet));  //reference is an external third book (not the source o the dest), so we do not have to chan+ge FirstSheet or LastSheet.

		}

        internal string GetExternName(int ExternSheetIndex, int ExternNameIndex, out string ExternBookName, out int SheetIndexInOtherFile)
        {
            TExternRef reference = FExternRefs[ExternSheetIndex];
            TSupBookRecord Supbook = FSupBooks[reference.SupBookRecord];
			ExternBookName = Supbook.BookName();
			SheetIndexInOtherFile = Supbook.FExternNameList[ExternNameIndex - 1].SheetIndexInOtherFile;
            return Supbook.FExternNameList[ExternNameIndex - 1].Name;
        }
    }
}
