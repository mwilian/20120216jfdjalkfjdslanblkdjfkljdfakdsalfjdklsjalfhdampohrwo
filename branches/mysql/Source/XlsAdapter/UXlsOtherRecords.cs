using System;
using System.Diagnostics;
using FlexCel.Core;

namespace FlexCel.XlsAdapter
{
	/// <summary>
	/// Begin of File. Every section on an XLS file should Start with a TBOFRecord and end
	/// with a <see cref="TEOFRecord"/> record. This record is never encrypted.
	/// </summary>
	internal class TBOFRecord : TxBaseRecord
	{
        static readonly byte[] WorkbookBof = new byte[]  { 0x00, 0x06, 0x05, 0x00, 0xAA, 0x1F, 0xCD, 0x07, 0xC9, 0x00, 0x01, 0x00, 0x06, 0x04, 0x00, 0x00 };

        internal TBOFRecord (int aId, byte[] aData): base(aId, aData)
		{
			  if (GetWord(0)!= (int)xlr.BofVersion) XlsMessages.ThrowException(XlsErr.ErrInvalidVersion);

		}

        internal static TBOFRecord CreateEmptyWorkbook(int XlsxBofVer)
        {
            //note that xlsx uses 5 for Excel 2010, biff uses 6 (!)
            byte[] aData = (byte[]) WorkbookBof.Clone();

            if (XlsxBofVer >= 5) aData[13] = 0x06;

            return new TBOFRecord((int)xlr.BOF, aData);
        }

        internal static TBOFRecord CreateEmptyWorksheet(TSheetType SheetType)
        {
            byte[] aData = (byte[])WorkbookBof.Clone();

            switch (SheetType)
            {
                case TSheetType.Worksheet:
                case TSheetType.Dialog:
                    BitOps.SetWord(aData, 2, 0x0010);
                    break;

                case TSheetType.Chart:
                    BitOps.SetWord(aData, 2, 0x0020);
                    break;

                case TSheetType.Macro:
                    BitOps.SetWord(aData, 2, 0x0040);
                    break;

                //default: //This is an unsupported sheet. We will save the bof everywhere else.
                //    XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
                //break;
            }
            return new TBOFRecord((int)xlr.BOF, aData);
        }


		internal int BOFType
		{
			get
			{
				return GetWord(2);
			}
		}

        internal int BiffVersion
        {
            get
            {
                return GetWord(4);
            }
        }

        internal TExcelFileFormat BiffFileFormat()
        {
            switch (Data[13])
            {
                case 0:
                case 1:
                case 2:
                case 3: return TExcelFileFormat.v2003;
                case 4: return TExcelFileFormat.v2007;

                default:
                    return TExcelFileFormat.v2010;
            }
        }

        private static readonly byte[] Excel2003Build = { 0x1C, 0x20, 0xCD, 0x07, 0xC9, 0xC0, 0x00, 0x00, 0x06 };
        private static readonly byte[] Excel2003Version = { 3 };
        private static readonly byte[] Excel2007Build = { 0xAA, 0x1F, 0xCD, 0x07, 0xC9, 0x00, 0x01, 0x00, 0x06 };
        private static readonly byte[] Excel2007Version = { 4 };

        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
        {
            if ((SaveData.ExcludedRecords & TExcludedRecords.Version) != 0) return; //Note that this will invalidate the size, but it doesn't matter as this is not saved for real use. We could write blanks here if we wanted to keep the offsets right.

            byte[] ExcelBuild; byte[] ExcelVersion;
            if (SaveData.BiffVersion == TXlsBiffVersion.Excel2003)
            {
                ExcelBuild = Excel2003Build;
                ExcelVersion = Excel2003Version;
            }
            else
            {
                ExcelBuild = Excel2007Build;
                ExcelVersion = Excel2007Version;
            }
			
            Workbook.WriteHeader((UInt16)Id, (UInt16) Data.Length);
			Workbook.WriteRaw(Data, 4);
			Workbook.WriteRaw(ExcelBuild, ExcelBuild.Length);
			
			Workbook.WriteRaw(ExcelVersion, 1);
			Workbook.WriteRaw(Data, 14, Data.Length - 14);
        }

        internal override void SaveToPxl(TPxlStream PxlStream, int Row, TPxlSaveData SaveData)
        {
            base.SaveToPxl (PxlStream, Row, SaveData);
            int bType = BOFType;
            Debug.Assert(bType == 0x05 || bType == 0x10, "Only streams valid on Pxl are Globals and Worksheets");
            PxlStream.WriteByte((byte) pxl.BOF);

            PxlStream.Write16(0x010F); //Pxl 2.0
            PxlStream.Write16((UInt16)bType);
        }


		internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
		{
			XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
		}

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
        }

    }

    /// <summary>
    /// End of File. Every section on an XLS file should Start with a <see cref="TBOFRecord"/>  and end
    /// with a TEOFRecord record
    /// </summary>
    internal class TEOFRecord : TxBaseRecord
    {
        internal TEOFRecord(): base((int)xlr.EOF, new byte[0])
        {
        }

        internal TEOFRecord (int aId, byte[] aData): base(aId, aData){}

		internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
		{
			ws.sEOF = this;
			Loader.Eof = true;
		}

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.sEOF = this;
        }

        internal override void SaveToPxl(TPxlStream PxlStream, int Row, TPxlSaveData SaveData)
        {
            base.SaveToPxl (PxlStream, Row, SaveData);
            PxlStream.WriteByte((byte)pxl.EOF);
        }
    }

	/// <summary>
	/// If this record is present, the file is an xlt template.
	/// </summary>
	internal class TTemplateRecord : TxBaseRecord
	{
		internal TTemplateRecord (int aId, byte[] aData): base(aId, aData){}

		internal static void SaveNewRecord(IDataStream Workbook)
		{
			Workbook.WriteHeader((UInt16)xlr.TEMPLATE, 0);
		}
		
		internal static int GetSize(bool IsTemplate)
		{
			if (IsTemplate) return XlsConsts.SizeOfTRecordHeader; else return 0;
		}

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.IsXltTemplate = true;
        }
	}


    /// <summary>
    /// Never encrypted
    /// </summary>
    internal class TInterfaceHdrRecord: TxBaseRecord
    {
        internal TInterfaceHdrRecord (int aId, byte[] aData): base(aId, aData){}

        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
        {
            Workbook.WriteHeader((UInt16)Id, (UInt16) Data.Length);
            Workbook.WriteRaw(Data, Data.Length);
        }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.LoadingInterfaceHdr = true;
            base.LoadIntoWorkbook(Globals, WorkbookLoader);
        }

    }

    /// <summary>
    /// End of interface section.
    /// </summary>
    internal class TInterfaceEndRecord: TxBaseRecord
    {
        internal TInterfaceEndRecord (int aId, byte[] aData): base(aId, aData){}

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            base.LoadIntoWorkbook(Globals, WorkbookLoader);
            Globals.LoadingInterfaceHdr = false;
        }
    }

	/// <summary>
	/// Text box.
	/// </summary>
	internal class TTXORecord : TxBaseRecord
	{
		internal TTXORecord (int aId, byte[] aData): base(aId, aData){}

		/// <summary>
		/// Creates a new empty txo.
		/// </summary>
		internal TTXORecord (): base((int)xlr.TXO, new byte[18])
		{
			SetWord(0, 0x212);
		}
	}


	/// <summary>
	/// Bound sheets. The file position is not encrypted.
	/// </summary>
	internal class TBoundSheetRecord: TxBaseRecord
	{
        internal string XlsxRelationshipId;

		internal TBoundSheetRecord(int aId, byte[] aData): base(aId, aData){}
		internal TBoundSheetRecord(int aOptionFlags, string aName): base((int)xlr.BOUNDSHEET, new byte[0])
		{
			TExcelString Xs= new TExcelString(TStrLenLength.is8bits, aName, null, false);
			Data=new byte[6 + Xs.TotalSize()];
			SetCardinal(0, 0);
			SetWord(4, aOptionFlags );
			Xs.CopyToPtr( Data, 6 );
		}

        internal TBoundSheetRecord(int aOptionFlags, string aName, string aXlsxRelationshipId)
            : this(aOptionFlags, aName)
        {
            XlsxRelationshipId = aXlsxRelationshipId;
        }


		internal int OptionFlags {get {return GetWord(4);} set{ SetWord(4,value);}}
        internal byte OptionFlags1 {get {return Data[4];} set{ Data[4]=value;}}
        internal byte OptionFlags2 {get {return Data[5];} set{ Data[5]=value;}}
		internal void SetOffset(long offset)
		{
			SetCardinal(0, offset);
		}

		internal string SheetName
		{
			get
			{
				int Ofs=6;
				TxBaseRecord Rec= this;
				TExcelString Xs = new TExcelString(TStrLenLength.is8bits, ref Rec, ref Ofs);
				return Xs.Data;
			}
			set
			{
				TExcelString Xs= new TExcelString(TStrLenLength.is8bits, value, null, false);
				byte[] NewData=new byte[6 + Xs.TotalSize()];
				Array.Copy(Data, 0, NewData, 0, 6);
				Data=NewData;
				Xs.CopyToPtr( Data, 6 );
			}
		}

        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
        {
            Workbook.WriteHeader((UInt16)Id, (UInt16) Data.Length);
            byte[] Pos=new byte[4]; 
            Array.Copy(Data,0,Pos,0,4);
            Workbook.WriteRaw( Pos, Pos.Length);
            Workbook.Write( Data, 4, Data.Length-4);
        }

        internal override void SaveToPxl(TPxlStream PxlStream, int Row, TPxlSaveData SaveData)
        {
            base.SaveToPxl (PxlStream, Row, SaveData);
            if (OptionFlags2 != 0x00) return; //Only worksheets.
            PxlStream.WriteByte((byte)pxl.BOUNDSHEET);
            PxlStream.WriteByte(0); //reserved
            PxlStream.WriteString8(SheetName);
        }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.BoundSheets.AddFromFile(this);
        }

	}

    /// <summary>
    /// Name for the VBA sheet asociated to a sheet. This record is REQUIRED if there are macros, and can't be repeated.
    /// </summary>
    internal class TCodeNameRecord: TxBaseRecord
    {
        internal TCodeNameRecord(int aId, byte[] aData): base(aId, aData){}
        internal TCodeNameRecord(string aName): base((int)xlr.CODENAME, new byte[0])
        {
            TExcelString Xs= new TExcelString(TStrLenLength.is16bits, aName, null, false);
            Data=new byte[Xs.TotalSize()];
            Xs.CopyToPtr( Data, 0 );
        }

        internal string SheetName
        {
            get
            {
                int Ofs=0;
                TxBaseRecord Rec= this;
                TExcelString Xs = new TExcelString(TStrLenLength.is16bits, ref Rec, ref Ofs);
                return Xs.Data;
            }
            set
                //Important: This method changes the size of the record without notifying it's parent list
                //It's necessary to adapt the Totalsize in the parent list.
            {
                TExcelString Xs= new TExcelString(TStrLenLength.is16bits, value, null, false);
                Data=new byte[Xs.TotalSize()];
                Xs.CopyToPtr( Data, 0 );
            }
        }

		internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
		{
			ws.CodeNameRecord = this;
		}

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.CodeNameRecord = this;
        }

    }

    /// <summary>
    /// When present, there are macros on the sheet.
    /// </summary>
    internal class TObProjRecord: TxBaseRecord
    {
        internal TObProjRecord(int aId, byte[] aData): base(aId, aData){}

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.ObjProj = this;
        }
    }

    /// <summary>
    /// No forms or class modules in the VBA stream. 
    /// </summary>
    internal class TObNoMacrosRecord : TxBaseRecord
    {
        internal TObNoMacrosRecord(int aId, byte[] aData)
            : base(aId, aData)
        {
        }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.ObNoMacros = this;
        }
    }

    /// <summary>
    /// Excel 9 Record. 
    /// </summary>
    internal class TExcel9FileRecord : TxBaseRecord
    {
        internal TExcel9FileRecord(int aId, byte[] aData)
            : base(aId, aData)
        {
        }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.Excel9File = this;
        }
    }

    internal class TDSFRecord : TNotStorableRecord
    {
        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            //Do nothing, should be ignored
        }

        internal static long StandardSize()
        {
            return XlsConsts.SizeOfTRecordHeader + 2;
        }

        internal static void SaveDSF(IDataStream DataStream)
        {
            DataStream.WriteHeader((UInt16)xlr.DSF, 2);
            DataStream.Write16(0); //we won't use double stream files, as the content in the other stream will probably be out of date.
        }
    }

    /// <summary>
    /// Size of the OLE object when Excel is an OLE server.
    /// </summary>
    internal class TOleObjectSizeRecord : TxBaseRecord
    {
        internal TOleObjectSizeRecord(int aId, byte[] aData) : base(aId, aData) { }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.OleObjectSize = this;
        }
    }

    /// <summary>
    /// Version of Excel that recalculated this file.
    /// </summary>
    internal class TRecalcIdRecord : TxBaseRecord
    {
        internal TRecalcIdRecord(int aId, byte[] aData) : base(aId, aData) { }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.RecalcId = this;
        }
    }

    /// <summary>
    /// Web publishing options.
    /// </summary>
    internal class TWebPubRecord : TxBaseRecord
    {
        internal TWebPubRecord(int aId, byte[] aData) : base(aId, aData) { }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.WebPub.Add(this);
        }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.WebPub.Add(this);
        }
    
        internal void ArrangeInsertRange(TXlsCellRange CellRange,int aRowCount,int aColCount,TSheetInfo SheetInfo)
        {
 	        //Should adapt the range.
        }
    
        internal void ArrangeMoveRange(TXlsCellRange CellRange,int NewRow,int NewCol,TSheetInfo SheetInfo)
        {
        }
    }

    /// <summary>
    /// Natural language formula that was deleted. Not saved by office 2007
    /// </summary>
    internal class TLelRecord : TxBaseRecord
    {
        internal TLelRecord(int aId, byte[] aData) : base(aId, aData) { }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.Lel.Add(this);
        }
    }

    /// <summary>
    /// Any FNGroup record. If later we need to specialize this they should derive drom this.
    /// </summary>
    internal class TFnGroupRecord : TxBaseRecord
    {
        internal TFnGroupRecord(int aId, byte[] aData) : base(aId, aData) { }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.FnGroups.Add(this);
        }
    }

    internal class TDocRouteRecord : TxBaseRecord
    {
        internal TDocRouteRecord(int aId, byte[] aData) : base(aId, aData) { }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.DocRoute.Add(this);
        }
    }

    internal class TRecipNameRecord : TxBaseRecord
    {
        internal TRecipNameRecord(int aId, byte[] aData) : base(aId, aData) { }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.DocRoute.Add(this);
        }
    }

    internal class TUserBViewRecord : TxBaseRecord
    {
        internal TUserBViewRecord(int aId, byte[] aData) : base(aId, aData) { }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.UserBView.Add(this);
        }

        internal Guid CustomView
        {
            get
            {
                byte[] bCustomView = new byte[16];
                Array.Copy(Data, 8, bCustomView, 0, 16);
                return new Guid(bCustomView);
            }
        }
    }

    internal class TMetaDataRecord : TxBaseRecord
    {
        internal TMetaDataRecord(int aId, byte[] aData) : base(aId, aData) { }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.MetaData.Add(this);
        }
    }

    internal class TRTDRecord : TxBaseRecord
    {
        internal TRTDRecord(int aId, byte[] aData) : base(aId, aData) { }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.RealTimeData.Add(this);
        }
    }

    internal class TTabIdRecord : TxBaseRecord
    {
        internal TTabIdRecord(int aId, byte[] aData) : base(aId, aData) { }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.BoundSheets.AddTabIdFromFile(this);
        }
    }

    internal class TDConnRecord : TxBaseRecord
    {
        internal TDConnRecord(int aId, byte[] aData) : base(aId, aData) { }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.DConn.Add(this);
        }
    }

    /// <summary>
    /// Web options.
    /// </summary>
    internal class TWOptRecord : TxBaseRecord
    {
        internal TWOptRecord(int aId, byte[] aData) : base(aId, aData) { }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.WOpt = this;
        }
    }

    /// <summary>
    /// Web options.
    /// </summary>
    internal class TBookExtRecord : TxBaseRecord
    {
        internal TBookExtRecord(int aId, byte[] aData) : base(aId, aData) { }

        /// <summary>
        /// This creates a biff12 bookext, with length 22.
        /// </summary>
        internal TBookExtRecord()
            : base ((int)xlr.BOOKEXT,
            new byte[] {0x63, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x16, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02, 0x00})
        {
        }
        
        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.BookExt = this;
        }

        private void sbit(int p, bool value)
        {
            if (value) SetWord(16, GetWord(16) | (1 << p));
            else SetWord(16, GetWord(16) & ~(1 << p));
        }

        private bool bit(int p)
        {
            return (GetWord(16) & (1 << p)) != 0;
        }

        private void sbitex(int byt, int p, bool value)
        {
            if (byt >= Data.Length) return;
            if (value) Data[byt] = (byte) (Data[byt] | (1 << p));
            else Data[byt] = (byte)(Data[byt] & ~(1 << p));
        }

        private bool bitex(int byt, int p)
        {
            if (byt >= Data.Length) return false;
            return (Data[byt] & (1 << p)) != 0;
        }

        internal bool DontAutoRecover { get { return bit(0); } set { sbit(0, value); } }
        internal bool HidePivotList   { get { return bit(1); } set { sbit(1, value); } }
        internal bool FilterPrivacy   { get { return bit(2); } set { sbit(2, value); } }
        internal bool EmbedFactoids   { get { return bit(3); } set { sbit(3, value); } }

        internal bool SavedDuringRecovery { get { return bit(6); } set { sbit(6, value); } }
        internal bool CreatedViaMinimalSave { get { return bit(7); } set { sbit(7, value); } }
        internal bool OpenedViaDataRecovery { get { return bit(8); } set { sbit(8, value); } }
        internal bool OpenedViaSafeLoad { get { return bit(9); } set { sbit(9, value); } }

        internal bool BuggedUserAboutSolution { get { return bitex(20, 0); } set { sbitex(20, 0, value); } }
        internal bool ShowInkAnnotation { get { return bitex(20, 1); } set { sbitex(20, 1, value); } }

        internal bool PublishedBookItems { get { return bitex(21, 1); } set { sbitex(21, 1, value); } }
        internal bool ShowPivotChartFilter { get { return bitex(21, 2); } set { sbitex(21, 2, value); } }


    }


	/// <summary>
	/// Base for all records representing ranges.
	/// </summary>
	internal class TRangeRecord: TxBaseRecord
	{
		internal TRangeRecord(int aId, byte[] aData): base(aId, aData){}
	}

    /// <summary>
    /// A Cell merging.
    /// </summary>
	internal class TCellMergingRecord: TRangeRecord
	{
		internal TCellMergingRecord(int aId, byte[] aData): base(aId, aData){}

		internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
		{
            TMergedCells a = new TMergedCells();
            ws.MergedCells.Add(a);
			a.LoadFromStream(RecordLoader, this);
		}

	}

	/// <summary>
	/// Data validation
	/// </summary>
	internal class TDValRecord: TxBaseRecord
	{
		internal TDValRecord(int aId, byte[] aData): base(aId, aData){}

		internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
		{
			ws.DataValidation.LoadDVal(this);
		}
	}

	/// <summary>
	/// Data validation
	/// </summary>
	internal class TDVRecord: TRangeRecord
	{
		internal TDVRecord(int aId, byte[] aData): base(aId, aData){}

		internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
		{
			ws.DataValidation[ws.DataValidation.Add(new TDataValidation())].LoadFromStream(RecordLoader, this);
		}

	}


	/// <summary>
	/// SheetExt (sheet tab color).
	/// </summary>
    internal class TSheetExtRecord : TBaseRecord
    {
        internal int Id;
        internal TExcelColor SheetColor;
        internal bool CondFmtCalc;
        internal bool NotPublished;

        internal TSheetExtRecord(int aId, byte[] aData)
        {
            Id = aId;
            int ci = (BitOps.GetWord(aData, 16) & 0x7F);
            SheetColor = TExcelColor.FromBiff8ColorIndex(ci);

            if (aData.Length < 0x028)
            {
                CondFmtCalc = true;
                NotPublished = false;
            }
            else
            {
                int ci2 = (BitOps.GetWord(aData, 20) & 0x7F);
                if (ci2 == ci) //If it isn't then this was changed by Excel 2003, and we need to keep the new Excel 2003 color.
                {
                    SheetColor = TExcelColor.FromBiff8(aData, 24, 32, 28, false);
                }

                CondFmtCalc = (aData[20] & 0x80) != 0;
                NotPublished = (aData[21] & 0x01) != 0;
            }
        }

        internal TSheetExtRecord(TExcelColor aColor)
        {
            Id = (int)xlr.SHEETEXT;
            SheetColor = aColor;
            CondFmtCalc = true;
            NotPublished = false;
        }

        internal TExcelColor GetTabColor(IFlexCelPalette xls)
        {
            return SheetColor;
        }

        internal void SetTabColor(TExcelColor aColor)
        {
            SheetColor = aColor;
        }

        protected override TBaseRecord DoCopyTo(TSheetInfo SheetInfo)
        {
            return (TBaseRecord)MemberwiseClone();
        }

        internal override int GetId
        {
            get { return Id; }
        }

        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
        {
            Workbook.WriteHeader((UInt16)Id, (UInt16)TotalSizeNoHeaders());

            Workbook.Write16((UInt16)Id); //FrtHeader
            Workbook.Write(new byte[10], 10);

            Workbook.Write32((UInt32)TotalSizeNoHeaders());


            int ic = SheetColor.GetBiff8ColorIndex(SaveData.Palette, TAutomaticColor.SheetTab);
            if (ic < 0x08 || ic > 0x3f) ic = 0x7f;
            Workbook.Write32((UInt32)ic);

            UInt32 ic2 = (UInt32)ic;
            if (CondFmtCalc) ic2 |= 0x80;
            if (NotPublished) ic2 |= 0x100;
            Workbook.Write32(ic2);

            switch (SheetColor.ColorType)
            {
                case TColorType.RGB:
                    Workbook.Write32(0x02);
                    UInt32 RGB = (UInt32)(SheetColor.RGB);
                    UInt32 BGR = 0xFF000000 | (RGB & 0x00FF00) | ((RGB & 0xFF0000) >> 16) | ((RGB & 0x0000FF) << 16);

                    Workbook.Write32(BGR);
                    break;
                case TColorType.Automatic:
                    Workbook.Write32(0x00);
                    Workbook.Write32(0x00);
                    break;
                case TColorType.Theme:
                    Workbook.Write32(0x03);
                    Workbook.Write32((UInt32)SheetColor.Theme);
                    break;
                case TColorType.Indexed:
                    Workbook.Write32(0x01);
                    Workbook.Write32((UInt32)ic);
                    break;
                default:
                    XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
                    break;
            }

            Workbook.Write(BitConverter.GetBytes(SheetColor.Tint), 8);
        }

        internal override int TotalSizeNoHeaders()
        {
            return 40;
        }

        internal override int TotalSize()
        {
            return TotalSizeNoHeaders() + XlsConsts.SizeOfTRecordHeader;
        }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.SheetExt = this;
        }
    }


    /// <summary>
    /// Window Freeze
    /// </summary>
    internal class TPaneRecord: TBaseRecord
    {
        internal int Id;
        internal int RowSplit;
        internal int ColSplit;
        internal int FirstVisibleRow;
		internal int FirstVisibleCol;
		internal int ActivePane;

        internal TPaneRecord()
        {
            Id = (int)xlr.PANE;
        }

        internal TPaneRecord(int aId, byte[] aData)
            : base()
        {
            Id = aId;
            ColSplit = BitOps.GetWord(aData, 0);
            RowSplit = BitOps.GetWord(aData, 2);
            FirstVisibleRow = BitOps.GetWord(aData, 4);
            FirstVisibleCol = BitOps.GetWord(aData, 6);
            ActivePane = BitOps.GetWord(aData, 8);
        }

		internal override int GetId	{ get {	return Id; }}

		internal void EnsureSelectedVisible()
		{
			switch (ActivePane)
			{
				case 0:
					if (ColSplit <= 0) 
					{
						if (RowSplit <= 0)
						{
							ActivePane = 3;
							return;
						}
						ActivePane = 2;
						return;
					}
					if (RowSplit <= 0)
						ActivePane = 1;
					return;
				case 2:
					if (RowSplit <= 0)
						ActivePane = 3;
					return;
				case 1: 
					if (ColSplit <= 0)
						ActivePane = 3;
					return;
			}
		}

        internal TPanePosition ActivePaneForSelection()
        {
            if (RowSplit <= 0)
            {
                if (ColSplit <= 0) return TPanePosition.UpperLeft;
                return TPanePosition.UpperRight;
            }

            if (ColSplit <= 0)
            {
                return TPanePosition.LowerLeft;
            }

            return TPanePosition.LowerRight;
        }

        internal void ArrangeInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount, bool frozen)
        {
            if (!frozen) return;
            int r=RowSplit;
            if (aRowCount!=0 && CellRange.Left<=0 && CellRange.Right>=FlxConsts.Max_Columns && CellRange.Top<r)
            {
                RowSplit=Math.Max(r+CellRange.RowCount*aRowCount,CellRange.Top);
                FirstVisibleRow = RowSplit;
            }

            int c=ColSplit;
            if (aColCount!=0 && CellRange.Top<=0 && CellRange.Bottom>=FlxConsts.Max_Rows && CellRange.Left<c)
            {
                ColSplit=Math.Max(c+CellRange.ColCount*aColCount,CellRange.Left);
                FirstVisibleCol = ColSplit;
            }
        }

        internal override void SaveToPxl(TPxlStream PxlStream, int Row, TPxlSaveData SaveData)
        {
            base.SaveToPxl (PxlStream, Row, SaveData);
            PxlStream.WriteByte((byte) pxl.PANE);

            PxlStream.Write16((UInt16)ColSplit);
            PxlStream.Write16((UInt16)RowSplit);
            PxlStream.Write16((UInt16)FirstVisibleRow);
            PxlStream.Write16((UInt16)FirstVisibleCol);
            PxlStream.WriteByte((byte)ActivePane);
        }

        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
        {
            //We can't check for ColSplit/RowSplit since we don't know what it is. It doesn't really matter either.
            Biff8Utils.CheckRow(FirstVisibleRow);
            Biff8Utils.CheckCol(FirstVisibleCol);

            Workbook.WriteHeader((UInt16)Id, (UInt16)TotalSizeNoHeaders());
            Workbook.Write16((UInt16)ColSplit);
            Workbook.Write16((UInt16)RowSplit);
            Workbook.Write16((UInt16)FirstVisibleRow);
            Workbook.Write16((UInt16)FirstVisibleCol);
            Workbook.Write16((UInt16)ActivePane);            
        }

        internal override int TotalSize()
        {
            return TotalSizeNoHeaders() + XlsConsts.SizeOfTRecordHeader;
        }

        internal override int TotalSizeNoHeaders()
        {
            return 10;
        }

        protected override TBaseRecord DoCopyTo(TSheetInfo SheetInfo)
        {
            return (TBaseRecord) MemberwiseClone();
        }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.Window.Pane = this;
        }

    }


    /// <summary>
    /// Guts record for grouping and outline.
    /// </summary>
    internal class TGutsRecord: TxBaseRecord
    {
        internal bool RecalcNeeded;
        internal TGutsRecord(int aId, byte[] aData): base(aId, aData)
        {
            RecalcNeeded=false;
        }

        internal TGutsRecord(): this((int)xlr.GUTS, new byte[8])
        {
            RecalcNeeded = true;
        }

        internal int RowLevel
        {
            get
            {
                return GetWord(4);
            }
            set
            {
                if (value<=0)  
                {
                    SetWord(0,0);
                    SetWord(4,0);
                }
                else
                    if (value<8)
                {
                    SetWord(0, 17+(1+value)*12);
                    SetWord(4, 1+value);
                }
                else
                {
                    SetWord(0, 17+(1+7)*12);
                    SetWord(4, 1+7);
                }
            }
        }

        internal int ColLevel
        {
            get
            {
                return GetWord(6);
            }
            set
            {
                if (value<=0) 
                {
                    SetWord(2,0);
                    SetWord(6,0);
                }
                else
                    if (value<8)
                {
                    SetWord(2, 17+(1+value)*12);
                    SetWord(6, 1+value);
                }
                else
                {
                    SetWord(2, 17+(1+7)*12);
                    SetWord(6, 1+7);
                }
            }
        }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.SheetGlobals.Guts = this;            
        }
    }
}
