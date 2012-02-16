using System;
using FlexCel.Core;
using System.Diagnostics;

namespace FlexCel.XlsAdapter
{
    /// <summary>
    /// A blank record.
    /// </summary>
    internal class TBlankRecord: TCellRecord
    {
        internal TBlankRecord(int aId, byte[] aData, TBiff8XFMap XFMap) : base(aId, aData, XFMap) { }
        internal TBlankRecord(int aCol, int aXF): base((int)xlr.BLANK, aCol, aXF){}

        #region MulBlank
        internal override bool CanJoinNext(TCellRecord NextRecord, int MaxCol)
        {
            TBlankRecord b= (NextRecord as TBlankRecord);
            return (b!=null && b.Col==Col+1 && b.Col<=MaxCol);
        }

        internal override void SaveFirstMul(IDataStream Workbook, TSaveData SaveData, int Row, int JoinedRecordSize)
        {
            unchecked
            {
                Workbook.WriteHeader((UInt16)xlr.MULBLANK,(UInt16)(JoinedRecordSize-XlsConsts.SizeOfTRecordHeader));
                Workbook.Write32((UInt32)((UInt16)Row +(((UInt16)Col)<<16)));
            }
            Workbook.Write16(SaveData.GetBiff8FromCellXF(XF));
        }

        internal override void SaveMidMul(IDataStream Workbook, TSaveData SaveData)
        {
            Workbook.Write16(SaveData.GetBiff8FromCellXF(XF));
        }

        internal override void SaveLastMul(IDataStream Workbook, TSaveData SaveData)
        {
            Workbook.Write16(SaveData.GetBiff8FromCellXF(XF));
            Workbook.Write16((UInt16)Col);            
        }

        internal override int TotalSizeFirst()
        {
            return TotalSize();
        }

        internal override int TotalSizeMid()
        {
            return 2;
        }

        internal override int TotalSizeLast()
        {
            return 4;
        }
        #endregion

        internal override void SaveToPxl(TPxlStream PxlStream, int Row, TPxlSaveData SaveData)
        {
            base.SaveToPxl(PxlStream, Row, SaveData);
            if (!PxlRecordIsValid(Row)) return;

            PxlStream.WriteByte((byte)pxl.BLANK);
            PxlStream.Write16((UInt16) Row);
            PxlStream.WriteByte((byte) Col);
            PxlStream.Write16(SaveData.GetBiff8FromCellXF(XF));
        }

#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
        internal override void SaveToXlsx(TOpenXmlWriter DataStream, int Row, TCellList CellList, bool Dates1904)
        {
            //nothing, it gets empty
        }
#endif

    }

    internal class TBoolErrRecord: TCellRecord
    {
        internal byte BoolErr;
        internal byte ErrFlag;

        internal TBoolErrRecord(int aId, byte[] aData, TBiff8XFMap XFMap) : base(aId, aData, XFMap) { BoolErr = aData[6]; ErrFlag = aData[7]; }
        
        internal TBoolErrRecord(int aCol, int aXF, bool Value): base((int)xlr.BOOLERR, aCol, aXF)
        {
            ErrFlag = 0;
            if (Value) BoolErr = 1; else BoolErr = 0;
        }

        internal TBoolErrRecord(int aCol, int aXF, TFlxFormulaErrorValue Value)
            : base((int)xlr.BOOLERR, aCol, aXF)
        {
            ErrFlag = 1;
            BoolErr = (byte)Value;
        }

        internal TBoolErrRecord(int aCol, int aXF, byte aBoolErr, byte aErrFlag): base((int)xlr.BOOLERR, aCol, aXF)
        {
            BoolErr = aBoolErr;
            ErrFlag = aErrFlag;
        }

        private static TFlxFormulaErrorValue ErrCodeToFlxFormulaErrorValue(int ErrCode)
        {
            if (!(Enum.IsDefined(typeof(TFlxFormulaErrorValue), ErrCode))) return TFlxFormulaErrorValue.ErrNA;   
            return (TFlxFormulaErrorValue)ErrCode;
        }

        internal override object GetValue(ICellList Cells)
        {
            if (ErrFlag==0) 
                if (BoolErr==0) return false; 
                else return true; 
            else return ErrCodeToFlxFormulaErrorValue(BoolErr);

        }

        protected override TBaseRecord DoCopyTo(TSheetInfo SheetInfo)
        {
            TBoolErrRecord Result= (TBoolErrRecord)base.DoCopyTo(SheetInfo);
            Result.ErrFlag=ErrFlag;
            Result.BoolErr=BoolErr;
            return Result;
        }

        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
        {
            base.SaveToStream(Workbook, SaveData, Row);
            byte[]b=new byte[2];
            b[0]=BoolErr;
            b[1]=ErrFlag;
            Workbook.Write(b,2);
        }

        internal override void SaveToPxl(TPxlStream PxlStream, int Row, TPxlSaveData SaveData)
        {
            base.SaveToPxl (PxlStream, Row, SaveData);
            PxlStream.WriteByte((byte)pxl.BOOLERR);
            PxlStream.Write16((UInt16) Row);
            PxlStream.WriteByte((byte) Col);
            PxlStream.Write16(SaveData.GetBiff8FromCellXF(XF));
            byte[]b=new byte[2];
            b[0]=BoolErr;
            b[1]=ErrFlag;
            PxlStream.Write(b,0,2);
        }


        internal override int TotalSizeNoHeaders()
        {
            return base.TotalSizeNoHeaders()+2;
        }

#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
        internal override void SaveToXlsx(TOpenXmlWriter DataStream, int Row, TCellList CellList, bool Dates1904)
        {
            if (ErrFlag==0) 
            {
                DataStream.WriteAtt("t", "b");
                DataStream.WriteElement("v", BoolErr == 1);
            }
            else
            {
                DataStream.WriteAtt("t", "e");
                DataStream.WriteElement("v", TFormulaMessages.ErrString(ErrCodeToFlxFormulaErrorValue(BoolErr)));
            }
        }
#endif

    }

    /// <summary>
    /// A record containing an IEEE 8bytes double number.
    /// </summary>
    internal class TNumberRecord: TCellRecord
    {
        double NumValue;
        internal TNumberRecord(int aId, byte[] aData, TBiff8XFMap XFMap) : base(aId, aData, XFMap) { NumValue = BitConverter.ToDouble(aData, 6); }
        internal TNumberRecord(int aCol, int aXF): base((int)xlr.NUMBER, aCol, aXF){NumValue=0;}
        internal TNumberRecord(int aCol, int aXF, double aNumValue): base((int)xlr.NUMBER, aCol, aXF){NumValue = aNumValue;}

        internal override object GetValue(ICellList Cells)
        {
            return NumValue;
        }

        protected override TBaseRecord DoCopyTo(TSheetInfo SheetInfo)
        {
            TNumberRecord Result=(TNumberRecord)base.DoCopyTo(SheetInfo);
            Result.NumValue=NumValue;
            return Result;
        }

        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
        {
            base.SaveToStream(Workbook, SaveData, Row);
            Workbook.Write(BitConverter.GetBytes(NumValue), 8);
        }

        internal override void SaveToPxl(TPxlStream PxlStream, int Row, TPxlSaveData SaveData)
        {
            base.SaveToPxl (PxlStream, Row, SaveData);
            if (!PxlRecordIsValid(Row)) return;

            PxlStream.WriteByte((byte)pxl.NUMBER);
            PxlStream.Write16((UInt16) Row);
            PxlStream.WriteByte((byte) Col);
            PxlStream.Write16(SaveData.GetBiff8FromCellXF(XF));
            PxlStream.Write(BitConverter.GetBytes(NumValue),0, 8);
        }


        internal override int TotalSizeNoHeaders()
        {
            return base.TotalSizeNoHeaders()+8;
        }

#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
        internal override void SaveToXlsx(TOpenXmlWriter DataStream, int Row, TCellList CellList, bool Dates1904)
        {
            DataStream.WriteElement("v", NumValue);
        }
#endif
    }

    /// <summary>
    /// A record containing a compressed 8 bytes double number into 4 bytes.
    /// </summary>
    internal class TRKRecord: TCellRecord
    {
        internal Int32 RK;

        internal TRKRecord(int aId, byte[] aData, TBiff8XFMap XFMap) : base(aId, aData, XFMap) { RK = BitConverter.ToInt32(aData, 6); }
        internal TRKRecord(int aCol, int aXF, double Value): base((int)xlr.RK, aCol, aXF)
        {
            if (!EncodeRK(Value, ref RK)) FlxMessages.ThrowException(FlxErr.ErrInvalidCellValue, Value.ToString());
        }

		internal override object GetValue(ICellList Cells)
		{
			return GetValue();
		}
		
		internal object GetValue()
		{
			//Warning: Be careful with sign. Shl/Shr on pascal works different that ">>" and "<<" in c
			//when using negative numbers. Shr will add 0 at the left, ">>" will preserve sign.

			double d=0;
			if ((RK & 0x2) == 0x2) //integer
				d=RK >> 2; else
			{
				byte[] b= new byte[8];
				unchecked
				{
					UInt32 uRk=((UInt32)RK) & 0xfffffffc;
					BitConverter.GetBytes(uRk).CopyTo(b,4);
				}
				d=BitConverter.ToDouble(b,0);
			}

			if ((RK & 0x1) == 0x1) return d/100; else return d;
		}

        internal static bool EncodeRK(double Value, ref Int32 RK)
        {
            if (Value == 0.0) Value = 0.0;  //a IEEE number might have the negative bit set, and we want 0 to be 0, not -0.
            for (byte i=0; i<2;i++)
            {
                double d=Value*(1+99*i);
                byte[] pd=BitConverter.GetBytes(d);
                if ((BitConverter.ToUInt32(pd,0) ==0 ) && ((pd[4] & 3)==0))    //Type 0-2   30 bits IEEE float
                {
                    pd[4]+=i;
                    RK=BitConverter.ToInt32(pd,4);
                    return true;
                }

                long Mask= 0x1FFFFFFF;  //29 bits
                if ((Math.Floor(d)==d) && (d<=Mask) && (d>=-Mask-1))    //Type 1-3: 30 bits integer
                {
                    RK= (Convert.ToInt32(d) << 2) + i+2;
                    return true;
                }
            }
            return false;
        }

        internal static bool IsRK(double Value)
        {
            Int32 r=0;
            return EncodeRK(Value, ref r);
        }

        protected override TBaseRecord DoCopyTo(TSheetInfo SheetInfo)
        {
            TRKRecord Result= (TRKRecord)base.DoCopyTo(SheetInfo);
            Result.RK=RK;
            return Result;
        }

        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
        {
            base.SaveToStream(Workbook, SaveData, Row);
            Workbook.Write(BitConverter.GetBytes((Int32)RK),4);
        }

        internal override void SaveToPxl(TPxlStream PxlStream, int Row, TPxlSaveData SaveData)
        {
            base.SaveToPxl (PxlStream, Row, SaveData);
            if (!PxlRecordIsValid(Row)) return;

            PxlStream.WriteByte((byte)pxl.NUMBER);
            PxlStream.Write16((UInt16) Row);
            PxlStream.WriteByte((byte) Col);
            PxlStream.Write16(SaveData.GetBiff8FromCellXF(XF));
            Double NumValue = (double)GetValue();
            PxlStream.Write(BitConverter.GetBytes(NumValue),0, 8);
        }


        internal override int TotalSizeNoHeaders()
        {
            return base.TotalSizeNoHeaders()+4;
        }

        #region MulRk
        internal override bool CanJoinNext(TCellRecord NextRecord, int MaxCol)
        {
            TRKRecord b= (NextRecord as TRKRecord);
            return (b!=null && b.Col==Col+1 && b.Col<=MaxCol);
        }

        internal override void SaveFirstMul(IDataStream Workbook, TSaveData SaveData, int Row, int JoinedRecordSize)
        {
            unchecked
            {
                Workbook.WriteHeader((UInt16)xlr.MULRK, (UInt16)(JoinedRecordSize-XlsConsts.SizeOfTRecordHeader));
                Workbook.Write32((UInt32)((UInt16)Row +(((UInt16)Col)<<16)));
            }
            Workbook.Write16(SaveData.GetBiff8FromCellXF(XF));
            Workbook.Write(BitConverter.GetBytes((Int32)RK),4);
        }

        internal override void SaveMidMul(IDataStream Workbook, TSaveData SaveData)
        {
            Workbook.Write16(SaveData.GetBiff8FromCellXF(XF));
            Workbook.Write(BitConverter.GetBytes((Int32)RK),4);
        }

        internal override void SaveLastMul(IDataStream Workbook, TSaveData SaveData)
        {
            Workbook.Write16(SaveData.GetBiff8FromCellXF(XF));
            Workbook.Write(BitConverter.GetBytes((Int32)RK),4);
            Workbook.Write16((UInt16)Col);            
        }

        internal override int TotalSizeFirst()
        {
            return TotalSize();
        }

        internal override int TotalSizeMid()
        {
            return 2+4;
        }

        internal override int TotalSizeLast()
        {
            return 4+4;
        }
        #endregion

#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
        internal override void SaveToXlsx(TOpenXmlWriter DataStream, int Row, TCellList CellList, bool Dates1904)
        {
            DataStream.WriteElement("v", (double)GetValue());
        }
#endif
    }

    /// <summary>
    /// A record containing multiple records of the same type for the same row.
    /// This is an optimization for saving the file, we will convert it to normal base records.
    /// </summary>
    internal abstract class TMultipleValueRecord: TxBaseRecord
    {
        protected int CurrentCol;
        protected TBiff8XFMap XFMap;
        internal TMultipleValueRecord(int aId, byte[] aData, TBiff8XFMap aXFMap): base(aId, aData)
        {
            XFMap = aXFMap;
        }
        internal abstract bool Eof();
        internal abstract TCellRecord ExtractOneRecord();

		internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
		{
			ws.Cells.AddMultipleCells(this, rRow, RecordLoader.VirtualReader);
		}

    }


    /// <summary>
    /// Multiple blanks.
    /// </summary>
    internal class TMulBlankRecord: TMultipleValueRecord
    {
        internal TMulBlankRecord(int aId, byte[] aData, TBiff8XFMap aXFMap): base(aId, aData, aXFMap){}

        internal override bool Eof()
        {
            return 4+(CurrentCol+1)*2>=Data.Length;
        }

        internal override TCellRecord ExtractOneRecord()
        {
            //Row= GetWord(0));
            TCellRecord Result=new TBlankRecord(GetWord(2) + CurrentCol, XFMap.GetCellXF2007(GetWord(4 + CurrentCol*2)));
            CurrentCol++;
            return Result;
        }


    }

    /// <summary>
    /// Multiple RK values.
    /// </summary>
    internal class TMulRKRecord: TMultipleValueRecord
    {
        private int SizeOfRK=6;

        internal TMulRKRecord(int aId, byte[] aData, TBiff8XFMap aXFMap): base(aId, aData, aXFMap){}

        internal override bool Eof()
        {
            return 4+(CurrentCol+1)*SizeOfRK>=Data.Length;
        }

        internal override TCellRecord ExtractOneRecord()
        {
            int NewDataSize=10;
            byte[] NewData=new byte[NewDataSize];
            BitOps.SetWord(NewData, 0, GetWord(0));
            BitOps.SetWord(NewData, 2, GetWord(2) + CurrentCol);
            BitOps.SetWord(NewData, 4, GetWord(4 + CurrentCol*SizeOfRK));  //XF
            BitOps.SetCardinal(NewData, 6, GetCardinal(6 + CurrentCol*SizeOfRK));  //RK

            TCellRecord Result=new TRKRecord((int)xlr.RK, NewData, XFMap);
            CurrentCol++;
            return Result;
        }


    }

    /// <summary>
    /// Encapsulates a Cell record with the corresponding row, to store cell records on charts and unsuported sheets.
    /// </summary>
    internal class TCellRecordWrapper: TBaseRecord
    {
        internal int FRow;
        internal TCellRecord CellRecord;

        internal TCellRecordWrapper(TCellRecord aCellRecord, int aRow)
        {
            FRow=aRow;
            CellRecord=aCellRecord;
        }

        internal override int GetId
        {
            get { return CellRecord.GetId; }
        }
        internal override int TotalSize()
        {
            return CellRecord.TotalSize();
        }

        internal override int TotalSizeNoHeaders()
        {
            return CellRecord.TotalSizeNoHeaders();
        }

        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
        {
            CellRecord.SaveToStream(Workbook, SaveData, FRow);
        }

        protected override TBaseRecord DoCopyTo(TSheetInfo SheetInfo)
        {
            return new TCellRecordWrapper((TCellRecord)TCellRecord.Clone(CellRecord, SheetInfo), FRow);
        }

    }

}
