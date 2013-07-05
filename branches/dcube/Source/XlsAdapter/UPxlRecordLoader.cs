using System;
using System.IO;
using System.Text;
using System.Diagnostics;

using FlexCel.Core;
using System.Collections.Generic;

namespace FlexCel.XlsAdapter
{
    /// <summary>
    /// Supported pxl versions
    /// </summary>
    internal enum TPxlVersion
    {
        Undefined,
        v10,
        v20
    }

    internal enum TFmlaConvert
    {
        Biff7To8,
        Biff8To7
    }

    /// <summary>
    /// A class to load pxl (Pocket Excel) files. (2.0 and 1.0)
    /// </summary>
    internal class TPxlRecordLoader : TBinRecordLoader
    {
        private Stream DataStream;
        private TPxlVersion PxlVersion;
        private TExternSheetList FExternSheetList;
        private int FormatId;
        private ExcelFile FWorkbook;
        private byte[] BofData;

        private long DataStreamLength;
        private long DataStreamPosition;
        TBorderList BorderList;
        TPatternList PatternList;

        private int MainBookXFCount;

        internal TPxlRecordLoader(Stream aDataStream, TExternSheetList aExternSheetList, 
            TEncryptionData aEncryption, TSST aSST, IFlexCelFontList aFontList, TBorderList aBorderList,
            TPatternList aPatternList, ExcelFile aWorkbook, TBiff8XFMap aXFMap, int aMainBookXFCount, TNameRecordList aNames, TVirtualReader VirtualReader)
            : base(aSST, aFontList, aEncryption, aWorkbook.XlsBiffVersion, aXFMap, aNames, VirtualReader)
        {
            DataStream = aDataStream;
            FExternSheetList = aExternSheetList;
            FormatId = 233;
            FWorkbook = aWorkbook;
            BorderList = aBorderList;
            PatternList = aPatternList;
            DataStreamLength = DataStream.Length;
            DataStreamPosition = DataStream.Position;  //cached for performance.

            MainBookXFCount = aMainBookXFCount;
            
        }

        internal override bool Eof
        {
            get
            {
                return DataStreamPosition >= DataStreamLength;
            }
        }

        private void ShRead(Stream aDataStream, byte[] Data, int IniOfs, int Count)
        {
            Sh.Read(aDataStream, Data, IniOfs, Count);
            DataStreamPosition += Count;
        }

        internal bool CheckHeader()
        {
            BofData = new byte[5];

            if (DataStream.Length - DataStream.Position < BofData.Length) return false;
            ShRead(DataStream, BofData, 0, BofData.Length);
            if (BofData[0] != (int)pxl.BOF)
            {
                return false;
            }
            int VersionNumber = BitOps.GetWord(BofData, 1);
            if (VersionNumber != 0x0009 && VersionNumber != 0x010F && VersionNumber != 0x010E) return false;
            int StreamType = BitOps.GetWord(BofData, 3);
            if (StreamType != 0x005) return false;

            return true;
        }

        internal override void ReadHeader()
        {
            if (BofData != null)
            {
                RecordHeader.Data[0] = BofData[0];
                RecordHeader.Data[1] = 0;
            }
            else
            {
                ShRead(DataStream, RecordHeader.Data, 0, 1);
                RecordHeader.Data[1] = 0; //always 1 bit.
            }
            RecordHeader.Size = GetLength((pxl)RecordHeader.Id, PxlVersion);
            RecordHeader.Id = (int)GetId((pxl)RecordHeader.Id);
        }

        internal TExternSheetList ExternSheetList
        {
            get
            {
                return FExternSheetList;
            }
        }

        internal override TBaseRecord LoadUnsupportdedRecord()
        {
            XlsMessages.ThrowException(XlsErr.ErrExcelInvalid); //we can't load unsupported sheets in pxl
            return null;
        }

        internal override TBaseRecord LoadRecord(out int rRow, bool InGlobalSection)
        {
            int Id = RecordHeader.Id;
            byte[] Data = new byte[RecordHeader.Size];
            if (BofData != null)
            {
                Array.Copy(BofData, 1, Data, 0, Data.Length);
                BofData = null;
            }
            else
            {
                ShRead(DataStream, Data, 0, Data.Length);
            }
            TBaseRecord R = null;

            if (Encryption.Engine != null && Id != (int)pxl.BOF)
                Data = Encryption.Engine.Decode(Data, DataStream.Position - Data.Length, 0, Data.Length, Data.Length);  //Note that we do not care about BoundSheet, as it's data won't be used.

            if (Data.Length > 2) rRow = BitConverter.ToUInt16(Data, 0); else rRow = 0;

            switch (Id)
            {
                case (int)xlr.BOF:
                    R = new TBOFRecord((int)xlr.BOF, GetBofData(Data));
                    if (PxlVersion == TPxlVersion.Undefined)
                    {
                        switch (BitOps.GetWord(Data, 0))
                        {
                            case 0x0009: PxlVersion = TPxlVersion.v10; break;
                            case 0x010F: PxlVersion = TPxlVersion.v20; break;
                            case 0x010E: PxlVersion = TPxlVersion.v20; break;
                            default: XlsMessages.ThrowException(XlsErr.ErrPxlIsInvalid); break;
                        }
                    }
                    break;

                case (int)xlr.EOF: R = new TEOFRecord((int)xlr.EOF, GetEofData(Data)); break;
                case (int)xlr.FORMULA: R = TFormulaRecord.CreateFromBiff8(Names, (int)xlr.FORMULA, GetFormulaData(Data), null); break;

                case (int)xlr.BOUNDSHEET: R = new TBoundSheetRecord(0, GetBoundSheetName(Data)); break;

                case (int)xlr.BLANK: R = new TBlankRecord(Data[2], GetXFAt3(Data)); break;
                case (int)xlr.BOOLERR: R = new TBoolErrRecord(Data[2], GetXFAt3(Data), Data[5], Data[6]); break;
                case (int)xlr.NUMBER: R = new TNumberRecord(Data[2], GetXFAt3(Data), BitConverter.ToDouble(Data, 5)); break;
                case (int)xlr.STRING:
                    if (PxlVersion == TPxlVersion.v10)
                        R = new TStringRecord(ReadString(Data[0]));
                    else
                        R = new TStringRecord(ReadString(BitOps.GetWord(Data, 0))); //Wrong Docs! 

                    break;  //String record saves the result of a formula

                case (int)xlr.XF: R = new TXFRecord((int)xlr.XF, GetXFData(Data), BorderList, PatternList, null); XFCount++; break;
                case (int)xlr.FONT: R = new TFontRecord((int)xlr.FONT, GetFontData(Data)); break;
                case (int)xlr.xFORMAT: R = new TFormatRecord(ReadString(Data[0]), FormatId); FormatId++; break;

                case (int)xlr.LABEL:
                    TLabelSSTRecord SSR = new TLabelSSTRecord(Data[2], GetXFAt3(Data), SST, FontList, ReadString(BitOps.GetWord(Data, 5)));
                    R = SSR;
                    break;

                case (int)xlr.ROW: R = new TRowRecord((int)xlr.ROW, GetRowData(Data), null); break;
                case (int)xlr.NAME: R = TNameRecord.CreateFromBiff8(Names, (int)xlr.NAME, GetNameData(Data)); break;

                case (int)xlr.WINDOW1: R = new TWindow1Record((int)xlr.WINDOW1, GetWindow1Data(Data)); break;
                case (int)xlr.WINDOW2: R = new TWindow2Record((int)xlr.WINDOW2, GetWindow2Data(Data)); break;

                case (int)xlr.PANE: R = new TPaneRecord((int)xlr.PANE, GetPaneData(Data)); break;
                case (int)xlr.SELECTION: R = new TBiff8SelectionRecord((int)xlr.SELECTION, GetSelectionData(Data)); break;

                case (int)xlr.COLINFO: R = new TColInfoRecord((int)xlr.COLINFO, GetColInfoData(Data), null); break;
                case (int)xlr.DEFAULTROWHEIGHT: R = new TDefaultRowHeightRecord((int)xlr.DEFAULTROWHEIGHT, GetDefRowHeightData(Data)); break;
                case (int)xlr.DEFCOLWIDTH: R = new TDefColWidthRecord((int)xlr.DEFCOLWIDTH, GetDefColWidthData(Data)); break;

                case (int)xlr.FILEPASS:
                    XlsMessages.ThrowException(XlsErr.ErrFileIsPasswordProtected);
                    break;
                /*
                TFilePassRecord Fr = new TFilePassRecord(Id, Data, true); 
                if (Encryption.OnPassword!=null) 
                {
                    OnPasswordEventArgs ea= new OnPasswordEventArgs(Encryption.Xls);
                    Encryption.OnPassword(ea); 
                    Encryption.ReadPassword=ea.Password;
                }
                Encryption.Engine=Fr.CreateEncryptionEngine(Encryption.ReadPassword); R=null;break;
                */

                case (int)xlr.CODEPAGE: R = new TCodePageRecord(Id, Data); break;
                case (int)xlr.COUNTRY: R = new TCountryRecord(Id, Data); break;

                default: XlsMessages.ThrowException(XlsErr.ErrPxlIsInvalid); break;
            } //case

            //Peek at the next record...
            if (!Eof)
            {
                ReadHeader();
                int Id2 = RecordHeader.Id;

                switch (Id2)
                {
                    case (int)xlr.STRING:
                        if (!(R is TFormulaRecord) & !(R is TBiff8ShrFmlaRecord) & !(R is TArrayRecord) & !(R is TTableRecord)) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
                        break;
                }
            }
            else
            {
                //Array.Clear(RecordHeader.Data,0,RecordHeader.Length);
                RecordHeader.Id = (int)xlr.EOF;  //Return EOFs, so in case of bad formed files we don't get on an infinite loop.
                RecordHeader.Size = 0;
            }
            return R;
        }

        private int GetXFAt3(byte[] Data)
        {
            return GetXF(Data, 3);
        }

        private int GetXF(byte[] Data, int pos)
        {
            return TBiff8XFMap.GetPxlCellXF2007(BitOps.GetWord(Data, pos), MainBookXFCount);
        }

        #region Utils
        private string ReadString(int len)
        {
            if (PxlVersion == TPxlVersion.v10)
            {
                byte[] Data = new byte[len];
                ShRead(DataStream, Data, 0, Data.Length);
                return Encoding.ASCII.GetString(Data, 0, Data.Length);
            }
            else
            {
                byte[] Data = new byte[len * 2];
                ShRead(DataStream, Data, 0, Data.Length);
                return Encoding.Unicode.GetString(Data, 0, Data.Length);
            }
        }

        private static byte GetColor(byte Data)
        {
            if (Data > 127) return 0x40; //negative, this is automatic
            return (byte)((Data + 8) & 0x7F);
        }

        private byte[] ReadBiff7Formula(TExternSheetList ExternSheetList, Stream aDataStream, byte[] Data, int DataPos, int LenPos)
        {
            ShRead(aDataStream, Data, DataPos, Data.Length - DataPos);
            TBiff7FormulaConverter ResultList = new TBiff7FormulaConverter();
            ResultList.LoadBiff7(Data, DataPos, ExternSheetList, PxlVersion);

            int ResultSize = ResultList.Size;
            byte[] Result = new byte[DataPos + ResultSize];
            Array.Copy(Data, 0, Result, 0, DataPos);
            ResultList.CopyToPtr(Result, DataPos);
            BitOps.SetWord(Result, LenPos, ResultSize);
            return Result;
        }

        #endregion

        #region Conversion
        #region BOF
        private static byte[] GetBofData(byte[] Data)
        {
            return new byte[] { 0x00, 0x06, Data[2], Data[3], 0xD3, 0x18, 0xCD, 0x07, 0x00, 0x00, 0x00, 0x00, 0x06, 0x03, 0x00, 0x00 };

        }
        #endregion
        #region EOF
        private static byte[] GetEofData(byte[] Data)
        {
            return Data;

        }
        #endregion
        #region FORMULA
        private byte[] GetFormulaData(byte[] Data)
        {
            int CCELen = BitOps.GetWord(Data, 14);
            byte[] Result = new byte[22 + CCELen];
            Array.Copy(Data, 0, Result, 0, 3); //Row and col
            BitOps.SetWord(Result, 4, GetXFAt3(Data)); //XF
            Array.Copy(Data, 5, Result, 6, 8); //Result
            Result[14] = (byte)(Data[13] & 0x05b);
            //Array.Copy(Data, 14, Result, 20, 2); //length of cce

            Result = ReadBiff7Formula(ExternSheetList, DataStream, Result, 22, 20);
            return Result;
        }

        #endregion
        #region BoundSheet
        private string GetBoundSheetName(byte[] Data)
        {
            return ReadString(Data[1]);
        }

        #endregion
        #region Font
        private byte[] GetFontData(byte[] Data)
        {
            string FontName = ReadString(Data[Data.Length - 1]);
            TExcelString Xs = new TExcelString(TStrLenLength.is8bits, FontName, null, true);

            byte[] Result = new byte[Data.Length + Xs.TotalSize() - 1]; //doesn't include len.
            Array.Copy(Data, 0, Result, 0, Data.Length);
            BitOps.SetWord(Result, 4, GetColor(Data[4]));
            Xs.CopyToPtr(Result, Data.Length, false);
            return Result;
        }
        #endregion
        #region XF
        private byte[] GetXFData(byte[] Data)
        {
            byte[] Result = new byte[20];
            Array.Copy(Data, 0, Result, 0, 2);  //Font Index

            int PxlOffset = PxlVersion == TPxlVersion.v10 ? 2 : 0;
            int Findex = BitOps.GetWord(Data, 2 + PxlOffset);
            // if (Findex > 230 && Findex < 0xFFF0) Findex -= 230;  Index < 230 are internal
            BitOps.SetWord(Result, 2, Findex);//Format Index. This is not valid Biff8 Format index, but it will be fixed at MergeFromPxlXF

            Result[4] = 1;

            Result[9] = 0xFC;

            byte VAlign = (byte)((((Data[10 + PxlOffset] >> 4) & 0x3) - 1) & 0x3);
            Result[6] = (byte)((Data[10 + PxlOffset] & 0xF) | (VAlign << 4)); //Text Attributes. Vertical Align is -1

            if (Data[14] != 0xFF)
            {
                Result[17] = ((int)TFlxPatternStyle.Solid) << 1;
                Result[18] = (byte)(GetColor(Data[14]) | 0x80); //background color | 1 bit fgcolor
            }
            else
            {
                Result[18] = 0xC0;
            }

            Result[19] = 0x20; //Pattern fore color, not used on pxl.

            if (PxlVersion != TPxlVersion.v20)  //border color
            {
                BitOps.SetWord(Result, 14, (0x40 << 7) | (0x40));
                BitOps.SetWord(Result, 12, (0x40 << 7) | (0x40));
            }
            else
            {
                BitOps.SetWord(Result, 14, (GetColor(Data[18]) << 7) | (GetColor(Data[16])));
                BitOps.SetWord(Result, 12, (GetColor(Data[19]) << 7) | (GetColor(Data[17])));
            }

            if ((Data[11 + PxlOffset] & 0x2) != 0) Result[10] = 1; //Has left border
            if ((Data[11 + PxlOffset] & 0x8) != 0) Result[10] |= (1 << 4); //Has right border
            if ((Data[11 + PxlOffset] & 0x1) != 0) Result[11] |= 1; //Has top border
            if ((Data[11 + PxlOffset] & 0x4) != 0) Result[11] |= (1 << 4); //Has bottom border


            return Result;

        }

        #endregion
        #region Row
        private byte[] GetRowData(byte[] Data)
        {
            byte[] Result = new byte[20];
            Array.Copy(Data, 0, Result, 0, 2);  //Row Number
            if (PxlVersion == TPxlVersion.v10)
            {
                BitOps.SetWord(Result, 6, (Int32)(BitOps.GetWord(Data, 2) * FlxConsts.RowMult));
            }
            else
            {
                Array.Copy(Data, 2, Result, 6, 2);  //Row Height
            }
            Result[12] = 0x80 | 0x40;  //Option flags
            if ((Data[4] & 0x1) != 0)
            {
                Result[12] |= 0x60; //Row hidden.
            }

            BitOps.SetWord(Result, 18, GetXF(Data, 6)); //XF
            Array.Copy(Data, 6, Result, 18, 2);  //XF

            return Result;

        }

        #endregion
        #region Window1
        private static byte[] GetWindow1Data(byte[] Data)
        {
            byte[] Result = new byte[18];
            Array.Copy(Data, 2, Result, 10, 2);  //Worksheet index
            Result[8] |= 0x38;  //show tabs and scroolbars
            Result[16] = 0x58;  //tabs size
            Result[17] = 0x02;  //tabs size
            return Result;
        }

        #endregion
        #region Window2
        private static byte[] GetWindow2Data(byte[] Data)
        {
            byte[] Result = new byte[18];
            Array.Copy(Data, 0, Result, 2, 3);  //Top Row and Left column
            if (Data.Length > 3)
            {
                if ((Data[3] & 0x08) != 0) Result[0] |= 0x08;
                if ((Data[4] & 0x01) != 0) Result[1] |= 0x01;
            }
            Result[0] |= 0xB6;  //show gridlines and headers
            return Result;
        }

        #endregion
        #region Pane
        private static byte[] GetPaneData(byte[] Data)
        {
            byte[] Result = new byte[10];
            Array.Copy(Data, 0, Result, 0, 9);
            return Result;
        }

        #endregion
        #region Selection
        private static byte[] GetSelectionData(byte[] Data)
        {
            byte[] Result = new byte[15];
            Result[0] = 3; //pane.
            Array.Copy(Data, 6, Result, 1, 2);   //row of active cell
            Array.Copy(Data, 8, Result, 3, 1);   //column of active cell
            Result[7] = 1; //1 selection.

            Array.Copy(Data, 0, Result, 9, 2);   //first row of selection
            Array.Copy(Data, 3, Result, 11, 2);   //last row of selection
            Array.Copy(Data, 2, Result, 13, 1);   //first col of selection
            Array.Copy(Data, 5, Result, 14, 1);   //last col of selection

            return Result;
        }

        #endregion
        #region ColInfo
        private byte[] GetColInfoData(byte[] Data)
        {
            byte[] Result = new byte[12]; //doc says 11, but it is 12.
            if (PxlVersion == TPxlVersion.v10)
            {
                Array.Copy(Data, 0, Result, 0, 2);   //first column
                Array.Copy(Data, 0, Result, 2, 2);   //last column
                Array.Copy(GetColWidth(BitOps.GetWord(Data, 2) & 0x3FFF), 0, Result, 4, 2);   //colwidth                
                BitOps.SetWord(Result, 6, GetXF(Data, 4));   //XF
                Result[8] = (byte)((Data[3] & 0x40) >> 6); //option flags.
            }
            else
            {
                Array.Copy(Data, 0, Result, 0, 6);   //most information
                BitOps.SetWord(Result, 6, GetXF(Data, 6));   //XF
                Result[8] = (byte)(Data[8] & 1); //option flags.
            }
            return Result;
        }

        private byte[] GetColWidth(int wPixels)
        {
            if (PxlVersion == TPxlVersion.v10)
            {
                return BitConverter.GetBytes((UInt16)(wPixels * ExcelMetrics.ColMult(FWorkbook)));
            }
            else
                return BitConverter.GetBytes((UInt16)wPixels);
        }

        #endregion
        #region DefColWidth
        private static byte[] GetDefColWidthData(byte[] Data)
        {
            byte[] Result = new byte[2]; //Some information will be lost here, as xls97 format does not support some options.
            Result[0] = Data[3];   //Column Width. This is different from the docs.
            return Result;
        }
        #endregion
        #region DefRowHeight
        private static byte[] GetDefRowHeightData(byte[] Data)
        {
            byte[] Result = new byte[4];
            Array.Copy(Data, 2, Result, 2, 2);   //Row Height
            if ((Data[0] & 1) != 0) Result[0] = 0x2; //Row hidden
            return Result;
        }

        #endregion
        #region Name
        private byte[] ReadName(byte[] Data, int DataOfs, out bool IsWide)
        {
            IsWide = false;

            if (PxlVersion == TPxlVersion.v10)
            {
                byte[] Result = new byte[Data[DataOfs]];
                ShRead(DataStream, Result, 0, Result.Length); //name text
                return Result;
            }
            else
            {
                byte[] NameText = new byte[Data[DataOfs] * 2];
                ShRead(DataStream, NameText, 0, NameText.Length); //name text

                for (int i = 0; i < Data[DataOfs]; i++)
                {
                    if (NameText[i * 2 + 1] != 0)
                    {
                        IsWide = true;
                        break;
                    }
                }

                if (IsWide) return NameText;

                byte[] Result = new byte[Data[DataOfs]];
                for (int i = 0; i < Data[DataOfs]; i++)
                {
                    Result[i] = NameText[i * 2];
                }
                return Result;
            }

        }
        private byte[] GetNameData(byte[] Data)
        {
            bool NameIsWide;

            int DataOfs = PxlVersion == TPxlVersion.v20 ? 2 : 0;
            int ResultLength = 14; //fixed part.

            byte[] NameStr = ReadName(Data, DataOfs, out NameIsWide);

            ResultLength += NameStr.Length;
            ResultLength += BitOps.GetWord(Data, DataOfs + 1) + 1;


            byte[] Result = new byte[ResultLength];
            if (PxlVersion != TPxlVersion.v10)
            {
                Result[0] = (byte)(Data[0] & 0x1); //option flags.
            }

            Array.Copy(Data, DataOfs, Result, 3, 5);   //most information
            Array.Copy(Data, DataOfs + 3, Result, 8, 2);   //ixals again

            if (BitOps.GetWord(Data, DataOfs + 3) == 0xFFFF) Array.Clear(Result, 6, 4); //negative sheets on pxl are 0 here.

            if (NameIsWide) Result[14] = 1; //Result[14] is optionflags
            Array.Copy(NameStr, 0, Result, 15, NameStr.Length);

            Result = ReadBiff7Formula(ExternSheetList, DataStream, Result, 15 + NameStr.Length, 4);

            return Result;
        }

        #endregion

        #endregion

        #region Get Record Length
        internal static int GetLength(pxl Record, TPxlVersion PxlVersion)
        {
            switch (Record)
            {
                case pxl.BLANK: return 5;
                case pxl.BOF: return 4;
                case pxl.BOOLERR: return 7;
                case pxl.BOUNDSHEET: return 2;  // Variable
                case pxl.COLINFO: if (PxlVersion == TPxlVersion.v20) return 9; else return 6;
                case pxl.DEFAULTROWHEIGHT: return 4;
                case pxl.DEFCOLWIDTH: return 6;
                case pxl.EOF: return 0;
                case pxl.FILEPASS: return 14;
                case pxl.FONT: return 15; // Variable;
                case pxl.xFORMAT: return 1; // Variable;
                case pxl.FORMULA: return 16; // Variable;
                case pxl.LABEL: return 7; // Variable;
                case pxl.NAME: if (PxlVersion == TPxlVersion.v20) return 7; else return 5; // Variable
                case pxl.NUMBER: return 13;
                case pxl.PANE: return 9;
                case pxl.ROW: return 8;
                case pxl.SELECTION: return 9;
                case pxl.STRING: if (PxlVersion == TPxlVersion.v20) return 2; else return 2;// Variable
                case pxl.WINDOW1: return 4;
                case pxl.WINDOW2: if (PxlVersion == TPxlVersion.v20) return 5; else return 3;
                case pxl.XF: return 22;

                case pxl.CODEPAGE: return 2;
                case pxl.COUNTRY: return 4;

                default: XlsMessages.ThrowException(XlsErr.ErrPxlIsInvalidToken, (int)Record);
                    return -1; //just to keep compiler happy.
            }
        }
        internal static xlr GetId(pxl Record)
        {
            switch (Record)
            {
                case pxl.BLANK: return xlr.BLANK;
                case pxl.BOF: return xlr.BOF;
                case pxl.BOOLERR: return xlr.BOOLERR;
                case pxl.BOUNDSHEET: return xlr.BOUNDSHEET;
                case pxl.COLINFO: return xlr.COLINFO;
                case pxl.DEFAULTROWHEIGHT: return xlr.DEFAULTROWHEIGHT;
                case pxl.DEFCOLWIDTH: return xlr.DEFCOLWIDTH;
                case pxl.EOF: return xlr.EOF;
                case pxl.FILEPASS: return xlr.FILEPASS;
                case pxl.FONT: return xlr.FONT;
                case pxl.xFORMAT: return xlr.xFORMAT;
                case pxl.FORMULA: return xlr.FORMULA;
                case pxl.LABEL: return xlr.LABEL;
                case pxl.NAME: return xlr.NAME;
                case pxl.NUMBER: return xlr.NUMBER;
                case pxl.PANE: return xlr.PANE;
                case pxl.ROW: return xlr.ROW;
                case pxl.SELECTION: return xlr.SELECTION;
                case pxl.STRING: return xlr.STRING;
                case pxl.WINDOW1: return xlr.WINDOW1;
                case pxl.WINDOW2: return xlr.WINDOW2;
                case pxl.XF: return xlr.XF;

                case pxl.CODEPAGE: return xlr.CODEPAGE;
                case pxl.COUNTRY: return xlr.COUNTRY;

                default: XlsMessages.ThrowException(XlsErr.ErrPxlIsInvalidToken, (int)Record);
                    return xlr.EOF; //just to keep compiler happy.
            }
        }
        #endregion
    }

    #region Pxl Ids

    internal enum pxl
    {
        BLANK = 0x01,
        BOF = 0x09,
        BOOLERR = 0x05,
        BOUNDSHEET = 0x85,
        COLINFO = 0x7D,
        DEFAULTROWHEIGHT = 0x25,
        DEFCOLWIDTH = 0x55,
        EOF = 0x0A,
        FILEPASS = 0x2F,
        FONT = 0x31,
        xFORMAT = 0x1E,
        FORMULA = 0x06,
        LABEL = 0x04,
        NAME = 0x18,
        NUMBER = 0x03,
        PANE = 0x41,
        ROW = 0x08,
        SELECTION = 0x1D,
        STRING = 0x07,
        WINDOW1 = 0x3D,
        WINDOW2 = 0x3E,
        XF = 0xE0,

        CODEPAGE = 0x42,
        COUNTRY = 0x8C

    }
    #endregion

    #region Formula Converter

    internal class TFormulaErrorValue
    {
        internal int Token;

        internal TFormulaErrorValue(int aToken)
        {
            Token = aToken;
        }
    }

    internal class TBiff7FormulaConverter
    {
        List<byte[]> ParsedTokens;

        internal TBiff7FormulaConverter()
        {
            ParsedTokens = new List<byte[]>();
        }

        #region Biff 7 -> 8
        private static void ConvertRowsAndColumns7To8(byte[] Data, byte[] Result, ref int tPos, ref int rPos, bool IsArea)
        {
            //Rows
            int RowAndFlags = BitOps.GetWord(Data, tPos);
            int RowAndFlags2 = 0;
            BitOps.SetWord(Result, rPos, RowAndFlags & 0x3FFF);
            rPos += 2; tPos += 2;
            if (IsArea)
            {
                RowAndFlags2 = BitOps.GetWord(Data, tPos);
                BitOps.SetWord(Result, rPos, RowAndFlags2 & 0x3FFF);
                rPos += 2; tPos += 2;
            }

            //Columns
            Result[rPos] = Data[tPos];
            rPos++; tPos++;
            Result[rPos] = (byte)((RowAndFlags >> 8) & (~0x3F));
            rPos++;
            if (IsArea)
            {
                Result[rPos] = Data[tPos];
                rPos++; tPos++;
                Result[rPos] = (byte)((RowAndFlags2 >> 8) & (~0x3F));
                rPos++;
            }

        }

        private static void Convert3D7To8(TExternSheetList ExternSheetList, byte Token, byte[] Data, byte[] Result, ref int tPos, ref int rPos)
        {
            unchecked
            {
                Int16 ex = (Int16)BitOps.GetWord(Data, tPos);
                if (ex >= 0)
                    XlsMessages.ThrowException(XlsErr.ErrBadToken, Token);
            }

            tPos += 10;

            BitOps.SetWord(Result, rPos, ExternSheetList.AddExternSheet(BitOps.GetWord(Data, tPos), BitOps.GetWord(Data, tPos + 2)));
            tPos += 4;
            rPos += 2;

        }

        private static byte[] ConvertToBiff8(TExternSheetList ExternSheetList, byte Token, byte[] Data, ref int tPos)
        {

            byte[] Result;
            int rPos = 0;

            switch (TBaseParsedToken.CalcBaseToken((ptg)Token))
            {
                case ptg.Name:
                    Result = new byte[4];  //Wrong on Excel docs!
                    BitOps.SetWord(Result, 0, BitOps.GetWord(Data, tPos));
                    tPos += 14;
                    return Result;

                case ptg.NameX:
                    Result = new byte[6]; //This is actually 6
                    BitOps.SetWord(Result, 2, BitOps.GetWord(Data, tPos + 10));
                    tPos += 24;
                    return Result;

                case ptg.Ref:
                case ptg.RefN:
                    Result = new byte[4];
                    ConvertRowsAndColumns7To8(Data, Result, ref tPos, ref rPos, false);
                    return Result;

                case ptg.Area:
                case ptg.AreaN:
                    Result = new byte[8];
                    ConvertRowsAndColumns7To8(Data, Result, ref tPos, ref rPos, true);
                    return Result;

                case ptg.RefErr:
                    tPos += 3;
                    return new byte[4];

                case ptg.AreaErr:
                    tPos += 6;
                    return new byte[8];

                case ptg.Ref3d:
                case ptg.Ref3dErr:
                    Result = new byte[6];
                    Convert3D7To8(ExternSheetList, Token, Data, Result, ref tPos, ref rPos);
                    ConvertRowsAndColumns7To8(Data, Result, ref tPos, ref rPos, false);
                    return Result;

                case ptg.Area3d:
                case ptg.Area3dErr:
                    Result = new byte[10];
                    Convert3D7To8(ExternSheetList, Token, Data, Result, ref tPos, ref rPos);
                    ConvertRowsAndColumns7To8(Data, Result, ref tPos, ref rPos, true);
                    return Result;
            }

            XlsMessages.ThrowException(XlsErr.ErrBadToken, Token);
            return null;  //just to compile.
        }

        #endregion

        #region Biff 8 -> 7
        private static bool ConvertRowsAndColumns8To7(byte[] Data, byte[] ResultData, ref int tPos, ref int rPos, bool IsArea)
        {
            //Read Biff8
            int Row1 = BitOps.GetWord(Data, tPos);
            if (Row1 > FlxConsts.Max_PxlRows) return false;
            tPos += 2;
            int Row2 = 0;
            if (IsArea)
            {
                Row2 = BitOps.GetWord(Data, tPos);
                if (Row2 > FlxConsts.Max_PxlRows) return false;
                tPos += 2;
            }

            int ColAndFlags1 = BitOps.GetWord(Data, tPos);
            tPos += 2;
            int ColAndFlags2 = 0;
            if (IsArea)
            {
                ColAndFlags2 = BitOps.GetWord(Data, tPos);
                tPos += 2;
            }

            //Write Biff7
            BitOps.SetWord(ResultData, rPos, (Row1 & 0x3FFF) | (ColAndFlags1 & (~0x3FFF)));
            rPos += 2;
            if (IsArea)
            {
                BitOps.SetWord(ResultData, rPos, (Row2 & 0x3FFF) | (ColAndFlags2 & (~0x3FFF)));
                rPos += 2;
            }

            ResultData[rPos] = (byte)((ColAndFlags1) & 0xFF);
            rPos++;
            if (IsArea)
            {
                ResultData[rPos] = (byte)((ColAndFlags2) & 0xFF);
                rPos++;
            }

            return true;
        }

        private static bool Convert3D8To7(TReferences References, byte Token, byte[] Data, byte[] ResultData, ref int tPos, ref int rPos)
        {
            BitOps.SetWord(ResultData, rPos, 0xFFFF);
            rPos += 2;

            Array.Clear(ResultData, rPos, 8); //reserved.
            rPos += 8;

            int Sheet1, Sheet2;
            bool HasExternal; string ExternBookName;
            References.GetSheetsFromExternSheet(BitOps.GetWord(Data, tPos), out Sheet1, out Sheet2, out HasExternal, out ExternBookName);
            tPos += 2;

            if (Sheet1 > FlxConsts.Max_PxlSheets || Sheet2 > FlxConsts.Max_PxlSheets) return false;
            if (HasExternal) return false;
            BitOps.SetWord(ResultData, rPos, Sheet1);
            rPos += 2;
            BitOps.SetWord(ResultData, rPos, Sheet2);
            rPos += 2;

            return true;
        }

        private static byte[] ConvertToBiff7(TReferences References, byte Token, byte[] Data, ref int tPos)
        {

            byte[] Result;
            int rPos = 0;

            switch (TBaseParsedToken.CalcBaseToken((ptg)Token))
            {
                case ptg.Name:
                    Result = new byte[14];
                    BitOps.SetWord(Result, 0, BitOps.GetWord(Data, tPos));
                    tPos += 4; //Biff8 Name is Wrong on Excel docs!
                    return Result;

                case ptg.NameX:
                    Result = new byte[24];
                    BitOps.SetWord(Result, 0, 0xFFFF);
                    BitOps.SetWord(Result, 10, BitOps.GetWord(Data, tPos + 2));
                    tPos += 6;//This is actually 6
                    return Result;

                case ptg.Ref:
                case ptg.RefN:
                    Result = new byte[3];
                    if (!ConvertRowsAndColumns8To7(Data, Result, ref tPos, ref rPos, false)) return null;
                    return Result;

                case ptg.Area:
                case ptg.AreaN:
                    Result = new byte[6];
                    if (!ConvertRowsAndColumns8To7(Data, Result, ref tPos, ref rPos, true)) return null;
                    return Result;

                case ptg.RefErr:
                    tPos += 4;
                    return new byte[3];

                case ptg.AreaErr:
                    tPos += 8;
                    return new byte[6];

                case ptg.Ref3d:
                case ptg.Ref3dErr:
                    Result = new byte[17];
                    if (!Convert3D8To7(References, Token, Data, Result, ref tPos, ref rPos)) return null;
                    if (!ConvertRowsAndColumns8To7(Data, Result, ref tPos, ref rPos, false)) return null;
                    return Result;

                case ptg.Area3d:
                case ptg.Area3dErr:
                    Result = new byte[20];
                    if (!Convert3D8To7(References, Token, Data, Result, ref tPos, ref rPos)) return null;
                    if (!ConvertRowsAndColumns8To7(Data, Result, ref tPos, ref rPos, true)) return null;
                    return Result;
            }

            return null;
        }

        #endregion

        private void Flush(byte[] Data, ref int bPos, int tPos)
        {
            if (bPos >= tPos) return;
            byte[] Ptgs = new byte[tPos - bPos];
            Array.Copy(Data, bPos, Ptgs, 0, tPos - bPos);
            ParsedTokens.Add(Ptgs);
            bPos = tPos;
        }

        private TFormulaErrorValue Load(byte[] Data, int Pos, TFmlaConvert ConvertType, TPxlVersion PxlVersion, TExternSheetList ExternSheetList, TReferences References)
        {
            int tPos = Pos;
            int bPos = Pos;
            int fPos = Data.Length;

            while (tPos < fPos)
            {
                byte Token = Data[tPos];
                if (Token >= 0x3 && Token <= 0x16) tPos++;//XlsTokens.IsUnaryOp(Token)||XlsTokens.IsBinaryOp(Token) || Token==XlsTokens.tk_MissArg ;
                else
                    if (XlsTokens.Is_tk_Operand(Token))
                    {
                        tPos++;
                        Flush(Data, ref bPos, tPos);
                        if (ConvertType == TFmlaConvert.Biff7To8)
                        {
                            ParsedTokens.Add(ConvertToBiff8(ExternSheetList, Token, Data, ref tPos));
                        }
                        else
                        {
                            byte[] ConvertedData = ConvertToBiff7(References, Token, Data, ref tPos);
                            if (ConvertedData == null) return new TFormulaErrorValue(Token);
                            ParsedTokens.Add(ConvertedData);
                        }
                        bPos = tPos;
                    }
                    else
                        switch (Token)
                        {
                            case XlsTokens.tk_Str:
                                if (PxlVersion == TPxlVersion.v10)
                                {
                                    tPos++;
                                    Flush(Data, ref bPos, tPos);

                                    byte[] StrValue = new byte[Data[tPos] + 2];
                                    StrValue[0] = Data[tPos]; //String len.
                                    StrValue[1] = 0; //Not wide string.
                                    Array.Copy(Data, tPos + 1, StrValue, 2, StrValue.Length - 2);
                                    ParsedTokens.Add(StrValue);
                                    tPos += Data[tPos] + 1;
                                    bPos = tPos;
                                }
                                else
                                {
                                    tPos += 1 + (int)StrOps.GetStrLen(false, Data, tPos + 1, false, 0);
                                }
                                break;
                            case XlsTokens.tk_Err:
                            case XlsTokens.tk_Bool:
                                tPos += 1 + 1;
                                break;
                            case XlsTokens.tk_Int:
                            case 0x21:  //XlsTokens.Is_tk_Func(Token):
                            case 0x41:
                            case 0x61:
                                tPos += 1 + 2;
                                break;
                            case 0x22: //XlsTokens.Is_tk_FuncVar(Token):
                            case 0x42:
                            case 0x62:
                                tPos += 1 + 3;
                                break;
                            case XlsTokens.tk_Num:
                                tPos += 1 + 8;
                                break;
                            case XlsTokens.tk_Attr:
                                bool IgnoreAttr = false;
                                if ((Data[tPos + 1] & (0x2 | 0x4 | 0x8)) != 0) //optimized if, goto, optimized chose.
                                {
                                    Flush(Data, ref bPos, tPos);
                                    IgnoreAttr = true;
                                }

                                if ((Data[tPos + 1] & 0x04) == 0x04) tPos += (BitOps.GetWord(Data, tPos + 2) + 1) * 2;
                                tPos += 1 + 3;

                                if (IgnoreAttr)
                                {
                                    bPos = tPos; //ignore the attribute, as it contains offsets to the formula that will change.
                                }
                                break;
                            case XlsTokens.tk_Table:
                                return new TFormulaErrorValue(Token);
                            case XlsTokens.tk_MemFunc:
                            case XlsTokens.tk_MemFunc + 0x20:
                            case XlsTokens.tk_MemFunc + 0x40:
                                tPos += 1 + 2; //+ GetWord(Data, tPos+1);
                                break;
                            case XlsTokens.tk_MemArea:
                            case XlsTokens.tk_MemArea + 0x20:
                            case XlsTokens.tk_MemArea + 0x40:
                            case XlsTokens.tk_MemErr:
                            case XlsTokens.tk_MemErr + 0x20:
                            case XlsTokens.tk_MemErr + 0x40:
                            case XlsTokens.tk_MemNoMem:
                            case XlsTokens.tk_MemNoMem + 0x20:
                            case XlsTokens.tk_MemNoMem + 0x40:
                                tPos += 1 + 6; //+ GetWord(Data, tPos+1);
                                break;
                            default:
                                return new TFormulaErrorValue(Token);
                        }
            }//while        
            Flush(Data, ref bPos, tPos);
            return null;
        }


        internal void LoadBiff7(byte[] Data, int Pos, TExternSheetList ExternSheetList, TPxlVersion PxlVersion)
        {
            TFormulaErrorValue Err = Load(Data, Pos, TFmlaConvert.Biff7To8, PxlVersion, ExternSheetList, null); //No references to load Biff7
            if (Err != null)
                XlsMessages.ThrowException(XlsErr.ErrBadToken, Err.Token);
        }

        internal TFormulaErrorValue LoadBiff8(byte[] Data, int Pos, TReferences References)
        {
            return Load(Data, Pos, TFmlaConvert.Biff8To7, TPxlVersion.Undefined, null, References); //No ExternSheetList to load Biff8
        }

        internal int Size
        {
            get
            {
                int Result = 0;
                for (int i = ParsedTokens.Count - 1; i >= 0; i--)
                {
                    Result += ParsedTokens[i].Length;
                }
                return Result;
            }
        }

        internal void CopyToPtr(byte[] Result, int Pos)
        {
            int aPos = Pos;
            for (int i = 0; i < ParsedTokens.Count; i++)
            {
                byte[] pt = ParsedTokens[i];
                Array.Copy(pt, 0, Result, aPos, pt.Length);
                aPos += pt.Length;
            }
            Debug.Assert(aPos == Result.Length, "Formula Result has not completely been filled");
        }


    }
    #endregion


    #region ExternSheet
    internal class TExternSheetList : List<TExternSheetEntry>
    {
        internal int AddExternSheet(int aFirstSheet, int aLastSheet)
        {
            for (int i = 0; i < Count; i++)
            {
                if (
                    this[i].FirstSheet == aFirstSheet &&
                    this[i].LastSheet == aLastSheet
                    )
                    return i;
            }
            Add(new TExternSheetEntry(aFirstSheet, aLastSheet));
            return Count - 1;
        }

    }

    internal class TExternSheetEntry
    {
        internal int FirstSheet;
        internal int LastSheet;

        internal TExternSheetEntry(int aFirstSheet, int aLastSheet)
        {
            FirstSheet = aFirstSheet;
            LastSheet = aLastSheet;
        }
    }
    #endregion
}
