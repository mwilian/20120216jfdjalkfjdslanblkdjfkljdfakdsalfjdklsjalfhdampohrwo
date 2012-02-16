using System;
using FlexCel.Core;

namespace FlexCel.XlsAdapter
{
     /// <summary>
    /// Excel 95 XOR Encryption.
    /// </summary>
    internal class TXorEncryption: TEncryptionEngine
    {
        private UInt16 FKey;
        private UInt16 FHash;
        private byte[] KeySeq;

        internal TXorEncryption(int Key, int Hash, string Password)
        {
			if (Password.Length == 0) Password = XlsConsts.EmptyExcelPassword; //This is used when no password is given.
			FHash=(UInt16)Hash;
            FKey=(UInt16)Key;
            KeySeq= CalcKeySeq(Key, Password);
        }

        internal TXorEncryption(string Password)
        {
            FHash=(UInt16)CalcHash(Password);
            FKey=(UInt16)CalcKey(Password);
            KeySeq= CalcKeySeq(FKey, Password);
        }

        internal override bool CheckHash(string Password)
        {
            if (Password.Length == 0) Password = XlsConsts.EmptyExcelPassword; //This is used when no password is given.
            return CalcHash(Password) == FHash && CalcKey(Password) == FKey;
        }

        internal override byte[] Decode(byte[] Data, long StreamPosition, int StartPos, int Count, int RecordLen)
        {
            unchecked
            {
                int KeyIndex = (int)((StreamPosition + RecordLen) & 0xF);
                byte[] NewData = new byte[Data.Length];

                for (int i=StartPos; i<StartPos+Count; i++)
                {
                    NewData[i]= (byte)(( (Data[i]<<3) | (Data[i]>>5))  ^ KeySeq[KeyIndex]);
                    KeyIndex=(KeyIndex+1) & 0xF;
                }
            
                return NewData;
            }
        }

        internal override byte[] Encode(byte[] Data, long StreamPosition, int StartPos, int Count, int RecordLen)
        {
            unchecked
            {
                int KeyIndex = (int)((StreamPosition+RecordLen) & 0xF);
                byte[] NewData = new byte[Data.Length];

                for (int i=StartPos; i<StartPos+Count; i++)
                {
                    byte t= (byte)(Data[i] ^ KeySeq[KeyIndex]);
                    NewData[i]= (byte)(( t>>3) | (t<<5));
                    KeyIndex=(KeyIndex+1) & 0xF;
                }
            
                return NewData;
            }
        }

        internal override int GetFilePassRecordLen()
        {
            return 10;
        }

        internal override byte[] GetFilePassRecord()
        {
            byte[] Result = new byte[GetFilePassRecordLen()];
            BitOps.SetWord(Result, 0, (UInt16)(xlr.FILEPASS));
            BitOps.SetWord(Result, 2, GetFilePassRecordLen()-XlsConsts.SizeOfTRecordHeader);
            BitOps.SetWord(Result, 4, 0);
            BitOps.SetWord(Result, 6, FKey);
            BitOps.SetWord(Result, 8, FHash);

            return Result;
        }


        /// <summary>
        /// This routine does not call Encode(Byte[]...) for performance reasons. We want it to be fast.
        /// </summary>
        /// <param name="Data"></param>
        /// <param name="StreamPosition"></param>
        /// <param name="RecordLen"></param>
        /// <returns></returns>
        internal override UInt16 Encode(UInt16 Data, long StreamPosition, int RecordLen)
        {
            unchecked
            {
                int KeyIndex = (int)((StreamPosition+RecordLen) & 0xF);
                UInt16 NewData = 0;

                int Mask= 0x00;
                for (int i=0;i<2;i++)
                {
                    byte t= (byte)((Data >> Mask) ^ KeySeq[KeyIndex]);
                    NewData |= (UInt16)((( t>>3) | (t<<5) & 0xFF) << Mask);
                    KeyIndex=(KeyIndex+1) & 0xF;
                    Mask += 8;
                }
            
                return NewData;
            }
        }

        /// <summary>
        /// This routine does not call Encode(Byte[]...) for performance reasons. We want it to be fast.
        /// </summary>
        /// <param name="Data"></param>
        /// <param name="StreamPosition"></param>
        /// <param name="RecordLen"></param>
        /// <returns></returns>
        internal override UInt32 Encode(UInt32 Data, long StreamPosition, int RecordLen)
        {
            unchecked
            {
                int KeyIndex = (int)((StreamPosition + RecordLen) & 0xF);
                UInt32 NewData = 0;

                int Mask= 0x00;
                for (int i=0;i<4;i++)
                {
                    byte t= (byte)((Data >> Mask) ^ KeySeq[KeyIndex]);
                    NewData |= (UInt32)((( t>>3) | (t<<5) & 0xFF) << Mask);
                    KeyIndex=(KeyIndex+1) & 0xF;
                    Mask += 8;
                }          
                return NewData;
            }
        }


        internal static int CalcHash(string Password)
        {
            int Result=0;
            int CharCount=Password.Length;

            if (CharCount>15)
                XlsMessages.ThrowException(XlsErr.ErrPasswordTooLong);

            for (int i=0; i<CharCount;i++)
            {
                char c=Password[i];
                unchecked
                {
                    int c2 = c<< (i+1);
                    Result ^=(c2 & 0x7FFF) | (c2 >> 0xF);
                }
            }

            return Result ^ CharCount ^ 0xCE4B;

        }

        internal static int CalcKey(string Password)
        {
            int Result=0;
            int KeyBase=0x8000;
            int KeyFinal=0xFFFF;
            int CharCount=Password.Length;

            if (CharCount>15)
                XlsMessages.ThrowException(XlsErr.ErrPasswordTooLong);

            for (int i=0; i<CharCount; i++)
            {
                byte c= (byte)(Password[CharCount-i-1] & 0x7F);

                for (int BitIndex=0; BitIndex<8; BitIndex++)
                {
                    KeyBase = ((KeyBase << 1) | (( KeyBase & 0xFFFF) >> (16-1))) & 0xFFFF;
                    if ((KeyBase & 1) == 1) KeyBase ^=0x1020;
                    KeyFinal = ((KeyFinal << 1) | (( KeyFinal & 0xFFFF) >> (16-1))) & 0xFFFF;
                    if ((KeyFinal & 1) == 1) KeyFinal ^=0x1020;

                    if ((c & 1) == 1) Result ^= KeyBase;
                    c>>=1;
                }

            }
            return Result ^ KeyFinal;

        }

        internal static byte[] CalcKeySeq(int Key, string Password)
        {
            byte[] Result= new byte[16];
            int CharCount=Password.Length;

            if (CharCount>15)
                XlsMessages.ThrowException(XlsErr.ErrPasswordTooLong);

            for (int i=0;i<CharCount;i++) Result[i]=(byte)Password[i];

            byte[] Remaining={0xBB, 0xFF, 0xFF, 0xBA, 0xFF, 0xFF, 0xB9, 0x80, 0x00, 0xBE, 0x0F, 0x00, 0xBF, 0x0F, 0x00};
            for (int i=CharCount; i< Result.Length; i++) Result[i]=Remaining[i-CharCount];

            unchecked
            {
                byte KeyLower= (byte)Key;
                byte KeyUpper= (byte)(Key >> 8);

                for (int SeqIndex=0; SeqIndex<16; SeqIndex+=2)
                {
                    Result[SeqIndex] ^= KeyLower;
                    Result[SeqIndex+1] ^= KeyUpper;
                }

                for (int i=0;i< Result.Length; i++)
                    Result[i] = (byte) ((Result[i]<<2) | (Result[i]>>6));
            }

            return Result;
        }
    }
}
