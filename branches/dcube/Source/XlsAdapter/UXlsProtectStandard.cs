using System;
using System.Text;
using FlexCel.Core;

namespace FlexCel.XlsAdapter
{
     /// <summary>
    /// Excel 97-2000 Standard Encryption.
    /// </summary>
    internal class TStandardEncryption: TEncryptionEngine
    {
        /// <summary>
        /// RC4 State variables.
        /// </summary>
        private byte[] Permutation;
        private int x;
        private int y;

        /// <summary>
        /// State variables.
        /// </summary>
        int Block;
        long LastStreamPos;
        byte[] Digest;

        /// <summary>
        /// FilePass info.
        /// </summary>
        byte[] FDocId;
        byte[] FSalt;
        byte[] FHashedSalt;


        private const int RekeyBlock = 0x400;

        private TStandardEncryption()
        {
            Permutation = new byte[256];
            Block = -1;
            LastStreamPos = 0;
        }

        internal TStandardEncryption(byte[] DocId, byte[] Salt, byte[] HashedSalt): this()
        {
            FDocId = DocId;
            FSalt = Salt;
            FHashedSalt = HashedSalt;
        }

        internal override bool CheckHash(string Password)
        {
			if (Password.Length ==0) Password=XlsConsts.EmptyExcelPassword; //This is used when no password is given.
			return CheckPassword(Password, FDocId, FSalt, FHashedSalt);
        }

        internal TStandardEncryption(string Password, bool Testing): this()
        {
            Random rnd =  new Random();
            if (Testing)
                rnd= new Random(76); //We need repeatable reads

            FDocId= new byte[16];
            for (int i=0;i<16;i++) FDocId[i]=(byte)rnd.Next(256);

            FSalt= new byte[16];
            for (int i=0;i<16;i++) FSalt[i]=(byte)rnd.Next(256);

            //{0xCB ,0x32 ,0x8D ,0xB2 ,0xF1 ,0x64 ,0x75 ,0xC7 ,0x95 ,0x05 ,0xD0 ,0xFF ,0xEB ,0x4A ,0x0B ,0xBA};
            //FSalt = new byte[16] {0xF5 ,0xDB ,0xE6 ,0xE3 ,0x3F ,0x18 ,0x90 ,0x9F ,0x29 ,0x54 ,0xC8 ,0x80 ,0x9B ,0x37 ,0xBD ,0x4A};
            //FHashedSalt= new byte[16]{0x62 ,0x49 ,0x60 ,0x20 ,0xF8 ,0xF6 ,0x19 ,0x81 ,0x21 ,0x78 ,0xE8 ,0x71 ,0xC0 ,0xCE ,0x3E ,0xC2};
            CalcDigest(Password, FDocId);

            FHashedSalt = CalcHashedSalt(FSalt);
        }


        internal override byte[] Decode(byte[] Data, long StreamPosition, int StartPos, int Count, int RecordLen)
        {
            byte[] NewData = new byte[Data.Length];

            unchecked
            {
                //Sync the stream
                SkipBytes (StreamPosition);
                
                Array.Copy(Data, StartPos, NewData, StartPos, Count);

                int st=StartPos; long pos=StreamPosition; int len= Count;
                while (Block != (pos + len) / RekeyBlock) 
                {
                    int step = RekeyBlock - (int)(pos % RekeyBlock);
                    rc4 (NewData, st, step);
                    st += step;
                    pos += step;
                    len -= step;
                    Block++;
                    MakeKey ((UInt32)Block, Digest);
                }

                rc4 (NewData, st, len);          
            }

            LastStreamPos=StreamPosition+Count;
            return NewData;
        }

        internal override byte[] Encode(byte[] Data, long StreamPosition, int StartPos, int Count, int RecordLen)
        {
            unchecked
            {
                //RC4 is simetric.
                return Decode(Data, StreamPosition, StartPos, Count, RecordLen);
            }
        }

        internal override int GetFilePassRecordLen()
        {
            return 58;
        }

        internal override byte[] GetFilePassRecord()
        {
            byte[] Result = new byte[GetFilePassRecordLen()];
            BitOps.SetWord(Result, 0, (UInt16)(xlr.FILEPASS));
            BitOps.SetWord(Result, 2, GetFilePassRecordLen()-XlsConsts.SizeOfTRecordHeader);
            BitOps.SetWord(Result, 4, 1);
            BitOps.SetWord(Result, 6, 1);
            BitOps.SetWord(Result, 8, 1);
            Array.Copy(FDocId,0, Result, 10, 16);
            Array.Copy(FSalt,0, Result, 26, 16);
            Array.Copy(FHashedSalt,0, Result, 42, 16);

            return Result;
        }

        internal override UInt16 Encode(UInt16 Data, long StreamPosition, int RecordLen)
        {
            byte[] aData= BitConverter.GetBytes(Data);
            return BitConverter.ToUInt16(Encode(aData, StreamPosition, 0, aData.Length, RecordLen),0);
        }

        internal override UInt32 Encode(UInt32 Data, long StreamPosition, int RecordLen)
        {
            byte[] aData= BitConverter.GetBytes(Data);
            return BitConverter.ToUInt32(Encode(aData, StreamPosition, 0, aData.Length, RecordLen),0);
        }

        void SkipBytes (long Position)
        {
            unchecked
            {
                int count = (int)(Position- LastStreamPos);

                int myblock = (int) (Position / RekeyBlock);
                if (myblock != Block) 
                {
                    Block=myblock;
                    MakeKey ((UInt32)myblock, Digest);
                    count = (int)(Position % RekeyBlock);
                }

                rc4 (null, 0, count);
            }
        }


        private static byte[] GetPassword(string Password)
        {
            if (Password.Length>15)
                XlsMessages.ThrowException(XlsErr.ErrPasswordTooLong);
            byte[] Result= new byte[64];

            Encoding.Unicode.GetBytes(Password,0, Password.Length, Result, 0);

            Result[2*Password.Length] = 0x80;
            Result[56] = (byte)(Password.Length << 4);

            return Result;
        }

        private void MakeKey(UInt32 aBlock, byte[] aDigest)
        {
            byte[] Pass= new byte[64];

            /* 40 bit of hashed password, set by verify_password () */
            Array.Copy(aDigest, 0, Pass, 0, 5);

            /* put block number in byte 6...9 */
            BitOps.SetCardinal(Pass, 5, aBlock);
            Pass[9] = 0x80;
            Pass[56] = 0x48;

			using (FlexCel.Core.MD5 MainContext= new MD5())
			{
				byte[] Hash = MainContext.ComputeHashEncryption(Pass);
				PrepareKey (Hash);
			}
        }

        private void CalcDigest(string Password, byte[] aDocId)
        {
            unchecked
            {
                byte[] Pass=GetPassword(Password);

				byte[] Md5Hash = null;
				using (FlexCel.Core.MD5 MainContext= new MD5())
				{
					Md5Hash = MainContext.ComputeHashEncryption(Pass);
				}


                //This is a modified MD5 hash.
                int offset = 0;
                int keyoffset = 0;
                int tocopy = 5;

				using (FlexCel.Core.MD5 ValContext= new MD5())
				{
					while (offset != 16) 
					{
						if ((64 - offset) < 5)
							tocopy = 64 - offset;

						Array.Copy (Md5Hash, keyoffset, Pass, offset,  tocopy);
						offset += tocopy;

						if (offset == 64) 
						{
							ValContext.HashCore (Pass, 0, 64);
							keyoffset = tocopy;
							tocopy = 5 - tocopy;
							offset = 0;
							continue;
						}

						keyoffset = 0;
						tocopy = 5;
						Array.Copy (aDocId, 0, Pass, offset, 16);
						offset += 16;
					}
				
					// Fix (zero) all but first 16 bytes 
					Pass[16] = 0x80;
					for (int i=17; i<64;i++) Pass[i]=0;
					Pass[56] = 0x80;
					Pass[57] = 0x0A;

					ValContext.HashCore (Pass, 0, 64);
					Digest = ValContext.HashFinalEncryption();
				}
            }
        }

        private bool CheckPassword(string Password, byte[] aDocId, byte[] aSalt, byte[] aHashedSalt)
        {
            unchecked
            {
                CalcDigest(Password, aDocId);
                /* Generate 40-bit RC4 key from 128-bit hashed password */
                MakeKey (0, Digest);

                byte[] Salt= new byte[64];
                Array.Copy(aSalt, 0, Salt, 0,  16);
                rc4 (Salt, 0, 16);

                byte[] HashedSalt= new byte[16];
                Array.Copy(aHashedSalt, 0, HashedSalt, 0,  16);
                rc4 (HashedSalt, 0, 16);

                Salt[16] = 0x80;
                for (int i=17; i<64;i++) Salt[i]=0;
                Salt[56] = 0x80;

				using (FlexCel.Core.MD5 MainContext= new MD5())
				{
					byte[] CalcHashedSalt = MainContext.ComputeHashEncryption(Salt);

					for (int i=0;i<16;i++) 
						if ( CalcHashedSalt[i] != HashedSalt[i]) return false;
				}
            }

            return true;
        }
        
        private byte[] CalcHashedSalt(byte[] aSalt)
        {
            unchecked
            {
                /* Generate 40-bit RC4 key from 128-bit hashed password */
                MakeKey (0, Digest);

                byte[] Salt= new byte[64];
                Array.Copy(aSalt, 0, Salt, 0,  16);
                rc4 (Salt, 0, 16);

                Salt[16] = 0x80;
                for (int i=17; i<64;i++) Salt[i]=0;
                Salt[56] = 0x80;

				using (FlexCel.Core.MD5 MainContext= new MD5())
				{
					byte[] HashedSalt = MainContext.ComputeHashEncryption(Salt);
					rc4 (HashedSalt, 0, 16);

					return HashedSalt;
				}
            }
        }

        private void PrepareKey(byte[] KeyData)
        {
            unchecked
            {
                for (int i = 0; i < Permutation.Length; i++)
                    Permutation[i] = (byte)i;
        
                x = 0;
                y = 0;

                int j = 0;
    
                for (int i = 0; i < Permutation.Length; i++)
                {
                    j = (KeyData[i % KeyData.Length] + Permutation[i] + j) % Permutation.Length;
                    //Swap.
                    byte tmp = Permutation[i];
                    Permutation[i] = Permutation[j];
                    Permutation[j] = tmp;
                }
            }
        }

        private void rc4 (byte[] Buffer, int start, int Len)
        {
            unchecked
            {
                for (int i = start; i < start+Len; i++)
                {
                    x = (x + 1) % 256;
                    y = (Permutation[x] + y) % 256;
       
                    byte tmp = Permutation[x];
                    Permutation[x] = Permutation[y];
                    Permutation[y] = tmp;

                    int xorIndex = (Permutation[x] + Permutation[y]) % 256;

                    if (Buffer!=null) Buffer[i] ^= Permutation[xorIndex];
                }
            }
        }
   
    }
}
