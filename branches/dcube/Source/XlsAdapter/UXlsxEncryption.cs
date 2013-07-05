using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using FlexCel.Core;
using System.Xml;
using System.Globalization;
using System.Reflection;


namespace FlexCel.XlsAdapter
{
    internal static class TEncryptionUtils
    {
        public const int AgileSegmentSize = 4096;

        public static AesManaged CreateEngine(TEncryptionParameters EncryptionParameters)
        {
            AesManaged Engine = new AesManaged(); //using

            Engine.KeySize = CalcKeySizeInBits(EncryptionParameters);

            Engine.Padding = EncryptionParameters.Padding;
            Engine.Mode = EncryptionParameters.ChainingMode;
            return Engine;
        }

        private static int CalcKeySizeInBits(TEncryptionParameters EncryptionParameters)
        {
            switch (EncryptionParameters.Algorithm)
            {
                case TEncryptionAlgorithm.AES_128:
                    return 0x80;

                case TEncryptionAlgorithm.AES_192:
                    return 0xC0;

                case TEncryptionAlgorithm.AES_256:
                    return 0x100;

            }
            XlsMessages.ThrowException(XlsErr.ErrNotSupportedEncryption);
            return 0;
        }

        public static byte[] PadArray(byte[] Value, int PadSize, byte PadValue)
        {
            if (Value.Length == PadSize) return Value;

            byte[] Result = new byte[PadSize];
            if (Value.Length < Result.Length)
            {
                Array.Copy(Value, Result, Value.Length);
                for (int i = Value.Length; i < Result.Length; i++)
                {
                    Result[i] = PadValue;
                }
            }
            else
            {
                Array.Copy(Value, Result, Result.Length);
            }

            return Result;
        }

        public static void CopyStream(Stream InputStream, Stream OutputStream, long MaxSize)
        {
            byte[] buffer = new byte[8192];
            int read;
            while ((read = InputStream.Read(buffer, 0, buffer.Length)) > 0)
            {
                if (MaxSize > 0 && OutputStream.Length + read > MaxSize)
                {
                    OutputStream.Write(buffer, 0, (int)(MaxSize - OutputStream.Length));
                    return;
                }
                OutputStream.Write(buffer, 0, read);
            }
        }

        internal static void CopyStream(Stream SourceStream, TOle2File DataStream)
        {
            byte[] buffer = new byte[8192];
            int read;
            while ((read = SourceStream.Read(buffer, 0, buffer.Length)) > 0)
            {
                DataStream.Write(buffer, 0, read);
            }
        }

        public static byte[] GetRandom(int Size)
        {
            byte[] Result = new byte[Size];
            RNGCryptoServiceProvider rng = new RNGCryptoServiceProvider();
            rng.GetBytes(Result);
            return Result;
        }
    }

    internal class TEncryptionParameters
    {
        public TEncryptionAlgorithm Algorithm;
        public PaddingMode Padding = PaddingMode.Zeros;
        public CipherMode ChainingMode;

        public static TEncryptionParameters CreateStandard(TEncryptionAlgorithm aAlgorithm)
        {
            TEncryptionParameters Result = new TEncryptionParameters();
            Result.Algorithm = aAlgorithm;
            Result.ChainingMode = CipherMode.ECB;

            return Result;
        }
    }

    internal abstract class TEncryptionKey
    {
        public int SpinCount;
        public int BlockSize;
        public int KeySizeInBytes;
        public byte[] Salt;
        public byte[] Password;

        private byte[] PrecomputedKey; //we will store this here to avoid multiple spincount loops (100000 in Excel 2010)
        public byte[] Key;
        public byte[] IV;

        public readonly bool VariableIV;

        public TEncryptionKey(bool aVariableIV)
        {
            VariableIV = aVariableIV;
        }

        protected static byte[] Concat(byte[] h1, byte[] h2)
        {
            if (h1 == null) return h2;
            if (h2 == null) return h1;
            byte[] Result = new byte[h1.Length + h2.Length];
            Array.Copy(h1, Result, h1.Length);
            Array.Copy(h2, 0, Result, h1.Length, h2.Length);
            return Result;
        }

        public static HashAlgorithm CreateHasher()
        {
            return SHA1.Create();
        }

        internal long HashSizeBytes()
        {
            using (HashAlgorithm hasher = CreateHasher())
            {
                return hasher.HashSize / 8;
            }
        }

        public void PreCalcKey()
        {
            //doesn't work, as it uses PBKDF2 and office uses PBKDF1
            //Rfc2898DeriveBytes keyGenerator = new Rfc2898DeriveBytes(Encoding.Unicode.GetBytes(Password), Salt, 50000); //salt in standard is 16 bytes.
            //key = keyGenerator.GetBytes(16);

            byte[] hash = Concat(Salt, Password);
            using (HashAlgorithm hasher = CreateHasher())
            {
                hash = hasher.ComputeHash(hash); // H0

                for (int i = 0; i < SpinCount; i++)
                {
                    byte[] hash1 = new byte[4 + hash.Length];
                    BitOps.SetCardinal(hash1, 0, i);
                    Array.Copy(hash, 0, hash1, 4, hash.Length);
                    hash = hasher.ComputeHash(hash1);
                }
            }

            PrecomputedKey =hash;
        }

        public void CalcKey(byte[] KeyBlockKey, byte[] IVBlockKey)
        {
            //Final H
            using (HashAlgorithm hasher = CreateHasher())
            {
                byte[] hfinal = hasher.ComputeHash(Concat(PrecomputedKey, KeyBlockKey));
                Key = DeriveKey(hasher, hfinal);
                IV = DeriveIV(hasher, IVBlockKey);
            }
        }

        public abstract void CalcDataIV(long SegNum);
        protected abstract byte[] DeriveKey(HashAlgorithm hasher, byte[] hfinal);
        protected abstract byte[] DeriveIV(HashAlgorithm hasher, byte[] BlockKey);

        internal byte[] Hash(byte[] value)
        {
            using (HashAlgorithm hasher = CreateHasher())
            {
                return hasher.ComputeHash(value);
            }
        }
    }

    internal class TStandardEncryptionKey: TEncryptionKey
    {
        public TStandardEncryptionKey(byte[] aSalt, int aKeySize): base(false)
        {
            SpinCount = 50000;
            BlockSize = 0x10;
            KeySizeInBytes = aKeySize;
            Salt = aSalt;
        }

        public static readonly byte[] BlockKey = new byte[4];

        protected override byte[] DeriveIV(HashAlgorithm hasher, byte[] BlockKey)
        {
            return new byte[16]; //doesn't matter in ECB
        }
        protected override byte[] DeriveKey(HashAlgorithm hasher, byte[] hfinal)
        {
            int cbRequiredKeyLength = BlockSize;

            byte[] X1 = GetX(hasher, hfinal, 0x36);
            byte[] X2 = GetX(hasher, hfinal, 0x5C);

            byte[] Result = new byte[cbRequiredKeyLength];

            for (int i = 0; i < Result.Length; i++)
            {
                if (i < X1.Length) Result[i] = X1[i];
                else Result[i] = X2[i - X1.Length];
            }

            return Result;
        }

        private static byte[] GetX(HashAlgorithm hasher, byte[] hfinal, byte initValue)
        {
            int cbHash = hfinal.Length; // 20;
            byte[] X1 = new byte[64];
            for (int i = 0; i < X1.Length; i++)
            {
                X1[i] = initValue;
            }

            for (int i = 0; i < cbHash; i++)
            {
                X1[i] ^= hfinal[i];
            }

            return hasher.ComputeHash(X1);
        }

        public override void CalcDataIV(long SegNum)
        {
            //nothing here.
        } 
    }

    internal class TAgileEncryptionKey : TEncryptionKey
    {
        public static readonly byte[] VerifierHashInputBlockKey = new byte[] { 0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79 };
        public static readonly byte[] VerifierHashValueBlockKey = new byte[] { 0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e };
        public static readonly byte[] VerifierKeyValueBlockKey =  new byte[] { 0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6 };

        public TAgileEncryptionKey()
            : base(true)
        {
        }

        public static TAgileEncryptionKey CreateForWriting(string aPassword, int KeySizeInBytes)
        {
            const int SaltSize = 16;
            TAgileEncryptionKey Result = new TAgileEncryptionKey();
            Result.SpinCount = 100000;
            Result.BlockSize = 16;
            Result.KeySizeInBytes = KeySizeInBytes;
            Result.Salt = TEncryptionUtils.GetRandom(SaltSize);

            if (aPassword != null)
            {
                Result.Password = Encoding.Unicode.GetBytes(aPassword);
                Result.PreCalcKey();
            }

            return Result;
        }

        public override void CalcDataIV(long SegNum)
        {
            byte[] BlockKey = new byte[4];
            BitOps.SetCardinal(BlockKey, 0, SegNum);
            CalcDataIV(BlockKey);
        }

        public void CalcDataIV(byte[] BlockKey)
        {
            using (HashAlgorithm hasher = CreateHasher())
            {
                IV = DeriveIV(hasher, BlockKey);
            }
        }

        protected override byte[] DeriveKey(HashAlgorithm hasher, byte[] hfinal)
        {
            return TEncryptionUtils.PadArray(hfinal, KeySizeInBytes, 0x36);
        }

        protected override byte[] DeriveIV(HashAlgorithm hasher, byte[] BlockKey)
        {
            byte[] ResultHash = Salt;
            if (BlockKey != null)
            {
                ResultHash = hasher.ComputeHash(Concat(Salt, BlockKey));
            }
            return TEncryptionUtils.PadArray(ResultHash, BlockSize, 0x36);
        }
    }

    internal abstract class TEncryptionVerifier
    {
        internal virtual bool VerifyPass(string Password, TEncryptionParameters EncParams, TEncryptionKey Key)
        {
            if (Password == null) XlsMessages.ThrowException(XlsErr.ErrInvalidPassword);
            if (Password.Length > 255) XlsMessages.ThrowException(XlsErr.ErrPasswordTooLong);

            Key.Password = Encoding.Unicode.GetBytes(Password);
            Key.PreCalcKey();

            return true;
        }

        protected static byte[] DecryptBytes(byte[] EncryptedVerifier, ICryptoTransform Decryptor, int MaxSize)
        {
            byte[] DecriptedVerifier;
            using (MemoryStream ms = new MemoryStream(EncryptedVerifier))
            {
                using (CryptoStream cs = new CryptoStream(ms, Decryptor, CryptoStreamMode.Read))
                {
                    DecriptedVerifier = GetBytesFromCryptoStream(cs, MaxSize);
                }
            }
            return DecriptedVerifier;
        }

        public static byte[] EncryptBytes(byte[] DecryptedData, ICryptoTransform Encryptor, int PadMultiple)
        {
            byte[] EncryptedData;
            using (MemoryStream ms = new MemoryStream(DecryptedData))
            {
                using (CryptoStream cs = new CryptoStream(ms, Encryptor, CryptoStreamMode.Read))
                {
                    EncryptedData = GetBytesFromCryptoStream(cs, -1);
                }
            }
            return EncryptedData;
        }


        private static byte[] GetBytesFromCryptoStream(CryptoStream cs, int MaxSize)
        {
            using (MemoryStream ms = new MemoryStream()) //not the most efficient code in the world, but should be used for small things
            {
                TEncryptionUtils.CopyStream(cs, ms, MaxSize);
                return ms.ToArray();
            }
        }
    }

    internal class TStandardEncryptionVerifier: TEncryptionVerifier
    {
        public byte[] EncryptedVerifier;
        public int VerifierHashSizeBytes;
        public byte[] EncryptedVerifierHash;

        internal override bool VerifyPass(string Password, TEncryptionParameters EncParams, TEncryptionKey Key)
        {
            if (!base.VerifyPass(Password, EncParams, Key)) return false;

            Key.CalcKey(TStandardEncryptionKey.BlockKey, null);
            byte[] DecriptedVerifier;
            byte[] DecriptedVerifierHash;
            using (AesManaged Engine = TEncryptionUtils.CreateEngine(EncParams))
            {
                using (ICryptoTransform Decryptor = Engine.CreateDecryptor(Key.Key, Key.IV))
                {
                    DecriptedVerifier = DecryptBytes(EncryptedVerifier, Decryptor, -1);
                    DecriptedVerifierHash = DecryptBytes(EncryptedVerifierHash, Decryptor, -1);
                }
            }

            using (HashAlgorithm hasher = TEncryptionKey.CreateHasher())
            {
                byte[] DecriptedVerifierHash2 = hasher.ComputeHash(DecriptedVerifier);
                if (!FlxUtils.CompareMem(DecriptedVerifierHash, 0, DecriptedVerifierHash2, 0, VerifierHashSizeBytes))
                {
                    return false;
                }
            }

            return true;

        }
            
    }

    internal class TAgileEncryptionVerifier : TEncryptionVerifier
    {
        public byte[] EncryptedVerifierHashValue;
        public byte[] EncryptedVerifierHashInput;
        public byte[] EncryptedKeyValue;

        internal override bool VerifyPass(string Password, TEncryptionParameters EncParams, TEncryptionKey Key)
        {
            if (!base.VerifyPass(Password, EncParams, Key)) return false;

            using (AesManaged Engine = TEncryptionUtils.CreateEngine(EncParams))
            {
                byte[] DecriptedVerifierHashInput;
                Key.CalcKey(TAgileEncryptionKey.VerifierHashInputBlockKey, null);
                using (ICryptoTransform Decryptor = Engine.CreateDecryptor(Key.Key, Key.IV))
                {
                    DecriptedVerifierHashInput = DecryptBytes(EncryptedVerifierHashInput, Decryptor, Key.Salt.Length); //this is the value padded to a blocksize multiple. We want only the Salt.Length initial bytes.
                    DecriptedVerifierHashInput = Key.Hash(DecriptedVerifierHashInput);
                }

                byte[] DecriptedVerifierHashValue;
                Key.CalcKey(TAgileEncryptionKey.VerifierHashValueBlockKey, null);
                using (ICryptoTransform Decryptor = Engine.CreateDecryptor(Key.Key, Key.IV))
                {
                    DecriptedVerifierHashValue = DecryptBytes(EncryptedVerifierHashValue, Decryptor, DecriptedVerifierHashInput.Length); //this is the 20 byte value of the hash + 12 "0" so it goes up to 32. (32 is 2*blocksize)
                }

                if (!FlxUtils.CompareMem(DecriptedVerifierHashValue, DecriptedVerifierHashInput))
                {
                    return false;
                }

                byte[] DecriptedKeyValue;
                Key.CalcKey(TAgileEncryptionKey.VerifierKeyValueBlockKey, null);
                using (ICryptoTransform Decryptor = Engine.CreateDecryptor(Key.Key, Key.IV))
                {
                    DecriptedKeyValue = DecryptBytes(EncryptedKeyValue, Decryptor, Key.KeySizeInBytes);
                }

                Key.Key = DecriptedKeyValue;

            }

            return true;

        }
    }

    internal static class EncryptedDocReader
    {
        internal static bool IsValidFile(Stream aStream)
        {
            using (TOle2File DataStream = new TOle2File(aStream, true))
            {
                if (DataStream.NotXls97) return false;
                if (!DataStream.SelectStream(XlsxConsts.EncryptionInfoString, true)) return false;
                if (!DataStream.SelectStream(XlsxConsts.ContentString, true)) return false;
            }

            return true;
        }

        internal static Stream Decrypt(Stream aStream, TEncryptionData Encryption)
        {
            using (TOle2File DataStream = new TOle2File(aStream, false))
            {
                DataStream.SelectStream(XlsxConsts.EncryptionInfoString);
                byte[] RecordHeader = new byte[4 * 2];
                DataStream.Read(RecordHeader, RecordHeader.Length);
                int vMajor = BitOps.GetWord(RecordHeader, 0);
                int vMinor = BitOps.GetWord(RecordHeader, 2);


                if ((vMajor == 0x03 || vMajor == 0x04) && vMinor == 0x02)
                {
                    long Flags = BitOps.GetCardinal(RecordHeader, 4);
                    if (Flags == 0x10) XlsMessages.ThrowException(XlsErr.ErrNotSupportedEncryption); //external encryption

                    return ReadStandardEncryptionInfo(DataStream, Encryption);
                }
                else if (vMajor == 4 && vMinor == 4 && BitOps.GetCardinal(RecordHeader, 4) == 0x040)
                {
                    return ReadAgileEncryptionInfo(DataStream, Encryption);
                }

                XlsMessages.ThrowException(XlsErr.ErrNotSupportedEncryption);
                return null;
            }
        }

        private static Stream ReadAgileEncryptionInfo(TOle2File DataStream, TEncryptionData Encryption)
        {
            byte[] Enc = new byte[DataStream.Length - DataStream.Position];
            DataStream.Read(Enc, Enc.Length);
            
            TEncryptionParameters DataEncParams = new TEncryptionParameters();
            TAgileEncryptionKey DataKey = new TAgileEncryptionKey();
            
            TAgileEncryptionVerifier KeyVerifier = new TAgileEncryptionVerifier();
            TEncryptionParameters KeyEncParams = new TEncryptionParameters();
            TAgileEncryptionKey KeyKey = new TAgileEncryptionKey();

            using (MemoryStream ms = new MemoryStream(Enc))
            {
                using (XmlReader xml = XmlReader.Create(ms))
                {
                    xml.ReadStartElement("encryption"); //goes to keyData

                    ReadAgileCipherParams(xml, DataEncParams, DataKey);

                    xml.ReadStartElement("keyData"); //goes to dataIntegrity

                    //We are not checking data integrity at the moment.
                    //DataIntegrity.EncryptedHMacKey = Convert.FromBase64String(xml.GetAttribute("encryptedHmacKey"));
                    //DataIntegrity.EncryptedHmacValue = Convert.FromBase64String(xml.GetAttribute("encryptedHmacValue"));

                    xml.ReadStartElement("dataIntegrity"); //goes to keyEncryptors
                    xml.ReadStartElement("keyEncryptors"); //goes to keyEncryptor
                    xml.ReadStartElement("keyEncryptor"); //goes to encryptedKey 
                    KeyKey.SpinCount = Convert.ToInt32(xml.GetAttribute("spinCount"), CultureInfo.InvariantCulture);
                    ReadAgileCipherParams(xml, KeyEncParams, KeyKey);
                    KeyVerifier.EncryptedVerifierHashInput = Convert.FromBase64String(xml.GetAttribute("encryptedVerifierHashInput"));
                    KeyVerifier.EncryptedVerifierHashValue = Convert.FromBase64String(xml.GetAttribute("encryptedVerifierHashValue"));
                    KeyVerifier.EncryptedKeyValue = Convert.FromBase64String(xml.GetAttribute("encryptedKeyValue"));
                }
            }

            CheckPassword(Encryption, KeyVerifier, KeyEncParams, KeyKey);
            DataKey.Key = KeyKey.Key;
            DataKey.Password = KeyKey.Password;
            DataKey.CalcDataIV(0);
            return DecryptStream(DataStream, DataEncParams, DataKey);
        }

        private static void ReadAgileCipherParams(XmlReader xml, TEncryptionParameters EncParams, TEncryptionKey Key)
        {
            int SaltSize = Convert.ToInt32(xml.GetAttribute("saltSize"), CultureInfo.InvariantCulture);
            int BlockSize = Convert.ToInt32(xml.GetAttribute("blockSize"), CultureInfo.InvariantCulture);
            Key.BlockSize = BlockSize;
            if (BlockSize != 0x10) XlsMessages.ThrowException(XlsErr.ErrNotSupportedEncryption);

            int KeyBits = Convert.ToInt32(xml.GetAttribute("keyBits"), CultureInfo.InvariantCulture);
            Key.KeySizeInBytes = KeyBits / 8;

            switch (KeyBits)
            {
                case 128: EncParams.Algorithm = TEncryptionAlgorithm.AES_128; break;
                case 192: EncParams.Algorithm = TEncryptionAlgorithm.AES_192; break;
                case 256: EncParams.Algorithm = TEncryptionAlgorithm.AES_256; break;

                default:
                    XlsMessages.ThrowException(XlsErr.ErrNotSupportedEncryption);
                    break;
            }
            string CipherAlgo = xml.GetAttribute("cipherAlgorithm");
            if (!string.Equals(CipherAlgo, "AES", StringComparison.InvariantCulture)) XlsMessages.ThrowException(XlsErr.ErrNotSupportedEncryption);

            string CipherChaining = xml.GetAttribute("cipherChaining");
            switch (CipherChaining)
            {
                case "ChainingModeCBC": EncParams.ChainingMode = CipherMode.CBC; break;
                case "ChainingModeCFB": EncParams.ChainingMode = CipherMode.CFB; break;
                default:
                    XlsMessages.ThrowException(XlsErr.ErrNotSupportedEncryption); break;
            }

            string HashAlgorithm = xml.GetAttribute("hashAlgorithm");
            if (HashAlgorithm != "SHA1" && HashAlgorithm != "SHA-1") XlsMessages.ThrowException(XlsErr.ErrNotSupportedEncryption);

            Key.Salt = Convert.FromBase64String(xml.GetAttribute("saltValue"));
            if (Key.Salt == null || SaltSize != Key.Salt.Length) XlsMessages.ThrowException(XlsErr.ErrNotSupportedEncryption);
        }

        private static Stream ReadStandardEncryptionInfo(TOle2File DataStream, TEncryptionData Encryption)
        {
            byte[] RecordHeaderLen = new byte[4];
            DataStream.Read(RecordHeaderLen, RecordHeaderLen.Length);

            long EncryptionHeaderSize = BitOps.GetCardinal(RecordHeaderLen, 0);

            byte[] EncryptionHeader = new byte[EncryptionHeaderSize];
            DataStream.Read(EncryptionHeader, EncryptionHeader.Length);
            long AlgId = BitOps.GetCardinal(EncryptionHeader, 8);
            long KeyBits = BitOps.GetCardinal(EncryptionHeader, 16);

            TEncryptionParameters EncParams = TEncryptionParameters.CreateStandard(GetStandardEncAlg(AlgId));

            byte[] VerifierBytes = new byte[DataStream.Length - DataStream.Position];
            DataStream.Read(VerifierBytes, VerifierBytes.Length);
            TStandardEncryptionVerifier Verifier = ReadStandardVerifier(VerifierBytes);
            TEncryptionKey Key = new TStandardEncryptionKey(ReadStandardSalt(VerifierBytes), (int)KeyBits / 8);

            CheckPassword(Encryption, Verifier, EncParams, Key);

            return DecryptStream(DataStream, EncParams, Key);

        }

        private static byte[] ReadStandardSalt(byte[] VerifierBytes)
        {
            int SaltSize = (int)BitOps.GetCardinal(VerifierBytes, 0);
            byte[] Salt = new byte[SaltSize];
            Array.Copy(VerifierBytes, 4, Salt, 0, Salt.Length);

            return Salt;
        }

        private static TEncryptionAlgorithm GetStandardEncAlg(long AlgId)
        {
            switch (AlgId)
            {
                case 0x0000660E: return TEncryptionAlgorithm.AES_128;
                case 0x0000660F: return TEncryptionAlgorithm.AES_192;
                case 0x00006610: return TEncryptionAlgorithm.AES_256;
            }

            XlsMessages.ThrowException(XlsErr.ErrFileIsNotSupported);
            return TEncryptionAlgorithm.AES_256; //just to compile
        }

        private static TStandardEncryptionVerifier ReadStandardVerifier(byte[] Verifier)
        {
            TStandardEncryptionVerifier Result = new TStandardEncryptionVerifier();

            Result.EncryptedVerifier = new byte[16];
            Array.Copy(Verifier, 20, Result.EncryptedVerifier, 0, Result.EncryptedVerifier.Length);

            Result.VerifierHashSizeBytes = (int)BitOps.GetCardinal(Verifier, 20 + 16) / 8;
            Result.EncryptedVerifierHash = new byte[32];
            Array.Copy(Verifier, 20 + 16 + 4, Result.EncryptedVerifierHash, 0, Result.EncryptedVerifierHash.Length);

            return Result;
        }


        private static void CheckPassword(TEncryptionData Encryption, TEncryptionVerifier Verifier, TEncryptionParameters EncParams, TEncryptionKey Key)
        {
            if (Verifier.VerifyPass(XlsConsts.EmptyExcelPassword, EncParams, Key))
            {
                return; //workbook password protected
            }
            
            string Password = Encryption.ReadPassword;
            if (Encryption.OnPassword != null)
            {
                OnPasswordEventArgs ea = new OnPasswordEventArgs(Encryption.Xls);
                Encryption.OnPassword(ea);
                Encryption.ReadPassword = ea.Password;
                Password = ea.Password;
            }

            if (!Verifier.VerifyPass(Password, EncParams, Key))
            {
                XlsMessages.ThrowException(XlsErr.ErrInvalidPassword);
            }
        }

        private static Stream DecryptStream(TOle2File DataStream, TEncryptionParameters EncParams, TEncryptionKey Key)
        {
            DataStream.SelectStream(XlsxConsts.ContentString);
            byte[] EncryptedSize = new byte[8];
            DataStream.Read(EncryptedSize, EncryptedSize.Length);
            AesManaged Engine = null;
            ICryptoTransform Decryptor = null;
            try
            {
                Engine = TEncryptionUtils.CreateEngine(EncParams);
                if (!Key.VariableIV)Decryptor = Engine.CreateDecryptor(Key.Key, Key.IV);
                return new TXlsxCryptoStreamReader(DataStream, BitOps.GetCardinal(EncryptedSize, 0), Engine, Decryptor, Key);
            }
            catch
            {
                if (Engine != null) ((IDisposable)Engine).Dispose();
                if (Decryptor != null) Decryptor.Dispose();
                throw;
            }

        }
    }

    internal class TXlsxCryptoStreamReader : Stream
    {
        const int SegmentSize = TEncryptionUtils.AgileSegmentSize; //required by Agile, in standard it could be other number. It could be any multiple too.
        const int DataStreamStart = 8; //first 8 bytes are the size.
        long FStreamLen;
        long FPosition;
        byte[] CurrentSegment;
        long CurrentSegmentNo = -1;
        TOle2File DataStream;
        ICryptoTransform Decryptor;
        AesManaged EncEngine;
        TEncryptionKey Key;

        public TXlsxCryptoStreamReader(TOle2File aDataStream, long aStreamLen, AesManaged aEncEncgine, ICryptoTransform aDecryptor, TEncryptionKey aKey)
        {
            DataStream = aDataStream;
            FStreamLen = aStreamLen;
            EncEngine = aEncEncgine;
            Decryptor = aDecryptor;
            CurrentSegment = new byte[SegmentSize];
            Key = aKey;
        }

        public override bool CanRead
        {
            get { return true; }
        }

        public override bool CanSeek
        {
            get { return true; }
        }

        public override bool CanWrite
        {
            get { return false; }
        }

        public override void Flush()
        {
            
        }

        public override long Length
        {
            get { return FStreamLen; }
        }

        public override long Position
        {
            get
            {
                return FPosition;
            }
            set
            {
                if (value < 0) throw new IOException(XlsMessages.GetString(XlsErr.ErrExcelInvalid)); 
                if (value > Length) throw new IOException(XlsMessages.GetString(XlsErr.ErrEofReached, FPosition - Length));
                FPosition = value;
            }
        }

        private void ReadSegment(long SegmentNumber)
        {
            if (SegmentNumber == CurrentSegmentNo) return;
            DataStream.Position = DataStreamStart + SegmentNumber * SegmentSize;

            byte[] CurrentSegmentEnc = new byte[Math.Min(SegmentSize, DataStream.Length - DataStream.Position)];
            DataStream.Read(CurrentSegmentEnc, CurrentSegmentEnc.Length);

            using (MemoryStream msDecrypt = new MemoryStream(CurrentSegmentEnc))
            {
                ICryptoTransform RealDecryptor = Decryptor;
                try
                {
                    if (RealDecryptor == null)
                    {
                        Key.CalcDataIV(SegmentNumber);
                        RealDecryptor = EncEngine.CreateDecryptor(Key.Key, Key.IV);
                    }
                    using (CryptoStream cs = new CryptoStream(msDecrypt, RealDecryptor, CryptoStreamMode.Read))
                    {
                        cs.Read(CurrentSegment, 0, CurrentSegmentEnc.Length);
                    }
                }
                finally
                {
                    if (RealDecryptor != Decryptor) RealDecryptor.Dispose();
                }
            }

            CurrentSegmentNo = SegmentNumber;
        }

        private static long CalcSegmentNo(long posi)
        {
            return posi / SegmentSize;
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            if (buffer == null) return 0;
            long aPos = FPosition;
            FPosition += count; //first so it checks it is valid.
            
            long StartSeg = CalcSegmentNo(aPos);
            long destoffs = offset;
            int BytesRead = 0;
            while (BytesRead < count)
            {
                ReadSegment(StartSeg);
                int BytesToRead = count - BytesRead;
                int csegOffs = (int)aPos % SegmentSize;
                if (csegOffs + BytesToRead > CurrentSegment.Length) BytesToRead = CurrentSegment.Length - csegOffs;
                Array.Copy(CurrentSegment, csegOffs, buffer, destoffs, BytesToRead);
                BytesRead += BytesToRead;
                destoffs += BytesToRead;
                aPos += BytesToRead;

                StartSeg++;
            }

            return count;
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            switch (origin)
            {
                case SeekOrigin.Begin:
                    Position = offset;
                    break;
                case SeekOrigin.Current:
                    Position += offset;
                    break;
                case SeekOrigin.End:
                    Position = Length - offset;
                    break;
                default:
                    break;
            }
            return Position;
        }

        public override void SetLength(long value)
        {
            //We can't write in a reader
            throw new NotImplementedException();
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            //We can't write in a reader
            throw new NotImplementedException();
        }

        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);
            if (disposing)
            {
                if (Decryptor != null)
                {
                    Decryptor.Dispose();
                    Decryptor = null;
                }
                if (EncEngine != null)
                {
                    ((IDisposable)EncEngine).Dispose();
                    EncEngine = null;
                }
            }
        }
    }

    internal class TXlsxCryptoStreamWriter: Stream
    {
        Stream WorkingStream;
        long WorkStreamZeroPos;
        TProtection Protection;
        Stream TargetStream;

        private static Stream GetEmptyEncryptedFile()
        {
            return Assembly.GetExecutingAssembly().GetManifestResourceStream("FlexCel.XlsAdapter.EmptyEncryptedFile.xlsx");
        }

        internal TXlsxCryptoStreamWriter(Stream aTargetStream, TProtection aProtection)
        {
            WorkingStream = new MemoryStream(); // we could use a TOle2File directly here, and then TOle2File.GetStreamForWriting(), but as we need to switch the stream at the end anyway to write the DataValidation info, we would have the tream in mem twice anyway.
            WorkStreamZeroPos = 8;
            Protection = aProtection;
            TargetStream = aTargetStream;
        }

        #region Clean up
        private void FinishEncryption()
        {
            TEncryptionParameters EncParams = GetEncryptionParams(Protection);
            AesManaged EncEngine = TEncryptionUtils.CreateEngine(EncParams);

            TAgileEncryptionKey DataKey = TAgileEncryptionKey.CreateForWriting(null, EncEngine.KeySize / 8);
            DataKey.Key = TEncryptionUtils.GetRandom(DataKey.KeySizeInBytes);

            TAgileEncryptionKey KeyKey = TAgileEncryptionKey.CreateForWriting(Protection.OpenPassword, EncEngine.KeySize / 8);

            byte[] WorkLen = new byte[8];
            BitOps.SetCardinal(WorkLen, 0, WorkingStream.Length - WorkStreamZeroPos);

            EncryptStream(EncEngine, DataKey, WorkLen);

            using (MemoryStream ms = new MemoryStream())
            {
                using (TOle2File Ole2File = new TOle2File(GetEmptyEncryptedFile()))
                {
                    Ole2File.PrepareForWrite(ms, XlsxConsts.EncryptionInfoString, new string[0]);
                    CreateInfoStream(Ole2File, EncEngine, EncParams, KeyKey, DataKey);
                }
                
                ms.Position = 0;

                using (TOle2File Ole2File = new TOle2File(ms))
                {
                    Ole2File.PrepareForWrite(TargetStream, XlsxConsts.ContentString, new string[0]);
                    WorkingStream.Position = 0;
                    TEncryptionUtils.CopyStream(WorkingStream, Ole2File);
                }
            }
        }

        private void EncryptStream(AesManaged EncEngine, TEncryptionKey DataKey, byte[] WorkLen)
        {
            WorkingStream.Position = 0;
            WorkingStream.Write(WorkLen, 0, WorkLen.Length);
            int read;
            byte[] CurrentSegment = new byte[TEncryptionUtils.AgileSegmentSize];
            int SegmentNumber = 0;
            while ((read = WorkingStream.Read(CurrentSegment, 0, CurrentSegment.Length)) > 0)
            {
                DataKey.CalcDataIV(SegmentNumber);
                SegmentNumber++;

                using (ICryptoTransform Encryptor = EncEngine.CreateEncryptor(DataKey.Key, DataKey.IV))
                {
                    using (MemoryStream outms = new MemoryStream()) //will be closed by Cryptostream
                    {
                        using (CryptoStream cs = new CryptoStream(outms, Encryptor, CryptoStreamMode.Write))
                        {
                            cs.Write(CurrentSegment, 0, read);
                        }

                        WorkingStream.Position -= read;
                        byte[] outArr = outms.ToArray(); //ToArray works with stream closed, and cs will close the stream.
                        WorkingStream.Write(outArr, 0, outArr.Length);
                    }
                }
            }
        }
        
        
        private static TEncryptionParameters GetEncryptionParams(TProtection Protection)
        {
            TEncryptionParameters Result = new TEncryptionParameters();
            Result.Algorithm = Protection.EncryptionAlgorithmXlsx;
            Result.ChainingMode = CipherMode.CBC;
            Result.Padding = PaddingMode.Zeros;
            return Result;
        }

        private void CreateInfoStream(TOle2File DataStream, AesManaged EncEngine, TEncryptionParameters EncParams, TAgileEncryptionKey KeyKey, TAgileEncryptionKey DataKey)
        {
            DataStream.Write16(0x0004);
            DataStream.Write16(0x0004);
            DataStream.Write32(0x00040);

            byte[] InfoStreamXml = GetInfoStreamXml(EncEngine, EncParams, KeyKey, DataKey);
            DataStream.Write(InfoStreamXml, InfoStreamXml.Length);
            byte[] pad = new byte[4098 - InfoStreamXml.Length]; //Our Ole2 implementation will fill a sector with 0 so it doesn't go to the ministream. Those 0 will confuse Excel, so we will write spaces.
            for (int i = 0; i < pad.Length; i++)
            {
                pad[i] = 32;
            }
            DataStream.Write(pad, pad.Length);
        }

        private byte[] GetInfoStreamXml(AesManaged EncEngine, TEncryptionParameters EncParams, TAgileEncryptionKey KeyKey, TAgileEncryptionKey DataKey)
        {
            const string KeyEncryptorPasswordNamespace = "http://schemas.microsoft.com/office/2006/keyEncryptor/password";
            using (MemoryStream ms = new MemoryStream())
            {               
                XmlWriterSettings Settings = new XmlWriterSettings();
                Settings.Encoding = new System.Text.UTF8Encoding(false);
                using (XmlWriter xml = XmlWriter.Create(ms, Settings))
                {
                    xml.WriteStartDocument(true);
                    xml.WriteStartElement("encryption", "http://schemas.microsoft.com/office/2006/encryption");
                    xml.WriteAttributeString("xmlns", "p", null, KeyEncryptorPasswordNamespace); 
                    xml.WriteStartElement("keyData");
                    WriteAgileCipherParams(xml, EncParams, DataKey);
                    xml.WriteEndElement();

                    xml.WriteStartElement("dataIntegrity");

                    byte[] HMacKeyBlockKey = new byte[] { 0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6 };
                    DataKey.CalcDataIV(HMacKeyBlockKey);
                    byte[] HMacKey = TEncryptionUtils.GetRandom((int)DataKey.HashSizeBytes());
                    using (ICryptoTransform Encryptor = EncEngine.CreateEncryptor(DataKey.Key, DataKey.IV))
                    {
                        byte[] EncryptedHMacKey = TAgileEncryptionVerifier.EncryptBytes(HMacKey, Encryptor, DataKey.BlockSize);
                        xml.WriteAttributeString("encryptedHmacKey", Convert.ToBase64String(EncryptedHMacKey));
                    }
                    
                    HMAC HMacCalc = new HMACSHA1(HMacKey);
                    WorkingStream.Position = 0;
                    byte[] HMac = HMacCalc.ComputeHash(WorkingStream);
                    byte[] HMacValBlockKey = new byte[] { 0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33 };
                    DataKey.CalcDataIV(HMacValBlockKey);
                    using (ICryptoTransform Encryptor = EncEngine.CreateEncryptor(DataKey.Key, DataKey.IV))
                    {
                        byte[] EncryptedHMacValue = TAgileEncryptionVerifier.EncryptBytes(HMac, Encryptor, DataKey.BlockSize);
                        xml.WriteAttributeString("encryptedHmacValue", Convert.ToBase64String(EncryptedHMacValue));
                    }

                    xml.WriteEndElement();
                    xml.WriteStartElement("keyEncryptors");
                    xml.WriteStartElement("keyEncryptor");
                    xml.WriteAttributeString("uri", KeyEncryptorPasswordNamespace);
                    xml.WriteStartElement("encryptedKey", KeyEncryptorPasswordNamespace);


                    xml.WriteAttributeString("spinCount", Convert.ToString(KeyKey.SpinCount, CultureInfo.InvariantCulture));
                    WriteAgileCipherParams(xml, EncParams, KeyKey);

                    byte[] RandData = TEncryptionUtils.GetRandom(KeyKey.Salt.Length);
                    KeyKey.CalcKey(TAgileEncryptionKey.VerifierHashInputBlockKey, null);
                    using (ICryptoTransform Encryptor = EncEngine.CreateEncryptor(KeyKey.Key, KeyKey.IV))
                    {
                        byte[] EncryptedVerifierHashInput = TAgileEncryptionVerifier.EncryptBytes(RandData, Encryptor, KeyKey.BlockSize);
                        xml.WriteAttributeString("encryptedVerifierHashInput", Convert.ToBase64String(EncryptedVerifierHashInput));
                    }

                    
                    KeyKey.CalcKey(TAgileEncryptionKey.VerifierHashValueBlockKey, null);
                    using (ICryptoTransform Encryptor = EncEngine.CreateEncryptor(KeyKey.Key, KeyKey.IV))
                    {
                        byte[] EncryptedVerifierHashValue = TAgileEncryptionVerifier.EncryptBytes(KeyKey.Hash(RandData), Encryptor, KeyKey.BlockSize);
                        xml.WriteAttributeString("encryptedVerifierHashValue", Convert.ToBase64String(EncryptedVerifierHashValue));
                    }

                    KeyKey.CalcKey(TAgileEncryptionKey.VerifierKeyValueBlockKey, null);
                    using (ICryptoTransform Encryptor = EncEngine.CreateEncryptor(KeyKey.Key, KeyKey.IV))
                    {
                        byte[] EncryptedKeyValue = TAgileEncryptionVerifier.EncryptBytes(DataKey.Key, Encryptor, KeyKey.BlockSize);
                        xml.WriteAttributeString("encryptedKeyValue", Convert.ToBase64String(EncryptedKeyValue));
                    }
                    xml.WriteEndElement();
                    xml.WriteEndElement();
                    xml.WriteEndElement();

                    xml.WriteEndElement();
                    xml.WriteEndDocument();
                }
                return ms.ToArray();
            }
        }

        private static void WriteAgileCipherParams(XmlWriter xml, TEncryptionParameters EncParams, TAgileEncryptionKey Key)
        {
            xml.WriteAttributeString("saltSize", Convert.ToString(Key.Salt.Length, CultureInfo.InvariantCulture));
            xml.WriteAttributeString("blockSize", Convert.ToString(Key.BlockSize, CultureInfo.InvariantCulture));
            xml.WriteAttributeString("keyBits", Convert.ToString(Key.KeySizeInBytes * 8, CultureInfo.InvariantCulture));
            xml.WriteAttributeString("hashSize", Convert.ToString(Key.HashSizeBytes(), CultureInfo.InvariantCulture)); //sha1 hash size

            xml.WriteAttributeString("cipherAlgorithm", "AES");
            switch (EncParams.ChainingMode)
            {
                case CipherMode.CBC:
                    xml.WriteAttributeString("cipherChaining", "ChainingModeCBC");
                    break;

                case CipherMode.CFB:
                    xml.WriteAttributeString("cipherChaining", "ChainingModeCFB");
                    break;

                case CipherMode.CTS:
                case CipherMode.ECB:
                case CipherMode.OFB:
                default:
                    XlsMessages.ThrowException(XlsErr.ErrNotSupportedEncryption); break;
            }

            xml.WriteAttributeString("hashAlgorithm", "SHA1");
            xml.WriteAttributeString("saltValue", Convert.ToBase64String(Key.Salt));
        }
        #endregion

        #region Stream
        public override bool CanRead
        {
            get { return WorkingStream.CanRead; }
        }

        public override bool CanSeek
        {
            get { return WorkingStream.CanSeek; }
        }

        public override bool CanWrite
        {
            get { return WorkingStream.CanWrite; }
        }

        public override void Flush()
        {
            WorkingStream.Flush();
        }

        public override long Length
        {
            get { long Result = WorkingStream.Length - WorkStreamZeroPos; if (Result < 0) return 0; else return Result; }
        }

        public override long Position
        {
            get
            {
                return WorkingStream.Position - WorkStreamZeroPos;
            }
            set
            {
                WorkingStream.Position = value + WorkStreamZeroPos;
            }
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            return WorkingStream.Read(buffer, offset, count);
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            switch (origin)
            {
                case SeekOrigin.Begin:
                    return WorkingStream.Seek(offset + WorkStreamZeroPos, origin);

                case SeekOrigin.Current:
                    return WorkingStream.Seek(offset, origin);

                case SeekOrigin.End:
                    return WorkingStream.Seek(offset, origin);
            }

            return Position;
        }

        public override void SetLength(long value)
        {
            WorkingStream.SetLength(value + WorkStreamZeroPos);
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            WorkingStream.Write(buffer, offset, count);
        }

        public override int ReadByte()
        {
            return WorkingStream.ReadByte();
        }

        public override void WriteByte(byte value)
        {
            WorkingStream.WriteByte(value);
        }
        #endregion

        #region IDisposable
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                FinishEncryption();
                if (WorkingStream != null)
                {
                    WorkingStream.Dispose();
                    WorkingStream = null;
                }
            }
            base.Dispose(disposing);
        }

        #endregion

    }

}
