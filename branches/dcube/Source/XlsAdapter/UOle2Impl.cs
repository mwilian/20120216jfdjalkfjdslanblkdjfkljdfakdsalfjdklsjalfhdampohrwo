using System;
using System.Text;
using System.IO;
using System.Diagnostics;

using FlexCel.Core;
using System.Globalization;
using System.Collections.Generic;

	/*
     * OLE2 File format implementation.
	 * This file is used by UOLE2Stream to provide an uniform access layer to OLE2 Compound documents.
	 * Honoring FlexCel tradition, this file is targeted to be "one" api to modify, instead of "two" apis, one for read and one for write. 
	 */
namespace FlexCel.XlsAdapter
{
    #region Ole2File
    internal enum STGTY
    {
        INVALID = 0,
        STORAGE = 1,
        STREAM = 2,
        LOCKBYTES = 3,
        PROPERTY = 4,
        ROOT = 5,
    }
    internal enum DECOLOR
    {
        RED = 0,
        BLACK = 1,
    }


    /// <summary>
    /// Header sector. It has a fixed size of 512 bytes.
    /// On this implementation, we don't save the first 109 DIF entries, as they will be saved by the DIF Sector.
    /// </summary>
    internal class TOle2Header
    {
        internal bool NotXls97;
        internal byte[] Data;

        internal const int HeaderSize = 512; //This is fixed on the header sector.
        
        /// <summary>
        /// 109
        /// </summary>
        internal const int DifsInHeader = 109;
        
        /// <summary>
        /// 109*4
        /// </summary>
        internal const int DifEntries = DifsInHeader * 4;  //Difs don't really belong here. 
        internal const UInt32 ENDOFCHAIN = 0xFFFFFFFE;
        internal const UInt32 DIFSECT = 0xFFFFFFFC;
        internal const UInt32 FATSECT = 0xFFFFFFFD;
        internal const UInt32 FREESECT = 0xFFFFFFFF;
        internal long StartOfs;

        private readonly byte[] FileSignature = new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };

        #region Cache
        internal UInt32 SectorSize;
        internal int uSectorShift;
        internal UInt32 ulMiniSectorCutoff;
        #endregion

        /// <summary>
        /// Creates the Header reading the data from a stream.
        /// </summary>
        /// <param name="aStream"></param>
        /// <param name="AvoidExceptions"></param>
        internal TOle2Header(Stream aStream, bool AvoidExceptions)
        {
            StartOfs = aStream.Position;
            Data = new byte[HeaderSize - DifEntries];
            if (aStream.Length - StartOfs < Data.Length)
            {
                if (AvoidExceptions)
                {
                    NotXls97 = true;
                    return;
                }
                throw new IOException(XlsMessages.GetString(XlsErr.ErrFileIsNotXLS));
            }
            Sh.Read(aStream, Data, 0, Data.Length, false);
            if (!CompareArray(Data, FileSignature, FileSignature.Length))
            {
                if (AvoidExceptions)
                {
                    NotXls97 = true;
                    return;
                }
                throw new IOException(XlsMessages.GetString(XlsErr.ErrFileIsNotXLS));
            }

            uSectorShift = FuSectorShift;
            SectorSize = FSectorSize;
            ulMiniSectorCutoff = FulMiniSectorCutoff;
        }

        internal void Save(Stream aStream)
        {
            Sh.Write(aStream, Data, 0, Data.Length);
        }

        private static bool CompareArray(byte[] a1, byte[] a2, int length)
        {
            for (int i = 0; i < length; i++) if (a1[i] != a2[i]) return false;
            return true;
        }

        internal int uDIFEntryShift()
        {
            return (uSectorShift - 2);
        }

        internal long SectToStPos(long Sect)
        {
            return (Sect << uSectorShift) + HeaderSize + StartOfs;
        }
        internal long SectToStPos(long Sect, long Ofs)
        {
            return (Sect << uSectorShift) + HeaderSize + Ofs;
        }

        private Int32 FuSectorShift { get { return BitConverter.ToUInt16(Data, 0x001E); } }  //UInt16 has a bug with mono

        private UInt32 FSectorSize { get { return ((UInt32)1 << FuSectorShift); } }
        internal Int32 uMiniSectorShift { get { return BitConverter.ToUInt16(Data, 0x0020); } }
        internal UInt32 MiniSectorSize { get { return ((UInt32)1 << uMiniSectorShift); } }
        internal UInt32 csectDir { get { return BitConverter.ToUInt32(Data, 0x0028); } set { BitConverter.GetBytes(value).CopyTo(Data, 0x0028); } }
        
        internal UInt32 csectFat { get { return BitConverter.ToUInt32(Data, 0x002C); } set { BitConverter.GetBytes(value).CopyTo(Data, 0x002C); } }
        internal UInt32 sectDirStart { get { return BitConverter.ToUInt32(Data, 0x0030); } set { BitConverter.GetBytes(value).CopyTo(Data, 0x0030); } }

        internal UInt32 FulMiniSectorCutoff { get { return BitConverter.ToUInt32(Data, 0x0038); } }
        internal UInt32 sectMiniFatStart { get { return BitConverter.ToUInt32(Data, 0x003C); } set { BitConverter.GetBytes(value).CopyTo(Data, 0x003C); } }
        internal UInt32 csectMiniFat { get { return BitConverter.ToUInt32(Data, 0x0040); } set { BitConverter.GetBytes(value).CopyTo(Data, 0x0040); } }
        internal UInt32 sectDifStart { get { return BitConverter.ToUInt32(Data, 0x0044); } set { BitConverter.GetBytes(value).CopyTo(Data, 0x0044); } }
        internal UInt32 csectDif { get { return BitConverter.ToUInt32(Data, 0x0048); } set { BitConverter.GetBytes(value).CopyTo(Data, 0x0048); } }

    }

    /// <summary>
    /// FAT Table stored as a list of ints.
    /// </summary>
    internal class TOle2FAT : UInt32List
    {
        private TOle2Header Header;

        /// <summary>
        /// Use Create() to create an instance. This way we avoid calling virtual methods on a constructor.
        /// </summary>
        private TOle2FAT() { }

        /// <summary>
        /// Creates a FAT integer list from the data on a stream. 
        /// We have to read the DIF to access the actual FAT sectors.
        /// </summary>
        /// <param name="aHeader">The header record</param>
        /// <param name="aStream">Stream to read the FAT. When null, an Empty FAT will be created.</param>
        internal static TOle2FAT Create(TOle2Header aHeader, Stream aStream)
        {
            TOle2FAT Result = new TOle2FAT();
            Result.Header = aHeader;
            if (aStream != null)
            {
                Result.Capacity = (int)(aHeader.csectFat << Result.uFATEntryShift()) + TOle2Header.DifsInHeader + 16;  //This is, number of fat sectors*(SectorSize/4)+109+ extra_just_in_case

                byte[] DifSect0 = new byte[TOle2Header.DifEntries];
                aStream.Seek(Result.Header.StartOfs + TOle2Header.HeaderSize - TOle2Header.DifEntries, SeekOrigin.Begin);
                Sh.Read(aStream, DifSect0, 0, DifSect0.Length, false);

                Result.LoadDifSector(DifSect0, 0, TOle2Header.DifEntries, aStream); //First 109 DIF records are on the header.

                //if there are Dif sectors, load them.
                UInt32 DifPos = Result.Header.sectDifStart;
                byte[] DifSect = new byte[Result.Header.SectorSize];
                for (UInt32 i = 0; i < Result.Header.csectDif; i++)
                {
                    if (DifPos == TOle2Header.ENDOFCHAIN) throw new IOException(XlsMessages.GetString(XlsErr.ErrExcelInvalid));

                    aStream.Seek(Result.Header.SectToStPos(DifPos), SeekOrigin.Begin);
                    Sh.Read(aStream, DifSect, 0, DifSect.Length, false);
                    Result.LoadDifSector(DifSect, 0, Result.Header.SectorSize - 4, aStream);
                    DifPos = BitConverter.ToUInt32(DifSect, (int)Result.Header.SectorSize - 4);
                }
            }
            //Some sanity checks
            //not really... sometimes it is not. if (DifPos!=TOle2Header.ENDOFCHAIN) throw new IOException(XlsMessages.GetString(XlsErr.ErrExcelInvalid));

            return Result;
        }

        internal int uFATEntryShift()
        {
            return (Header.uSectorShift - 2);
        }

        internal long GetNextSector(long Sect)
        {
            return this[(int)Sect];
        }


        private long LastFindSectorOfs = -1;
        private long LastFindSectorStart = -1;
        private long LastFindSectorRes = 0;
        internal long FindSector(long StartSect, long SectOfs)
        {
            long NewSect = StartSect;
            long RealSectOfs = SectOfs;
            if ((LastFindSectorStart == StartSect) && (SectOfs >= LastFindSectorOfs))  //Optimization for sequential read.
            {
                NewSect = LastFindSectorRes;
                RealSectOfs -= LastFindSectorOfs;
            }

            for (int i = 0; i < RealSectOfs; i++)
            {
                NewSect = this[(int)NewSect];
            }

            LastFindSectorStart = StartSect;
            LastFindSectorOfs = SectOfs;
            LastFindSectorRes = NewSect;
            return NewSect;
        }

        private void LoadDifSector(byte[] data, UInt32 inipos, UInt32 endpos, Stream aStream)
        {
            byte[] FatSect = new byte[Header.SectorSize];
            int FatEntries = 1 << uFATEntryShift();

            for (UInt32 i = inipos; i < endpos; i += 4)
            {
                UInt32 FatId = BitConverter.ToUInt32(data, (int)i);
                if (FatId == TOle2Header.ENDOFCHAIN) return;
                if (FatId == TOle2Header.FREESECT)
                {
                    //We have to keep track of the FAT position.
                    for (int k = 0; k < FatEntries; k++) Add(TOle2Header.FREESECT);
                    continue;
                }
                aStream.Seek(Header.SectToStPos(FatId), SeekOrigin.Begin);
                Sh.Read(aStream, FatSect, 0, FatSect.Length, false);
                LoadFatSector(FatSect);
            }
        }

        private void LoadFatSector(byte[] data)
        {
            UInt32 HeaderSectorSize = Header.SectorSize;
            for (int i = 0; i < HeaderSectorSize; i += 4)
            {
                UInt32 Sect = BitConverter.ToUInt32(data, i);
                //No, we have to load it the same. if (Sect== TOle2Header.FREESECT) continue;
                Add(Sect);
            }
        }
    }

    /// <summary>
    /// MINIFAT Table stored as a list of ints.
    /// </summary>
    internal class TOle2MiniFAT : UInt32List
    {
        private TOle2Header Header;

        /// <summary>
        /// Use Create() to create an instance. This way we avoid calling virtual methods on a constructor.
        /// </summary>
        private TOle2MiniFAT() { }

        /// <summary>
        /// Creates a MiniFAT integer list from the data on a stream. 
        /// </summary>
        /// <param name="aHeader"></param>
        /// <param name="aStream"></param>
        /// <param name="aFAT"></param>
        internal static TOle2MiniFAT Create(TOle2Header aHeader, Stream aStream, TOle2FAT aFAT)
        {
            TOle2MiniFAT Result = new TOle2MiniFAT();
            Result.Header = aHeader;
            Result.Capacity = (int)(aHeader.csectMiniFat << (aHeader.uSectorShift - 2)) + 16;  //This is, number of minifat sectors*(SectorSize/4)+ extra_just_in_case

            if (aStream != null)
            {
                byte[] MiniFatSect = new byte[aHeader.SectorSize];
                long MiniFatPos = aHeader.sectMiniFatStart;
                for (UInt32 i = 0; i < aHeader.csectMiniFat; i++)
                {
                    if (MiniFatPos == TOle2Header.ENDOFCHAIN) throw new IOException(XlsMessages.GetString(XlsErr.ErrExcelInvalid));

                    aStream.Seek(aHeader.SectToStPos(MiniFatPos), SeekOrigin.Begin);
                    Sh.Read(aStream, MiniFatSect, 0, MiniFatSect.Length, false);
                    Result.LoadMiniFatSector(MiniFatSect);
                    MiniFatPos = aFAT.GetNextSector(MiniFatPos);
                }
            }

            return Result;
        }

        internal long GetNextSector(long Sect)
        {
            return this[(int)Sect];
        }

        internal long FindSector(long StartSect, long SectOfs)
        {
            long NewSect = StartSect;
            for (int i = 0; i < SectOfs; i++)
            {
                NewSect = this[(int)NewSect];
            }
            return NewSect;
        }


        private void LoadMiniFatSector(byte[] data)
        {
            for (int i = 0; i < Header.SectorSize; i += 4)
            {
                UInt32 Sect = BitConverter.ToUInt32(data, i);
                //NO!  Has to be loaded anyway. if (Sect== TOle2Header.FREESECT)continue;
                Add(Sect);
            }
        }
    }

    /// <summary>
    /// A semi-sector containing 1 Directory entry. 
    /// </summary>
    internal class TOle2Directory
    {
        internal const int DirectorySize = 128;
        internal byte[] Data;
        internal long ulSize;


        internal TOle2Directory(byte[] aData)
        {
            Data = aData;
            ulSize = xulSize;
        }

        internal int NameSize
        {
            get { return GetNameSize(Data, 0); }
            set
            {
                if ((value < 0) || (value > 62)) XlsMessages.ThrowException(XlsErr.ErrTooManyEntries, value, 62);
                Data[0x0040] = (byte)(value + 2);
            }
        }
        internal string Name
        {
            get { return GetName(Data, 0); }
            set
            {
                string aValue = value.PadRight(32, '\u0000');
                NameSize = aValue.Length * 2;
                Encoding.Unicode.GetBytes(aValue, 0, aValue.Length, Data, 0);
            }
        }

        internal void Save(Stream aStream)
        {
            xulSize = ulSize;
            Sh.Write(aStream, Data, 0, Data.Length);
        }


        internal static int GetNameSize(byte[] Data, int StartPos)
        {
            int nl = Data[0x0040 + StartPos]; if ((nl < 2) || (nl > 64)) return 0; else return nl - 2;
        }
        internal static string GetName(byte[] Data, int StartPos)
        {
            return Encoding.Unicode.GetString(Data, StartPos, GetNameSize(Data, StartPos));
        }
        internal static STGTY GetType(byte[] Data, int StartPos)
        {
            return (STGTY)Data[0x0042 + StartPos];
        }
        internal static long GetSectStart(byte[] Data, int StartPos)
        {
            //return BitConverter.ToUInt32(Data, 0x0074+StartPos);
            unchecked
            {
                return (UInt32)(Data[0x0074 + StartPos] + (Data[0x0075 + StartPos] << 8) + (Data[0x0076 + StartPos] << 16) + (Data[0x0077 + StartPos] << 24));
            }
        }
        internal static void SetSectStart(byte[] Data, int StartPos, long value)
        {
            //BitConverter.GetBytes((UInt32)value).CopyTo(Data,0x0074+StartPos);
            unchecked
            {
                int tPos = 0x0074 + StartPos;
                Data[tPos] = (byte)value;
                Data[tPos + 1] = (byte)((UInt32)value >> 8);
                Data[tPos + 2] = (byte)((UInt32)value >> 16);
                Data[tPos + 3] = (byte)((UInt32)value >> 24);
            }
        }
        internal static long GetSize(byte[] Data, int StartPos)
        {
            // return BitConverter.ToUInt32(Data, 0x0078+StartPos);
            unchecked
            {
                return (UInt32)(Data[0x0078 + StartPos] + (Data[0x0079 + StartPos] << 8) + (Data[0x007A + StartPos] << 16) + (Data[0x007B + StartPos] << 24));
            }
        }
        internal static void SetSize(byte[] Data, int StartPos, long value)
        {
            //BitConverter.GetBytes((UInt32)value).CopyTo(Data,0x0078+StartPos);
            unchecked
            {
                int tPos = 0x0078 + StartPos;
                Data[tPos] = (byte)value;
                Data[tPos + 1] = (byte)((UInt32)value >> 8);
                Data[tPos + 2] = (byte)((UInt32)value >> 16);
                Data[tPos + 3] = (byte)((UInt32)value >> 24);
            }
        }

        internal static void Clear(byte[] Data, int StartPos)
        {
            Array.Clear(Data, StartPos, 64 + 2 //Clear name and name length.
                                        + 1  //StgType invalid
                                        + 1  //DeColor
                );

            //Data[StartPos+64+2]=0;

            byte[] Ones = BitConverter.GetBytes((UInt32)0xFFFFFFFF);
            Ones.CopyTo(Data, StartPos + 0x44);  //Left Sibling
            Ones.CopyTo(Data, StartPos + 0x48);  //Right Sibling
            Ones.CopyTo(Data, StartPos + 0x4C);  //Child Sibling
            Array.Clear(Data, StartPos + 0x50, TOle2Directory.DirectorySize - 0x50); //All else
        }

        internal static int GetLeftSid(byte[] Data, int StartPos)
        {
            return BitConverter.ToInt32(Data, 0x0044 + StartPos);
        }
        internal static void SetLeftSid(byte[] Data, int StartPos, int value)
        {
            BitConverter.GetBytes((Int32)value).CopyTo(Data, 0x0044 + StartPos);
        }

        internal static int GetRightSid(byte[] Data, int StartPos)
        {
            return BitConverter.ToInt32(Data, 0x0048 + StartPos);
        }
        internal static void SetRightSid(byte[] Data, int StartPos, int value)
        {
            BitConverter.GetBytes((Int32)value).CopyTo(Data, 0x0048 + StartPos);
        }

        internal static int GetChildSid(byte[] Data, int StartPos)
        {
            return BitConverter.ToInt32(Data, 0x004C + StartPos);
        }
        internal static void SetChildSid(byte[] Data, int StartPos, int value)
        {
            BitConverter.GetBytes((Int32)value).CopyTo(Data, 0x004C + StartPos);
        }

        internal static DECOLOR GetColor(byte[] Data, int StartPos)
        {
            return (DECOLOR)Data[0x0043 + StartPos];
        }

        internal static void SetColor(byte[] Data, int StartPos, DECOLOR value)
        {
            Data[0x0043 + StartPos] = (byte)value;
        }



        internal STGTY ObjType { get { return GetType(Data, 0); } set { Data[0x0042] = (byte)value; } }
        internal long SectStart { get { return GetSectStart(Data, 0); } set { SetSectStart(Data, 0, value); } }

        internal long xulSize { get { return GetSize(Data, 0); } set { SetSize(Data, 0, value); } }

        internal int LeftSid
        {
            get
            {
                return GetLeftSid(Data, 0);
            }
            set
            {
                SetLeftSid(Data, 0, value);
            }
        }

        internal int RightSid
        {
            get
            {
                return GetRightSid(Data, 0);
            }
            set
            {
                SetRightSid(Data, 0, value);
            }
        }

        internal int ChildSid
        {
            get
            {
                return GetChildSid(Data, 0);
            }
            set
            {
                SetChildSid(Data, 0, value);
            }
        }
    }

    /// <summary>
    /// Hold a sector in memory.
    /// </summary>
    internal class TSectorBuffer
    {
        private byte[] Data;
        private bool Changed = false;
        private long FSectorId = -1;
        private TOle2Header Header;
        private Stream DataStream;

        internal long SectorId { get { return FSectorId; } }
        internal TSectorBuffer(TOle2Header aHeader, Stream aStream)
        {
            Header = aHeader;
            DataStream = aStream;
            Data = new byte[Header.SectorSize];
            Changed = false;
            FSectorId = -1;
        }

        internal void Load(long SectNo)
        {
            if (Changed) Save();
            if (SectNo == FSectorId) return;
            DataStream.Seek(Header.SectToStPos(SectNo), SeekOrigin.Begin);
            FSectorId = -1; //It is invalid until we read the data.
            Sh.Read(DataStream, Data, 0, Data.Length, false);
            FSectorId = SectNo;
        }
        internal void Save()
        {
            if (Changed)
            {
                DataStream.Seek(Header.SectToStPos(FSectorId), SeekOrigin.Begin);
                Sh.Write(DataStream, Data, 0, Data.Length);
                Changed = false;
            }
        }

        internal void Read(byte[] aBuffer, long BufferPos, ref long nRead, long StartPos, long Count, long SectorSize)
        {
            if (Count > SectorSize - StartPos) nRead = SectorSize - StartPos; else nRead = Count;
            Array.Copy(Data, (int)StartPos, aBuffer, (int)BufferPos, (int)nRead);  //The (int) are to be compatible with CF
        }
    }

    internal class TDirEntryList : List<TOneDirEntry>
    {
    }

    internal class TOneDirEntry
    {
        internal string Name;
        internal int LeftSid;
        internal int RightSid;
        internal int ChildSid;

        internal bool Deleted;
        internal DECOLOR Color;
        internal STGTY DirType;

        internal TOle2Directory Ole2Dir;

        internal TOneDirEntry(string aName, int aLeftSid, int aRightSid, int aChildSid, DECOLOR aColor, STGTY aDirType, TOle2Directory aOle2Dir)
        {
            Name = aName;
            LeftSid = aLeftSid;
            RightSid = aRightSid;
            ChildSid = aChildSid;
            Deleted = false;
            Color = aColor;
            DirType = aDirType;
            Ole2Dir = aOle2Dir;
        }
    }

    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////	

    /// <summary>
    /// Class encapsulating an OLE2 file. FAT is kept in memory, data is read/written from/to disk.
    /// </summary>
    internal class TOle2File : IDisposable, IDataStream
    {
        internal bool NotXls97;

        private Stream FStream;
        private TOle2Header Header;
        private TOle2FAT FAT;
        private TOle2MiniFAT MiniFAT;
        private TSectorBuffer SectorBuffer;
        private TOle2Directory ROOT;
        //private TOle2DirList DirList;

        private TEncryptionData FEncryption;

        private string TOle2FileStr = "TOle2File";
        /// <summary>
        /// Opens an EXISTING OLE2 Stream. There is no support for creating a new one, you can only modify existing ones.
        /// </summary>
        /// <param name="aStream">The stream with the data</param>
        internal TOle2File(Stream aStream)
            : this(aStream, false)
        {
        }

        /// <summary>
        /// Opens an EXISTING OLE2 Stream, without throwing an exception if it is a wrong file. (On this case the error is logged into the Notxls97 variable)
        /// There is no support for creating a new one, you can only modify existing ones.
        /// </summary>
        /// <param name="aStream">The stream with the data</param>
        /// <param name="AvoidExceptions">If true, no Exception will be raised when the file is not OLE2.</param>
        internal TOle2File(Stream aStream, bool AvoidExceptions)
        {
            FStream = aStream;
            long StreamPosition = aStream.Position;

            Header = new TOle2Header(FStream, AvoidExceptions);
            if (Header.NotXls97)
            {
                NotXls97 = true;
                FStream.Position = StreamPosition;
                return;
            }
            FAT = TOle2FAT.Create(Header, FStream);
            MiniFAT = TOle2MiniFAT.Create(Header, FStream, FAT);
            ROOT = FindRoot();
            SectorBuffer = new TSectorBuffer(Header, FStream);
            FEncryption = new TEncryptionData(String.Empty, null, null);
        }

        #region IDisposable Members
        private bool disposed = false;
        internal void Close()
        {
            Dispose();
        }

        public void Dispose()
        {
            Dispose(true);
            // Take yourself off the Finalization queue 
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            lock (this)
            {

                // Check to see if Dispose has already been called.
                if (disposing && !this.disposed)
                {
                    FinishStream();
                    disposed = true;
                }
            }
        }

        #endregion


        public TEncryptionData Encryption
        {
            get { return FEncryption; }
        }

        internal TOle2Directory FindDir(string DirName)
        {
            if (disposed) throw new ObjectDisposedException(TOle2FileStr);

            byte[] Data = new byte[Header.SectorSize];
            long DirSect = Header.sectDirStart;
            while (DirSect != TOle2Header.ENDOFCHAIN)
            {
                FStream.Seek(Header.SectToStPos(DirSect), SeekOrigin.Begin);
                Sh.Read(FStream, Data, 0, Data.Length, false);
                for (int k = 0; k < Header.SectorSize; k += TOle2Directory.DirectorySize)
                {
                    if (String.Equals(TOle2Directory.GetName(Data, k), DirName, StringComparison.InvariantCultureIgnoreCase))
                    {
                        byte[] nd = new byte[TOle2Directory.DirectorySize];
                        Array.Copy(Data, k, nd, 0, nd.Length);
                        return new TOle2Directory(nd);
                    }
                }
                DirSect = FAT.GetNextSector(DirSect);

            }
            return null;
        }
        internal TOle2Directory FindDirWithPath(string DirName)
        {
            TDirEntryList Dirs;
            TOneDirEntry R = FindDirWithPath(DirName, out Dirs);
            if (R == null) return null;
            return R.Ole2Dir;
        }

        internal TOneDirEntry FindDirWithPath(string DirName, out TDirEntryList Dirs)
        {
            Dirs = null;
            if (disposed) throw new ObjectDisposedException(TOle2FileStr);
            if (DirName.Length < 1) return null;

            Dirs = ListDirs(true);
            string[] PathToFind = DirName.Substring(1).Split((char)0);

            return FindPath(Dirs, Dirs[0], PathToFind, 0);

        }

        private TOneDirEntry FindPath(TDirEntryList Dirs, TOneDirEntry ParentDir, string[] PathToFind, int PathPos)
        {
            if (PathPos >= PathToFind.Length) return null;

            if (String.Equals(ParentDir.Name, PathToFind[PathPos], StringComparison.InvariantCultureIgnoreCase))
            {
                if (PathPos == PathToFind.Length - 1) return ParentDir;
                if (ParentDir.ChildSid < 0) return null;
                return FindPath(Dirs, Dirs[ParentDir.ChildSid], PathToFind, PathPos + 1);
            }

            if (ParentDir.LeftSid >= 0)
            {
                TOneDirEntry Result = FindPath(Dirs, Dirs[ParentDir.LeftSid], PathToFind, PathPos);
                if (Result != null) return Result;
            }

            if (ParentDir.RightSid >= 0)
            {
                TOneDirEntry Result = FindPath(Dirs, Dirs[ParentDir.RightSid], PathToFind, PathPos);
                if (Result != null) return Result;
            }
            return null;
        }

        private void FillDir(TDirEntryList Dirs, List<String> ResultList, int Position, string Path)
        {
            if (Position >= Dirs.Count || Position < 0) return;

            TOneDirEntry dir = Dirs[Position];
            if (dir.DirType == STGTY.STREAM) ResultList.Add(Path + (char)0 + dir.Name);
            if (dir.LeftSid >= 0) FillDir(Dirs, ResultList, dir.LeftSid, Path);
            if (dir.RightSid >= 0) FillDir(Dirs, ResultList, dir.RightSid, Path);
            if (dir.ChildSid >= 0) FillDir(Dirs, ResultList, dir.ChildSid, Path + (char)0 + dir.Name);
            
        }

        internal String[] ListStreams()
        {
            List<string> ResultList = new List<string>();
            TDirEntryList Dirs = ListDirs(false);
            FillDir(Dirs, ResultList, 0, "");

            return ResultList.ToArray();
        }

        internal TOle2Directory FindRoot()
        {
            if (disposed) throw new ObjectDisposedException(TOle2FileStr);

            byte[] Data = new byte[TOle2Directory.DirectorySize];
            long DirSect = Header.sectDirStart;
            FStream.Seek(Header.SectToStPos(DirSect), SeekOrigin.Begin);
            Sh.Read(FStream, Data, 0, Data.Length, false);
            return new TOle2Directory(Data);
        }

        #region STREAM

        private TOle2Directory DIR = null;
        private long StreamPos;
        private bool PreparedForWrite = false;
        private long DIRStartPos = -1;

        /// <summary>
        /// StreamName might be a single name, in which case it will be the first stream with that name, or a path starting with (char)0.
        /// and where every folder is separated with (char)0
        /// </summary>
        /// <param name="StreamName"></param>
        internal void SelectStream(string StreamName)
        {
            SelectStream(StreamName, false);
        }

        internal bool SelectStream(string StreamName, bool AvoidExceptions)
        {
            if (disposed) throw new ObjectDisposedException(TOle2FileStr);

            if (PreparedForWrite) XlsMessages.ThrowException(XlsErr.ErrInvalidStream, StreamName);
            if (StreamName.StartsWith(((char)0).ToString())) DIR = FindDirWithPath(StreamName); else DIR = FindDir(StreamName);
            if (DIR == null)
            {
                if (AvoidExceptions) return false;
                XlsMessages.ThrowException(XlsErr.ErrFileIsNotXLS);
            }
            StreamPos = 0;

            return true;
        }

        internal long Length
        {
            get
            {
                if (DIR == null || DIR.ObjType != STGTY.STREAM) return 0; else return DIR.ulSize;
            }
        }

        public long Position
        {
            get
            {
                /*if (PreparedForWrite) return FStream.Position-  Header.SectToStPos(DIR.SectStart);
                    else return StreamPos;
				*/
                if (PreparedForWrite)
                    return DIR.ulSize;
                return StreamPos;
            }
            set
            {
                if (PreparedForWrite || StreamPos < 0 || StreamPos > Length) XlsMessages.ThrowException(XlsErr.ErrInvalidStream, String.Empty);
                StreamPos = value;
            }
        }

        internal bool Eof
        {
            get
            {
                if (disposed) throw new ObjectDisposedException(TOle2FileStr);
                if (PreparedForWrite) XlsMessages.ThrowException(XlsErr.ErrInvalidStream, String.Empty);
                return StreamPos >= Length;
            }
        }

        internal bool NextEof(int Count)
        {
            if (disposed) throw new ObjectDisposedException(TOle2FileStr);
            if (PreparedForWrite) XlsMessages.ThrowException(XlsErr.ErrInvalidStream, String.Empty);
            return StreamPos + Count >= Length;
        }

        public string FileName
        {
            get
            {
                if (disposed) throw new ObjectDisposedException(TOle2FileStr);
                if (FStream is FileStream) return ((FileStream)FStream).Name;
                else return String.Empty;
            }
        }

        internal void Read(byte[] aBuffer, int Count)
        {
            if (aBuffer.Length == 0) return;  // this is needed to avoid reading into a free record.

            if (DIR == null) return; //No stream selected
            if (PreparedForWrite) XlsMessages.ThrowException(XlsErr.ErrInvalidStream, String.Empty);
            long DIRulSize = DIR.ulSize;
            if (StreamPos + Count > DIRulSize) throw new IOException(XlsMessages.GetString(XlsErr.ErrEofReached, StreamPos + Count - DIRulSize)); //Reading past the end of the stream.
            if ((DIRulSize < Header.ulMiniSectorCutoff))
            {
                //Read from the MiniFat.
                long MiniSectsOn1Sect = 1 << (Header.uSectorShift - Header.uMiniSectorShift);


                //Find the minifat Sector number we have to read
                long MiniFatSectorOfs = StreamPos >> Header.uMiniSectorShift;
                long ActualMiniFatSector = MiniFAT.FindSector(DIR.SectStart, MiniFatSectorOfs);

                //Now, find this minifat sector into the MiniStream

                long SectorOfs = ActualMiniFatSector >> (Header.uSectorShift - Header.uMiniSectorShift);  // MiniFAT/8
                long MiniStreamSector = FAT.FindSector(ROOT.SectStart, SectorOfs);

                long nRead = 0;
                long TotalRead = 0;
                while (TotalRead < Count)
                {
                    SectorBuffer.Load(MiniStreamSector);
                    long MiniOffset = ((ActualMiniFatSector % MiniSectsOn1Sect) << Header.uMiniSectorShift);
                    long MiniStart = StreamPos % Header.MiniSectorSize + MiniOffset;
                    SectorBuffer.Read(aBuffer, TotalRead, ref nRead, MiniStart, Count - TotalRead, Header.MiniSectorSize + MiniOffset);
                    StreamPos += nRead;
                    TotalRead += nRead;
                    if (TotalRead < Count)
                    {
                        ActualMiniFatSector = MiniFAT.GetNextSector(ActualMiniFatSector);
                        SectorOfs = ActualMiniFatSector >> (Header.uSectorShift - Header.uMiniSectorShift);  // MiniFAT/8
                        MiniStreamSector = FAT.FindSector(ROOT.SectStart, SectorOfs);
                    }
                }


            }
            else
            {
                //Read from a normal sector
                long SectorOfs = (StreamPos >> Header.uSectorShift);
                long ActualSector = FAT.FindSector(DIR.SectStart, SectorOfs);

                long nRead = 0;
                long TotalRead = 0;
                while (TotalRead < Count)
                {
                    SectorBuffer.Load(ActualSector);
                    SectorBuffer.Read(aBuffer, TotalRead, ref nRead, StreamPos % Header.SectorSize, Count - TotalRead, Header.SectorSize);
                    StreamPos += nRead;
                    if (TotalRead < Count)
                    {
                        TotalRead += nRead;
                        ActualSector = FAT.GetNextSector(ActualSector);
                    }
                }
            }
        }

        /// <summary>
        /// Writes to the stream sequentially. No seek or read allowed while writing.
        /// </summary>
        /// <param name="Buffer">The data.</param>
        /// <param name="Count">number of bytes to write.</param>
        public void WriteRaw(byte[] Buffer, int Count)
        {
            if (disposed) throw new ObjectDisposedException(TOle2FileStr);

            if (DIR == null) return; //No stream selected
            if (!PreparedForWrite) XlsMessages.ThrowException(XlsErr.ErrInvalidStream, String.Empty);
            Sh.Write(FStream, Buffer, 0, Count);
            DIR.ulSize += Count;
        }

        public void Write(byte[] Buffer, int Count)
        {
            if (FEncryption != null && FEncryption.Engine != null)
                Buffer = FEncryption.Engine.Encode(Buffer, Position, 0, Buffer.Length, FEncryption.ActualRecordLen);
            WriteRaw(Buffer, Count);
        }

        public void WriteRaw(byte[] Buffer, int StartPos, int Count)
        {
            if (disposed) throw new ObjectDisposedException(TOle2FileStr);

            if (DIR == null) return; //No stream selected
            if (!PreparedForWrite) XlsMessages.ThrowException(XlsErr.ErrInvalidStream, String.Empty);
            Sh.Write(FStream, Buffer, StartPos, Count);
            DIR.ulSize += Count;
        }

        public void WriteHeader(UInt16 Id, UInt16 Len)
        {
            byte[] Header = { (byte)(Id & 0xFF), (byte)((Id >> 8) & 0xFF), (byte)(Len & 0xFF), (byte)((Len >> 8) & 0xFF) };
            WriteRaw(Header, Header.Length);
            FEncryption.ActualRecordLen = Len;

        }

        public void Write(byte[] Buffer, int StartPos, int Count)
        {
            if (disposed) throw new ObjectDisposedException(TOle2FileStr);

            if (DIR == null) return; //No stream selected
            if (!PreparedForWrite) XlsMessages.ThrowException(XlsErr.ErrInvalidStream, String.Empty);
            if (FEncryption != null && FEncryption.Engine != null)
                Buffer = FEncryption.Engine.Encode(Buffer, Position, StartPos, Count, FEncryption.ActualRecordLen);

            Sh.Write(FStream, Buffer, StartPos, Count);
            DIR.ulSize += Count;
        }

        public void Write16(UInt16 Buffer)
        {
            if (disposed) throw new ObjectDisposedException(TOle2FileStr);

            if (DIR == null) return; //No stream selected
            if (!PreparedForWrite) XlsMessages.ThrowException(XlsErr.ErrInvalidStream, String.Empty);

            if (FEncryption != null && FEncryption.Engine != null)
                Buffer = FEncryption.Engine.Encode(Buffer, Position, FEncryption.ActualRecordLen);

            //byte[] p = BitConverter.GetBytes(Buffer);
            //Sh.Write(FStream, p, 0, p.Length);
            //DIR.ulSize+=p.Length;
            unchecked
            {
                FStream.WriteByte((byte)Buffer);
                FStream.WriteByte((byte)(Buffer >> 8));
            }
            DIR.ulSize += 2;
        }

        public void Write32(UInt32 Buffer)
        {
            if (disposed) throw new ObjectDisposedException(TOle2FileStr);

            if (FEncryption != null && FEncryption.Engine != null)
                Buffer = FEncryption.Engine.Encode(Buffer, Position, FEncryption.ActualRecordLen);

            if (DIR == null) return; //No stream selected
            if (!PreparedForWrite) XlsMessages.ThrowException(XlsErr.ErrInvalidStream, String.Empty);
            //byte[] p = BitConverter.GetBytes(Buffer);
            //Sh.Write(FStream, p, 0, p.Length);
            //DIR.ulSize+=p.Length;
            unchecked
            {
                FStream.WriteByte((byte)Buffer);
                FStream.WriteByte((byte)((UInt32)Buffer >> 8));
                FStream.WriteByte((byte)((UInt32)Buffer >> 16));
                FStream.WriteByte((byte)((UInt32)Buffer >> 24));
            }
            DIR.ulSize += 4;

        }

        public void WriteRow(int Row)
        {
            Biff8Utils.CheckRow(Row);
            unchecked
            {
                Write16((UInt16)Row);
            }
        }

        public void WriteCol(int Col)
        {
            Biff8Utils.CheckCol(Col);
            unchecked
            {
                Write16((UInt16)Col);
            }
        }

        public void WriteColByte(int Col)
        {
            Biff8Utils.CheckCol(Col);
            unchecked
            {
                Write(new byte[] { (byte)Col }, 1);
            }
        }

        /// <summary>
        /// When using this one, remember to call FinishStreamWriting once you are over.
        /// </summary>
        /// <returns></returns>
        public Stream GetStreamForWriting()
        {
            return FStream;
        }

        public void FinishStreamWriting()
        {
            FStream.Position = FStream.Length;
            DIR.ulSize = FStream.Length -  Header.SectToStPos(DIR.SectStart);
        }

        internal static bool FindString(string s, string[] list)
        {
            for (int i = 0; i < list.Length; i++)
                if (s.Equals(list[i])) return true;
            return false;
        }

        /// <summary>
        /// Only seeks forward, no reverse. Not really optimized either, don't use in heavy places.
        /// </summary>
        /// <param name="Offset"></param>
        /// <returns></returns>
        internal void SeekForward(long Offset)
        {
            if (Position > Offset) XlsMessages.ThrowException(XlsErr.ErrInvalidPropertySector);

            if (Offset > Position)
            {
                byte[] Tmp = new byte[Offset - Position];
                Read(Tmp, Tmp.Length);
            }
        }

        private void MarkDeleted(int i, TDirEntryList Result, int Level)
        {
            if (Result[i].Deleted) return;
            Result[i].Deleted = true;
            if (Result[i].ChildSid >= 0) MarkDeleted(Result[i].ChildSid, Result, Level + 1);
            if ((Level > 0) && (Result[i].LeftSid >= 0)) MarkDeleted(Result[i].LeftSid, Result, Level + 1);
            if ((Level > 0) && (Result[i].RightSid >= 0)) MarkDeleted(Result[i].RightSid, Result, Level + 1);
        }

        private static void DeleteNode(TDirEntryList Result, ref int ParentLeaf)
        {
            if ((Result[ParentLeaf].LeftSid < 0) && (Result[ParentLeaf].RightSid < 0)) //It is a final node
            {
                ParentLeaf = -1;
                return;
            }
            if ((Result[ParentLeaf].LeftSid < 0)) //Only right branch.
            {
                ParentLeaf = Result[ParentLeaf].RightSid;
                return;
            }
            if ((Result[ParentLeaf].RightSid < 0)) //Only left branch.
            {
                ParentLeaf = Result[ParentLeaf].LeftSid;
                return;
            }

            //Leaf has both branchs.
            //Relabel the node as its successor and delete the successor
            ///////////////////////////////////////////////////////////////////
            //Example: Delete node 3 here
            //           10
            //        3       
            //     2      6     
            //          4    7
            //            5    
            // We need to relabel 4 as 3, and hang 5 from 6.
            ///////////////////////////////////////////////////////////////////

            //Find the next node. (once to the right and then always left)
            int NextNode = Result[ParentLeaf].RightSid;
            int PreviousNode = -1;
            while (Result[NextNode].LeftSid >= 0)
            {
                PreviousNode = NextNode;
                NextNode = Result[NextNode].LeftSid;
            }

            //Rename it.
            Result[NextNode].LeftSid = Result[ParentLeaf].LeftSid;  //LeftSid is always-1, we are at the left end.
            if (PreviousNode >= 0)  //If parentNode=-1, we are at the first node (6 on the example) and we don't have to fix the right part.
            {
                if (Result[NextNode].RightSid >= 0) Result[PreviousNode].LeftSid = Result[NextNode].RightSid; else Result[PreviousNode].LeftSid = -1;
                Result[NextNode].RightSid = Result[ParentLeaf].RightSid;
            }

            ParentLeaf = NextNode;
        }

        private void FixNode(TDirEntryList Result, ref int ParentNode)
        {
            while ((ParentNode > 0) && (Result[ParentNode].Deleted))
                DeleteNode(Result, ref ParentNode);
            if (ParentNode < 0) return;
            if (Result[ParentNode].LeftSid >= 0) FixNode(Result, ref Result[ParentNode].LeftSid);
            if (Result[ParentNode].RightSid >= 0) FixNode(Result, ref Result[ParentNode].RightSid);
            if (Result[ParentNode].ChildSid >= 0) FixNode(Result, ref Result[ParentNode].ChildSid);
        }

        private TDirEntryList ReadDirs(string[] DeletedStorages, ref bool PaintItBlack)
        {
            TDirEntryList Result = ListDirs(false);
            // Tag deleted storages and its children.
            for (int i = 1; i < Result.Count; i++)  //Skip 0, we can't delete root.
            {
                if (FindString(Result[i].Name, DeletedStorages))
                {
                    MarkDeleted(i, Result, 0);
                    if (Result[i].Color == DECOLOR.BLACK) PaintItBlack = true;
                }
            }

            //Now that we know the deletes, delete the nodes from the red/black tree.
            int FakeParent = 0;
            FixNode(Result, ref FakeParent);
            Debug.Assert(FakeParent == 0, "Can't delete root");

            return Result;
        }

        private TDirEntryList ListDirs(bool AddOle2Dir)
        {
            TDirEntryList Result = new TDirEntryList();
            long DirSect = Header.sectDirStart;
            byte[] DirSector = new byte[Header.SectorSize];
            while (DirSect != TOle2Header.ENDOFCHAIN)
            {
                FStream.Seek(Header.SectToStPos(DirSect), SeekOrigin.Begin);
                //Read the whole sector, tipically 4 DIR entries.
                Sh.Read(FStream, DirSector, 0, DirSector.Length, false);
                for (int i = 0; i < DirSector.Length; i += TOle2Directory.DirectorySize)
                {
                    TOle2Directory Ole2Dir = null;
                    if (AddOle2Dir)
                    {
                        byte[] nd = new byte[TOle2Directory.DirectorySize];
                        Array.Copy(DirSector, i, nd, 0, nd.Length);
                        Ole2Dir = new TOle2Directory(nd);
                    }

                    Result.Add(new TOneDirEntry(TOle2Directory.GetName(DirSector, i),
                        TOle2Directory.GetLeftSid(DirSector, i),
                        TOle2Directory.GetRightSid(DirSector, i),
                        TOle2Directory.GetChildSid(DirSector, i),
                        TOle2Directory.GetColor(DirSector, i),
                        TOle2Directory.GetType(DirSector, i),
                        Ole2Dir));
                }

                DirSect = FAT.GetNextSector(DirSect);
            }
            return Result;
        }


        /// <summary>
        /// This method copies the contents of the ole stream to a new one, and clears the OStreamName 
        /// stream, leaving it ready to be written.
        /// </summary>
        /// <param name="OutStream">The new Stream where we will write the data</param>
        /// <param name="OStreamName">Ole Stream Name (tipically "Workbook") that we will clear to write the new data.</param>
        /// <param name="DeleteStorages">Storages we are not going to copy to the new one. Used to remove macros.</param>
        internal void PrepareForWrite(Stream OutStream, string OStreamName, string[] DeleteStorages)
        {
            if (disposed) throw new ObjectDisposedException(TOle2FileStr);

            if (PreparedForWrite) XlsMessages.ThrowException(XlsErr.ErrInvalidStream, String.Empty);
            DIRStartPos = -1;
            int DIRStartOfs = -1;
            TOle2FAT NewFat = TOle2FAT.Create(Header, null);
            int LastDirPos = -1;
            long IniPos = OutStream.Position;
            byte[] DirSector = new byte[Header.SectorSize];
            byte[] DataSector = new byte[Header.SectorSize];
            DIR = null;
            bool PaintItBlack = false; //We are not going to mess with red/black here. If a recolor on the tree is needed, we will paint it all black.

            TDirEntryList DirEntries = null;
            if (DeleteStorages.Length > 0) //Find all the storages and streams to delete, and patch the others not to point to them.
                DirEntries = ReadDirs(DeleteStorages, ref PaintItBlack);

            long DirSect = Header.sectDirStart;
            OutStream.Seek(Header.SectToStPos(0, IniPos), SeekOrigin.Begin); //Advance to the first sector.
            int CurrentDirPos = 0;
            //Copy the Dir tree and their asociated streams. If stream is OStreamName, set its size to 0.
            while (DirSect != TOle2Header.ENDOFCHAIN)
            {
                FStream.Seek(Header.SectToStPos(DirSect), SeekOrigin.Begin);
                //Read the whole sector, tipically 4 DIR entries.
                Sh.Read(FStream, DirSector, 0, DirSector.Length, false);
                for (int i = 0; i < DirSector.Length; i += TOle2Directory.DirectorySize)
                {
                    STGTY SType = TOle2Directory.GetType(DirSector, i);

                    if (PaintItBlack)
                        TOle2Directory.SetColor(DirSector, i, DECOLOR.BLACK);
                    if ((DirEntries != null) && (!DirEntries[CurrentDirPos].Deleted))
                    {
                        //Fix the tree.
                        TOle2Directory.SetLeftSid(DirSector, i, DirEntries[CurrentDirPos].LeftSid);
                        TOle2Directory.SetRightSid(DirSector, i, DirEntries[CurrentDirPos].RightSid);
                        TOle2Directory.SetChildSid(DirSector, i, DirEntries[CurrentDirPos].ChildSid);
                    }


                    if ((DirEntries != null) && (DirEntries[CurrentDirPos].Deleted))
                    {
                        TOle2Directory.Clear(DirSector, i);
                    }
                    else
                        if (((SType == STGTY.STREAM) &&
                            ((TOle2Directory.GetSize(DirSector, i) >= Header.ulMiniSectorCutoff) || (String.Equals(TOle2Directory.GetName(DirSector, i), OStreamName, StringComparison.InvariantCultureIgnoreCase)))) ||
                            (SType == STGTY.ROOT))  //When ROOT, the stream is the MiniStream.  When Sectors reference the ministream, the data is not copied, as the whole ministream was copied with root.
                            if (String.Compare(TOle2Directory.GetName(DirSector, i), OStreamName, true, CultureInfo.InvariantCulture) != 0)
                            {
                                //Arrange FAT
                                long StreamSect = TOle2Directory.GetSectStart(DirSector, i);
                                long StreamSize = TOle2Directory.GetSize(DirSector, i);
                                long bRead = 0;
                                if (StreamSect != TOle2Header.ENDOFCHAIN) TOle2Directory.SetSectStart(DirSector, i, NewFat.Count);

                                while (StreamSect != TOle2Header.ENDOFCHAIN && bRead < StreamSize)
                                {
                                    //Copy old Sector to New sector
                                    FStream.Seek(Header.SectToStPos(StreamSect), SeekOrigin.Begin);
                                    Sh.Read(FStream, DataSector, 0, DataSector.Length, false);
                                    Sh.Write(OutStream, DataSector, 0, DataSector.Length);

                                    Debug.Assert(OutStream.Position - IniPos == Header.SectToStPos(NewFat.Count + 1) - Header.StartOfs, "New Stream IO Error");
                                    //Update The Fat
                                    StreamSect = FAT.GetNextSector(StreamSect);
                                    bRead += Header.SectorSize;
                                    if (StreamSect != TOle2Header.ENDOFCHAIN && bRead < StreamSize) NewFat.Add(((UInt32)NewFat.Count + 1)); else NewFat.Add(TOle2Header.ENDOFCHAIN);
                                }
                            }
                            else
                            {
                                TOle2Directory.SetSectStart(DirSector, i, TOle2Header.ENDOFCHAIN);
                                TOle2Directory.SetSize(DirSector, i, 0);
                                byte[] nd = new byte[TOle2Directory.DirectorySize];
                                Array.Copy(DirSector, i, nd, 0, nd.Length);
                                DIR = new TOle2Directory(nd);
                                DIRStartOfs = i;
                            }

                    CurrentDirPos++;
                }
                //Save the DIR Sector
                if (DIRStartOfs >= 0)
                {
                    DIRStartPos = OutStream.Position + DIRStartOfs;  //We must save the position here, just before writing the sector.
                    DIRStartOfs = -1;
                }
                Sh.Write(OutStream, DirSector, 0, DirSector.Length);
                //Add a new entry on the FAT for the new DIR sector.
                NewFat.Add(TOle2Header.ENDOFCHAIN);
                if (LastDirPos > 0) NewFat[LastDirPos] = ((UInt32)NewFat.Count - 1);  //Chain the last FAT DIR point to this.
                else Header.sectDirStart = ((UInt32)NewFat.Count - 1);
                DirSect = FAT.GetNextSector(DirSect);

                LastDirPos = NewFat.Count - 1;

            }
            if (DIR == null) XlsMessages.ThrowException(XlsErr.ErrInvalidStream, String.Empty);

            //Copy the MiniFat
            int LastMiniFatPos = -1;
            long MiniFatSect = Header.sectMiniFatStart;
            while (MiniFatSect != TOle2Header.ENDOFCHAIN)
            {
                FStream.Seek(Header.SectToStPos(MiniFatSect), SeekOrigin.Begin);
                //Read the whole sector, tipically 128 MiniFat entries.
                Sh.Read(FStream, DataSector, 0, DataSector.Length, false);
                Sh.Write(OutStream, DataSector, 0, DataSector.Length);

                NewFat.Add(TOle2Header.ENDOFCHAIN);
                if (LastMiniFatPos > 0) NewFat[LastMiniFatPos] = ((UInt32)NewFat.Count - 1);  //Chain the last FAT MiniFat point to this.
                else Header.sectMiniFatStart = ((UInt32)NewFat.Count - 1);
                MiniFatSect = FAT.GetNextSector(MiniFatSect);
                LastMiniFatPos = NewFat.Count - 1;

            }

            //Switch to the new Stream.
            FStream = OutStream;
            Header.StartOfs = IniPos;
            FAT = NewFat;
            //MiniFat stays the same.
            SectorBuffer = new TSectorBuffer(Header, FStream);
            ROOT = null;  //No need for it when writing.

            PreparedForWrite = true;
            DIR.SectStart = NewFat.Count;
            DIR.ulSize = 0;
        }

        private void FinishStream()
        {
            if (disposed) throw new ObjectDisposedException(TOle2FileStr);
            if (!PreparedForWrite) return;

            //Ensure Workbook has at least 4096 bytes, so it doesn't go to the MiniStream. 
            if (DIR.ulSize < Header.ulMiniSectorCutoff)
            {
                byte[] Data = new byte[Header.ulMiniSectorCutoff - DIR.ulSize];  //Filled with 0.
                WriteRaw(Data, Data.Length);
            }

            //Fill the rest of the sector with 0s. 
            if ((DIR.ulSize % Header.SectorSize) > 0)
            {
                byte[] Data = new byte[Header.SectorSize - (DIR.ulSize % Header.SectorSize)];  //Filled with 0.
                WriteRaw(Data, Data.Length);
                DIR.ulSize -= Data.Length; //This does not count to the final stream size
            }

            long DifSectorCount;
            long FATSectorCount;
            long StreamFatCount = (((DIR.ulSize - 1) >> Header.uSectorShift) + 1);
            FixHeaderSector(FAT, Header, StreamFatCount, out DifSectorCount, out FATSectorCount);

            //Save DIR
            FStream.Seek(DIRStartPos, SeekOrigin.Begin);
            DIR.Save(FStream);

            SaveHeader(FStream, Header);
            long StartFat = SaveDif(FStream, Header, FAT, StreamFatCount, DifSectorCount, FATSectorCount);
            SaveFAT(FStream, Header, FAT, StreamFatCount, DifSectorCount, FATSectorCount, StartFat);
        }

        private static long SaveDif(Stream FStream, TOle2Header Header, TOle2FAT FAT, long StreamFatCount, long DifSectorCount, long FATSectorCount)
        {
            //Save DIF in header
            byte[] DifInHeader = new byte[TOle2Header.DifEntries];
            long StartDif = FAT.Count + 1 + StreamFatCount;
            long StartFat = StartDif + DifSectorCount;
            int f = 0;
            if (FATSectorCount <= TOle2Header.DifsInHeader) f = (int)FATSectorCount; else f = TOle2Header.DifsInHeader;
            for (int i = 0; i < f; i++)
            {
                BitConverter.GetBytes((UInt32)(StartFat + i)).CopyTo(DifInHeader, i << 2);
            }

            //Docs say there should be an endofchain at the last slot of the last dif sector, but there is none. Only unused sectors (FFFF)
            for (int i = f << 2; i < DifInHeader.Length; i++) DifInHeader[i] = 0xFF;
            Sh.Write(FStream, DifInHeader, 0, DifInHeader.Length);

            byte[] DifSectorData = new byte[Header.SectorSize];
            //Save DIF Sectors.
            FStream.Seek(Header.SectToStPos(StartDif), SeekOrigin.Begin);
            for (int k = 0; k < DifSectorCount; k++)
            {
                int SectEnd = (int)Header.SectorSize - 4;
                if (k == DifSectorCount - 1) SectEnd = (int)((((FATSectorCount - TOle2Header.DifsInHeader - 1) % (Header.SectorSize / 4 - 1)) + 1) << 2);
                for (int i = 0; i < SectEnd; i += 4)
                {
                    BitConverter.GetBytes((UInt32)((StartFat + (i >> 2) + TOle2Header.DifsInHeader + k * (Header.SectorSize / 4 - 1)))).CopyTo(DifSectorData, i);
                }

                for (int i = SectEnd; i < Header.SectorSize - 4; i += 4)
                {
                    BitConverter.GetBytes((UInt32)TOle2Header.FREESECT).CopyTo(DifSectorData, i);
                }


                if (k == DifSectorCount - 1)
                    BitConverter.GetBytes((UInt32)TOle2Header.FREESECT).CopyTo(DifSectorData, (int)Header.SectorSize - 4);//Cast to int to be compatible with CF
                else
                    BitConverter.GetBytes((UInt32)(StartDif + k + 1)).CopyTo(DifSectorData, (int)Header.SectorSize - 4);//Cast to int to be compatible with CF
                Sh.Write(FStream, DifSectorData, 0, DifSectorData.Length);
            }
            return StartFat;
        }

        private static void SaveHeader(Stream FStream, TOle2Header Header)
        {
            //Save Header.
            FStream.Seek(Header.StartOfs, SeekOrigin.Begin);
            Header.Save(FStream);
        }

        private static void SaveFAT(Stream FStream, TOle2Header Header, TOle2FAT FAT, long StreamFatCount, long DifSectorCount, long FATSectorCount, long StartFat)
        {
            //Write FAT for unmodified storages/streams
            for (int k = 0; k < FAT.Count; k++)
                Sh.Write(FStream, BitConverter.GetBytes((UInt32)FAT[k]), 0, 4);

            //Write Stream FAT
            for (int k = 0; k < StreamFatCount; k++)
                Sh.Write(FStream, BitConverter.GetBytes((UInt32)(FAT.Count + k + 1)), 0, 4);
            Sh.Write(FStream, BitConverter.GetBytes((UInt32)TOle2Header.ENDOFCHAIN), 0, 4);

            //Write DIF FAT
            for (int k = 0; k < DifSectorCount; k++)
                Sh.Write(FStream, BitConverter.GetBytes((UInt32)TOle2Header.DIFSECT), 0, 4);

            //Write FAT FAT
            for (int k = 0; k < FATSectorCount; k++)
                Sh.Write(FStream, BitConverter.GetBytes((UInt32)TOle2Header.FATSECT), 0, 4);

            //Fill FAT sector with FF
            byte[] One = { 0xFF };
            for (int k = (int)(((StartFat + FATSectorCount) << 2) % Header.SectorSize); k < Header.SectorSize; k++)
                Sh.Write(FStream, One, 0, One.Length);
        }

        private static void FixHeaderSector(TOle2FAT FAT, TOle2Header Header, long ExtraFATSectors, out long DifSectorCount, out long FATSectorCount)
        {
            //Fix Header. Fat count &sect, dif count & sect.  //Minifat and Dir are already fixed.
            long OldDifSectorCount = 0;
            DifSectorCount = 0;
            FATSectorCount = 0;

            //Iterate to get the real dif/fat count. Adding a dif sector might add a fat sector, so it might add another dif... luckily this will converge really fast.
            //Also, the fat sectors should be included on the fat count. If it wasn't discrete, it would be a nice 3 x equation.
            do
            {
                long FATEntryCount0 = FAT.Count + 1 + ExtraFATSectors + DifSectorCount;
                long FatEntryDelta = ((FATEntryCount0 - 1) >> FAT.uFATEntryShift()) + 1;  //first guess
                long OldFatEntryDelta = 0;
                do
                {
                    OldFatEntryDelta = FatEntryDelta;
                    FatEntryDelta = ((FATEntryCount0 + FatEntryDelta - 1) >> FAT.uFATEntryShift()) + 1;
                    /* This converges, because FatEntryDelta>=OldFatEntryDelta
                     * To prove Fed[n+1]>=Fed[n], lets begin... (n=0): Fed[0]=0 <= Fed[1]=(0+FEC0)/128.
                     * (n=k):  if Fed[k]>=Fed[k-1] -> (n=k+1):  Fed[k+1]=(FEC0+Fed[n])/128  
                     * As Fed[n]>=Fed[n-1], FEC0>0 ->  (FEC0+Fed[n])/128)>=(FEC0+Fed[n-1])/128  ->
                     * 
                     * Fed[n+1]>=(FEC0+Fed[n-1])/128=Fed[n]  ;-)
                    */
                } while (FatEntryDelta != OldFatEntryDelta);

                FATSectorCount = ((FATEntryCount0 + FatEntryDelta - 1) >> FAT.uFATEntryShift()) + 1;

                OldDifSectorCount = DifSectorCount;
                if (FATSectorCount > TOle2Header.DifsInHeader)
                    DifSectorCount = (FATSectorCount - TOle2Header.DifsInHeader - 1) / (Header.SectorSize / 4 - 1) + 1;   //The last diff entry is a pointer to the new diff sector, so we have 127 slots, not 128.
            } while (OldDifSectorCount != DifSectorCount);

            Header.csectFat = (UInt32)FATSectorCount;
            Header.csectDif = (UInt32)DifSectorCount;
            if (DifSectorCount > 0) Header.sectDifStart = (UInt32)(FAT.Count + 1 + ExtraFATSectors); else Header.sectDifStart = TOle2Header.ENDOFCHAIN;
        }

        /// <summary>
        /// This method is destructive and will invalidate what is in the object. It will write a subtree of the folder tree
        /// converting the first storage (ParentStorage) to Root.
        /// </summary>
        /// <param name="ParentStorage"></param>
        /// <param name="OutStream"></param>
        internal void GetStorages(string ParentStorage, Stream OutStream)
        {
            if (disposed) throw new ObjectDisposedException(TOle2FileStr);

            TOle2FAT NewFat = TOle2FAT.Create(Header, null);
            TOle2MiniFAT NewMiniFat = TOle2MiniFAT.Create(Header, null, null);
            long IniPos = OutStream.Position;
            byte[] DataSector = new byte[Header.SectorSize];
            DIR = null;

            TDirEntryList DirEntries;

            TOneDirEntry CurrentDir = FindDirWithPath(ParentStorage, out DirEntries);
            if (CurrentDir == null) return;
            Debug.Assert(CurrentDir.DirType == STGTY.STORAGE, "This method must be called only for a directory");

            OutStream.Seek(Header.SectToStPos(0, IniPos), SeekOrigin.Begin); //Advance to the first sector, that is, sector 0.

            TOle2Directory NewRoot = new TOle2Directory(ROOT.Data);
            //Copy the Dir tree and their asociated streams.
            using (MemoryStream MiniStream = new MemoryStream())
            {
                CopyAllSectors(CurrentDir, DirEntries, OutStream, NewFat, DataSector, MiniStream, NewMiniFat, true);
                SaveMiniSectors(OutStream, MiniStream, NewFat, NewMiniFat, NewRoot);
            }
            
            long DifSectorCount;
            long FATSectorCount;
            long StreamFatCount = 0;
            FixHeaderSector(NewFat, Header, StreamFatCount, out DifSectorCount, out FATSectorCount);
            
            SaveMiniFat(OutStream, Header, NewFat, NewMiniFat);
            SaveDirs(OutStream, CurrentDir, DirEntries, Header, NewFat, NewRoot);

            SaveHeader(OutStream, Header);
            long StartFat = SaveDif(OutStream, Header, NewFat, StreamFatCount, DifSectorCount, FATSectorCount);
            SaveFAT(OutStream, Header, NewFat, StreamFatCount, DifSectorCount, FATSectorCount, StartFat);

        }

        private void SaveDirs(Stream OutStream, TOneDirEntry CurrentDir, TDirEntryList DirList, TOle2Header Header, TOle2FAT NewFat, TOle2Directory NewRoot)
        {
            int ChildSid = CurrentDir.Ole2Dir.ChildSid;
            CurrentDir.Ole2Dir.ChildSid = -1; //We will write this value later. header has only 1 child, no left/right siblings.
            long RootPos = OutStream.Position;
            NewRoot.Save(OutStream);
            int SavedDirs = 1;
            Header.sectDirStart = (UInt32)NewFat.Count;

            SaveChildren(OutStream, ChildSid, DirList, ref SavedDirs);

            long DirsPerSector = Header.SectorSize / TOle2Directory.DirectorySize;
            long DirSectors = 1 + (SavedDirs - 1) / DirsPerSector;
            
            //complete the sector
            long Remaining = SavedDirs % DirsPerSector;
            if (Remaining > 0)
            {
                long RemainingSectors = DirsPerSector - Remaining;
                byte[] EmptySec = new byte[TOle2Directory.DirectorySize];
                for (int i = 0; i < RemainingSectors; i++)
                {
                    OutStream.Write(EmptySec, 0, EmptySec.Length);
                }
            }

            AddFatSectors(NewFat, DirSectors);

            OutStream.Position = RootPos;
            NewRoot.ChildSid = SavedDirs - 1;
            NewRoot.Save(OutStream);
            OutStream.Seek(0, SeekOrigin.End);

            Header.csectDir = 0; //this is not supported in v3 ole files.
        }

        private void SaveChildren(Stream OutStream, int Sid, TDirEntryList DirList, ref int SavedDirs)
        {
            TOneDirEntry CurrentDir = DirList[Sid];
            int ChildSid = CurrentDir.ChildSid;
            if (ChildSid >= 0)
            {
                SaveChildren(OutStream, ChildSid, DirList, ref SavedDirs);
                CurrentDir.Ole2Dir.ChildSid = SavedDirs - 1;
            }
            int LeftSid = CurrentDir.LeftSid;
            if (LeftSid >= 0)
            {
                SaveChildren(OutStream, LeftSid, DirList, ref SavedDirs);
                CurrentDir.Ole2Dir.LeftSid = SavedDirs - 1;
            }
            int RightSid = CurrentDir.RightSid;
            if (RightSid >= 0)
            {
                SaveChildren(OutStream, RightSid, DirList, ref SavedDirs);
                CurrentDir.Ole2Dir.RightSid = SavedDirs - 1;
            }

            CurrentDir.Ole2Dir.Save(OutStream);
            SavedDirs++;
        }


        private void SaveMiniFat(Stream OutStream, TOle2Header Header, TOle2FAT NewFat, TOle2MiniFAT NewMiniFat)
        {
            Header.sectMiniFatStart = (UInt32)NewFat.Count;

            int MiniFatSize = NewMiniFat.Count * 4;
            for (int k = 0; k < NewMiniFat.Count; k++)
                Sh.Write(OutStream, BitConverter.GetBytes((UInt32)NewMiniFat[k]), 0, 4);

            //Fill MiniFAT sector with FF
            byte[] One = { 0xFF };
            long RemainingCount = MiniFatSize % Header.SectorSize;
            if (RemainingCount > 0)
            {
                long RemainingSize = (Header.SectorSize - RemainingCount);
                for (int k = 0; k < RemainingSize; k++)
                {
                    Sh.Write(OutStream, One, 0, One.Length);
                }
            }

            Header.csectMiniFat = 1 + ((UInt32)MiniFatSize - 1) / Header.SectorSize;


            //add fat sectors
            long RealSectorsUsed = 1 + (MiniFatSize - 1) / Header.SectorSize;
            AddFatSectors(NewFat, RealSectorsUsed);
        }

        private void SaveMiniSectors(Stream OutStream, MemoryStream MiniStream, TOle2FAT NewFat, TOle2MiniFAT NewMiniFat, TOle2Directory NewROOT)
        {
            if (MiniStream.Length == 0)
            {
                NewROOT.SectStart = TOle2Header.ENDOFCHAIN;
                NewROOT.xulSize = 0;
                return;
            }
            NewROOT.SectStart = NewFat.Count;
            NewROOT.xulSize = MiniStream.Length;
            Sh.Write(OutStream, MiniStream.ToArray(), 0, (Int32)MiniStream.Length);

            //complete ministream with 0.
            long MiniSectsOn1Sect = 1 << (Header.uSectorShift - Header.uMiniSectorShift);

            long RemainingCount = NewMiniFat.Count % MiniSectsOn1Sect;
            if (RemainingCount > 0)
            {
                long RemainingSize = (MiniSectsOn1Sect - RemainingCount) << Header.uMiniSectorShift;
                byte[] Zeros = new byte[RemainingSize];
                Sh.Write(OutStream, Zeros, 0, Zeros.Length);
            }

            //add fat sectors
            long RealSectorsUsed = 1 + (NewMiniFat.Count - 1) / MiniSectsOn1Sect;
            AddFatSectors(NewFat, RealSectorsUsed);
           

        }

        private static void AddFatSectors(TOle2FAT NewFat, long RealSectorsUsed)
        {
            for (long i = 0; i < RealSectorsUsed - 1; i++)
            {
                NewFat.Add((UInt32)NewFat.Count + 1);
            }

            NewFat.Add(TOle2Header.ENDOFCHAIN);
        }

        private void CopyAllSectors(TOneDirEntry CurrentDir, TDirEntryList DirEntries, Stream OutStream, TOle2FAT NewFat, 
            byte[] DataSector, MemoryStream MiniStream, TOle2MiniFAT NewMiniFAT, bool First)
        {
            if (CurrentDir == null) return;
            TOle2Directory.SetColor(CurrentDir.Ole2Dir.Data, 0, DECOLOR.BLACK);

            if (CurrentDir.DirType == STGTY.STREAM)
            {
                CopySectors(CurrentDir.Ole2Dir, OutStream, NewFat, DataSector, MiniStream, NewMiniFAT);
            }

            if (!First)
            {
                int LeftSid = CurrentDir.LeftSid;
                if (LeftSid >= 0)
                {
                    CopyAllSectors(DirEntries[LeftSid], DirEntries, OutStream, NewFat, DataSector, MiniStream, NewMiniFAT,false);
                }

                int RightSid = CurrentDir.RightSid;
                if (RightSid >= 0)
                {
                    CopyAllSectors(DirEntries[RightSid], DirEntries, OutStream, NewFat, DataSector, MiniStream, NewMiniFAT, false);
                }

            }

            int ChildSid = CurrentDir.ChildSid;
            if (ChildSid >= 0)
            {
                CopyAllSectors(DirEntries[ChildSid], DirEntries, OutStream, NewFat, DataSector, MiniStream, NewMiniFAT, false);
            }

        }

        private void CopySectors(TOle2Directory CurrentOleDir, Stream OutStream, TOle2FAT NewFat, byte[] DataSector, 
            MemoryStream MiniStream, TOle2MiniFAT NewMiniFAT)
        {
            if ((CurrentOleDir.xulSize < Header.ulMiniSectorCutoff))
            {
                CopyMiniSectors(CurrentOleDir, MiniStream, NewMiniFAT);
                return;
            }

            long StreamSect = CurrentOleDir.SectStart;
            long StreamSize = CurrentOleDir.xulSize;
            long bRead = 0;
            
            if (StreamSect != TOle2Header.ENDOFCHAIN) CurrentOleDir.SectStart = NewFat.Count;

            while (StreamSect != TOle2Header.ENDOFCHAIN && bRead < StreamSize)
            {
                //Copy old Sector to New sector
                FStream.Seek(Header.SectToStPos(StreamSect), SeekOrigin.Begin);
                Sh.Read(FStream, DataSector, 0, DataSector.Length, false);
                Sh.Write(OutStream, DataSector, 0, DataSector.Length);

                //Update The Fat
                StreamSect = FAT.GetNextSector(StreamSect);
                bRead += Header.SectorSize;
                if (StreamSect != TOle2Header.ENDOFCHAIN && bRead < StreamSize) NewFat.Add(((UInt32)NewFat.Count + 1)); else NewFat.Add(TOle2Header.ENDOFCHAIN);
            }
        }

        private void CopyMiniSectors(TOle2Directory CurrentOleDir, MemoryStream MiniStream, TOle2MiniFAT NewMiniFAT)
        {
            byte[] MiniSector = new byte[Header.MiniSectorSize];
            long MiniSectsOn1Sect = 1 << (Header.uSectorShift - Header.uMiniSectorShift);
            
            long StreamSect = CurrentOleDir.SectStart;
            if (StreamSect != TOle2Header.ENDOFCHAIN) CurrentOleDir.SectStart = NewMiniFAT.Count;

            long ActualMiniFatSector = StreamSect;

            while (ActualMiniFatSector != TOle2Header.ENDOFCHAIN)
            {
                long SectorOfs = ActualMiniFatSector >> (Header.uSectorShift - Header.uMiniSectorShift);  // MiniFAT/8
                long MiniStreamSector = FAT.FindSector(ROOT.SectStart, SectorOfs);
  
                long MiniOffset = ((ActualMiniFatSector % MiniSectsOn1Sect) << Header.uMiniSectorShift);

                FStream.Seek(Header.SectToStPos(MiniStreamSector) + MiniOffset, SeekOrigin.Begin);
                Sh.Read(FStream, MiniSector, 0, MiniSector.Length, true);
                Sh.Write(MiniStream, MiniSector, 0, MiniSector.Length);
                ActualMiniFatSector = MiniFAT.GetNextSector(ActualMiniFatSector);

                if (ActualMiniFatSector != TOle2Header.ENDOFCHAIN) NewMiniFAT.Add((UInt32)NewMiniFAT.Count + 1); else NewMiniFAT.Add(TOle2Header.ENDOFCHAIN);
            }
        }
    


        #endregion

    }
    #endregion

    #region IDataStream
    internal interface IDataStream
    {
        void WriteHeader(UInt16 Id, UInt16 Len);
        void WriteRow(int Row);
        void WriteCol(int Col);
        void WriteColByte(int Col);
        void Write16(UInt16 Buffer);
        void Write32(UInt32 Buffer);
        void Write(byte[] Buffer, int Count);
        void Write(byte[] Buffer, int StartPos, int Count);
        void WriteRaw(byte[] Buffer, int Count);
        void WriteRaw(byte[] Buffer, int StartPos, int Count);

        long Position { get; }
        string FileName { get; }
        TEncryptionData Encryption { get; }
    }
    #endregion

    #region MemOle2
    internal class MemOle2 : IDataStream, IDisposable
    {
        private Stream Ms;
        private TEncryptionData FEncryption;
        bool OwnsStream;

        internal MemOle2(): this(new MemoryStream())
        {
            OwnsStream = true;
        }

        internal MemOle2(Stream aStream)
        {
            FEncryption = new TEncryptionData(String.Empty, null, null);
            Ms = aStream;
        }

        internal byte[] GetBytes()
        {
            if (!(Ms is MemoryStream)) FlxMessages.ThrowException(FlxErr.ErrInternal);
            return ((MemoryStream)Ms).ToArray();
        }

        #region IDataStream Members

        public void WriteHeader(UInt16 Id, UInt16 Len)
        {
            byte[] Header = { (byte)(Id & 0xFF), (byte)((Id >> 8) & 0xFF), (byte)(Len & 0xFF), (byte)((Len >> 8) & 0xFF) };
            WriteRaw(Header, Header.Length);
        }

        public void WriteRow(int Row)
        {
            Biff8Utils.CheckRow(Row);
            unchecked
            {
                Write16((UInt16)Row);
            }
            
        }

        public void WriteCol(int Col)
        {
            Biff8Utils.CheckCol(Col);
            unchecked
            {
                Write16((UInt16)Col);
            }
        }

        public void WriteColByte(int Col)
        {
            Biff8Utils.CheckCol(Col);
            unchecked
            {
                Write(new byte[] { (byte)Col }, 1);
            }
        }

        public void Write16(ushort Buffer)
        {
            unchecked
            {
                Ms.WriteByte((byte)Buffer);
                Ms.WriteByte((byte)(Buffer >> 8));
            }
        }

        public void Write32(uint Buffer)
        {
            unchecked
            {
                Ms.WriteByte((byte)Buffer);
                Ms.WriteByte((byte)(Buffer >> 8));
                Ms.WriteByte((byte)(Buffer >> 16));
                Ms.WriteByte((byte)(Buffer >> 24));
            }
        }

        public void Write(byte[] Buffer, int Count)
        {
            Ms.Write(Buffer, 0, Count);
        }

        public void Write(byte[] Buffer, int StartPos, int Count)
        {
            Ms.Write(Buffer, StartPos, Count);
        }

        public void WriteRaw(byte[] Buffer, int Count)
        {
            Ms.Write(Buffer, 0, Count);
        }

        public void WriteRaw(byte[] Buffer, int StartPos, int Count)
        {
            Ms.Write(Buffer, StartPos, Count);
        }


        public long Position
        {
            get
            {
                return Ms.Position;
            }
        }

        public string FileName
        {
            get
            {
                if (Ms is FileStream) return ((FileStream)Ms).Name;
                else return String.Empty;
            }
        }

        public TEncryptionData Encryption
        {
            get { return FEncryption; }
        }

        #endregion

        #region IDisposable Members

        public void Dispose()
        {
            if (OwnsStream)
            {
                Ms.Close();
            }
            GC.SuppressFinalize(this);
        }

        #endregion
    }
    #endregion

}

