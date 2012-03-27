using System;
using FlexCel.Core;

namespace FlexCel.XlsAdapter
{
    /// <summary>
    /// Holds an encryption engine and a password. Engine has to be created on demand (and is polymorphical) so we need to store password in another place.
    /// </summary>
    internal class TEncryptionData
    {
        internal ExcelFile Xls;
        internal string ReadPassword;
        internal OnPasswordEventHandler OnPassword;
        internal TEncryptionEngine Engine;
        internal int ActualRecordLen;

        internal TEncryptionData(string aReadPassword, OnPasswordEventHandler aOnPassword, ExcelFile aXls)
        {
            ReadPassword= aReadPassword;
            OnPassword=aOnPassword;
            Xls=aXls;
        }

        internal int TotalSize()
        {
            if (Engine==null) return 0;
            return Engine.GetFilePassRecordLen();
        }
    }

    /// <summary>
    /// Base for all encrtyption engines.
    /// </summary>
    internal abstract class TEncryptionEngine
    {
        internal abstract bool CheckHash(string Password);
        internal abstract byte[] Decode(byte[] Data, long StreamPosition, int StartPos, int Count, int RecordLen);
        
        internal abstract byte[] Encode(byte[] Data, long StreamPosition, int StartPos, int Count, int RecordLen);
        internal abstract UInt16 Encode(UInt16 Data, long StreamPosition, int RecordLen);
        internal abstract UInt32 Encode(UInt32 Data, long StreamPosition, int RecordLen);

        internal abstract byte[] GetFilePassRecord();
        internal abstract int GetFilePassRecordLen();
    }

    /// <summary>
    /// Saves the records for file encryption.
    /// </summary>
    internal class TFileEncryption
    {
        internal TWriteProtRecord WriteProt;
        internal TWriteAccessRecord WriteAccess;
        internal TFileSharingRecord FileSharing;
        internal TMiscRecordList InterfaceHdr;

        internal TFileEncryption()
        {
            InterfaceHdr = new TMiscRecordList();
        }

        internal void Clear()
        {
            WriteProt = null;
            WriteAccess = null;
            FileSharing = null;
            InterfaceHdr.Clear();
        }

        internal long TotalSize()
        {
            long Result=InterfaceHdr.TotalSize;
            if (WriteProt!=null) Result+=WriteProt.TotalSize();
            if (WriteAccess!=null) Result+=WriteAccess.TotalSize();
            if (FileSharing!=null) Result+=FileSharing.TotalSize();
            
            return Result;
        }

        internal void SaveFirstPart(IDataStream DataStream, TSaveData SaveData)
        {
            if (WriteProt != null) WriteProt.SaveToStream(DataStream, SaveData,  0);            
		}

        internal void SaveSecondPart(IDataStream DataStream, TSaveData SaveData)
        {
            InterfaceHdr.SaveToStream(DataStream, SaveData,  0);
            if (WriteAccess != null) WriteAccess.SaveToStream(DataStream, SaveData,  0);
            if (FileSharing != null) FileSharing.SaveToStream(DataStream, SaveData,  0);
        }
    }

    /// <summary>
    /// Saves the records for file protection.
    /// </summary>
    internal class TWorkbookProtection
    {
        internal TWindowProtectRecord WindowProtect;
        internal TProtectRecord Protect;
        internal TPasswordRecord Password;
        internal TProt4RevRecord Prot4Rev;
        internal TProt4RevPassRecord Prot4RevPass;

        internal void Clear()
        {
            WindowProtect = null;
            Protect = null;
            Password = null;
            Prot4Rev = null;
            Prot4RevPass = null;
        }

        internal int TotalSize()
        {
            int Result = 0;
            if (WindowProtect != null) Result += WindowProtect.TotalSize();
            if (Protect != null) Result += Protect.TotalSize();
            if (Password != null) Result += Password.TotalSize();
            if (Prot4Rev != null) Result += Prot4Rev.TotalSize();
            if (Prot4RevPass != null) Result += Prot4RevPass.TotalSize();

            return Result;
        }

        internal void SaveToStream(IDataStream DataStream, TSaveData SaveData)
        {
            if (WindowProtect!= null) WindowProtect.SaveToStream(DataStream, SaveData, 0);
            if (Protect != null) Protect.SaveToStream(DataStream, SaveData, 0);
            if (Password != null) Password.SaveToStream(DataStream, SaveData, 0);
            if (Prot4Rev != null) Prot4Rev.SaveToStream(DataStream, SaveData, 0);
            if (Prot4RevPass != null) Prot4RevPass.SaveToStream(DataStream, SaveData, 0);
        }

    }

    /// <summary>
    /// Saves the records for sheet protection.
    /// </summary>
    internal class TSheetProtection
    {
        internal TProtectRecord Protect;
        internal TScenProtectRecord ScenProtect;
        internal TObjProtectRecord ObjProtect;
        internal TPasswordRecord Password;
        internal TSheetProtectRecord SheetProtect;

        internal void Clear()
        {
            Protect = null;
            ScenProtect = null;
            ObjProtect = null;
            Password = null;
            SheetProtect = null;
        }

        internal static TSheetProtection Clone(TSheetProtection Source, TSheetInfo SheetInfo)
        {
            TSheetProtection Result = new TSheetProtection();
            if (Source != null)
            {
                Result.Protect = (TProtectRecord) TBaseRecord.Clone(Source.Protect, SheetInfo);
                Result.ScenProtect = (TScenProtectRecord) TBaseRecord.Clone(Source.ScenProtect, SheetInfo);
                Result.ObjProtect = (TObjProtectRecord) TBaseRecord.Clone(Source.ObjProtect, SheetInfo);
                Result.Password = (TPasswordRecord) TBaseRecord.Clone(Source.Password, SheetInfo);
                Result.SheetProtect = (TSheetProtectRecord) TBaseRecord.Clone(Source.SheetProtect, SheetInfo);
            }
            return Result;
        }

        internal int TotalSizeFirst()
        {
            int Result = 0;
            if (Protect != null) Result += Protect.TotalSize();
            if (ScenProtect != null) Result += ScenProtect.TotalSize();
            if (ObjProtect != null) Result += ObjProtect.TotalSize();
            if (Password != null) Result += Password.TotalSize();

            return Result;
        }

        internal int TotalSizeDialogSheet()
        {
            int Result = 0;
            if (Protect != null) Result += Protect.TotalSize();
            if (Password != null) Result += Password.TotalSize();

            return Result;
        }

        internal int TotalSizeSecond()
        {
            int Result = 0;
            if (SheetProtect != null) Result += SheetProtect.TotalSize();

            return Result;
        }

        internal void SaveFirstPart(IDataStream DataStream, TSaveData SaveData)
        {
            if (Protect!=null) Protect.SaveToStream(DataStream, SaveData, 0);
            if (ScenProtect!= null) ScenProtect.SaveToStream(DataStream, SaveData, 0);
            if (ObjProtect!=null) ObjProtect.SaveToStream(DataStream, SaveData, 0);
            if (Password!=null) Password.SaveToStream(DataStream, SaveData, 0);
        }

        internal void SaveDialogSheetProtect(IDataStream DataStream, TSaveData SaveData)
        {
            if (Protect != null) Protect.SaveToStream(DataStream, SaveData, 0);
            if (Password != null) Password.SaveToStream(DataStream, SaveData, 0);
        }


        internal void SaveSecondPart(IDataStream DataStream, TSaveData SaveData)
        {
            if (SheetProtect!=null) SheetProtect.SaveToStream(DataStream, SaveData, 0);
        }
    }

}

