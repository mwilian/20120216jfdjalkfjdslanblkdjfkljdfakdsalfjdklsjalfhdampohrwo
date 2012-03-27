using System;
using FlexCel.Core;

namespace FlexCel.XlsAdapter
{
    #region File Encryption
    /// <summary>
    /// Filepass. When this record is present, file is encrypted. The encryption algorithm can change.
    /// </summary>
    internal class TFilePassRecord: TxBaseRecord
    {
        bool IsPxl;
        internal TFilePassRecord(int aId, byte[] aData, bool aIsPxl): base(aId, aData)
        {
            IsPxl = aIsPxl;
        }

        internal TEncryptionType EncryptionType
        {
            get
            {
                if (IsPxl) return TEncryptionType.Xor; //pxl.

                if (GetWord(0)==0) return TEncryptionType.Xor;
                if (GetWord(4)==2) return TEncryptionType.Strong;                
                return TEncryptionType.Standard;
            }
        }

        internal TEncryptionEngine CreateEncryptionEngine(string Password)
        {
            if (IsPxl) return new TXorEncryption(GetWord(0), GetWord(2), Password); //pxl.

            if (EncryptionType== TEncryptionType.Xor) return new TXorEncryption(GetWord(2), GetWord(4), Password);
            if (EncryptionType== TEncryptionType.Standard) return new TStandardEncryption(GetArray(6,16), GetArray(22,16), GetArray(38,16));
            XlsMessages.ThrowException(XlsErr.ErrNotSupportedEncryption);
            return null; //just to compile
        }
    }


    /// <summary>
    /// This record is used together with FileSharing to protect against writing.
    /// </summary>
    internal class TWriteProtRecord: TxBaseRecord
    {
        internal TWriteProtRecord(int aId, byte[] aData): base(aId, aData)
        {
        }

        internal TWriteProtRecord(): base((int) xlr.WRITEPROT, new byte[0])
        {
        }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.FileEncryption.WriteProt = this;
        }
    }

    /// <summary>
    /// Name of the user who saved the file.
    /// </summary>
    internal class TWriteAccessRecord: TxBaseRecord
    {
        internal TWriteAccessRecord(int aId, byte[] aData): base(aId, aData)
        {
        }

        internal string UserName
        {
            get
            {
                string Result = String.Empty;
                long StSize = 0;
                StrOps.GetSimpleString(true, Data, 0, false, 0, ref Result, ref StSize);
                return Result;
            }
            set
            {
                string v = value == null? String.Empty: value;
                if (v.Length > 31) v = v.Substring(0, 31);
                TExcelString Xs = new TExcelString(TStrLenLength.is16bits, v, null, false);
                Xs.CopyToPtr(Data, 0);
                for (int i = Xs.TotalSize();i < Data.Length;i++)
                {
                    Data[i] = 32;
                }
            }
        }

        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
        {
            if ((SaveData.ExcludedRecords & TExcludedRecords.WriteAccess) != 0) return; //Note that this will invalidate the size, but it doesnt matter as this is not saved for real use. We could write blanks here if we wanted to keep the offsets right.
            base.SaveToStream(Workbook, SaveData, Row);
        }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.FileEncryption.WriteAccess = this;
        } 
    }

    /// <summary>
    /// Password and settings for write protection.
    /// </summary>
    internal class TFileSharingRecord: TxBaseRecord
    {
        internal TFileSharingRecord(int aId, byte[] aData): base(aId, aData)
        {
        }

        internal TFileSharingRecord(bool aRecommendReadOnly, string aPassword, string aUser, bool PassIsHash): base((int)xlr.FILESHARING, new byte[0])
        {
            if (aUser==null) aUser=String.Empty;
            TExcelString Xs = new TExcelString(TStrLenLength.is16bits, aUser, null, false);
            Data= new byte[4+Xs.TotalSize()];
            RecommendReadOnly = aRecommendReadOnly;
            if (PassIsHash)
            {
                if (aPassword == null || aPassword.Length == 0)
                {
                    SetWord(2, 0);
                }
                else
                {
                    if (aPassword.Length <= 4) //we will only use standard pass, which is what Excel 2010 uses.
                    {
                        SetWord(2, TCompactFramework.HexToNumber(aPassword));
                    }
                }
            }
            else
            {
                SetPassword(aPassword);
            }
            Xs.CopyToPtr(Data,4);
        }

        internal bool RecommendReadOnly
        {
            get
            {
                return GetWord(0)==1;
            }
            set
            {
                if (value) SetWord(0,1); else SetWord(0,0);
            }
        }

        internal void SetPassword(string Pass)
        {
            if (Pass==null || Pass.Length==0)
                SetWord(2, 0);
            else
                SetWord(2, TXorEncryption.CalcHash(Pass));
        }

        internal int HashedPass
        {
            get
            {
                return GetWord(2);
            }
        }

        internal string User
        {
            get
            {
                string st = null;
                long stSize = 0;
                StrOps.GetSimpleString(true, Data, 4, false, 0, ref st, ref stSize);
                return st;
            }
        }


        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.FileEncryption.FileSharing = this;
        }
    }

    #endregion

    #region Common protection
    /// <summary>
    /// Base for all general protection records.
    /// </summary>
    internal class TBaseProtectRecord: TxBaseRecord
    {
        internal TBaseProtectRecord(int aId, byte[] aData): base(aId, aData)
        {
        }

        internal bool Protected
        {
            get
            {
                return GetWord(0)==1;
            }
            set
            {
                if (value) SetWord(0,1); else SetWord(0,0);
            }
        }
    }

    /// <summary>
    /// Workbook /worksheet is protected.
    /// </summary>
    internal class TProtectRecord: TBaseProtectRecord
    {
        internal TProtectRecord(int aId, byte[] aData): base(aId, aData){}
        internal TProtectRecord() : base((int)xlr.PROTECT, new byte[2]) { }
        
        internal TProtectRecord(bool Protect) : this()
        {
            Protected = Protect;
        }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.WorkbookProtection.Protect = this;
        }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.SheetProtection.Protect = this;
        }
    }

    /// <summary>
    /// Objects on Workbook /worksheet are protected.
    /// </summary>
    internal class TObjProtectRecord: TBaseProtectRecord
    {
        internal TObjProtectRecord(int aId, byte[] aData): base(aId, aData){}
        internal TObjProtectRecord(): base((int) xlr.OBJPROTECT, new byte[2]){}

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.SheetProtection.ObjProtect = this;
        }
    }

    /// <summary>
    /// Window configuration on Workbook /worksheet are protected.
    /// </summary>
    internal class TWindowProtectRecord : TBaseProtectRecord
    {
        internal TWindowProtectRecord(int aId, byte[] aData) : base(aId, aData) { }
        internal TWindowProtectRecord() : base((int)xlr.WINDOWPROTECT, new byte[2]) { }

        internal TWindowProtectRecord(bool Protect)
            : this()
        {
            Protected = Protect;
        }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.WorkbookProtection.WindowProtect = this;
        }
    }

    /// <summary>
    /// Scenarios on Workbook /worksheet are protected.
    /// </summary>
    internal class TScenProtectRecord: TBaseProtectRecord
    {
        internal TScenProtectRecord(int aId, byte[] aData): base(aId, aData){}
        internal TScenProtectRecord(): base((int) xlr.SCENPROTECT, new byte[2]){}

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.SheetProtection.ScenProtect = this;
        }
    }

    /// <summary>
    /// Password for protection.
    /// </summary>
    internal class TPasswordRecord: TxBaseRecord
    {
        internal TPasswordRecord(int aId, byte[] aData): base(aId, aData){}
        internal TPasswordRecord() : base((int)xlr.PASSWORD, new byte[2]) { }
        internal TPasswordRecord(int Hash) : this() 
        {
            SetWord(0, Hash);
        }

        internal void SetPassword(string Pass)
        {
            if (Pass==null || Pass.Length==0)
                SetWord(0, 0);
            else
                SetWord(0, TXorEncryption.CalcHash(Pass));
        }

        internal int GetHash()
        {
            return GetWord(0);
        }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.WorkbookProtection.Password = this;
        }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.SheetProtection.Password = this;
        }
    }

    internal enum SharedFeatureType
    {
        ISFPROTECTION = 0x0002,  //Specifies the enhanced protection type. A Shared Feature of this type is used to protect a shared workbook by restricting access to the areas of the workbook and to the available functionality. 
        ISFFEC2 = 0x0003, //  Specifies the ignored formula errors type. A Shared Feature of this type is used to specify the formula errors to be ignored. 
        ISFFACTOID = 0x0004, //  Specifies the smart tag type. A Shared Feature of this type is used to recognize certain types of entries (for example, proper names, dates/times, financial symbols) and flag them for action.  
        ISFLIST = 0x0005 // Specifies the list type. A Shared Feature of this type is used to describe a table within a sheet.
    }

    internal class TFeatHdrRecord : TxBaseRecord
    {
        protected TFeatHdrRecord(int aId, byte[] aData) : base(aId, aData) { }

        internal static TFeatHdrRecord Create(int aId, byte[] aData)
        {
            if (aData.Length < 14) return new TFeatHdrRecord(aId, aData);
            switch ((SharedFeatureType)aData[12])
            {
                case SharedFeatureType.ISFPROTECTION:
                    return new TSheetProtectRecord(aId, aData);

                case SharedFeatureType.ISFFEC2:
                    return new TFeatHdrRecord(aId, aData);
                
                case SharedFeatureType.ISFFACTOID:
                    return new TFeatHdrRecord(aId, aData);
                    
                case SharedFeatureType.ISFLIST:
                    return new TFeatHdrRecord(aId, aData);

                default:
                    return new TFeatHdrRecord(aId, aData);
            }
            
        }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.FeatHdr.Add(this);
        }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.Feat.Add(this);
        }
    }

    /// <summary>
    /// Only for excel xp and up. Note that this record has the same header as SmartTag :(
    /// </summary>
    internal class TSheetProtectRecord: TFeatHdrRecord
    {
        internal TSheetProtectRecord(int aId, byte[] aData): base(aId, aData)
        {
            if (Data.Length < 23) //This record might be smaller than 23
            {
                Data = new byte[23];
                Data[19] = 0x44;
                Array.Copy(aData, 0, Data, 0, aData.Length);
                FillFixedData(); //Record is not standard, so we "fix" it here just in case.
            }
        }

        internal void FillFixedData()
        {
            SetWord(0, 0x0867);
            SetCardinal(11, 0x01000200);
            SetCardinal(15, 0xFFFFFFFF);
        }

        internal TSheetProtectRecord(): base((int)xlr.FEATHDR, new byte[23])
        {
            FillFixedData();
            Data[19] = 0x44;
        }

        internal bool GetProtect(int i)
        {
            return (GetWord(19) & (1 <<i)) !=0;
        }

        internal void SetProtect(int i, bool value)
        {
            if (value)
                SetWord(19, GetWord(19) | (1 <<i));
            else
                SetWord(19, GetWord(19) & ~(1 <<i));
        }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.SheetProtection.SheetProtect = this;
        }
    }
    #endregion

    #region Workbook Protect
    /// <summary>
    /// Shared protection.
    /// </summary>
    internal class TProt4RevRecord: TBaseProtectRecord
    {
        internal TProt4RevRecord(int aId, byte[] aData): base(aId, aData){}
        internal TProt4RevRecord() : base((int)xlr.PROT4REV, new byte[2]) { }
        
        internal TProt4RevRecord(bool Protect) : this() 
        {
            Protected = Protect;
        }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.WorkbookProtection.Prot4Rev = this;
        }
    }

    /// <summary>
    /// Password for rev protection.
    /// </summary>
    internal class TProt4RevPassRecord : TxBaseRecord
    {
        internal TProt4RevPassRecord(int aId, byte[] aData) : base(aId, aData) { }
        internal TProt4RevPassRecord() : base((int)xlr.PROT4REVPASS, new byte[2]) { }

        internal TProt4RevPassRecord(int Hash) : this()
        {
            SetWord(0, Hash);
        }

        internal void SetPassword(string Pass)
        {
            if (Pass == null || Pass.Length == 0)
                SetWord(0, 0);
            else
                SetWord(0, TXorEncryption.CalcHash(Pass));
        }

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.WorkbookProtection.Prot4RevPass = this;
        }
    }

    #endregion
}
