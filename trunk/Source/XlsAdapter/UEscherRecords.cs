using System;
using System.IO;
using System.Text;
using FlexCel.Core;
using System.Diagnostics;
using System.Globalization;
using System.Collections.Generic;

namespace FlexCel.XlsAdapter
{
    /// <summary>
    /// An Escher Record Header, that is different from an Excel Record Header
    /// </summary>
    internal class TEscherRecordHeader
    {
        byte[] FData;


        internal TEscherRecordHeader()
        {
            FData = new byte[XlsEscherConsts.SizeOfTEscherRecordHeader];
        }

        internal TEscherRecordHeader(byte[] aData)
        {
            FData = aData;
        }

        internal TEscherRecordHeader(int aPre, int aId, long aSize)
        {
            FData = new byte[XlsEscherConsts.SizeOfTEscherRecordHeader];
            Pre=aPre;
            Id=aId;
            Size=aSize;
        }

        internal byte[] Data{ get {return FData;}}

        internal int Pre { get {return BitConverter.ToUInt16(FData,0);} set { BitConverter.GetBytes((UInt16)value).CopyTo(FData,0);}}
        internal int Id  { get {return BitConverter.ToUInt16(FData,2);} set { BitConverter.GetBytes((UInt16)value).CopyTo(FData,2);}}
        internal long Size  { get {return BitConverter.ToUInt32(FData,4);} set { BitConverter.GetBytes((UInt32)value).CopyTo(FData,4);}}

        internal int Length { get { return FData.Length;}}
    }


    internal class TEscherDwgCache
    {
        internal int MaxObjId;
        internal TEscherDgRecord Dg;
        internal TEscherSolverContainerRecord Solver;
        internal TEscherSpgrContainerRecord Patriarch;
        internal TEscherAnchorCache AnchorList;
        internal TEscherShapeCache Shape;
        internal TEscherObjCache Obj;
        internal TEscherOPTCache Blip;
        internal TEscherOptByNameCache OptByName;
        internal TRadioButtonCache RadioButtons;

        internal TEscherDwgCache(){}
        internal TEscherDwgCache(
            int aMaxObjId,
            TEscherDgRecord aDg,
            TEscherSolverContainerRecord aSolver,
            TEscherSpgrContainerRecord aPatriarch,
            TEscherAnchorCache aAnchorList,
            TEscherShapeCache aShape,
            TEscherObjCache aObj,
            TEscherOPTCache aBlip,
            TRadioButtonCache aRadioButtons,
            TEscherOptByNameCache aOptByName)
        {
            MaxObjId   = aMaxObjId;
            Dg         = aDg;
            Solver     = aSolver;
            Patriarch  = aPatriarch;
            AnchorList = aAnchorList;
            Shape      = aShape;
            Obj        = aObj;
            Blip       = aBlip;
            OptByName  = aOptByName; 
            RadioButtons = aRadioButtons;
        }
        
    }

    internal class TEscherDwgGroupCache
    {
        internal TEscherBStoreRecord BStore;
        internal TEscherDggRecord Dgg;
    }

    internal abstract class TEscherRecord: IComparable
    {
                       
        protected TEscherContainerRecord FParent;

        protected TEscherDwgCache DwgCache;
        protected TEscherDwgGroupCache DwgGroupCache;

        internal int Pre;
        internal int Id;
        internal int TotalDataSize;  
        internal int LoadedDataSize; 

        private TEscherRecord FCopiedTo;
        private TCopiedGen CopiedToGen;

        internal TEscherRecord(TEscherRecordHeader aEscherHeader, TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
        {
            LoadedDataSize=0;
            if (aEscherHeader!=null)
            {
                TotalDataSize=(int)aEscherHeader.Size;
                Id= aEscherHeader.Id;
                Pre= aEscherHeader.Pre;
            }
            DwgGroupCache= aDwgGroupCache;
            DwgCache= aDwgCache;
            FParent=aParent;
        }

        protected virtual TEscherRecord DoCopyTo(int RowOfs, int ColOfs, TEscherDwgCache NewDwgCache, TEscherDwgGroupCache NewDwgGroupCache, TSheetInfo SheetInfo)
        {
            //This should recreate all actions on the constructor. MemberwiseClone does NOT call the constructor.
            TEscherRecord Result = (TEscherRecord)MemberwiseClone();
            Result.DwgGroupCache = NewDwgGroupCache;
            Result.DwgCache = NewDwgCache;
            Result.FParent = null;

            if (FParent != null)
                if (FParent.CopiedTo(SheetInfo.CopiedGen) != null) Result.FParent = (TEscherContainerRecord)FParent.CopiedTo(SheetInfo.CopiedGen);
                else Result.FParent = FParent;
            FCopiedTo = Result;
            CopiedToGen = SheetInfo.CopiedGen;
            return Result;
        }

        internal virtual TEscherContainerRecord Parent {get {return FParent;} set{FParent=value;}}

        /// <summary>
        /// We don't have deterministic destructors here. So this works as a dispose call.
        /// </summary>
        internal virtual void Destroy()
        {
        }

        internal virtual void AfterCreate()
        {
            //Nothing here
        }

        internal TEscherRecord CopiedTo(TCopiedGen aCopiedToGen)
        {
            if (CopiedToGen != aCopiedToGen) { FCopiedTo = null; }
            return FCopiedTo;            
        }
        
        protected static void IncNextPos(ref int NextPos, int Size,ref int RealSize, TBreakList BreakList, int ContinueRecord, int ExtraDataLen)
        {
            if (NextPos> XlsConsts.MaxRecordDataSize+1-ExtraDataLen) XlsMessages.ThrowException(XlsErr.ErrInternal);
            NextPos += Size;
            RealSize += Size;
            while (NextPos>XlsConsts.MaxRecordDataSize+1-ExtraDataLen) 
            {
                NextPos -= XlsConsts.MaxRecordDataSize+1-ExtraDataLen;
                RealSize+= XlsConsts.SizeOfTRecordHeader+ExtraDataLen;  //continue record
                if (BreakList!=null) BreakList.Add(ContinueRecord, XlsConsts.MaxRecordDataSize+1-ExtraDataLen);
            }

        }

        protected static void CheckSplit(IDataStream DataStream, TSaveData SaveData, TBreakList BreakList)
        {
            if (DataStream.Position > BreakList.AcumSize()) XlsMessages.ThrowException(XlsErr.ErrInternal);
            if (DataStream.Position == BreakList.AcumSize())
            {
                WriteNewRecord(DataStream, SaveData, BreakList);
                BreakList.IncCurrent();
            }
        }

        private static void WriteNewRecord(IDataStream DataStream, TSaveData SaveData, TBreakList BreakList)
        {
            DataStream.WriteHeader((UInt16)BreakList.CurrentId, (UInt16)BreakList.CurrentSize);
            DataStream.Write(BreakList.GetExtraData, BreakList.ExtraDataLen());
        }


        protected int Instance()
        {
            return  (Pre >> 4)& 0x0FFF;
        }

        protected int Version()
        {
            return Pre & 0x0F;
        }

        internal virtual bool HasExternRefs()
        {
            return false;
        }
            
        internal abstract void Load(ref TxBaseRecord aRecord, ref int aPos, bool HeaderImage, bool aChartCoords);

        internal virtual void SaveToStream(IDataStream DataStream, TSaveData SaveData, TBreakList BreakList)
        {
            if (!Loaded()) XlsMessages.ThrowException(XlsErr.ErrEscherNotLoaded);

            TEscherRecordHeader Rs= new TEscherRecordHeader(Pre, Id, TotalSizeNoSplit()- XlsEscherConsts.SizeOfTEscherRecordHeader);
            CheckSplit(DataStream, SaveData, BreakList);
            int Remaining= (int)(BreakList.AcumSize() - DataStream.Position) ;
            if (Rs.Length>Remaining) 
            {
                DataStream.Write(Rs.Data, Remaining);
                CheckSplit(DataStream, SaveData, BreakList);
                DataStream.Write(Rs.Data, Remaining, Rs.Length-Remaining);
            }
            else DataStream.Write(Rs.Data, Rs.Length);

        }

        internal static TEscherRecord Clone(TEscherRecord Self, int RowOfs, int ColOfs, TEscherDwgCache NewDwgCache, TEscherDwgGroupCache NewDwgGroupCache, TSheetInfo SheetInfo) //this should be non-virtual. It allows you to obtain a clone, even if the object is null
        {
            if (Self==null) return null;   //for this to work, this cant be a virtual method
            else return Self.DoCopyTo(RowOfs, ColOfs, NewDwgCache, NewDwgGroupCache, SheetInfo);
        }

        internal virtual long TotalSizeNoSplit()
        {
            return XlsEscherConsts.SizeOfTEscherRecordHeader;
        }

        internal virtual bool Loaded()
        {
            if (LoadedDataSize > TotalDataSize) XlsMessages.ThrowException(XlsErr.ErrInternal);
            return TotalDataSize == LoadedDataSize;
        }

        internal static bool IsContainer(int aPre)
        {
            return (aPre & 0x000F ) == 0x000F;
        }

        internal virtual bool WaitingClientData(ref TClientType ClientType)
        {
            ClientType=TClientType.Null;
            return false;
        }

        internal virtual void AssignClientData(TBaseClientData aClientData, bool Copied)
        {
            XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
        }

        internal virtual void SplitRecords(ref int NextPos, ref int RealSize, ref int NextDwg, TBreakList BreakList, int ContinueRecord, int ExtraDataLen)
        {
            if (NextDwg > 0)
            {
                if (BreakList != null) BreakList.Add(NextDwg, NextPos);
                RealSize += XlsConsts.SizeOfTRecordHeader;
                NextPos = 0;
                NextDwg = -1;
            }

            IncNextPos(ref NextPos, XlsEscherConsts.SizeOfTEscherRecordHeader, ref RealSize, BreakList, ContinueRecord, ExtraDataLen);

        }

        internal virtual void ArrangeInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo, bool Forced)
        {
            //Nothing here
        }

        internal virtual void ArrangeMoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
        {
            //Nothing here
        }

        internal TEscherRecord FindRoot()
        {
            TEscherRecord Result=this;
            if (DwgCache==null) return Result;
            while ((Result != null) && (Result.FParent != DwgCache.Patriarch)) Result = Result.FParent;
            return Result;
        }

        internal virtual T FindRec<T>() where T: TEscherRecord
        {
            return null;
        }

        internal TEscherSpgrContainerRecord Patriarch()
        {
            if (DwgCache==null) return null; else
                return DwgCache.Patriarch;
        }

        internal TEscherRecord CopyDwg(int RowOfs, int ColOfs, TSheetInfo SheetInfo)
        {
            if ((DwgCache.Patriarch==null) || (FindRoot()==null)) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
            DwgCache.Patriarch.FContainedRecords.Add(TEscherRecord.Clone(FindRoot(), RowOfs, ColOfs, DwgCache, DwgGroupCache, SheetInfo));
            return CopiedTo(SheetInfo.CopiedGen);
        }

        internal virtual int CompareRec(TEscherRecord aRecord) //this is used for searching
        {
            if (Id < aRecord.Id) return -1;
            else if (aRecord.Id > Id) return 1;
            else
                if (Pre < aRecord.Pre) return -1;
                else if (Pre > aRecord.Pre) return 1;
                else
                    if (TotalDataSize < aRecord.TotalDataSize) return -1;
                    else if (TotalDataSize > aRecord.TotalDataSize) return 1;
                    else
                        return 0;
        }
        #region IComparable Members

        public int CompareTo(object obj)
        {
            return CompareRec((TEscherRecord) obj);
        }

        #endregion
    }

    /// <summary>
    /// An escher record containing data and not other Escher records
    /// </summary>
    internal class TEscherDataRecord: TEscherRecord
    {
        protected byte[]Data;
            
        internal TEscherDataRecord(TEscherRecordHeader aEscherHeader, TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent )
            : base(aEscherHeader, aDwgGroupCache, aDwgCache, aParent)
        {
            Data = new byte[TotalDataSize];
        }

        protected override TEscherRecord DoCopyTo(int RowOfs, int ColOfs, TEscherDwgCache NewDwgCache, TEscherDwgGroupCache NewDwgGroupCache, TSheetInfo SheetInfo)
        {
            //This should recreate all actions on the constructor.
            TEscherDataRecord Result = (TEscherDataRecord)base.DoCopyTo(RowOfs, ColOfs, NewDwgCache, NewDwgGroupCache, SheetInfo);
            if (Data != null) Result.Data = (byte[])Data.Clone();
            return Result;
        }

        internal override void Load(ref TxBaseRecord aRecord, ref int aPos, bool HeaderImage, bool aChartCoords)
        {
            if (TotalDataSize == 0) return;
            int RSize = aRecord.TotalSizeNoHeaders() - aPos;
            if (RSize <= 0) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
            if (TotalDataSize - LoadedDataSize < RSize) RSize = TotalDataSize - LoadedDataSize;
            if (LoadedDataSize + RSize > TotalDataSize) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
            BitOps.ReadMem(ref aRecord, ref aPos, Data, LoadedDataSize, RSize);
            LoadedDataSize += RSize;
        }

        internal override void SaveToStream(IDataStream DataStream, TSaveData SaveData, TBreakList BreakList)
        {
            base.SaveToStream(DataStream, SaveData, BreakList);
            if (TotalDataSize > 0)
            {
                int RemainingSize= TotalDataSize;
                while (RemainingSize > BreakList.AcumSize() - DataStream.Position)
                {
                    int FracSize = (int)(BreakList.AcumSize() - DataStream.Position);
                    CheckSplit(DataStream, SaveData, BreakList);
                    DataStream.Write(Data, TotalDataSize - RemainingSize, FracSize);

                    RemainingSize -= FracSize;
                } //while

                CheckSplit(DataStream, SaveData, BreakList);
                DataStream.Write(Data, TotalDataSize - RemainingSize, RemainingSize);
            }
        }

        internal override long TotalSizeNoSplit()
        {
            return base.TotalSizeNoSplit() + TotalDataSize;
        }

        internal override void SplitRecords(ref int NextPos, ref int RealSize, ref int NextDwg, TBreakList BreakList, int ContinueRecord, int ExtraDataLen)
        {
            base.SplitRecords(ref NextPos, ref RealSize, ref NextDwg, BreakList, ContinueRecord, ExtraDataLen);
            IncNextPos(ref NextPos, TotalDataSize, ref RealSize, BreakList, ContinueRecord, ExtraDataLen);
        }

        internal override int CompareRec(TEscherRecord aRecord) //this is used for searching
        {
            int Result = base.CompareRec(aRecord);
            if (Result == 0 && Data != null)
            {
                for (int i = 0; i < TotalDataSize; i++)
                {
                    Result = Data[i].CompareTo(((TEscherDataRecord)aRecord).Data[i]);
                    if (Result != 0) return Result;
                }
            }
            return Result;
        }

        internal void ClearData()
        {
            Array.Clear(Data, 0, Data.Length);
        }
    }

    /// <summary>
    /// Record containing other records. An escher record is a DataRecord or a ContainerRecord 
    /// </summary>
    class TEscherContainerRecord: TEscherRecord
    {
        internal TEscherRecordList FContainedRecords;
                
        internal TEscherContainerRecord(TEscherRecordHeader aEscherHeader, TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
            : base(aEscherHeader, aDwgGroupCache, aDwgCache, aParent)
        {
            Init();
        }

        private void Init()
        {
            FContainedRecords= new TEscherRecordList();
        }

        internal override void Destroy()
        {
            base.Destroy();
            FContainedRecords.Destroy();
        }

        protected override TEscherRecord DoCopyTo(int RowOfs, int ColOfs, TEscherDwgCache NewDwgCache, TEscherDwgGroupCache NewDwgGroupCache, TSheetInfo SheetInfo)
        {
            //This should recreate all actions on the constructor.
            TEscherContainerRecord Result = (TEscherContainerRecord) base.DoCopyTo(RowOfs, ColOfs, NewDwgCache, NewDwgGroupCache, SheetInfo);
            Result.Init();
            Result.FContainedRecords.CopyFrom(RowOfs, ColOfs, FContainedRecords, NewDwgCache, NewDwgGroupCache, SheetInfo);
            return Result;
        }

        internal TEscherRecordList ContainedRecords{ get { return FContainedRecords;}}

        internal override void Load(ref TxBaseRecord aRecord, ref int aPos, bool HeaderImage, bool aChartCoords)
        {
            int RSize=aRecord.TotalSizeNoHeaders();
            TEscherRecordHeader RecordHeader= new TEscherRecordHeader();
            if (aPos> RSize) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
            while ((! Loaded()) && (aPos<RSize))
            {
                if (aRecord.Continue == null && aPos == aRecord.Data.Length) return; //There is nothing more to read, we need to load the next record from disk. This can happen when reading nested MsObjs.

                if ((FContainedRecords.Count==0) || (LastRecord.Loaded()))
                {
                    BitOps.ReadMem(ref aRecord, ref aPos, RecordHeader.Data);

                    if (IsContainer(RecordHeader.Pre))
                        switch ((Msofbt)RecordHeader.Id)
                        {
                            case Msofbt.BstoreContainer:
                                FContainedRecords.Add(new TEscherBStoreRecord(RecordHeader, DwgGroupCache, DwgCache, this));
                                break;
                            case Msofbt.SpgrContainer:
                                FContainedRecords.Add(new TEscherSpgrContainerRecord(RecordHeader, DwgGroupCache, DwgCache, this));
                                break;
                            case Msofbt.SpContainer:
                                FContainedRecords.Add(new TEscherSpContainerRecord(RecordHeader, DwgGroupCache, DwgCache, this));
                                break;
                            case Msofbt.SolverContainer:
                                FContainedRecords.Add(new TEscherSolverContainerRecord(RecordHeader, DwgGroupCache, DwgCache, this));
                                break;
                            default:
                                FContainedRecords.Add(new TEscherContainerRecord(RecordHeader, DwgGroupCache, DwgCache, this));
                                break;
                        }
                    else
                        switch ((Msofbt)RecordHeader.Id)
                        {
                            case Msofbt.ClientData:
                                FContainedRecords.Add(new TEscherClientDataRecord(RecordHeader, DwgGroupCache, DwgCache, this));
                                break;
                            case Msofbt.ClientTextbox:
                                FContainedRecords.Add(new TEscherClientTextBoxRecord(RecordHeader, DwgGroupCache, DwgCache, this));
                                break;
                            case Msofbt.ClientAnchor:
                                if (HeaderImage)
                                    FContainedRecords.Add(new TEscherHeaderAnchorRecord(RecordHeader, DwgGroupCache, DwgCache, this));
                                else
                                    FContainedRecords.Add(new TEscherImageAnchorRecord(RecordHeader, DwgGroupCache, DwgCache, this));
                                break;
                            case Msofbt.ChildAnchor:
                                FContainedRecords.Add(new TEscherChildAnchorRecord(RecordHeader, DwgGroupCache, DwgCache, this));
                                break;

                            case Msofbt.BSE:
                                FContainedRecords.Add(new TEscherBSERecord(RecordHeader, DwgGroupCache, DwgCache, this));
                                break;
                            case Msofbt.Dg:
                                FContainedRecords.Add(new TEscherDgRecord(RecordHeader, DwgGroupCache, DwgCache, this));
                                break;
                            case Msofbt.Dgg:
                                FContainedRecords.Add(new TEscherDggRecord(RecordHeader, DwgGroupCache, DwgCache, this));
                                break;
                            case Msofbt.Sp:
                                FContainedRecords.Add(new TEscherSpRecord(RecordHeader, DwgGroupCache, DwgCache, this));
                                break;
                            case Msofbt.Spgr:
                                FContainedRecords.Add(new TEscherSpgrRecord(RecordHeader, DwgGroupCache, DwgCache, this));
                                break;
                            case Msofbt.OPT:
                                TEscherOPTRecord Opt = new TEscherOPTRecord(RecordHeader, DwgGroupCache, DwgCache, this);
                                FContainedRecords.Add(Opt);
                                Opt.AddShapeId();
                                break;
                            case Msofbt.SplitMenuColors:
                                FContainedRecords.Add(new TEscherSplitMenuRecord(RecordHeader, DwgGroupCache, DwgCache, this));
                                break;

                            case Msofbt.ConnectorRule:
                                FContainedRecords.Add(new TEscherConnectorRuleRecord(RecordHeader, DwgGroupCache, DwgCache, this));
                                break;
                            case Msofbt.AlignRule:
                                FContainedRecords.Add(new TEscherAlignRuleRecord(RecordHeader, DwgGroupCache, DwgCache, this));
                                break;
                            case Msofbt.ArcRule:
                                FContainedRecords.Add(new TEscherArcRuleRecord(RecordHeader, DwgGroupCache, DwgCache, this));
                                break;
                            case Msofbt.CalloutRule:
                                FContainedRecords.Add(new TEscherCallOutRuleRecord(RecordHeader, DwgGroupCache, DwgCache, this));
                                break;
                            default:
                                FContainedRecords.Add(new TEscherDataRecord(RecordHeader, DwgGroupCache, DwgCache, this));
                                break;
                        } //case
                }

                LastRecord.Load(ref aRecord, ref aPos, HeaderImage, aChartCoords);
                if (LastRecord.Loaded())
                {
                    LoadedDataSize+= (int) LastRecord.TotalSizeNoSplit();
                    LastRecord.AfterCreate();
                }
            }
        }

        internal override void SaveToStream(IDataStream DataStream, TSaveData SaveData, TBreakList BreakList)
        {
            base.SaveToStream(DataStream, SaveData, BreakList);
            FContainedRecords.SaveToStream(DataStream, SaveData, BreakList);
        }

        internal override long TotalSizeNoSplit()
        {
            return base.TotalSizeNoSplit() + FContainedRecords.TotalSizeNoSplit();
        }

        internal override bool WaitingClientData(ref TClientType ClientType)
        {
            if (FContainedRecords.Count==0) return false;
            else return LastRecord.WaitingClientData(ref ClientType);
        }

        internal override void AssignClientData(TBaseClientData aClientData, bool Copied)
        {
            LastRecord.AssignClientData(aClientData, Copied);
        }

        internal override bool HasExternRefs()
        {
            for (int i=0;i<FContainedRecords.Count;i++)
                if (FContainedRecords[i].HasExternRefs()) return true;
            return false;
        }


        internal TEscherRecord LastRecord
        { 
            get 
            {
                if (FContainedRecords.Count==0) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
                return FContainedRecords[FContainedRecords.Count-1];
            }
        }


        internal override void SplitRecords(ref int NextPos, ref int RealSize, ref int NextDwg, TBreakList BreakList, int ContinueRecord, int ExtraDataLen)
        {
            base.SplitRecords(ref NextPos, ref RealSize, ref NextDwg, BreakList, ContinueRecord, ExtraDataLen);
            for (int i = 0; i < FContainedRecords.Count; i++)
                FContainedRecords[i].SplitRecords(ref NextPos, ref RealSize, ref NextDwg, BreakList, ContinueRecord, ExtraDataLen);
        }

        internal override void ArrangeInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo, bool Forced)
        {
            base.ArrangeInsertRange(CellRange, aRowCount, aColCount, SheetInfo, Forced);
            for (int i = 0; i < FContainedRecords.Count; i++) FContainedRecords[i].ArrangeInsertRange(CellRange, aRowCount, aColCount, SheetInfo, Forced);
        }

        internal override void ArrangeMoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
        {
            base.ArrangeMoveRange (CellRange, NewRow, NewCol, SheetInfo);
            for (int i=0;i<FContainedRecords.Count;i++) FContainedRecords[i].ArrangeMoveRange(CellRange, NewRow,NewCol, SheetInfo);

        }

        /// <summary>
        ///  FindRec goes only one level down, it is not recursive.
        /// </summary>
        internal override T FindRec<T>()
        {
            for (int i = 0; i < FContainedRecords.Count; i++)
            {
                T Result = FContainedRecords[i] as T;
                if (Result != null) return Result;
            }

            return null;
        }

        internal static TEscherContainerRecord Navigate(TEscherDwgCache RecordCache, int ObjIndex, string ObjectPath)
        {
            string OriginalObjectPath = ObjectPath;
            if (ObjectPath != null)
            {
                if (ObjectPath.StartsWith(FlxConsts.ObjectPathAbsolute, StringComparison.InvariantCulture))
                {
                    ReadAbsolutePath(ref ObjIndex, ref ObjectPath);
                }
                else
                    if (ObjectPath.StartsWith(FlxConsts.ObjectPathObjName, StringComparison.InvariantCulture))
                    {
                        TEscherOPTRecord Opt = RecordCache.OptByName.FindObj(ObjectPath.Substring(1));
                        if (Opt == null) XlsMessages.ThrowException(XlsErr.ErrObjectNotFound, OriginalObjectPath);
                        return Opt.Parent as TEscherContainerRecord;
                    }
                    else
                        if (ObjectPath.StartsWith(FlxConsts.ObjectPathSpId, StringComparison.InvariantCulture))
                        {
                            TEscherOPTRecord Opt = RecordCache.OptByName.FindObjByShapeId(ObjectPath.Substring(1));
                            if (Opt == null) XlsMessages.ThrowException(XlsErr.ErrObjectNotFound, OriginalObjectPath);
                            return Opt.Parent as TEscherContainerRecord;
                        }

            }

            if (ObjIndex < 0 || ObjIndex >= RecordCache.AnchorList.Count)  XlsMessages.ThrowException(XlsErr.ErrObjectNotFound, OriginalObjectPath);
            TEscherContainerRecord Root = RecordCache.AnchorList.VisibleItem(ObjIndex).FindRoot() as TEscherContainerRecord;
            if (Root == null) return null;
            return Root.Navigate(ObjectPath) as TEscherContainerRecord;
        }

        private static void ReadAbsolutePath(ref int ObjIndex, ref string ObjectPath)
        {
            ObjectPath = ObjectPath.Substring(1);
            int p = ObjectPath.IndexOf(FlxConsts.ObjectPathSeparator);
            if (p >= 0)
            {
                string s = ObjectPath.Substring(0, p);
                if (!TCompactFramework.TryParse(s, out ObjIndex)) ObjIndex = -1;
                ObjIndex--;
                ObjectPath = ObjectPath.Substring(p + 1);
            }
            else
            {
                if (!TCompactFramework.TryParse(ObjectPath, out ObjIndex)) ObjIndex = -1;
                ObjIndex--;
                ObjectPath = null;
            }
        }

        /// <summary>
        /// Navigates to a child object.
        /// </summary>
        /// <param name="ObjectPath"></param>
        /// <returns></returns>
        internal TEscherRecord Navigate(string ObjectPath)
        {
            return Navigate(ObjectPath, ObjectPath);
        }

        private TEscherRecord Navigate(string ObjectPath, string OriginalPath)
        {
            if (ObjectPath==null || ObjectPath.Length==0) return this;
            int p = ObjectPath.IndexOf(FlxConsts.ObjectPathSeparator);
            string s = ObjectPath;
            string remaining= null;
            if (p>=0)
            {
                s = ObjectPath.Substring(0, p);
                remaining = ObjectPath.Substring(p+1);
            }
            int SearchId;
            if (!TCompactFramework.TryParse(s, out SearchId)) SearchId = -1;
            if (SearchId<0 || SearchId>= ContainedRecords.Count) XlsMessages.ThrowException(XlsErr.ErrObjectNotFound, OriginalPath);

            TEscherContainerRecord cr = ContainedRecords[SearchId] as TEscherContainerRecord;
            if (cr == null) XlsMessages.ThrowException(XlsErr.ErrObjectNotFound, OriginalPath);
            return cr.Navigate(remaining, OriginalPath);
        }
    }

    /// <summary>
    /// SP Record.
    /// </summary>
    internal class TEscherSpContainerRecord: TEscherContainerRecord
    {
        internal TEscherSpContainerRecord(TEscherRecordHeader aEscherHeader, TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
            : base(aEscherHeader, aDwgGroupCache, aDwgCache, aParent){}

        internal TEscherSpRecord SP;
        internal TEscherOPTRecord Opt;
        internal TEscherBaseClientAnchorRecord ClientAnchor;
        internal TEscherChildAnchorRecord ChildAnchor;

        internal int Row
        {
            get
            {
                if (ClientAnchor!=null) return ClientAnchor.Row1; else return 0;
            }
        }

        internal int Col
        {
            get
            {
                if (ClientAnchor!=null) return ClientAnchor.Col1; else return 0;
            }
        }

    }

    /// <summary>
    /// Spgr Record.
    /// </summary>
    internal class TEscherSpgrContainerRecord: TEscherContainerRecord
    {
        internal TEscherSpgrContainerRecord(TEscherRecordHeader aEscherHeader, TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
            : base(aEscherHeader, aDwgGroupCache, aDwgCache, aParent)
        {
            Init();
        }

        private void Init()
        {
            if (DwgCache.Patriarch==null) DwgCache.Patriarch=this;
        }

        internal override void Destroy()
        {
            base.Destroy();
            if (DwgCache.Patriarch==this) DwgCache.Patriarch=null;
        }

        protected override TEscherRecord DoCopyTo(int RowOfs, int ColOfs, TEscherDwgCache NewDwgCache, TEscherDwgGroupCache NewDwgGroupCache, TSheetInfo SheetInfo)
        {
            //This should recreate all actions on the constructor.
            TEscherSpgrContainerRecord Result= (TEscherSpgrContainerRecord)base.DoCopyTo(RowOfs, ColOfs, NewDwgCache, NewDwgGroupCache, SheetInfo);
            //Result.Init();
            if (DwgCache.Patriarch == this) Result.Init(); //We want to keep correpondent patriarchs.
            return Result;
        }

    }

    /// <summary>
    /// Client Data Record. It holds an OBJ inside
    /// </summary>
    internal class TEscherClientDataRecord: TEscherDataRecord
    {
        internal TBaseClientData ClientData;

        internal TEscherClientDataRecord(TEscherRecordHeader aEscherHeader, TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
            : base(aEscherHeader, aDwgGroupCache, aDwgCache, aParent)
        {
            Init();
        }

        private void Init()
        {
            if (DwgCache.Obj != null) DwgCache.Obj.Add(this);
        }

        /// <summary>
        /// CreateFromData
        /// </summary>
        /// <param name="aDwgGroupCache"></param>
        /// <param name="aDwgCache"></param>
        /// <param name="aParent"></param>
        internal TEscherClientDataRecord(TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
            : base(new TEscherRecordHeader(0,(int)Msofbt.ClientData,0), aDwgGroupCache, aDwgCache, aParent)
        {
            LoadedDataSize=0;
            Init();
        }

        internal override void Destroy()
        {
            base.Destroy();
            if (DwgCache.Obj != null) DwgCache.Obj.Remove(this);
            if (DwgCache.RadioButtons != null) DwgCache.RadioButtons.Remove(this);
        }

        protected override TEscherRecord DoCopyTo(int RowOfs, int ColOfs, TEscherDwgCache NewDwgCache, TEscherDwgGroupCache NewDwgGroupCache, TSheetInfo SheetInfo)
        {
            TEscherClientDataRecord Result = (TEscherClientDataRecord)base.DoCopyTo(RowOfs, ColOfs, NewDwgCache, NewDwgGroupCache, SheetInfo);
            Result.Init();
            Result.AssignClientData(TBaseClientData.Clone(ClientData, SheetInfo), true);

            /*if (NewDwgCache == DwgCache || NeedsNewObjId)*/ Result.ClientData.ArrangeId(ref DwgCache.MaxObjId);
            Result.ArrangeCopyRange(RowOfs, ColOfs, SheetInfo);

            return Result;
        }

        internal override void SaveToStream(IDataStream DataStream, TSaveData SaveData, TBreakList BreakList)
        {
            base.SaveToStream(DataStream, SaveData, BreakList);
            int StreamPos= (int)DataStream.Position;
            if (ClientData != null)
            {
                SaveData = CalcAnchorIfNeeded(SaveData);
                ClientData.SaveToStream(DataStream, SaveData);
            }
            BreakList.AddToZeroPos((int)(DataStream.Position-StreamPos));
        }

        private TSaveData CalcAnchorIfNeeded(TSaveData SaveData)
        {
            TMsObj ms = ClientData as TMsObj;
            if (ms != null && ms.ObjType == TObjectType.ListBox)
            {
                TClientAnchor Anchor = (Parent.FindRec<TEscherImageAnchorRecord>() as TEscherImageAnchorRecord).GetAnchor();
                SaveData.ObjectHeight = Anchor.CalcImageHeightInternal(new RowColSize(1, 1, SaveData.SavingSheet));
                SaveData.FixPage = true;

            }
            return SaveData;
        }

        internal override bool WaitingClientData(ref TClientType ClientType)
        {
            ClientType= TClientType.TMsObj; 
            return (base.Loaded()) && (ClientData==null);
        }

        internal override void SplitRecords(ref int NextPos, ref int RealSize, ref int NextDwg, TBreakList BreakList, int ContinueRecord, int ExtraDataLen)
        {
            base.SplitRecords(ref NextPos, ref RealSize, ref NextDwg, BreakList, ContinueRecord, ExtraDataLen);
            if (ClientData!=null) RealSize += (int)ClientData.TotalSize();
            NextDwg=(int)xlr.MSODRAWING;
        }

        internal override void AssignClientData(TBaseClientData aClientData, bool Copied)
        {
            ClientData = aClientData;
            if (ClientData!=null) 
            {
                if (ClientData.Id > DwgCache.MaxObjId) DwgCache.MaxObjId = ClientData.Id;

                if (DwgCache.RadioButtons != null && aClientData.ObjRecord() == TClientType.TMsObj)
                {
                    DwgCache.RadioButtons.Add(this, Copied);
                }
            }
        }

        internal void ArrangeCopyRange(int RowOfs, int ColOfs, TSheetInfo SheetInfo)
        {
            if (ClientData!=null) ClientData.ArrangeCopyRange(RowOfs, ColOfs, SheetInfo);
        }

        internal override void ArrangeInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo, bool Forced)
        {
            base.ArrangeInsertRange(CellRange, aRowCount, aColCount, SheetInfo, Forced);
            if (ClientData!=null) ClientData.ArrangeInsertRange(CellRange, aRowCount, aColCount, SheetInfo);
        }

        internal override void ArrangeMoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
        {
            base.ArrangeMoveRange (CellRange, NewRow, NewCol, SheetInfo);
            if (ClientData!=null) ClientData.ArrangeMoveRange(CellRange, NewRow, NewCol, SheetInfo);
        }


        internal override bool HasExternRefs()
        {
            if (ClientData==null) return true;
            else return ClientData.HasExternRefs();
        }

        internal int ObjId
        {
            get
            {
                if (ClientData!=null) return ClientData.Id; else return 0;
            }
        }

        internal bool IsFirstRadioButton
        {
            get
            {
                TMsObj ms = ClientData as TMsObj;
                if (ms == null) return false;
                return ms.GetRbFirstInGroup();
            }
            set
            {
                TMsObj ms = ClientData as TMsObj;
                if (ms == null) return;
                ms.SetRbFirstInGroup(value);
            }
        }

        internal int NextRbId
        {
            get
            {
                TMsObj ms = ClientData as TMsObj;
                if (ms == null) return 0;
                return ms.GetRbNextId();
            }
            set
            {
                TMsObj ms = ClientData as TMsObj;
                if (ms == null) return;
                ms.SetRbNextId(value);
            }
        }

    }

    /// <summary>
    /// Split Menu Record.
    /// </summary>
    internal class TEscherSplitMenuRecord: TEscherDataRecord
    {
        internal TEscherSplitMenuRecord(TEscherRecordHeader aEscherHeader, TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
            : base(aEscherHeader, aDwgGroupCache, aDwgCache, aParent)
        {
        }

        /// <summary>
        /// Create from data
        /// </summary>
        internal TEscherSplitMenuRecord(TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
            : base(new TEscherRecordHeader(0x0040,(int)Msofbt.SplitMenuColors,16), aDwgGroupCache, aDwgCache, aParent)
        {
            byte[] aData={0x0D, 0x00, 0x00, 0x08, 0x0C, 0x00, 0x00, 0x08, 0x17, 0x00, 0x00, 0x08, 0xF7, 0x00, 0x00, 0x10};
            Data=aData;
            LoadedDataSize=TotalDataSize;
        }
    }

    /// <summary>
    /// Spgr Record. Has the relative bounds on a grouped shape
    /// </summary>
    internal class TEscherSpgrRecord: TEscherDataRecord
    {
        internal TEscherSpgrRecord(TEscherRecordHeader aEscherHeader, TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
            : base(aEscherHeader, aDwgGroupCache, aDwgCache, aParent)
        {
        }

        internal TEscherSpgrRecord(TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent, TDrawingPoint p1, int Height, int Width)
            : base(new TEscherRecordHeader(0x01, (int)Msofbt.Spgr, 0x10), aDwgGroupCache, aDwgCache, aParent)
        {
            BitOps.SetCardinal(Data, 0, p1.X.Emu);
            BitOps.SetCardinal(Data, 4, p1.Y.Emu);
            BitOps.SetCardinal(Data, 8, p1.X.Emu + Width);
            BitOps.SetCardinal(Data, 12, p1.Y.Emu + Height);
            LoadedDataSize = TotalDataSize;
        }

        internal int[] Bounds
        {
            get
            {
                unchecked
                {
                    return new int[]
                   {
                       (int)BitOps.GetCardinal(Data,0),
                       (int)BitOps.GetCardinal(Data,4),
                       (int)BitOps.GetCardinal(Data,8),
                       (int)BitOps.GetCardinal(Data,12)
                   };
                }
            }
        }
    }

    internal abstract class TEscherBaseClientAnchorRecord: TEscherRecord
    {
        internal int Row1;
        internal int Col1;

        internal TEscherBaseClientAnchorRecord(TEscherRecordHeader aEscherHeader, TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
            : base(aEscherHeader, aDwgGroupCache, aDwgCache, aParent)
        {
            Init();
        }
    
        protected void Init()
        {
            if (DwgCache.AnchorList!=null) DwgCache.AnchorList.Add(this);
            if (FParent != null) ((TEscherSpContainerRecord)FParent ).ClientAnchor=this;
        }
    
        internal override void Destroy()
        {
            base.Destroy();
            if (DwgCache.AnchorList!= null) DwgCache.AnchorList.Remove(this);
            if (FParent != null) ((TEscherSpContainerRecord)FParent).ClientAnchor=null;
        }

        internal virtual bool AllowCopy(bool IncludeDontMoveAndResize, int FirstRow, int LastRow, int FirstCol, int LastCol, out bool IsInRange)
        {
            IsInRange = false;
            return false;
        }

        internal virtual bool AllowDelete(int FirstRow, int LastRow, int FirstCol, int LastCol)
        {
            return false;
        }

        internal virtual void SaveObjectCoords(TSheet sSheet)
        {
        }

        internal virtual void RestoreObjectCoords(TSheet dSheet)
        {
        }

        #region Biff8 Read/save
        internal abstract byte[] GetBiff8Data(TSheet aSheet);

        private static int GetUInt16(byte[] Data, int Pos, int Max)
        {
            int Result = Data[Pos] + (Data[Pos + 1] << 8);
            if (Result > Max) return Max;
            return Result;
        }

        private static UInt32 GetUInt32(byte[] Data, int Pos, UInt32 Max)
        {
            long Result = Data[Pos] + (Data[Pos + 1] << 8) + (Data[Pos + 2] << 16 + (Data[Pos + 3] << 24));
            if (Result > Max) return Max;
            return (UInt32)Result;
        }

        protected void SetVar(ref int v, byte[] Data, int dPos, int MaxValue, int RSize)
        {
            if (dPos < LoadedDataSize) return; //already loaded.
            if (dPos - LoadedDataSize + 2 > RSize) return; //not yet loaded.
            v = GetUInt16(Data, dPos - LoadedDataSize, MaxValue);
        }

        protected void SetVar32(ref UInt32 v, byte[] Data, int dPos, UInt32 MaxValue, int RSize)
        {
            if (dPos < LoadedDataSize) return; //already loaded.
            if (dPos - LoadedDataSize + 2 > RSize) return; //not yet loaded.
            v = GetUInt32(Data, dPos - LoadedDataSize, MaxValue);
        }

        internal override void SaveToStream(IDataStream DataStream, TSaveData SaveData, TBreakList BreakList)
        {
            base.SaveToStream(DataStream, SaveData, BreakList);
            byte[] Data = GetBiff8Data(SaveData.SavingSheet);
            if (TotalDataSize > 0)
            {
                int RemainingSize = TotalDataSize;
                while (RemainingSize > BreakList.AcumSize() - DataStream.Position)
                {
                    int FracSize = (int)(BreakList.AcumSize() - DataStream.Position);
                    CheckSplit(DataStream, SaveData, BreakList);
                    DataStream.Write(Data, TotalDataSize - RemainingSize, FracSize);

                    RemainingSize -= FracSize;
                } //while

                CheckSplit(DataStream, SaveData, BreakList);
                DataStream.Write(Data, TotalDataSize - RemainingSize, RemainingSize);
            }
        }

        internal override long TotalSizeNoSplit()
        {
            return base.TotalSizeNoSplit() + TotalDataSize;
        }

        internal override void SplitRecords(ref int NextPos, ref int RealSize, ref int NextDwg, TBreakList BreakList, int ContinueRecord, int ExtraDataLen)
        {
            base.SplitRecords(ref NextPos, ref RealSize, ref NextDwg, BreakList, ContinueRecord, ExtraDataLen);
            IncNextPos(ref NextPos, TotalDataSize, ref RealSize, BreakList, ContinueRecord, ExtraDataLen);
        }
        #endregion

    }

    internal class TEscherHeaderAnchorRecord: TEscherBaseClientAnchorRecord
    {
        UInt32 Width;
        UInt32 Height;
        internal TEscherHeaderAnchorRecord(TEscherRecordHeader aEscherHeader, TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
            : base(aEscherHeader, aDwgGroupCache, aDwgCache, aParent)
        {
        }

        /// <summary>
        /// Create from Data
        /// </summary>
        internal TEscherHeaderAnchorRecord(THeaderOrFooterAnchor aAnchor, TEscherRecordHeader aEscherHeader, TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
            : base(aEscherHeader, aDwgGroupCache, aDwgCache, aParent)
        {
            Width = (UInt32)aAnchor.Width;
            Height = (UInt32)aAnchor.Height;
            LoadedDataSize=TotalDataSize;
        }

        internal THeaderOrFooterAnchor GetHeaderAnchor()
        {
            return new THeaderOrFooterAnchor(Width, Height); 
        }

        protected override TEscherRecord DoCopyTo(int RowOfs, int ColOfs, TEscherDwgCache NewDwgCache, TEscherDwgGroupCache NewDwgGroupCache, TSheetInfo SheetInfo)
        {
            TEscherHeaderAnchorRecord Result= (TEscherHeaderAnchorRecord) base.DoCopyTo(RowOfs, ColOfs, NewDwgCache, NewDwgGroupCache, SheetInfo);
            Result.Init();
            ((TEscherSpContainerRecord)Result.FParent).ClientAnchor=Result;
            return Result;
        }

        internal void SetAnchor(THeaderOrFooterAnchor aAnchor)
        {
            Width = (UInt32)aAnchor.Width;
            Height = (UInt32)aAnchor.Height;
        }

        #region Biff8

        internal override byte[] GetBiff8Data(TSheet aSheet)
        {
            byte[] Result = new byte[8];
            BitConverter.GetBytes(Width).CopyTo(Result, 0);
            BitConverter.GetBytes(Height).CopyTo(Result, 4);
            return Result;
        }

        internal override int CompareRec(TEscherRecord aRecord)
        {
            TEscherHeaderAnchorRecord a = aRecord as TEscherHeaderAnchorRecord;
            if (a == null) return -1;

            int Result = base.CompareRec(aRecord); if (Result != 0) return Result;
            Result = Width.CompareTo(a.Width); if (Result != 0) return Result;
            Result = Height.CompareTo(a.Height); if (Result != 0) return Result;
            return 0;
        }

        internal override void Load(ref TxBaseRecord aRecord, ref int aPos, bool HeaderImage, bool aChartCoords)
        {
            if (TotalDataSize == 0) return;
            int RSize = aRecord.TotalSizeNoHeaders() - aPos;
            if (RSize <= 0) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
            if (TotalDataSize - LoadedDataSize < RSize) RSize = TotalDataSize - LoadedDataSize;
            if (LoadedDataSize + RSize > TotalDataSize) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);

            byte[] Data = new byte[RSize];
            BitOps.ReadMem(ref aRecord, ref aPos, Data, LoadedDataSize, RSize);

            SetVar32(ref Width, Data, 0, UInt32.MaxValue, RSize);
            SetVar32(ref Height, Data, 4, UInt32.MaxValue, RSize);

            LoadedDataSize += RSize;
        }
        #endregion
    }

    /// <summary>
    /// Client Anchor Record for an image.
    /// </summary>
    internal class TEscherImageAnchorRecord: TEscherBaseClientAnchorRecord
    {
        #region Variables
        int Flag;
        int Row2, Col2, Dx1, Dy1, Dx2, Dy2;
        bool ChartCoords;
        TAbsoluteAnchorRect SaveRect;
        #endregion

        #region Constructor
        internal TEscherImageAnchorRecord(TEscherRecordHeader aEscherHeader, TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
            : base(aEscherHeader, aDwgGroupCache, aDwgCache, aParent)
        {
        }

        /// <summary>
        /// Create from Data
        /// </summary>
        internal static TEscherImageAnchorRecord CreateFromData(TClientAnchor aAnchor, TEscherRecordHeader aEscherHeader, TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent, TSheet sSheet)
        {
            TEscherImageAnchorRecord Result = new TEscherImageAnchorRecord(aEscherHeader, aDwgGroupCache, aDwgCache, aParent);
            Result.SetAnchor(aAnchor, null);
            
            Result.LoadedDataSize= Result.TotalDataSize;
            Result.SaveObjectCoords(sSheet);
            return Result;
        }

        protected override TEscherRecord DoCopyTo(int RowOfs, int ColOfs, TEscherDwgCache NewDwgCache, TEscherDwgGroupCache NewDwgGroupCache, TSheetInfo SheetInfo)
        {

            TEscherImageAnchorRecord Result= (TEscherImageAnchorRecord) base.DoCopyTo(RowOfs, ColOfs, NewDwgCache, NewDwgGroupCache, SheetInfo);
            Result.Init();
            ((TEscherSpContainerRecord)Result.FParent).ClientAnchor=Result;
            if (!ChartCoords)
            {
                Result.Row1 += RowOfs;
                Result.Col1 += ColOfs;
                Result.Row2 += RowOfs;
                Result.Col2 += ColOfs;
            }

            if (SaveRect != null) Result.SaveRect = (TAbsoluteAnchorRect) SaveRect.Clone();

            Result.RestoreObjectCoords(SheetInfo.DestSheet);

            return Result;
        }
        #endregion

        #region Biff8
        internal override byte[] GetBiff8Data(TSheet aSheet)
        {
            byte[] Result = new byte[18];
            BitConverter.GetBytes((UInt16)Flag).CopyTo(Result, 0);

            int Row1Bis = Row1;
            int Row2Bis = Row2;
            int Col1Bis = Col1;
            int Col2Bis = Col2;
            int Dx1Bis = Dx1;
            int Dy1Bis = Dy1;
            int Dx2Bis = Dx2;
            int Dy2Bis = Dy2;

            if (ChartCoords)
            {
                Biff8Utils.CheckChart(ref Row1Bis);
                Biff8Utils.CheckChart(ref Col1Bis);
                Biff8Utils.CheckChart(ref Row2Bis);
                Biff8Utils.CheckChart(ref Col2Bis);
            }
            else
            {
                Biff8Utils.CheckRow(Row1Bis); //This gets checked before, so we don't move all imags below to the last row.

                //Comments by default are saved in Col+1. So a file with comment in IV won't save if we don't decrease this.
                int ColForCheck = Col1Bis;
                if (ColForCheck == FlxConsts.Max_Columns97_2003 + 1) { ColForCheck --; } 
                Biff8Utils.CheckCol(ColForCheck);

                //Instead of raising an Exception, we will just decrease Row2 and Col2 if they don't fit.
                MoveImageToBottomRight(aSheet, ref Row1Bis, ref Row2Bis, ref Col1Bis, ref Col2Bis, ref Dx1Bis, ref Dy1Bis, ref Dx2Bis, ref Dy2Bis);
                
                Biff8Utils.CheckRow(Row2Bis);
                Biff8Utils.CheckCol(Col2Bis);

            }

            BitConverter.GetBytes((UInt16)Col1Bis).CopyTo(Result, 2);
            BitConverter.GetBytes((UInt16)Dx1Bis).CopyTo(Result, 4);
            BitConverter.GetBytes((UInt16)Row1Bis).CopyTo(Result, 6);
            BitConverter.GetBytes((UInt16)Dy1Bis).CopyTo(Result, 8);

            BitConverter.GetBytes((UInt16)Col2Bis).CopyTo(Result, 10);
            BitConverter.GetBytes((UInt16)Dx2Bis).CopyTo(Result, 12);
            BitConverter.GetBytes((UInt16)Row2Bis).CopyTo(Result, 14);
            BitConverter.GetBytes((UInt16)Dy2Bis).CopyTo(Result, 16);
            return Result;        
        }

        private void MoveImageToBottomRight(TSheet aSheet, ref int Row1Bis, ref int Row2Bis, ref int Col1Bis, ref int Col2Bis, ref int Dx1Bis, ref int Dy1Bis, ref int Dx2Bis, ref int Dy2Bis)
        {
            if (Row2Bis > FlxConsts.Max_Rows97_2003)
            {
                if (aSheet != null)
                {
                    long y1 = 0, y2 = 0;
                    CalcAbsRow(aSheet, Row1, Row1, Dy1, out y1);
                    CalcAbsRow(aSheet, Row1, Row2, Dy2, out y2);

                    long Dummy;
                    CalcRowAndDy(aSheet, FlxConsts.Max_Rows97_2003 + 1, false, y1 - y2, out Row1Bis, out Dy1Bis, out Dummy);
                }

                Row2Bis = FlxConsts.Max_Rows97_2003;
                Dy2Bis = 255;
            }
            if (Col2Bis > FlxConsts.Max_Columns97_2003)
            {
                if (aSheet != null)
                {
                    long x1 = 0, x2 = 0;
                    CalcAbsCol(aSheet, Col1, Col1, Dx1, out x1);
                    CalcAbsCol(aSheet, Col1, Col2, Dx2, out x2);

                    long Dummy;
                    CalcColAndDx(aSheet, FlxConsts.Max_Columns97_2003 + 1, false, x1 - x2, out Col1Bis, out Dx1Bis, out Dummy);
                }
                Col2Bis = FlxConsts.Max_Columns97_2003;
                Dx2Bis = 1024;
            }
        }


        private TFlxAnchorType AnchorType
        {
            get
            {
                switch (Flag & 3)
                {
                    case 1: //flag = 1 is not documented, but happens on AutoFilters. AutoFilters should move.
                    case 0: return TFlxAnchorType.MoveAndResize;
                    case 3: return TFlxAnchorType.DontMoveAndDontResize;
                }
                return TFlxAnchorType.MoveAndDontResize;
            }
        }

        #endregion

        #region Convert to absolute
        private void CalcAbsCol(TSheet Workbook, int StartCol, int Col, int Deltax, out long x)
        {
            Debug.Assert(!ChartCoords);
            x =0;
            for (int c=StartCol; c<Col; c++)            
                x+=Workbook.GetColWidth(c, true) * 1024;

            for (int c=StartCol; c>Col; c--)  //negative deltas
                x-=Workbook.GetColWidth(c - 1, true) * 1024;


            x+=Workbook.GetColWidth(Col, true)*Deltax;
        }

        private void CalcAbsRow(TSheet Workbook, int StartRow, int Row, int Deltay, out long y)
        {
            Debug.Assert(!ChartCoords);
            y = 0;
            for (int r=StartRow; r<Row; r++)            
                y+=Workbook.GetRowHeight(r, true) * 255;

            for (int r=StartRow; r>Row; r--)  //negative deltas
                y-=Workbook.GetRowHeight(r - 1, true) * 255;

            y+=Workbook.GetRowHeight(Row, true)*Deltay;
        }

        private void CalcColAndDx(TSheet Workbook, int StartCol, bool FixedCol, long RectX, out int Column, out int Deltax, out long Acumx)
        {
            Debug.Assert(!ChartCoords);
            int Col = StartCol;
            long x=0;
            long Lastx =0;

            if (!FixedCol)
            {
                if (RectX < 0)
                {
                    while (Col>0 && x> RectX)
                    {
                        x-=Workbook.GetColWidth(Col - 1, true) * 1024;
                        Lastx = x;
                        Col--;
                    }
                    Column=Col;
                }
                else
                {
                    while (Col<=FlxConsts.Max_Columns && x<= RectX)
                    {
                        Lastx = x;
                        x+=Workbook.GetColWidth(Col, true) * 1024;
                        Col++;
                    }
                    Column=Col-1;
                }
            }
            else
            {
                Column = StartCol;
            }

            Acumx = Lastx;

            if (Column<0)
            {
                Column=0;
                Deltax=0;
            }
            else
            {
                float fw = (float)Workbook.GetColWidth(Column, true);
                if (Workbook.GetColWidth(Column, true)>0)
                    Deltax = (int) Math.Round((float)(RectX-Lastx) / fw);
                else Deltax=0;

                if (Deltax>1024) Deltax=1024;
                if (Deltax<0) Deltax= 0; //only happens when col = 0;
            }
        }

        private void CalcRowAndDy(TSheet Workbook, int StartRow, bool FixedRow, long RectY, out int RowFinal, out int Deltay, out long Acumy)
        {
            Debug.Assert(!ChartCoords);
            int Row = StartRow;
            long y=0;
            long Lasty =0;

            if (!FixedRow)
            {
                if (RectY < 0)
                {
                    while (Row> 0 && y> RectY)
                    {
                        y-=Workbook.GetRowHeight(Row - 1, true)* 255;
                        Lasty = y;
                        Row--;
                    }
                    RowFinal=Row;
                }
                else
                {
                    while (Row<=FlxConsts.Max_Rows && y<= RectY)
                    {
                        Lasty = y;
                        y+=Workbook.GetRowHeight(Row, true) * 255;
                        Row++;
                    }
                    RowFinal=Row-1;
                }
            }
            else
            {
                RowFinal = Row;
            }
            Acumy = Lasty;


            if (RowFinal<0)
            {
                RowFinal=0;
                Deltay=0;
            }
            else
            {
                float fw = (float)Workbook.GetRowHeight(RowFinal, true);
                if (Workbook.GetRowHeight(RowFinal, true)>0)
                    Deltay = (int) Math.Round((float)(RectY-Lasty) / fw);
                else Deltay=0;

                if (Deltay>255) Deltay=255;
                if (Deltay<0) Deltay= 0; //only happens when row = 0;
            }
        }

        #endregion

        #region Save and restore coords
        internal override void SaveObjectCoords(TSheet sSheet)
        {
            if (ChartCoords) return;
            if (sSheet == null) return;
            base.SaveObjectCoords (sSheet);

            switch (AnchorType)
            {
                case TFlxAnchorType.MoveAndDontResize:
                {
                    long x1=0, x2=0, y1=0, y2=0;
                    CalcAbsCol(sSheet, Col1, Col1, Dx1, out x1); //We do not need to calculate the full coordinates here. We only need to save the size.
                    CalcAbsRow(sSheet, Row1, Row1, Dy1, out y1);
                    CalcAbsCol(sSheet, Col1, Col2, Dx2, out x2);
                    CalcAbsRow(sSheet, Row1, Row2, Dy2, out y2);

            
                    SaveRect = new TAbsoluteAnchorRect(AnchorType, x1, y1, x2, y2);
                    break;
                }
                case TFlxAnchorType.DontMoveAndDontResize:
                {
                    long x1=0, x2=0, y1=0, y2=0;
                    CalcAbsCol(sSheet, 0, Col1, Dx1, out x1); 
                    CalcAbsRow(sSheet, 0, Row1, Dy1, out y1);

                    long tmpx1, tmpy1;
                    CalcAbsCol(sSheet, Col1, Col1, Dx1, out tmpx1); //Calculate how much of x is on the starting row. This is inexpensive.
                    CalcAbsRow(sSheet, Row1, Row1, Dy1, out tmpy1);

                    CalcAbsCol(sSheet, Col1, Col2, Dx2, out x2);
                    x2 += x1 - tmpx1;
                    CalcAbsRow(sSheet, Row1, Row2, Dy2, out y2);
                    y2 += y1 - tmpy1;

            
                    SaveRect = new TAbsoluteAnchorRect(AnchorType, x1, y1, x2, y2);
                    break;
                }
            }
        }

        internal override void RestoreObjectCoords(TSheet dSheet)
        {
            if (ChartCoords) return;
            if (dSheet == null) return;
            switch (AnchorType)
            {
                case TFlxAnchorType.MoveAndDontResize:
                {
                    int Row, Col, Dx, Dy; long Dummy;
                    long w1= SaveRect.x2 - SaveRect.x1;
                    long h1 = SaveRect.y2 - SaveRect.y1;

                    if (dSheet.GetColWidth(Col1, true) != 0) //If width is 0, do not change anything. So the column will be restored and the image will keep its position.
                    {
                        CalcColAndDx(dSheet, Col1, true, SaveRect.x1, out Col, out Dx, out Dummy); Col1 = Col; Dx1 = Dx;  //If the column was made wider, recalc the new dx. Do NOT change column width.
                    }

                    if (dSheet.GetRowHeight(Row1, true) != 0)
                    {
                        CalcRowAndDy(dSheet, Row1, true, SaveRect.y1, out Row, out Dy, out Dummy); Row1 = Row; Dy1 = Dy;
                    }

                    long NewX1, NewY1;
                    CalcAbsCol(dSheet, Col1, Col1, Dx1, out NewX1);
                    CalcAbsRow(dSheet, Row1, Row1, Dy1, out NewY1);

                    CalcColAndDx(dSheet, Col1, false, w1 + NewX1, out Col, out Dx, out Dummy); Col2 = Col; Dx2 = Dx;
                    CalcRowAndDy(dSheet, Row1, false, h1 + NewY1, out Row, out Dy, out Dummy); Row2 = Row; Dy2 = Dy;

                    SaveObjectCoords(dSheet); //Save the new x1, y1. They might have been modified.
                    break;
                }
                case TFlxAnchorType.DontMoveAndDontResize:
                {
                    int Row, Col, Dx, Dy; long Acumx, Acumy, Dummy;
                    CalcColAndDx(dSheet, 0, false, SaveRect.x1, out Col, out Dx, out Acumx); Col1 = Col; Dx1 = Dx;
                    CalcColAndDx(dSheet, Col, false, SaveRect.x2 - Acumx, out Col, out Dx, out Dummy); Col2 = Col; Dx2 = Dx;
                    CalcRowAndDy(dSheet, 0, false, SaveRect.y1, out Row, out Dy, out Acumy); Row1 = Row; Dy1 = Dy;
                    CalcRowAndDy(dSheet, Row, false, SaveRect.y2 - Acumy, out Row, out Dy, out Dummy); Row2 = Row; Dy2 = Dy;
                    break;
                }
            }
        }

        internal TAbsoluteAnchorRect SaveCommentCoords(TSheet sSheet, int aRow, int aCol)
        {
            if (ChartCoords)
            {
                return new TAbsoluteAnchorRect(AnchorType, Col1, Row1, Col2, Row2);
            }

            long x1=0, x2=0, y1=0, y2=0;
            CalcAbsCol(sSheet, aCol + 1, Col1, Dx1, out x1); //We do not need to calculate the full coordinates here. We only need to save the size.
            CalcAbsRow(sSheet, aRow, Row1, Dy1, out y1);
            CalcAbsCol(sSheet, aCol + 1, Col2, Dx2, out x2);
            CalcAbsRow(sSheet, aRow, Row2, Dy2, out y2);

            
            return new TAbsoluteAnchorRect(AnchorType, x1, y1, x2, y2);
        }

        internal void RestoreCommentCoords(TAbsoluteAnchorRect aSaveRect, TSheet dSheet, int aRow, int aCol)
        {
            if (ChartCoords) return;
            int Row, Col, Dx, Dy; long Dummy;
            CalcColAndDx(dSheet, aCol + 1, false, aSaveRect.x1, out Col, out Dx, out Dummy); Col1 = Col; Dx1 = Dx;
            CalcRowAndDy(dSheet, aRow, false, aSaveRect.y1, out Row, out Dy, out Dummy); Row1 = Row; Dy1 = Dy;
            CalcColAndDx(dSheet, aCol + 1, false, aSaveRect.x2, out Col, out Dx, out Dummy); Col2 = Col; Dx2 = Dx;
            CalcRowAndDy(dSheet, aRow, false, aSaveRect.y2, out Row, out Dy, out Dummy); Row2 = Row; Dy2 = Dy;
        }
        #endregion

        #region Get / Set Anchor
        internal TClientAnchor GetAnchor()
        {
            return new TClientAnchor(ChartCoords, (TFlxAnchorType)Flag, Row1, Dy1, Col1, Dx1, Row2, Dy2, Col2, Dx2);
        }

        internal void SetAnchor(TClientAnchor aAnchor, TSheet sSheet)
        {
            Flag=(int)aAnchor.AnchorType;
            Row1=aAnchor.Row1;
            Dy1=aAnchor.Dy1;
            Row2=aAnchor.Row2;
            Dy2=aAnchor.Dy2;

            Col1=aAnchor.Col1;
            Dx1=aAnchor.Dx1;
            Col2=aAnchor.Col2;
            Dx2=aAnchor.Dx2;
            ChartCoords = aAnchor.ChartCoords;
            if (sSheet != null) SaveObjectCoords(sSheet);
        }
        #endregion

        #region Insert And move
        internal override void ArrangeInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo, bool Forced)
        {
            if (ChartCoords) return;
            if (Row2 < CellRange.Top || Col2 < CellRange.Left) return; //Image is not affected by the insert. If range is full, then image is above/to the left. If range is not full, rows will not move.

            base.ArrangeInsertRange(CellRange, aRowCount, aColCount, SheetInfo, Forced);
            if (SheetInfo.SourceFormulaSheet != SheetInfo.InsSheet) return;
            int Af = Flag;
            if (Forced) Af = 2;
            switch (Af & 3)
            {
                case 0: //move and resize
                    //Rows
                    if ((Row1 >= CellRange.Top) && (Col1 >= CellRange.Left) && (Col1 <= CellRange.Right))
                    {
                        int dr = Row1 + aRowCount * CellRange.RowCount;
                        Row1 = BitOps.GetIncMaxMin(Row1, dr - Row1, FlxConsts.Max_Rows, CellRange.Top);
                        if (dr < CellRange.Top) Dy1 = 0; //We deleted rows
                    }
                    if ((Col1 >= CellRange.Left) && (Col1 <= CellRange.Right)) //We always use col1 to know if to move the image or not.
                    {
                        int dr = Row2 + aRowCount * CellRange.RowCount;
                        Row2 = BitOps.GetIncMaxMin(Row2, dr - Row2, FlxConsts.Max_Rows, CellRange.Top);
                        if (dr < CellRange.Top) Dy2 = 0;
                    }

                    //Columns
                    if ((Col1 >= CellRange.Left) && (Row1 >= CellRange.Top) && (Row1 <= CellRange.Bottom))
                    {
                        int dc = Col1 + aColCount * CellRange.ColCount;
                        Col1 = BitOps.GetIncMaxMin(Col1, dc - Col1, FlxConsts.Max_Columns, CellRange.Left);
                        if (dc < CellRange.Left) Dx1 = 0; //We deleted columns
                    }
                    if ((Row1 >= CellRange.Top) && (Row1 <= CellRange.Bottom)) //We always use row1 to know if to move the image or not.
                    {
                        int dc = Col2 + aColCount * CellRange.ColCount;
                        Col2 = BitOps.GetIncMaxMin(Col2, dc - Col2, FlxConsts.Max_Columns, CellRange.Left);
                        if (dc < CellRange.Left) Dx2 = 0;
                    }
                    break;
                case 1: // not documented, but happens on AutoFilters. AutoFilters should move.
                case 2: //move  
                    if (((Row1 >= CellRange.Top) && (Col1 >= CellRange.Left) && (Col1 <= CellRange.Right)) || Forced)
                    {
                        int dr = Row1;
                        Row1 = BitOps.GetIncMaxMin(Row1, aRowCount * CellRange.RowCount, FlxConsts.Max_Rows, CellRange.Top);
                        Row2 = BitOps.GetIncMaxMin(Row2, Row1 - dr, FlxConsts.Max_Rows, Row1);
                    }
                    if (((Col1 >= CellRange.Left) && (Row1 >= CellRange.Top) && (Row1 <= CellRange.Bottom)) || Forced)
                    {
                        int dc = Col1;
                        Col1 = BitOps.GetIncMaxMin(Col1, aColCount * CellRange.ColCount, FlxConsts.Max_Columns, CellRange.Left);
                        Col2 = BitOps.GetIncMaxMin(Col2, Col1 - dc, FlxConsts.Max_Columns, Col1);
                    }

                    bool InsertingRows = aRowCount != 0;
                    bool IsFullRange = (CellRange.Left <= 0 && CellRange.Right >= FlxConsts.Max_Columns) || (CellRange.Top <= 0 && CellRange.Bottom >= FlxConsts.Max_Rows);
                    bool ImageInsideInsertedRange = (Row1 >= CellRange.Top && Row2 <= CellRange.Bottom && Col1 >= CellRange.Left && Col2 <= CellRange.Right);
                    bool ImageBelowRange = Row1 > CellRange.Bottom && Row2 < FlxConsts.Max_Rows; 
                    bool ImageRightToRange = Col1 > CellRange.Right && Col2 < FlxConsts.Max_Columns;
                    if (
                        (IsFullRange && (ImageInsideInsertedRange || ImageRightToRange || ImageBelowRange))
                        || (!IsFullRange && (InsertingRows && ImageRightToRange) || (!InsertingRows && ImageBelowRange))
                        )
                    { }//no need here.
                    else RestoreObjectCoords(SheetInfo.DestSheet);
                    break;

                case 3: //dont move
                    RestoreObjectCoords(SheetInfo.DestSheet);
                    break;
            } //case
        }

        internal void ArrangeMoveRange(int RowOfs, int ColOfs, TSheetInfo SheetInfo)
        {
            if (ChartCoords) return;
            if (SheetInfo.SourceFormulaSheet != SheetInfo.InsSheet) return;
            switch (AnchorType)
            {
                case TFlxAnchorType.MoveAndResize: 
                case TFlxAnchorType.MoveAndDontResize:
                    Row1=BitOps.GetIncMaxMin(Row1, RowOfs, FlxConsts.Max_Rows, 0);
                    Row2=BitOps.GetIncMaxMin(Row2, RowOfs, FlxConsts.Max_Rows, 0);
                    Col1=BitOps.GetIncMaxMin(Col1, ColOfs, FlxConsts.Max_Columns, 0);
                    Col2=BitOps.GetIncMaxMin(Col2, ColOfs, FlxConsts.Max_Columns, 0);
                    break;

                case TFlxAnchorType.DontMoveAndDontResize:
                    break;
            } //case
        }
        #endregion

        #region LoadFrom /SaveTo stream
        internal override int CompareRec(TEscherRecord aRecord)
        {
            TEscherImageAnchorRecord a = aRecord as TEscherImageAnchorRecord;
            if (a == null) return -1;

            int Result = base.CompareRec(aRecord); if (Result != 0) return Result;
            Result = Flag.CompareTo(a.Flag); if (Result != 0) return Result;
            Result = Row1.CompareTo(a.Row1); if (Result != 0) return Result;
            Result = Col1.CompareTo(a.Col1); if (Result != 0) return Result;
            Result = Row2.CompareTo(a.Row2); if (Result != 0) return Result;
            Result = Col2.CompareTo(a.Col2); if (Result != 0) return Result;
            Result = Dx1.CompareTo(a.Dx1); if (Result != 0) return Result;
            Result = Dy1.CompareTo(a.Dy1); if (Result != 0) return Result;
            Result = Dx2.CompareTo(a.Dx2); if (Result != 0) return Result;
            Result = Dy2.CompareTo(a.Dy2); if (Result != 0) return Result;
            Result = ChartCoords.CompareTo(a.ChartCoords); if (Result != 0) return Result;
            return 0;
        }

        internal override void Load(ref TxBaseRecord aRecord, ref int aPos, bool HeaderImage, bool aChartCoords)
        {
            if (TotalDataSize == 0) return;
            int RSize = aRecord.TotalSizeNoHeaders() - aPos;
            if (RSize <= 0) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
            if (TotalDataSize - LoadedDataSize < RSize) RSize = TotalDataSize - LoadedDataSize;
            if (LoadedDataSize + RSize > TotalDataSize) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);

            byte[] Data = new byte[RSize];
            BitOps.ReadMem(ref aRecord, ref aPos, Data, LoadedDataSize, RSize);

            SetVar(ref Flag, Data, 0, UInt16.MaxValue, RSize);
            SetVar(ref Col1, Data, 2, UInt16.MaxValue, RSize);
            SetVar(ref Dx1, Data, 4, 1024, RSize);
            SetVar(ref Row1, Data, 6, UInt16.MaxValue, RSize);
            SetVar(ref Dy1, Data, 8, 255, RSize);

            SetVar(ref Col2, Data, 10, UInt16.MaxValue, RSize);
            SetVar(ref Dx2, Data, 12, 1024, RSize);
            SetVar(ref Row2, Data, 14, UInt16.MaxValue, RSize);
            SetVar(ref Dy2, Data, 16, 255, RSize);

            ChartCoords = aChartCoords;

            LoadedDataSize += RSize;
        }
        #endregion


        internal override bool AllowCopy(bool IncludeDontMoveAndResize, int FirstRow, int LastRow, int FirstCol, int LastCol, out bool IsInRange)
        {

            if (ChartCoords) { IsInRange = true; return false; }
            IsInRange = (Row1 >= FirstRow) && (Row2 <= LastRow)
                && (Col1 >= FirstCol) && (Col2 <= LastCol);

            if (IncludeDontMoveAndResize) //see not to copy comments
            {
                TEscherClientDataRecord obj = (TEscherClientDataRecord)Parent.FindRec<TEscherClientDataRecord>();
                if (obj != null)
                {
                    TMsObj msobj = obj.ClientData as TMsObj;
                    if (msobj != null)
                    {
                        if (msobj.ObjType == TObjectType.Comment) return false; //commments shouldn't be copied.
                    }
                        
                }
            }
                

            return ( ((Flag & 3)==0) ||((Flag & 3)==2) || (IncludeDontMoveAndResize && ((Flag & 3) == 3))) 
                && IsInRange;
        }

        internal override bool AllowDelete(int FirstRow, int LastRow, int FirstCol, int LastCol)
        {
            if (ChartCoords) return false;
            return (AnchorType != TFlxAnchorType.DontMoveAndDontResize)
                && (Row1>=FirstRow) && (Row2<=LastRow)
                && (Col1>=FirstCol) && (Col2<=LastCol);
        }
    }

    internal class TEscherChildAnchorRecord: TEscherDataRecord
    {

        internal TEscherChildAnchorRecord(TEscherRecordHeader aEscherHeader, TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
            : base(aEscherHeader, aDwgGroupCache, aDwgCache, aParent)
        {
            Init();
        }

        internal static TEscherChildAnchorRecord CreateFromData(TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, 
            TEscherContainerRecord aParent, TDrawingPoint p1, int Height, int Width)
        {
            TEscherChildAnchorRecord Result = new TEscherChildAnchorRecord(new TEscherRecordHeader(0x00, (int)Msofbt.ChildAnchor, 0x10), aDwgGroupCache, aDwgCache, aParent);
            Result.Dx1 = (int)p1.X.Emu;
            Result.Dy1 = (int)p1.Y.Emu;
            Result.Dx2 = (int)p1.X.Emu + Width;
            Result.Dy2 = (int)p1.Y.Emu + Height;
            Result.LoadedDataSize = Result.TotalDataSize;

            return Result;
        }
    
        protected void Init()
        {
            if (FParent != null) ((TEscherSpContainerRecord)FParent ).ChildAnchor=this;
        }
    
        internal override void Destroy()
        {
            base.Destroy();
            if (FParent != null) ((TEscherSpContainerRecord)FParent).ChildAnchor=null;
        }

        protected override TEscherRecord DoCopyTo(int RowOfs, int ColOfs, TEscherDwgCache NewDwgCache, TEscherDwgGroupCache NewDwgGroupCache, TSheetInfo SheetInfo)
        {
            TEscherChildAnchorRecord Result= (TEscherChildAnchorRecord) base.DoCopyTo(RowOfs, ColOfs, NewDwgCache, NewDwgGroupCache, SheetInfo);
            Result.Init();
            ((TEscherSpContainerRecord)Result.FParent).ChildAnchor=Result;
            return Result;
        }



        internal int Dx1 {get {return BitConverter.ToInt32(Data,0);} set{BitConverter.GetBytes((Int32)value).CopyTo(Data,0);}}
        internal int Dy1 {get {return BitConverter.ToInt32(Data,4);} set{BitConverter.GetBytes((Int32)value).CopyTo(Data,4);}}
        internal int Dx2 {get {return BitConverter.ToInt32(Data,8);} set{BitConverter.GetBytes((Int32)value).CopyTo(Data,8);}}
        internal int Dy2 {get {return BitConverter.ToInt32(Data,12);} set{BitConverter.GetBytes((Int32)value).CopyTo(Data,12);}}

    }

    /// <summary>
    /// BStore Record.
    /// </summary>
    internal class TEscherBStoreRecord: TEscherContainerRecord
    {
        internal TEscherBStoreRecord(TEscherRecordHeader aEscherHeader, TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
            : base(aEscherHeader, aDwgGroupCache, aDwgCache, aParent)
        {
            Init();
        }

        private void Init()
        {
            if (DwgGroupCache.BStore==null) DwgGroupCache.BStore=this; else XlsMessages.ThrowException(XlsErr.ErrBStoreDuplicated);
        }

        protected override TEscherRecord DoCopyTo(int RowOfs, int ColOfs, TEscherDwgCache NewDwgCache, TEscherDwgGroupCache NewDwgGroupCache, TSheetInfo SheetInfo)
        {
            TEscherBStoreRecord Result= (TEscherBStoreRecord) base.DoCopyTo(RowOfs, ColOfs, NewDwgCache, NewDwgGroupCache, SheetInfo);
            Result.Init();
            return Result;
        }


        internal override void Destroy()
        {
            base.Destroy();
            DwgGroupCache.BStore=null;
        }

        internal void AddRef(int BlipPos)
        {
            if ((BlipPos<1)||(BlipPos> FContainedRecords.Count)) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
            ((TEscherBSERecord)FContainedRecords[BlipPos-1]).AddRef();
        }

        internal void Release(int BlipPos)
        {
            if ((BlipPos<1)||(BlipPos> FContainedRecords.Count)) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
            ((TEscherBSERecord)FContainedRecords[BlipPos-1]).Release();
        }

        internal void FixBSEPositions()
        {
            for (int i = 0; i < FContainedRecords.Count; i++) ((TEscherBSERecord)FContainedRecords[i]).BStorePos = i + 1;
        }

        internal override void SaveToStream(IDataStream DataStream, TSaveData SaveData, TBreakList BreakList)
        {
            FixBSEPositions();
            base.SaveToStream(DataStream, SaveData, BreakList);
        }
    }

    /// <summary>
    /// BSE Record. It holds an individual image.
    /// </summary>
    internal class TEscherBSERecord: TEscherDataRecord
    {
        internal int BStorePos;

        internal TEscherBSERecord(TEscherRecordHeader aEscherHeader, TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
            : base(aEscherHeader, aDwgGroupCache, aDwgCache, aParent)
        {
        }

        internal void AddRef()
        {
            BitOps.IncCardinal(Data,24,1);
        }

        internal void Release()
        {
            BitOps.IncCardinal(Data,24,-1);
            if ((References==0) && (DwgGroupCache.BStore!=null)) 
                DwgGroupCache.BStore.ContainedRecords.Remove(this); //When refs=0 , delete from bstore

        }

        internal long References{ get { return BitOps.GetCardinal(Data, 24);}}

        /// <summary>
        /// Search by signature
        /// </summary>
        /// <param name="aRecord">Another record to compare to</param>
        /// <returns>-1 if this less than aRecord, 0 if equal, 1 if bigger </returns>
        internal override int CompareRec(TEscherRecord aRecord)
        {
            //We can't just compare the data of the 2 records, because cRef can be different
            //no inherited

            if (TotalDataSize< aRecord.TotalDataSize) return -1; else if (TotalDataSize> aRecord.TotalDataSize) return 1; else
            {
                TEscherBSERecord bse = (TEscherBSERecord)aRecord;
                for (int i=2;i<16+2;i++)
                    if (Data[i]<bse.Data[i]) return -1;
                    else
                        if (Data[i]>bse.Data[i]) return 1;
                return 0;
            }
        }

        internal void CopyFromData(byte[] BSEHeader, TEscherRecordHeader BlipHeader, Stream BlipData)
        {
            if (36+BlipData.Length + BlipHeader.Length != TotalDataSize) XlsMessages.ThrowException(XlsErr.ErrInternal);
            BSEHeader.CopyTo(Data,0);
            BlipHeader.Data.CopyTo(Data,36);
            Sh.Read(BlipData, Data, 36+ BlipHeader.Length, (int) BlipData.Length);
            LoadedDataSize=TotalDataSize;
        }

        internal void SaveGraphicToStream(Stream aData, ref TXlsImgType aDataType)
        {
            SaveGraphicToStream(Data, aData, ref aDataType);
        }	
    
        /// <summary>
        /// Converts and internal Excel image to a standard format. This is mostly for internal use, but might be necessary
        /// when parsing shape properties.
        /// </summary>
        /// <param name="Data">Source Image Data.</param>
        /// <param name="DestData">Destination Stream.</param>
        /// <param name="aDataType">Data type of the converted image.</param>
        public static void SaveGraphicToStream(byte[] Data, Stream DestData, ref TXlsImgType aDataType)
        {
            switch ((msoblip)Data[0])
            {
                case msoblip.EMF  : aDataType=TXlsImgType.Emf;break;
                case msoblip.WMF  : aDataType=TXlsImgType.Wmf;break;
                case msoblip.JPEG : aDataType=TXlsImgType.Jpeg;break;
                case msoblip.PNG  : aDataType=TXlsImgType.Png;break;
                case msoblip.DIB  : aDataType=TXlsImgType.Bmp;break;
                default           : aDataType=TXlsImgType.Unknown;break;
            } //case

            SaveGraphicToStream(Data, 36, DestData, aDataType);
        }

        public static void SaveGraphicToStream(byte[] Data, int DataStartPos, Stream DestData, TXlsImgType aDataType)
        {
            const int BI_RGB       = 0;
            /*const int BI_RLE8      = 1;
            const int BI_RLE4      = 2;
            const int BI_BITFIELDS = 3;
            const int BI_JPEG      = 4;
            const int BI_PNG       = 5;*/
            
            int HeadOfs=16;
            if ((aDataType == TXlsImgType.Jpeg) || (aDataType == TXlsImgType.Png) || (aDataType == TXlsImgType.Bmp)) HeadOfs=17;

            int st = DataStartPos+XlsEscherConsts.SizeOfTEscherRecordHeader+HeadOfs;

            if (aDataType == TXlsImgType.Bmp)
            {
                byte[] BmpHead= new byte[14];  //This is an BITMAPFILEHEADER struct, see win32 doc.
                BitOps.SetWord(BmpHead, 0, 0x4D42); //bitmap type ("BM")
                BitOps.SetCardinal(BmpHead, 2, (UInt32)(14 +Data.Length-st));
                
                int BitOfs = BitOps.GetWord(Data, st);
                int BitsPerPixel;
                int BitsPerPixelUsed;
                int Compression = 0;
                int PaletteSize = 4;
                if (BitOfs <= 12) //BITMAPCOREHEADER
                {
                    BitsPerPixel = BitOps.GetWord(Data, st + 10);
                    BitsPerPixelUsed = BitsPerPixel;
                    PaletteSize = 3;
                }
                else
                {
                    BitsPerPixel = BitOps.GetWord(Data, st + 14);
                    Compression = BitOps.GetWord(Data, st + 12);
                    BitsPerPixelUsed = BitOps.GetWord(Data, st + 30);
                    if (BitsPerPixelUsed == 0 && BitsPerPixel != 24) BitsPerPixelUsed = BitsPerPixel;  // 24 bpp images do not have palette, unless BitPixelsUsed != 0
                }

                if (BitsPerPixel < 16)
                {
                    BitOfs += PaletteSize * ((1 << BitsPerPixelUsed)); 
                }
                else 
                {
                    if (Compression != BI_RGB && BitsPerPixelUsed > 0)
                        BitOfs += PaletteSize * ((1 << BitsPerPixelUsed));
                }

                BitOps.SetWord(BmpHead, 10, 14 + BitOfs);
                DestData.Write(BmpHead, 0, BmpHead.Length);
            }

            if (aDataType == TXlsImgType.Wmf)
            {
                XlsMetafiles.ToWMF(Data, st, DestData, false);
            }
            else
                if (aDataType == TXlsImgType.Emf)
            {
                XlsMetafiles.ToWMF(Data, st, DestData, true);
            }
            else
                //bitmap.
                DestData.Write(Data, st, Data.Length-st);
        }

    }

    /// <summary>
    /// DG Record.
    /// </summary>
    internal class TEscherDgRecord: TEscherDataRecord
    {
        internal TEscherDgRecord(TEscherRecordHeader aEscherHeader, TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
            : base(aEscherHeader, aDwgGroupCache, aDwgCache, aParent)
        {
            Init();
        }

        private void Init()
        {
            if (DwgCache.Dg==null) DwgCache.Dg=this; else XlsMessages.ThrowException(XlsErr.ErrDgDuplicated);
        }

        /// <summary>
        /// Create from data
        /// </summary>
        internal TEscherDgRecord(long ShapeCount, int DgId, long FirstId, TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
            : base(new TEscherRecordHeader(DgId << 4,(int)Msofbt.Dg,8), aDwgGroupCache, aDwgCache, aParent)
        {
            Init();
            BitOps.SetCardinal(Data, 0, ShapeCount);
            BitOps.SetCardinal(Data, 4, FirstId + 1);
            LoadedDataSize=TotalDataSize;
        }

        internal override void Destroy()
        {
            DwgGroupCache.Dgg.DestroyClusters(Instance());
            base.Destroy();
            DwgCache.Dg=null;
        }

        protected override TEscherRecord DoCopyTo(int RowOfs, int ColOfs, TEscherDwgCache NewDwgCache, TEscherDwgGroupCache NewDwgGroupCache, TSheetInfo SheetInfo)
        {
            TEscherDgRecord Result= (TEscherDgRecord)base.DoCopyTo(RowOfs, ColOfs, NewDwgCache, NewDwgGroupCache, SheetInfo);
            Result.Init();
            BitOps.SetCardinal(Result.Data, 0, 0); //We are reseting this group.
            int DgId; long FirstId;
            Result.DwgGroupCache.Dgg.GetNewDgIdAndCluster(out DgId, out FirstId);
            Result.Pre = DgId << 4;
            BitOps.SetCardinal(Result.Data, 4, FirstId + 1);  //0x401 for id = 1, 0x801 for id = 2, etc.
            return Result;
        }

        internal long IncMaxShapeId()
        {
            BitOps.IncCardinal(Data,0,1);
            long LastImageId = BitOps.GetCardinal(Data,4);
            long Result = DwgGroupCache.Dgg.AddImage(Instance(), LastImageId);
            BitOps.SetCardinal(Data,4,Result); //this should be done even if lastimageid was bigger than result. (it is not possible on this implementation, since the new cluster will always be bigger). But if it were and we lost our lastid, there would be no way to get it again.
            return Result;
        }

        internal void DecShapeCount()
        {
            BitOps.IncCardinal(Data,0,-1);
            DwgGroupCache.Dgg.RemoveImage(Instance());
        }
    }


    /// <summary>
    /// DGG Record.
    /// </summary>
    internal class TEscherDggRecord: TEscherDataRecord
    {
        internal TEscherDggRecord(TEscherRecordHeader aEscherHeader, TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
            : base(aEscherHeader, aDwgGroupCache, aDwgCache, aParent)
        {
            Init();
        }

        private void Init()
        {
            if (DwgGroupCache.Dgg==null) DwgGroupCache.Dgg=this; else XlsMessages.ThrowException(XlsErr.ErrDggDuplicated);
        }

        /// <summary>
        /// Create from data
        /// </summary>
        internal TEscherDggRecord(TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
            : base(new TEscherRecordHeader(0x0000,(int)Msofbt.Dgg,16), aDwgGroupCache, aDwgCache, aParent)
        {
            Init();

            /* MaxShapeId is the maximum shape id in all clusters + 1. As every cluster hax at most 0x400 shapes, if we have 1 cluster it will be 0x402, with 2 it will be 0x802, 3 0xc02, etc.
             * FIdClCount is the number of clusters + 1. 
             * ShapesSaved is the number of shapes in all clusters. 
             * DwgSaved is one per Dwg record, and we have one dwg record per sheet. So this is the number of sheets with dwg info.
             * 
             * After that comes an array with 8 bytes registers, one per cluster. 
             * The first 4 bytes identify the dwg the cluster represents (it is the same number as the Instance() value of the corresponding Dwg record)
             * The second is the number of images in the cluster, and it must be <= than 0x400.
             * */
            BitOps.SetCardinal(Data,0,2); // FDgg.MaxShapeId:=2;
            BitOps.SetCardinal(Data,4,1); // FDgg.FIDclCount:=1;
            BitOps.SetCardinal(Data,8,0); // FDgg.ShapesSaved:=1;
            BitOps.SetCardinal(Data,12,0); // FDgg.DwgSaved:=1;

            LoadedDataSize=16;

        }

        internal override void Destroy()
        {
            base.Destroy();
            DwgGroupCache.Dgg=null;
        }

        protected override TEscherRecord DoCopyTo(int RowOfs, int ColOfs, TEscherDwgCache NewDwgCache, TEscherDwgGroupCache NewDwgGroupCache, TSheetInfo SheetInfo)
        {
            TEscherDggRecord Result= (TEscherDggRecord)base.DoCopyTo(RowOfs, ColOfs, NewDwgCache, NewDwgGroupCache, SheetInfo);
            Result.Init();
            return Result;
        }

        private long DwgSaved { get {return BitOps.GetCardinal(Data,12);} set {BitOps.SetCardinal(Data,12,value);}}

        private void GetNewCluster(ref int DgId, bool GetNewId, long ShapeCount, out long FirstShapeId)
        {
            //Note that excel will show Picture 1, 2... etc in the name box starting with 
            //0x400 * (pos in this record of first dgid) + 1.  For example, if pos of dgid of the record is 2, 0x801 will be picture 1.

            //find last unused cluster and id.  Unused clusters have a dgid of 0.
            int Found = -1;
            if (GetNewId) 
            {
                DgId = 1; // only if adding a new id.

                for (int i = 16; i + 7 < Data.Length; i+= 8)
                {
                    long id = BitOps.GetCardinal(Data, i);
                    if (Found< 0 && id == 0)
                    {
                        Found = i;
                    }
                
                    if (id >= DgId) DgId = (int)(id + 1);
                }
            }
            else
            {
                //if we are adding a cluster to an existing drawing, we cannot use any slot lower than the last used by that drawing.
                for (int i = Data.Length - 8; i>= 16; i-= 8)
                {
                    long id = BitOps.GetCardinal(Data, i);
                    if (Found< 0 && id == 0)
                    {
                        Found = i;
                    }
                
                    if (id == DgId) break; //do not keep searching.
                }

            }


            if (Found < 0) // no empty clusters, grow the thing
            {
                byte[] NewData = new byte[Data.Length + 8];
                Array.Copy(Data, 0, NewData, 0, Data.Length);

                Found = Data.Length;
                Data = NewData;
                TotalDataSize += 8;
                LoadedDataSize += 8;
                BitOps.IncCardinal(Data,4,1); // FDgg.FIDclCount;
            }

            /*Not needed, it is updated when we add a new image.long MaxShapeId = BitOps.GetCardinal(Data, 0);
            MaxShapeId = ((MaxShapeId / 0x400) + 1) * 0x400;
            BitOps.SetCardinal(Data, 0, MaxShapeId + 1); //MaxShapeId*/

            BitOps.SetCardinal(Data, Found, DgId);
            BitOps.SetCardinal(Data, Found + 4, ShapeCount); //it will be incremented to 1 when we add the first shape.

            FirstShapeId = ((Found - 16) / 8 + 1) * 0x400;
        }


        internal void GetNewDgIdAndCluster(out int DgId, out long FirstShapeId)
        {
            DgId = -1;
            GetNewCluster(ref DgId, true, 0, out FirstShapeId);
            DwgSaved++;
        }

        private void AddNewCluster(int DgId, long ShapeCount, out long FirstShapeId)
        {
            GetNewCluster(ref DgId, false, ShapeCount, out FirstShapeId);
        }

        internal void DestroyClusters(int DgId)
        {
            for (int i = 16; i + 7 < Data.Length; i+= 8)
            {
                if (BitOps.GetCardinal(Data, i) == DgId)
                {
                    BitOps.SetCardinal(Data, i, 0);
                }
            }

            DwgSaved--;
        }

        internal long AddImage(int DgId, long LastImageId)
        {
            long Result = -1;
            int ExpectedCluster = (int)((LastImageId) / 0x400) - 1;
            int ExpectedClusterPos = 16 + ExpectedCluster * 8;

            if (ExpectedClusterPos >= 16 && ExpectedClusterPos <= Data.Length - 8) 
            {
                long ExpectedDgId = BitOps.GetCardinal(Data, ExpectedClusterPos);
                if (ExpectedDgId == DgId)
                {
                    long IdInCluster = BitOps.GetCardinal(Data, ExpectedClusterPos + 4);
                    if (IdInCluster < 0x400)
                    {
                        Result = (ExpectedCluster + 1) * 0x400 + BitOps.GetCardinal(Data, ExpectedClusterPos + 4);
                        BitOps.IncCardinal(Data, ExpectedClusterPos + 4, 1);
                    }
                }
            }

            //Cluster not found or empty, start another.
            if (Result < 0)
            {
                AddNewCluster(DgId, 1, out Result);
            }

            BitOps.IncCardinal(Data,8,1); // FDgg.ShapesSaved+=1;
            long MaxShapeId = BitOps.GetCardinal(Data, 0);
            if (Result + 1 > MaxShapeId) BitOps.SetCardinal(Data, 0, Result + 1);
            
            return Result;
        }

        internal void RemoveImage(int DgId)
        {
            BitOps.IncCardinal(Data,8,-1); // FDgg.ShapesSaved-=1;
        }

        internal bool IsEmpty()
        {
            return DwgSaved <= 0;
        }


    }

    /// <summary>
    /// SP Record.
    /// </summary>
    internal class TEscherSpRecord: TEscherDataRecord
    {
        internal TEscherSpRecord(TEscherRecordHeader aEscherHeader, TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
            : base(aEscherHeader, aDwgGroupCache, aDwgCache, aParent)
        {
            Init();
        }

        private void Init()
        {
            if (DwgCache.Shape!=null) DwgCache.Shape.Add(this);
            if (FParent != null) ((TEscherSpContainerRecord)FParent).SP=this;
        }

        /// <summary>
        /// Only use to create a temp object for searching! SpRecord is not valid.
        /// </summary>
        /// <param name="aShapeId"></param>
        internal TEscherSpRecord(long aShapeId)
            : base(null, null, null, null)
        {
            Data=new byte[4];
            ShapeId=aShapeId;
        }

        /// <summary>
        /// Create from data
        /// </summary>
        internal TEscherSpRecord(int Pre, long aShapeId, TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache,
            TEscherContainerRecord aParent, bool IsGroup, bool HasParent, bool IsPatriarch, bool HasShapeType, bool HasAnchor)
            : base(new TEscherRecordHeader(Pre, (int)Msofbt.Sp, 8), aDwgGroupCache, aDwgCache, aParent)
        {
            Init();
            ShapeId = aShapeId;

            int Flags = 0;
            if (IsGroup) Flags |= 0x01;
            if (HasParent) Flags |= 0x02;
            if (IsPatriarch) Flags |= 0x004;
            if (HasAnchor) Flags |= 0x0200;
            if (HasShapeType) Flags |= 0x0800;

            BitOps.SetCardinal(Data, 4, Flags);
            LoadedDataSize = 8;
        }

        internal override void Destroy()
        {
            base.Destroy();
            int Index=-1;
            if (DwgCache.Dg!=null) DwgCache.Dg.DecShapeCount();
            if (DwgCache.Solver!=null) DwgCache.Solver.DeleteRef(this);
            if (DwgCache.Shape != null)
            {
                if (DwgCache.Shape.Find(ShapeId, ref Index))
                {
                    DwgCache.Shape.Delete(Index);
                }
            }
            if (FParent != null) ((TEscherSpContainerRecord)FParent).SP=null;

            //MADE: Delete all references in connectors with shapedest= self;
        }

        protected override TEscherRecord DoCopyTo(int RowOfs, int ColOfs, TEscherDwgCache NewDwgCache, TEscherDwgGroupCache NewDwgGroupCache, TSheetInfo SheetInfo)
        {
            TEscherSpRecord Result = (TEscherSpRecord)base.DoCopyTo(RowOfs, ColOfs, NewDwgCache, NewDwgGroupCache, SheetInfo);
            Result.Init();
            //This can't be done. We could skip changing the spid if copying the whole thing from a sheet to another, but we might be copying only parts of it. We should need a parameter (like copyingFullSheet) for this.
            //if (NewDwgCache==DwgCache || NewDwgGroupCache!=DwgGroupCache) //We are copying to the same sheet or to a different workbook. When copying to another sheet we will copy everything, so we don't need to inc shapeid.
            Result.ShapeId = Result.DwgCache.Dg.IncMaxShapeId();
            return Result;
        }

        internal long ShapeId { get { return BitOps.GetCardinal(Data,0);} set {BitOps.SetCardinal(Data,0,value);}}
        internal long Flags { get { return BitOps.GetCardinal(Data,4);} set {BitOps.SetCardinal(Data,4,value);}}

        internal TShapeType ShapeType { get { return (TShapeType)(Pre >> 4); } }

        internal override void SaveToStream(IDataStream DataStream, TSaveData SaveData, TBreakList BreakList)
        {
            int SavePre = Pre;
            long SaveFlags = Flags;
            try
            {
                if (ShapeType >= TShapeType.LineInv)
                {
                    if (ShapeType == TShapeType.Trapezoid2007)
                    {
                        Pre = Pre & 0x0F | ((int)TShapeType.Trapezoid << 4);
                        if ((Flags & 0x80) == 0) Flags |= 0x80; else Flags &= ~0x80; //trapezoid is inverted in 2007
                    }
                    else
                    {
                        Pre &= 0x0F;
                    }
                }
                base.SaveToStream(DataStream, SaveData, BreakList);
            }
            finally
            {
                Pre = SavePre;
                Flags = SaveFlags;
            }
            
        }
    }

    /// <summary>
    /// Solver Container Record.
    /// </summary>
    internal class TEscherSolverContainerRecord: TEscherContainerRecord
    {
        internal long MaxRuleId;

        internal TEscherSolverContainerRecord(TEscherRecordHeader aEscherHeader, TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
            : base(aEscherHeader, aDwgGroupCache, aDwgCache, aParent)
        {
            Init();
        }

        private void Init()
        {
            if (DwgCache.Solver==null) DwgCache.Solver=this; else XlsMessages.ThrowException(XlsErr.ErrSolverDuplicated);
        }

        internal override void Destroy()
        {
            DwgCache.Solver=null;
        }

        protected override TEscherRecord DoCopyTo(int RowOfs, int ColOfs, TEscherDwgCache NewDwgCache, TEscherDwgGroupCache NewDwgGroupCache, TSheetInfo SheetInfo)
        {
            TEscherSolverContainerRecord Result=(TEscherSolverContainerRecord) base.DoCopyTo(RowOfs, ColOfs, NewDwgCache, NewDwgGroupCache, SheetInfo);
            Result.Init();
            return Result;
        }

        internal override void SaveToStream(IDataStream DataStream, TSaveData SaveData, TBreakList BreakList)
        {
            Pre = 0x0F | (FContainedRecords.Count << 4);
            base.SaveToStream(DataStream, SaveData, BreakList);
        }

        internal long IncMaxRuleId()
        {
            MaxRuleId+=2;
            return MaxRuleId;
        }

        internal void CheckMax(long aRuleId)
        {
            if (MaxRuleId<aRuleId) MaxRuleId=aRuleId;
        }

        internal void DeleteRef(TEscherSpRecord Shape)
        {
            for (int i=FContainedRecords.Count-1; i>=0; i--) 
                if (((TRuleRecord)FContainedRecords[i]).DeleteRef(Shape)) FContainedRecords.Delete(i);
        }

        internal void FixPointers()
        {
            for (int i=0; i< FContainedRecords.Count;i++)
                ((TRuleRecord)FContainedRecords[i]).FixPointers();
        }

        internal void ArrangeCopyRange(int RowOfs, int ColOfs, TSheetInfo SheetInfo)
        {
            for (int i=0; i< FContainedRecords.Count;i++)
                ((TRuleRecord)FContainedRecords[i]).ArrangeCopyRange(RowOfs, ColOfs, SheetInfo);
        }
    }

    enum EmptyBlip
    {
        Empty
    }

    class TOptComplexData
    {
        internal readonly bool SetBlip;
        internal readonly byte[] Data;

        public TOptComplexData(bool aSetBlip, byte[] aData)
        {
            SetBlip = aSetBlip;
            Data = aData;
        }

        internal TOptComplexData Clone()
        {
            byte[] bNew = new byte[Data.Length];
            Array.Copy(Data, 0, bNew, 0, Data.Length);
            return new TOptComplexData(SetBlip, bNew);
        }
    }

    /// <summary>
    ///  OfficeArtFOPT Record.
    /// </summary>
    internal class TEscherOPTRecord: TEscherDataRecord
    {
        #region Variables
        int? FixedRow; //for searching
        bool? LockByDefault;
        List<TShapeOption> Blips;
        internal long PosInList; //no need to keep in sync, will be updated before used with Blip.FixOPTPositions

        SortedList<TShapeOption, object> Records; //can have TOptComplexData, Int32, EmptyBlip and BSERecords
        
        TShapeFill ShapeFill; //keep a more complete shape fill for xlsx.
        TShapeLine ShapeLine;
        TShapeGeom ShapeGeom;
        TShapeFont ShapeFont;
        TShapeEffects ShapeEffects;
        TEffectProperties EffectProps;
        TDrawingRichString TextExt;
        TDrawingHyperlink HLinkClick;
        TDrawingHyperlink HLinkHover;
        TBodyPr BodyPr;
        string LstStyle;
        #endregion

        #region Constructors
        internal TEscherOPTRecord(TEscherRecordHeader aEscherHeader, TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
            : base(aEscherHeader, aDwgGroupCache, aDwgCache, aParent)
        {
            Init();
            LockByDefault = null;
        }

        private void Init()
        {
            SetParent(FParent);
            Blips = new List<TShapeOption>(1);
            Records = new SortedList<TShapeOption, object>();
            LockByDefault = false;
            ShapeFill = null;
            ShapeLine = null;
            ShapeGeom = null;
        }

        internal override TEscherContainerRecord Parent
        {
            get
            {
                return base.Parent;
            }
            set
            {
                SetParent(value);
            }
        }

        private void SetParent(TEscherContainerRecord value)
        {
            base.Parent = value;
            if ((FParent != null) && (FParent is TEscherSpContainerRecord)) ((TEscherSpContainerRecord)FParent).Opt = this;
        }

        /// <summary>
        /// Creates a new object for searching. Do not use in general, the object is invalid!
        /// </summary>
        /// <param name="aRow"></param>
        internal TEscherOPTRecord(int aRow)
            : base(null, null, null, null)
        {
            FixedRow = aRow;
        }

        private static TEscherOPTRecord CreateEmpty(TEscherOPTRecord ResultData, byte[] DefaultData)
        {
            ResultData.Init();
            ResultData.Data = DefaultData;
            ResultData.AfterCreate();
            ResultData.LoadedDataSize = ResultData.TotalDataSize;
            return ResultData;
        }

        internal static TEscherOPTRecord CreateFromDataImg(byte[] aPict, TXlsImgType aPicType, TBaseImageProperties Props, string ShapeName, TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
        {
            TEscherOPTRecord Result = new TEscherOPTRecord(new TEscherRecordHeader(
                0x0033, // 3 properties.
                (int)Msofbt.OPT,
                6 * 3
                ),
            aDwgGroupCache, aDwgCache, aParent);

            byte[] DefaultData ={
                                   0x04, 0x41, 0x00, 0x00, 0x00, 0x00, // pib  
                                   0xBF, 0x01, 0x00, 0x00, 0x01, 0x00,  //fNoFillHitTest
                                   0xBF, 0x03, 0x00, 0x00, 0x08, 0x00};   //Print

            BitOps.SetCardinal(DefaultData, 2, Result.AddImg(aPict, aPicType) + 1);

            Result = CreateEmpty(Result, DefaultData);
            Result.LockByDefault = true;
            Result.FileName = Props.FileName;
            Result.ShapeName = ShapeName;
            Result.TransparentColor = Props.TransparentColor;
            Result.Brightness = Props.Brightness;
            Result.Contrast = Props.Contrast;
            Result.Gamma = Props.Gamma;
            if (!Props.PreferRelativeSize) Result.PreferRelativeSize = Props.PreferRelativeSize;
            if (!Props.LockAspectRatio) Result.LockAspectRatio = Props.LockAspectRatio;
            if (Props.BiLevel) Result.BiLevel = Props.BiLevel;
            if (Props.Grayscale) Result.Grayscale = Props.Grayscale;

            Result.CropArea = Props.CropArea;
            Result.LoadedDataSize=Result.TotalDataSize;
            return Result;
        }

        internal static TEscherOPTRecord CreateFromDataNote(TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent, int DataNoteDummy)
        {
            TEscherOPTRecord Result = new TEscherOPTRecord(new TEscherRecordHeader(
                0x0093, // 9 properties.
                (int)Msofbt.OPT,
                6 * 9
                ),
            aDwgGroupCache, aDwgCache, aParent);

            byte[] DefaultData ={
                                   0x80, 0x00, 0x00, 0x00, 0x00, 0x00,  //IdText. Nothing, as we dont know what goes here
                                   0xBF, 0x00, 0x08, 0x00, 0x08, 0x00,  //fFitTextToShape
                                   0x58, 0x01, 0x00, 0x00, 0x00, 0x00,  //cxk
                                   0x81, 0x01, 0x50, 0x00, 0x00, 0x08,  //fillcolor
                                   0x83, 0x01, 0x50, 0x00, 0x00, 0x08,  //fillbackcolor
                                   0xBF, 0x01, 0x10, 0x00, 0x11, 0x00,  //fNoFillHitTest
                                   0x01, 0x02, 0x00, 0x00, 0x00, 0x00,  //shadow color
                                   0x3F, 0x02, 0x03, 0x00, 0x03, 0x00,  //shadowObscured
                                   0xBF, 0x03, 0x02, 0x00, 0x0A, 0x00};   //Print

            return CreateEmpty(Result, DefaultData);
        }

        internal static TEscherOPTRecord CreateFromDataShape(TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
        {
            TEscherOPTRecord Result = new TEscherOPTRecord(new TEscherRecordHeader(
                0x0063, // 6 properties.
                (int)Msofbt.OPT,
                6 * 6
                ),
            aDwgGroupCache, aDwgCache, aParent);

            byte[] DefaultData ={
                                           0xBF, 0x00, 0x08, 0x00, 0x08, 0x00,  //fFitTextToShape
                                           0x81, 0x01, 0x50, 0x00, 0x00, 0x08,  //fillcolor
                                           0x83, 0x01, 0x50, 0x00, 0x00, 0x08,  //fillbackcolor
                                           0xBF, 0x01, 0x10, 0x00, 0x11, 0x00,  //fNoFillHitTest
                                           0xC0, 0x01, 0x00, 0x00, 0x00, 0x00,  //line color
                                           0x01, 0x02, 0x00, 0x00, 0x00, 0x00   //shadow color
                               };

            return CreateEmpty(Result, DefaultData);
        }

        internal static TEscherOPTRecord CreateFromDataGroup(TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
        {
            TEscherOPTRecord Result = new TEscherOPTRecord(new TEscherRecordHeader(
                0x0033, // 3 properties.
                (int)Msofbt.OPT,
                6 * 3
                ),
            aDwgGroupCache, aDwgCache, aParent);

            byte[] DefaultData ={
                                           0x04, 0x00, 0x00, 0x00, 0x00, 0x00,  // rotation
                                           0x7F, 0x00, 0x04, 0x00, 0x04, 0x00,  // fLockAgainstGrouping ->locktext
                                           0xBF, 0x03, 0x00, 0x00, 0x00, 0x00   // fPrint 
                               };

            return CreateEmpty(Result, DefaultData);
        }
        
        internal static TEscherOPTRecord CreateFromDataGlobalGroup(TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
        {
            TEscherOPTRecord Result = new TEscherOPTRecord(new TEscherRecordHeader(
                0x0033, // 3 properties.
                (int)Msofbt.OPT,
                6 * 3
                ),
            aDwgGroupCache, aDwgCache, aParent);

            byte[] DefaultData = { 0xBF, 0x00, 0x08, 0x00, 0x08, 0x00, 
                                   0x81, 0x01, 0x09, 0x00, 0x00, 0x08, 
                                   0xC0, 0x01, 0x40, 0x00, 0x00, 0x08 };

            return CreateEmpty(Result, DefaultData);
        }

        internal static TEscherOPTRecord CreateFromDataAutoFilter(TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
        {
            TEscherOPTRecord Result = new TEscherOPTRecord(new TEscherRecordHeader(
                0x0043, // 4 properties.
                (int)Msofbt.OPT,
                6 * 4
                ),
            aDwgGroupCache, aDwgCache, aParent);
              
            byte[] DefaultData={
                                   0x7F, 0x00, 0x04, 0x01, 0x04, 0x01, 
                                   0xBF, 0x00, 0x08, 0x00, 0x08, 0x00, 
                                   0xFF, 0x01, 0x00, 0x00, 0x08, 0x00, 
                                   0xBF, 0x03, 0x00, 0x00, 0x02, 0x00
                               };

            return CreateEmpty(Result, DefaultData);
        }

        internal static TEscherOPTRecord CreateFromDataRadioOrCheckbox(TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
        {
            TEscherOPTRecord Result = new TEscherOPTRecord(new TEscherRecordHeader(
                0x00A3, // 10 properties.
                (int)Msofbt.OPT,
                6 * 10
                ),
            aDwgGroupCache, aDwgCache, aParent);

            byte[] DefaultData ={
                                   0x7F, 0x00, 0x00, 0x01, 0x00, 0x01, //protection: locked
                                   0x80, 0x00, 0x00, 0x00, 0x00, 0x00,  //IdText. Nothing, as we dont know what goes here
                                   0x85, 0x00, 0x01, 0x00, 0x00, 0x00, // wraptext
                                   0x8B, 0x00, 0x02, 0x00, 0x00, 0x00, // Txdir: context
                                   0xBF, 0x00, 0x08, 0x00, 0x1A, 0x00, // text boolean props: auto text margin
                                   0x7F, 0x01, 0x29, 0x00, 0x29, 0x00, // geometry boolean properties
                                   0x81, 0x01, 0x41, 0x00, 0x00, 0x08, //text margin at the left
                                   0xBF, 0x01, 0x00, 0x00, 0x10, 0x00, //fill style bool props
                                   0xC0, 0x01, 0x40, 0x00, 0x00, 0x08, //line color
                                   0xFF, 0x01, 0x00, 0x00, 0x08, 0x00  // line style bool props.
                               };

            return CreateEmpty(Result, DefaultData);
        }

        internal static TEscherOPTRecord CreateFromDataButton(TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
        {
            TEscherOPTRecord Result = new TEscherOPTRecord(new TEscherRecordHeader(
                0x00A3, // 10 properties.
                (int)Msofbt.OPT,
                6 * 10
                ),
            aDwgGroupCache, aDwgCache, aParent);

            byte[] DefaultData ={
                                   0x7F, 0x00, 0x00, 0x01, 0x00, 0x01, //protection: locked
                                   0x80, 0x00, 0x00, 0x00, 0x00, 0x00,  //IdText. Nothing, as we dont know what goes here
                                   0x85, 0x00, 0x01, 0x00, 0x00, 0x00, // wraptext
                                   0x8B, 0x00, 0x02, 0x00, 0x00, 0x00, // Txdir: context
                                   0xBF, 0x00, 0x08, 0x00, 0x1A, 0x00, // text boolean props: auto text margin
                                   0x81, 0x01, 0x43, 0x00, 0x00, 0x08, // fill color
                                   0x83, 0x01, 0x43, 0x00, 0x00, 0x08, // fill back color
                                   0xBF, 0x01, 0x11, 0x00, 0x11, 0x00, //fill style bool props
                                   0xC0, 0x01, 0x40, 0x00, 0x00, 0x08, //line color
                                   0xBF, 0x03, 0x08, 0x00, 0x08, 0x00  // text bool props.
                               };

            return CreateEmpty(Result, DefaultData);
        }

        internal static TEscherOPTRecord CreateFromDataGroupBox(TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
        {
            TEscherOPTRecord Result = new TEscherOPTRecord(new TEscherRecordHeader(
                0x0083, // 8 properties.
                (int)Msofbt.OPT,
                6 * 8
                ),
            aDwgGroupCache, aDwgCache, aParent);
                                                                    
            byte[] DefaultData ={
                                   0x7F, 0x00, 0x00, 0x01, 0x00, 0x01, //protection: locked
                                   0x80, 0x00, 0x00, 0x00, 0x00, 0x00,  //IdText. Nothing, as we dont know what goes here
                                   0x85, 0x00, 0x01, 0x00, 0x00, 0x00, // wraptext
                                   0x8B, 0x00, 0x02, 0x00, 0x00, 0x00, // Txdir: context
                                   0xBF, 0x00, 0x08, 0x00, 0x1A, 0x00, // text boolean props: auto text margin
                                   0xBF, 0x01, 0x00, 0x00, 0x10, 0x00, //fill style bool props
                                   0xC0, 0x01, 0x40, 0x00, 0x00, 0x08, //line color
                                   0xFF, 0x01, 0x08, 0x00, 0x08, 0x00  // line style bool props.
                               };

            return CreateEmpty(Result, DefaultData);
        }

        internal static TEscherOPTRecord CreateFromDataComboBox(TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
        {
            TEscherOPTRecord Result = new TEscherOPTRecord(new TEscherRecordHeader(
                0x0053, // 5 properties.
                (int)Msofbt.OPT,
                6 * 5
                ),
            aDwgGroupCache, aDwgCache, aParent);

            byte[] DefaultData ={
                                   0x7F, 0x00, 0x00, 0x01, 0x00, 0x01, //protection: locked
                                   0x85, 0x00, 0x01, 0x00, 0x00, 0x00, // wraptext
                                   0xBF, 0x00, 0x08, 0x00, 0x08, 0x00, // text boolean props: auto text margin
                                   0xC0, 0x01, 0x40, 0x00, 0x00, 0x08, //line color
                                   0xFF, 0x01, 0x00, 0x00, 0x08, 0x00  // line style bool props.
                               };

            return CreateEmpty(Result, DefaultData);
        }

        internal static TEscherOPTRecord CreateFromDataListBox(TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
        {
            TEscherOPTRecord Result = new TEscherOPTRecord(new TEscherRecordHeader(
                0x0043, // 4 properties.
                (int)Msofbt.OPT,
                6 * 4
                ),
            aDwgGroupCache, aDwgCache, aParent);

            byte[] DefaultData ={
                                   0x7F, 0x00, 0x00, 0x01, 0x00, 0x01, //protection: locked
                                   0xBF, 0x00, 0x08, 0x00, 0x08, 0x00, // text boolean props: auto text margin
                                   0xC0, 0x01, 0x40, 0x00, 0x00, 0x08, //line color
                                   0xFF, 0x01, 0x00, 0x00, 0x08, 0x00  // line style bool props.
                               };

            return CreateEmpty(Result, DefaultData);
        }

        internal static TEscherOPTRecord CreateFromDataLabel(TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
        {
            TEscherOPTRecord Result = new TEscherOPTRecord(new TEscherRecordHeader(
                0x0093, // 9 properties.
                (int)Msofbt.OPT,
                6 * 9
                ),
            aDwgGroupCache, aDwgCache, aParent);

            byte[] DefaultData ={
                                   0x7F, 0x00, 0x00, 0x01, 0x00, 0x01, //protection: locked
                                   0x80, 0x00, 0x00, 0x00, 0x00, 0x00,  //IdText. Nothing, as we dont know what goes here
                                   0x85, 0x00, 0x01, 0x00, 0x00, 0x00, // wraptext
                                   0x8B, 0x00, 0x02, 0x00, 0x00, 0x00, // Txdir: context
                                   0xBF, 0x00, 0x08, 0x00, 0x1A, 0x00, // text boolean props: auto text margin
                                   0x81, 0x01, 0x41, 0x00, 0x00, 0x08,
                                   0xBF, 0x01, 0x00, 0x00, 0x10, 0x00, //fill style bool props
                                   0xC0, 0x01, 0x40, 0x00, 0x00, 0x08, //line color
                                   0xFF, 0x01, 0x00, 0x00, 0x08, 0x00  // line style bool props.
                               };

            return CreateEmpty(Result, DefaultData);
        }

        internal static TEscherOPTRecord CreateFromDataSpinScroll(TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
        {
            TEscherOPTRecord Result = new TEscherOPTRecord(new TEscherRecordHeader(
                0x0033, // 3 properties.
                (int)Msofbt.OPT,
                6 * 3
                ),
            aDwgGroupCache, aDwgCache, aParent);

            byte[] DefaultData ={
                                   0x7F, 0x00, 0x04, 0x01, 0x04, 0x01, //protection: locked
                                   0xBF, 0x00, 0x08, 0x00, 0x08, 0x00, // text boolean props: auto text margin
                                   0xC0, 0x01, 0x40, 0x00, 0x00, 0x08, //line color
                               };

            return CreateEmpty(Result, DefaultData);
        }
        #endregion

        internal override void Destroy()
        {
            if (Blips.Count > 0)
            {
                for (int i = 0; i < Blips.Count; i++)
                {
                    GetBlip(Blips[i]).Release();
                }
                if (DwgCache.Blip != null) DwgCache.Blip.Remove(this);
            }
            DwgCache.OptByName.Remove(this, true);
            if ((FParent != null) && (FParent is TEscherSpContainerRecord)) ((TEscherSpContainerRecord)FParent).Opt = null;
        }

        protected override TEscherRecord DoCopyTo(int RowOfs, int ColOfs, TEscherDwgCache NewDwgCache, TEscherDwgGroupCache NewDwgGroupCache, TSheetInfo SheetInfo)
        {
            TEscherOPTRecord Result = (TEscherOPTRecord)base.DoCopyTo(RowOfs, ColOfs, NewDwgCache, NewDwgGroupCache, SheetInfo);
            Result.Init();
            if ((DwgCache.Blip != null) && (Blips.Count > 0)) NewDwgCache.Blip.Add(Result);

            Result.LockByDefault = LockByDefault;
            Result.FixedRow = FixedRow;

            foreach (KeyValuePair<TShapeOption, object> prop in Records)
            {
                Result.Records[prop.Key] = CloneProp(prop.Value);
            }

            if (NewDwgGroupCache != DwgGroupCache)  //We are copying to another file
            {
                TEscherBStoreRecord NewBStore = NewDwgGroupCache.BStore;
                Debug.Assert(NewBStore != null, "BStore can't be null");

                for (int i = 0; i < Blips.Count; i++)
                {
                    TEscherBSERecord NewBSE = (TEscherBSERecord)(TEscherBSERecord.Clone(GetBlip(Blips[i]), RowOfs, ColOfs, NewDwgCache, NewDwgGroupCache, SheetInfo));
                    int Index = -1;
                    if (!NewBStore.ContainedRecords.Find(NewBSE, ref Index))
                        NewBStore.ContainedRecords.Insert(Index, NewBSE);
                    else NewBSE.Destroy();
                    Result.Records[Blips[i]] = (TEscherBSERecord)NewBStore.ContainedRecords[Index];
                }
            }

            for (int i = 0; i < Blips.Count; i++)
            {
                Result.Blips.Add(Blips[i]);
                Result.GetBlip(Result.Blips[i]).AddRef();
            }


            NewDwgCache.OptByName.AddShapeName(Result);
            NewDwgCache.OptByName.AddShapeId(Result);

            if (ShapeLine != null) Result.ShapeLine = ShapeLine.Clone();
            if (ShapeFill != null) Result.ShapeFill = ShapeFill.Clone();
            if (ShapeGeom != null) Result.ShapeGeom = ShapeGeom.Clone();
            if (ShapeFont != null) Result.ShapeFont = ShapeFont.Clone();
            if (ShapeEffects != null) Result.ShapeEffects = ShapeEffects.Clone();
            Result.EffectProps = EffectProps;
            Result.TextExt = TextExt;
            Result.HLinkClick = HLinkClick;
            Result.HLinkHover = HLinkHover;
            Result.BodyPr = BodyPr;
            Result.LstStyle = LstStyle;
            return Result;
        }

        internal override void AfterCreate()
        {
            int ins = Instance();
            int tPos = 0; long ComplexOfs = ins * 6;
            for (int i = 0; i < ins; i++)
            {
                if (tPos + 6 > TotalDataSize) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
                int Pid = BitOps.GetWord(Data, tPos);
                TShapeOption ShapeOpt = (TShapeOption)(Pid & 0x3FFF);

                if ((Pid & (1 << 15)) != 0) //Complex property
                {
                    long len = GetComplexLen(ShapeOpt, Data, tPos + 2, ComplexOfs);
                    byte[] b = new byte[len];
                    Array.Copy(Data, (int)ComplexOfs, b, 0, (int)len);
                    ComplexOfs += len;
                    Records.Add(ShapeOpt, new TOptComplexData((Pid & (1 << 14)) != 0, b));
                }

                else if ((Pid & (1 << 14)) != 0) //blip
                {
                    //Images default lock to true. We need to read this here to know if this record is an image.
                    LockByDefault = true;

                    int BStorePos = (int)BitOps.GetCardinal(Data, tPos + 2) - 1;
                    if (BStorePos >= 0)
                    {
                        Blips.Add(ShapeOpt);
                        Records[ShapeOpt] = (TEscherBSERecord)DwgGroupCache.BStore.ContainedRecords[BStorePos];
                        if ((DwgCache.Blip != null) && (Blips.Count == 1)) DwgCache.Blip.Add(this);
                    }
                    else
                    {
                        Records[ShapeOpt] = EmptyBlip.Empty;
                    }

                }
                else //long prop
                {
                    Records[ShapeOpt] = BitOps.GetInt32(Data, tPos + 2);
                }


                //Goto Next
                tPos += 6;
            }

            if (DwgCache != null && DwgCache.OptByName != null) DwgCache.OptByName.AddShapeName(this);
            Data = null;
        }

        internal static bool IsIMSOArray(TShapeOption Id)
        {
            switch (Id)
            {
                case TShapeOption.dgmConstrainBounds:
                case TShapeOption.fillShadeColors:
                case TShapeOption.lineLeftDashStyle:
                case TShapeOption.lineRightDashStyle:
                case TShapeOption.lineTopDashStyle:
                case TShapeOption.lineBottomDashStyle:
                case TShapeOption.lineDashStyle:
                case TShapeOption.pAdjustHandles:
                case TShapeOption.pConnectionSites:
                case TShapeOption.pConnectionSitesDir:
                case TShapeOption.pInscribe:
                case TShapeOption.pRelationTbl:
                case TShapeOption.pSegmentInfo:
                case TShapeOption.pVertices:
                case TShapeOption.pWrapPolygonVertices:
                case TShapeOption.tableRowProperties:
                    return true;
            }
            return false;
        }

        internal static long GetComplexLen(TShapeOption ShapeOpt, byte[] BiffData, int tPos, long ComplexOfs)
        {
            long Result = BitOps.GetCardinal(BiffData, tPos);
            if (Result > 0 && IsIMSOArray(ShapeOpt))
            {
                unchecked
                {
                    int ElementSize = (short)BitOps.GetWord(BiffData, (int)ComplexOfs + 4);
                    if (ElementSize < 0) ElementSize = -ElementSize / 4; //When ElementSize < 0, it seems like it is the count in half-bytes.
                    return BitOps.GetWord(BiffData, (int)ComplexOfs) * ElementSize + 6; 
                    //ImsoArrays have a 6 byte header, and sometimes those bytes are included on the length and sometimes not.
                    //This depends on the version of Excel, newer Excels do it right. So we need to make sure we read old ones, but we can save the correct value.
                }
            }
            return Result;
        }

        internal TShapeOptionList ShapeOptions()
        {
            if (DwgCache != null && DwgCache.Blip != null) DwgCache.Blip.FixOPTPositions();
            TShapeOptionList Result = new TShapeOptionList();

            foreach (KeyValuePair<TShapeOption, object> item in Records)
            {
                TEscherBSERecord bse = item.Value as TEscherBSERecord;
                if (bse != null)
                {
                    Result.Add(item.Key, PosInList);
                }
                else

                    if (item.Value is EmptyBlip)
                    {
                        Result.Add(item.Key, (long)0);
                    }
                    else if (item.Value is Int32)
                    {
                        if (item.Key == TShapeOption.pictureTransparent)
                        {
                            Result.Add(item.Key, TransparentColor);
                        }
                        else
                        {
                            unchecked
                            {
                                Result.Add(item.Key, (long)((UInt32)((Int32)item.Value)));
                            }
                        }
                    }
                    else
                    {
                        Result.Add(item.Key, ClonePropForOpt(item.Value));
                    }
            }
            return Result;
        }

        private object CloneProp(object p)
        {
            TOptComplexData b = p as TOptComplexData;
            if (b != null)
            {
                return b.Clone();
            }

            return p;
        }

        private object ClonePropForOpt(object p)
        {
            TOptComplexData b = p as TOptComplexData;
            if (b != null)
            {
                return b.Clone().Data;
            }

            return p;
        }

        #region GetProps
        internal int GetInt(TShapeOption so, int DefaultValue)
        {
            object o;
            if (!Records.TryGetValue(so, out o)) return DefaultValue;
            return Convert.ToInt32(o);
        }

        private uint GetUInt(TShapeOption so, uint DefaultValue)
        {
            object o;
            if (!Records.TryGetValue(so, out o)) return DefaultValue;
            unchecked
            {
                return (UInt32)((Int32)(o));
            }
        }

        private double Get1616(TShapeOption so, double DefaultValue)
        {
            object o;
            if (!Records.TryGetValue(so, out o)) return DefaultValue;
            Int32 value = (Int32)o;
            unchecked
            {
                return (short)((value >> 16)) + (value & 0xFFFF) / 65536f;
            }
        }

        private TEscherBSERecord GetBlip(TShapeOption so)
        {
            return Records[so] as TEscherBSERecord;
        }

        private bool GetBool(TShapeOption so, UInt16 mask, bool? DefaultValue)
        {
            UInt32 v = GetUInt(so, 0);

            bool UseValue = (v & ((UInt32)mask << 16)) != 0;
            if (UseValue) return (v & mask) == mask;

            if (DefaultValue.HasValue && DefaultValue.Value) return true;
            return false;
        }

        private bool GetBool(TShapeOption so, UInt32 mask)
        {
            return (GetUInt(so, 0) & mask) == mask;
        }

        private string GetString(TShapeOption so)
        {
            byte[] StrData = GetByteArrayProp(so);
            if (StrData == null) return null;
            string Result;
            if (TShapeOptionList.IsASCII(so))
            {
                Result = Encoding.ASCII.GetString(StrData, 0, StrData.Length);
            }
            else
                if (TShapeOptionList.IsUTF8(so))
                {
                    Result = Encoding.UTF8.GetString(StrData, 0, StrData.Length);
                }
                else Result = Encoding.Unicode.GetString(StrData, 0, StrData.Length);

            int k = Result.Length - 1;
            while (k >= 0 && Result[k] == (char)0) k--;
            return Result.Substring(0, k + 1);
        }

        private byte[] GetByteArrayProp(TShapeOption so)
        {
            object o;
            if (!Records.TryGetValue(so, out o)) return null;
            TOptComplexData cpx = o as TOptComplexData;
            if (cpx == null) return null;
            return cpx.Data;
        }

        #endregion

        #region SetProps
        int PropLen(TShapeOption Id)
        {
            object obj;
            if (!Records.TryGetValue(Id, out obj)) return 0;
            TOptComplexData b = obj as TOptComplexData;
            if (b == null) return 6;
            return 6 + b.Data.Length;

        }
        internal void RemoveProperty(TShapeOption Id)
        {
            int len = PropLen(Id);
            if (len > 0)
            {
                TotalDataSize -= len;
                LoadedDataSize -= len;
                Pre -= 0x10;
                Records.Remove(Id);
            }
        }

        internal void RemoveProperties(TShapeOption First, TShapeOption Last)
        {
            //Why no binarysearch in sorted list??  Well... it doesn't matter, there aren't too many records to search here.
            for (int i = Records.Count - 1; i >= 0; i--)
            {
                TShapeOption key = Records.Keys[i];
                if (key < First) return;
                if (key > Last) continue;

                RemoveProperty(key);
            }
        }


        internal void SetStringProperty(TShapeOption Id, string Value)
        {
            if (string.IsNullOrEmpty(Value))
            {
                RemoveProperty(Id);
                return;
            }

            byte[] StrData;
            if (TShapeOptionList.IsUTF8(Id))
            {
                StrData = Encoding.UTF8.GetBytes(Value + (char)0);
            }
            else if (TShapeOptionList.IsASCII(Id))
            {
                StrData = Encoding.ASCII.GetBytes(Value + (char)0);
            }
            else
            {
                StrData = Encoding.Unicode.GetBytes(Value + (char) 0);
            }
            SetByteArrayProperty(Id, StrData);

        }

        internal void SetHLinkProperty(TShapeOption Id, THyperLink NewLink)
        {
            if (NewLink == null)
            {
                RemoveProperty(Id);
                return;
            }
            
            THLinkRecord tmp = THLinkRecord.CreateNew(new TXlsCellRange(0, 0, 0, 0), NewLink);
            byte[] NewValue2 = tmp.GetHLinkStream();
            SetByteArrayProperty(Id, NewValue2);
        }

        internal void SetIntProperty(TShapeOption Id, Int32 NewValue, Int32? DefaultValue)
        {
            if (DefaultValue.HasValue && NewValue == DefaultValue.Value)
            {
                RemoveProperty(Id);
                return;
            }

            if (!Records.ContainsKey(Id))
            {
                LoadedDataSize += 6;
                TotalDataSize += 6;
                Pre += 0x10;
            }
            Records[Id] = (Int32)(NewValue);
        }

        internal void SetUIntProperty(TShapeOption Id, UInt32 NewValue, UInt32? DefaultValue)
        {
            unchecked
            {
                SetIntProperty(Id, (Int32)NewValue, (Int32?)DefaultValue);
            }
        }

        internal void SetLongProperty(TShapeOption Id, long NewValue)
        {
            unchecked
            {
                SetIntProperty(Id, (Int32)NewValue, null);
            }
        }

        internal void SetBoolProperty(TShapeOption Id, UInt32 NewValue, bool AndValue, bool OrValue)
        {
            if (AndValue)
            {
                UInt32 OldValue = GetUInt(Id, 0);
                NewValue = NewValue & OldValue;
            }
            if (OrValue)
            {
                UInt32 OldValue = GetUInt(Id, 0);
                NewValue = NewValue | OldValue;
            }

            SetUIntProperty(Id, NewValue, null);
        }

        internal void SetByteArrayProperty(TShapeOption Id, byte[] NewValue)
        {
            object OldOpt;
            bool SetBlip = true;
            if (Records.TryGetValue(Id, out OldOpt))
            {
                TOptComplexData oo = OldOpt as TOptComplexData;
                if (oo != null) SetBlip = oo.SetBlip;
            }
            RemoveProperty(Id);
            if (NewValue == null)
            {
                return;
            }

            Records[Id] = new TOptComplexData(SetBlip, NewValue);

            int len = PropLen(Id);
            LoadedDataSize += len;
            TotalDataSize += len;
            Pre += 0x10;
        }
        #endregion

        #region Properties
        internal string ShapeName 
        { 
            get 
            {
                return GetString(TShapeOption.wzName);
            }
            set
            {
                string v = value;
                if (string.IsNullOrEmpty(value)) v = null;
                if (v == ShapeName) return;

                DwgCache.OptByName.Remove(this, false);
                SetStringProperty(TShapeOption.wzName, v);
                DwgCache.OptByName.AddShapeName(this);
            }
        }

        internal TShapeType ShapeType
        {
            get
            {

                TEscherSpRecord Sp = Parent.FindRec<TEscherSpRecord>() as TEscherSpRecord;
                return Sp.ShapeType;
            }
        }

        internal bool Visible
        {
            get
            {
                return !GetBool(TShapeOption.fPrint, 0x020002);
            }
            set
            {
                if (!value)
                {
                    SetBoolProperty(TShapeOption.fPrint, 0x00020002, false, true);
                }
                else
                {
                    SetBoolProperty(TShapeOption.fPrint, ~0x00000002u, true, false);
                }
            }
        }

        internal TBwMode BwMode
        {
            get
            {
                return (TBwMode)GetInt(TShapeOption.bWMode, 0);
            }

        }

        internal bool PreferRelativeSize
        {
            get
            {
                return GetBool(TShapeOption.fBackground, 0x10, LockByDefault);
            }
            set
            {
                if (value)
                {
                    SetBoolProperty(TShapeOption.fBackground, 0x00100010, false, true);
                }
                else
                {
                    SetBoolProperty(TShapeOption.fBackground, ~0x00100010u, true, false);
                    SetBoolProperty(TShapeOption.fBackground, 0x00100000, false, true);
                }
            }
        }

        internal bool LockAspectRatio
        {
            get
            {
                return GetBool(TShapeOption.fLockAgainstGrouping, 0x80, LockByDefault);
            }
            set
            {
                if (value)
                {
                    SetBoolProperty(TShapeOption.fLockAgainstGrouping, 0x00800080, false, true);
                }
                else
                {
                    SetBoolProperty(TShapeOption.fLockAgainstGrouping, ~0x00800080u, true, false);
                    SetBoolProperty(TShapeOption.fLockAgainstGrouping, 0x00800000, false, true);
                }
            }
        }

        internal bool Grayscale
        {
            get
            {
                return GetBool(TShapeOption.pictureActive, 0x040004);
            }
            set
            {
                if (value)
                {
                    SetBoolProperty(TShapeOption.pictureActive, 0x00040004, false, true);
                }
                else
                {
                    SetBoolProperty(TShapeOption.pictureActive, ~0x00040004u, true, false);
                }
            }
        }

        internal bool BiLevel
        {
            get
            {
                return GetBool(TShapeOption.pictureActive, 0x020002);
            }
            set
            {
                if (value)
                {
                    SetBoolProperty(TShapeOption.pictureActive, 0x00020002, false, true);
                }
                else
                {
                    SetBoolProperty(TShapeOption.pictureActive, ~0x00020002u, true, false);
                }
            }
        }

        internal string FileName
        {
            get 
            {
                return GetString(TShapeOption.pibName);
            }
            set
            {
                SetStringProperty(TShapeOption.pibName, value);
            }
        }

        internal string AltText
        {
            get
            {
                return GetString(TShapeOption.wzDescription);
            }
        }

        internal TCropArea CropArea
        {
            get 
            {
                TCropArea Result = new TCropArea();
                object v;
                unchecked
                {
                    if (Records.TryGetValue(TShapeOption.cropFromTop, out v)) Result.CropFromTop = (Int32)v;
                    if (Records.TryGetValue(TShapeOption.cropFromLeft, out v)) Result.CropFromLeft = (Int32)v;
                    if (Records.TryGetValue(TShapeOption.cropFromBottom, out v)) Result.CropFromBottom = (Int32)v;
                    if (Records.TryGetValue(TShapeOption.cropFromRight, out v)) Result.CropFromRight = (Int32)v;
                }

                return Result;
            }
            set
            {
                if (value == null)
                {
                    RemoveProperty(TShapeOption.cropFromTop);
                    RemoveProperty(TShapeOption.cropFromLeft);
                    RemoveProperty(TShapeOption.cropFromBottom);
                    RemoveProperty(TShapeOption.cropFromRight);
                }
                else
                {
                    SetIntProperty(TShapeOption.cropFromTop, value.CropFromTop, 0);
                    SetIntProperty(TShapeOption.cropFromLeft, value.CropFromLeft, 0);
                    SetIntProperty(TShapeOption.cropFromBottom, value.CropFromBottom, 0);
                    SetIntProperty(TShapeOption.cropFromRight, value.CropFromRight, 0);
                }
            }
        }

        internal long TransparentColor
        {
            get
            {
                unchecked
                {
                    long Result = GetUInt(TShapeOption.pictureTransparent, (UInt32)FlxConsts.NoTransparentColor);
                    if (Result == (UInt32)FlxConsts.NoTransparentColor) return FlxConsts.NoTransparentColor;
                    return Result;
                }
            }
            set
            {
                unchecked
                {
                    SetUIntProperty(TShapeOption.pictureTransparent, (UInt32)value, (UInt32)FlxConsts.NoTransparentColor);
                }
            }
        }

        internal bool HasFill
        {
            get
            {
                 return GetBool(TShapeOption.fNoFillHitTest, 0x100010);
            }
        }

        internal bool HasLine
        {
            get
            {
                return GetBool(TShapeOption.fNoLineDrawDash, 0x080008);
            }
        }

        internal int Brightness
        {
            get
            {
                return GetInt(TShapeOption.pictureBrightness, FlxConsts.DefaultBrightness);
            }
            set
            {
                SetIntProperty(TShapeOption.pictureBrightness, value, FlxConsts.DefaultBrightness);
            }
        }

        internal int Contrast
        {
            get 
            {
                return GetInt(TShapeOption.pictureContrast, FlxConsts.DefaultContrast); 
            }
            set
            {
                SetIntProperty(TShapeOption.pictureContrast, value, FlxConsts.DefaultContrast);
            }
        }

        internal int Gamma
        {
            get
            {
                return GetInt(TShapeOption.pictureGamma, FlxConsts.DefaultGamma);
            }
            set
            {
                SetIntProperty(TShapeOption.pictureGamma, value, FlxConsts.DefaultGamma);
            }
        }

        internal double Rotation
        {
            get
            {
                return Get1616(TShapeOption.Rotation, 0);
            }
        }


        #endregion

        private int AddImg(byte[] ImgData, TXlsImgType DataType)
        {
            TEscherBStoreRecord BStore = DwgGroupCache.BStore;
            Debug.Assert(BStore != null, "BStore can't be null");
            TEscherBSERecord BSE = EscherGraphToBSE.Convert(ImgData, DataType, DwgGroupCache, DwgCache);
            int Result = 0;
            if (!BStore.ContainedRecords.Find(BSE, ref Result))
                BStore.ContainedRecords.Insert(Result, BSE);
            else BSE.Destroy();
            ((TEscherBSERecord)BStore.ContainedRecords[Result]).AddRef();
            return Result;
        }

        internal override void SaveToStream(IDataStream DataStream, TSaveData SaveData, TBreakList BreakList)
        {
            Data = GetBiff8Data();
            try
            {
                base.SaveToStream(DataStream, SaveData, BreakList);
            }
            finally
            {
                Data = null;
            }
        }

        private byte[] GetBiff8Data()
        {
            byte[] Result = new byte[TotalDataSize];
            int MainPos = 0;
            int ComplexPos = 6 * Records.Count;
            foreach (KeyValuePair<TShapeOption, object> item in Records)
            {
                TOptComplexData b = item.Value as TOptComplexData;
                if (b != null)
                {
                    int BlipBit = b.SetBlip ? (1 << 14) : 0;
                    BitOps.SetWord(Result, MainPos, (int)item.Key | (1 << 15) | BlipBit);
                    BitOps.SetCardinal(Result, MainPos + 2, b.Data.Length);
                    Array.Copy(b.Data, 0, Result, ComplexPos, b.Data.Length);

                    ComplexPos += b.Data.Length;
                    MainPos += 6;
                    continue;
                }

                TEscherBSERecord bse = item.Value as TEscherBSERecord;
                if (bse != null)
                {
                    BitOps.SetWord(Result, MainPos, (int)item.Key | (1 << 14));
                    BitOps.SetCardinal(Result, MainPos + 2, bse.BStorePos);
                    MainPos += 6;
                    continue;
                }

                if (item.Value is EmptyBlip)
                {
                    BitOps.SetWord(Result, MainPos, (int)item.Key | (1 << 14));
                    BitOps.SetCardinal(Result, MainPos + 2, 0);
                    MainPos += 6;
                    continue;
                }

                BitOps.SetWord(Result, MainPos, (int)item.Key);
                BitOps.SetCardinal(Result, MainPos + 2, (Int32)item.Value);             
                MainPos += 6;

            }

            Debug.Assert(MainPos == 6 * Records.Count);
            Debug.Assert(ComplexPos == Result.Length);
            return Result;
        }

        #region Anchor
        internal int Row
        {
            get
            {
                if (FixedRow.HasValue) return FixedRow.Value;  //For searching.
                TEscherRecord Fr = FindRoot();
                TEscherSpContainerRecord SpContainer = Fr as TEscherSpContainerRecord;
                if ((DwgCache.Patriarch == null) || (Fr == null) || (SpContainer.ClientAnchor == null)) return 0;
                else return SpContainer.Row;
            }
        }

        internal int Col
        {
            get
            {
                TEscherRecord Fr = FindRoot();
                TEscherSpContainerRecord SpContainer = Fr as TEscherSpContainerRecord;
                if ((DwgCache.Patriarch == null) || (Fr == null) || (SpContainer.ClientAnchor == null)) return 0;
                else return SpContainer.Col;
            }
        }

        private static TClientAnchor CalcChildCoords(int[] ParentCoords, TEscherChildAnchorRecord ChildAnchor)
        {
            if (ParentCoords == null) return new TClientAnchor(); //This should only happen when using GetImageProps directly.
            TClientAnchor Result = new TClientAnchor();

            double Width = ParentCoords[2] - ParentCoords[0];
            double Height = ParentCoords[3] - ParentCoords[1];

            if (Width > 0 && Height > 0)
            {
                Result.ChildAnchor = new TChildAnchor(
                        (ChildAnchor.Dx1 - ParentCoords[0]) / Width,
                        (ChildAnchor.Dy1 - ParentCoords[1]) / Height,
                        (ChildAnchor.Dx2 - ParentCoords[0]) / Width,
                        (ChildAnchor.Dy2 - ParentCoords[1]) / Height
                );
            }

            return Result;
        }

        internal TClientAnchor GetAnchor(ref int[] ParentCoords)
        {
            if (DwgCache.Patriarch==null) return new TClientAnchor();

            TEscherSpContainerRecord SpContainer = Parent as TEscherSpContainerRecord;

            if (SpContainer != null)
            {
                TClientAnchor Result;
                if (SpContainer.ClientAnchor == null)
                    Result = new TClientAnchor();
                else
                    Result = ((TEscherImageAnchorRecord)SpContainer.ClientAnchor).GetAnchor();

                TEscherSpgrContainerRecord SpgrContainer = SpContainer.Parent as TEscherSpgrContainerRecord;
                if (SpgrContainer != null && SpgrContainer.ContainedRecords.Count > 0)
                {
                    //Information on the shape is always on the first sp record of the group.
                    TEscherSpContainerRecord SpGlobalContainer = SpgrContainer.ContainedRecords[0] as TEscherSpContainerRecord;

                    if (SpContainer == SpGlobalContainer) //In this case, we need to set the new parent coords.
                    {
                        TEscherSpgrRecord Spgr = (TEscherSpgrRecord)SpContainer.FindRec<TEscherSpgrRecord>();
                        if (Spgr != null)
                        {
                            if (SpContainer.ChildAnchor != null)
                            {
                                Result = CalcChildCoords(ParentCoords, SpContainer.ChildAnchor);
                            }
                            ParentCoords = Spgr.Bounds;
                        }
                    }
                    else
                    {
                        if (SpContainer.ChildAnchor != null)
                        {
                            Result = CalcChildCoords(ParentCoords, SpContainer.ChildAnchor);
                        }
                    }
                }


                return Result;
            }           

            return new TClientAnchor();
        }

        internal THeaderOrFooterAnchor GetHeaderAnchor()
        {
            TEscherRecord Fr = FindRoot();
            TEscherSpContainerRecord SpContainer = (Fr as TEscherSpContainerRecord);

            if ((DwgCache.Patriarch == null) || (SpContainer == null) ||
                (SpContainer.ClientAnchor == null)) return new THeaderOrFooterAnchor(0, 0);

            else return ((TEscherHeaderAnchorRecord)SpContainer.ClientAnchor).GetHeaderAnchor();
        }

        internal void SetAnchor(TClientAnchor aAnchor, TSheet sSheet)
        {
            TEscherRecord Fr = FindRoot();
            TEscherSpContainerRecord SpContainer = (Fr as TEscherSpContainerRecord);

            if ((DwgCache.Patriarch == null) || (SpContainer == null) ||
                (SpContainer.ClientAnchor == null)) return;

            ((TEscherImageAnchorRecord)SpContainer.ClientAnchor).SetAnchor(aAnchor, sSheet);
        }

        internal void SetAnchor(THeaderOrFooterAnchor aAnchor)
        {
            TEscherRecord Fr=FindRoot();
            TEscherSpContainerRecord SpContainer = (Fr as TEscherSpContainerRecord);

            if ((DwgCache.Patriarch==null) || (SpContainer==null) || 
                (SpContainer.ClientAnchor==null)) return;

            ((TEscherHeaderAnchorRecord)SpContainer.ClientAnchor).SetAnchor(aAnchor);
        }
        #endregion

        #region BSE
        internal void ChangeRef(TEscherBSERecord aBSE)
        {
            if (Blips.Count < 1) XlsMessages.ThrowException(XlsErr.ErrChangingEscher);
            for (int i = 0; i < Blips.Count; i++)
            {
                if (GetBlip(Blips[i]) == aBSE) continue;
                aBSE.AddRef();
                GetBlip(Blips[i]).Release();
                Records[Blips[i]] = aBSE;
            }
        }

        internal void ReplaceImg(byte[] ImgData, TXlsImgType DataType)
        {
            TEscherBStoreRecord BStore = DwgGroupCache.BStore;
            Debug.Assert(BStore != null, "BStore can't be null");
            TEscherBSERecord BSE = EscherGraphToBSE.Convert(ImgData, DataType, DwgGroupCache, DwgCache);
            int Index = -1;
            if (!BStore.ContainedRecords.Find(BSE, ref Index))
                BStore.ContainedRecords.Insert(Index, BSE);
            else BSE.Destroy();
            ChangeRef((TEscherBSERecord)BStore.ContainedRecords[Index]);
        }

        internal bool HasBlip()
        {
            return Blips.Count > 0;
        }

        internal void GetImageFromStream(Stream ImgData, ref TXlsImgType DataType)
        {
            if (Blips.Count == 0) return;
            if (Blips.Count != 1) XlsMessages.ThrowException(XlsErr.ErrChangingEscher);
            GetBlip(Blips[0]).SaveGraphicToStream(ImgData, ref DataType);
        }

        internal long ReferencesCount()
        {
            if (Blips.Count != 1) XlsMessages.ThrowException(XlsErr.ErrChangingEscher);
            return GetBlip(Blips[0]).References;
        }
        #endregion

        #region ObjId
        internal TMsObj GetObj()
        {
            TEscherClientDataRecord obj = Parent.FindRec<TEscherClientDataRecord>();
            return obj.ClientData as TMsObj;
        }

        internal TTXO GetTXO()
        {
            TEscherClientTextBoxRecord obj = Parent.FindRec<TEscherClientTextBoxRecord>();
            if (obj == null) return null;
            return obj.ClientData as TTXO;
        }

        internal long ShapeId()
        {
            TEscherSpRecord Sp = (TEscherSpRecord)Parent.FindRec<TEscherSpRecord>();
            if (Sp == null) return -1;
            return Sp.ShapeId;
        }

        internal void AddShapeId()
        {
            if (DwgCache.OptByName == null) return;
            DwgCache.OptByName.AddShapeId(this);
        }
        #endregion

        #region Set Line and Fill
        internal void SetFillColor(IFlexCelPalette aPalette, TShapeFill aShapeFill, TSystemColor DefaultSysColor)
        {
            ShapeFill = aShapeFill;
            TSolidFill sf = aShapeFill == null? null: aShapeFill.GetFill(aPalette) as TSolidFill;
            long sfc = 0x08000000 | ColorUtil.GetSysColor(DefaultSysColor) + 56;
            SetBiff8Fill(aPalette, TShapeOption.fillColor, DefaultSysColor, sf, sfc);
            bool HasFill = aShapeFill == null ? false : aShapeFill.HasFill;
            SetBoolFillProp(HasFill, TShapeOption.fNoFillHitTest, 4);
        }

        internal void SetLineStyle(IFlexCelPalette aPalette, TShapeLine aShapeLine, TSystemColor DefaultSysColor)
        {
            ShapeLine = aShapeLine;
            TFillStyle LineFill = aShapeLine == null? null: aShapeLine.GetLineFill(aPalette);
            TSolidFill sf = LineFill as TSolidFill;
            long sfc = -1;
            SetBiff8Fill(aPalette, TShapeOption.lineColor, DefaultSysColor, sf, sfc);
            bool HasLine = aShapeLine == null || LineFill is TNoFill ? false : aShapeLine.HasLine;
            
            SetBoolFillProp(HasLine, TShapeOption.fNoLineDrawDash, 3);

            if (aShapeLine != null)
            {
                SetIntProperty(TShapeOption.lineWidth, aShapeLine.GetWidth(aPalette), TLineStyle.DefaultWidth);
                SetIntProperty(TShapeOption.lineDashing, (int)aShapeLine.GetDashing(aPalette), 0);

                TLineArrow HeadArrow = aShapeLine.GetHeadArrow(aPalette);
                SetIntProperty(TShapeOption.lineStartArrowhead, (int)HeadArrow.Style, 0);
                SetIntProperty(TShapeOption.lineStartArrowLength, (int)HeadArrow.Len, 1);
                SetIntProperty(TShapeOption.lineStartArrowWidth, (int)HeadArrow.Width, 1);
                TLineArrow TailArrow = aShapeLine.GetTailArrow(aPalette);
                SetIntProperty(TShapeOption.lineEndArrowhead, (int)TailArrow.Style, 0);
                SetIntProperty(TShapeOption.lineEndArrowLength, (int)TailArrow.Len, 1);
                SetIntProperty(TShapeOption.lineEndArrowWidth, (int)TailArrow.Width, 1);
            }
        }

        private void SetBiff8Fill(IFlexCelPalette aPalette, TShapeOption ShOpt, TSystemColor DefaultSysColor, TSolidFill sf, long sfc)
        {
            if (sf != null)
            {
                switch (sf.Color.ColorType)
                {
                    case TDrawingColorType.System:
                        int SysColor = ColorUtil.GetSysColor(sf.Color.System);
                        if (SysColor < 0) SysColor = ColorUtil.GetSysColor(DefaultSysColor);
                        sfc = 0x08000000 | (SysColor + 56);
                        break;
                    default:
                        sfc = ColorUtil.BgrToRgb(sf.Color.ToColor(aPalette).ToArgb());
                        sfc = sfc & 0x00FFFFFF; //fourth byte is control. Since this is an rgb color, we use 0.
                        break;
                }
            }

            if (sfc != -1) SetLongProperty(ShOpt, sfc);
        }        

        private void SetBoolFillProp(bool value, TShapeOption ShpOpt, int ShpPos)
        {

            if (value)
            {
                //Remove "no fill" bit:
                if (!ShapeOptions().AsBool(ShpOpt, true, ShpPos))
                {
                    UInt32 vtrue = 0x10001u << ShpPos;
                    SetBoolProperty(ShpOpt, vtrue, false, true);
                }
            }
            else
            {
                if (ShapeOptions().AsBool(ShpOpt, true, ShpPos))
                {
                    UInt32 vfalse = 0x00001u << ShpPos;
                    SetBoolProperty(ShpOpt, ~vfalse, true, false);
                    UInt32 vfalse2 = 0x10000u << ShpPos;
                    SetBoolProperty(ShpOpt, vfalse2, false, true);
                }
            }
        }
        #endregion

        #region Get Fill
        internal TShapeFill GetFillColor(ExcelFile xls)
        {
            if (ShapeFill != null) return ShapeFill.Clone();
            TFillType ShadeType = (TFillType)GetUInt(TShapeOption.fillType, 0);
            switch (ShadeType)
            {
                case TFillType.Pattern:
                case TFillType.Texture:
                case TFillType.Picture:
                    return GetTextureBrush((TFillType)ShadeType);

                case TFillType.Shade:
                case TFillType.ShadeCenter:
                case TFillType.ShadeScale:
                case TFillType.ShadeShape:
                case TFillType.ShadeTitle:
                    return GetGradientBrush(xls, (TFillType)ShadeType);
            }

            //Anything else:
            return GetSolidBrush(xls);
        }

        private TShapeFill GetTextureBrush(TFillType FillType)
        {
            return new TShapeFill(false, new TNoFill());
        }

        private TShapeFill GetGradientBrush(ExcelFile Workbook, TFillType FillType)
        {
            /* if (ShProp.ShapeOptions.AsBool(TShapeOption.fNoFillHitTest, false, 1))
             {
                 Coords2 = FlexCelRender.RectangleXY(
                     GetInt(ShProp.ShapeOptions[TShapeOption.fillRectLeft], 0),
                     GetInt(ShProp.ShapeOptions[TShapeOption.fillRectTop], 0),
                     GetInt(ShProp.ShapeOptions[TShapeOption.fillRectRight], 0),
                     GetInt(ShProp.ShapeOptions[TShapeOption.fillRectBottom], 0));
             }

             if (Coords2.Width <= 0 || Coords2.Height <= 0) return null; */

            double FillOpacity = Get1616(TShapeOption.fillOpacity, 1);
            TDrawingColor Color1 = ColorFromLong(TShapeOption.fillColor, 0xffffffff, FillOpacity, Workbook, true, 0);

            double FillBackOpacity = Get1616(TShapeOption.fillBackOpacity, 1);
            TDrawingColor Color2 = ColorFromLong(TShapeOption.fillBackColor, 0xffffffff, FillBackOpacity, Workbook, true, 0);

            //Y coords here go the other way around. so there are lots of "-" symbols.

            double FillAngle = Get1616(TShapeOption.fillAngle, 0);
            while (FillAngle <= -360) FillAngle += 360;
            while (FillAngle >= 360) FillAngle -= 360; //Excel 2007 behaves different here. An angle of 360 = 0. In 2003 it is -360

            //FillFocus does not work with multiple colors!
            double FillFocus = GetInt(TShapeOption.fillFocus, 0);

            bool InvertColors = false;
            if (FillFocus < -100) FillFocus = -100; //Focus on GDI+ must be between 0 and 1
            if (FillFocus > 100) FillFocus = 100;

            if (FillFocus < 0)
            {
                FillFocus += 100;
                InvertColors = true;
            }

            FillFocus = 1 - FillFocus / 100.0;

            if (FillAngle < 0)
            {
                FillAngle += 180;
                if (FillFocus > 0 && FillFocus < 1) InvertColors = !InvertColors; //there is a border case here
            }

            if (InvertColors)
            {
                SwapColors(ref FillOpacity, ref Color1, ref FillBackOpacity, ref Color2);
            }


            byte[] Blending = GetByteArrayProp(TShapeOption.fillShadeColors);

            bool RotateWithShape = true; //This value is located on an undocumented record (ID f122), It is not available on xls2000
            TFlipMode FlipMode = TFlipMode.XY;
            TDrawingGradientDef GradientDef;
            TDrawingGradientStop[] GradientStops;

            switch (FillType)
            {
                case TFillType.ShadeShape:
                    {
                        GradientDef = new TDrawingPathGradient(null, TPathShadeType.Rect);
                        GradientStops = GetBlending(Workbook, Blending, FillOpacity, FillBackOpacity, FillFocus, InvertColors, Color2, Color1);
                        break;
                    }

                case TFillType.ShadeCenter:
                    {
                        GradientDef = new TDrawingPathGradient(null, TPathShadeType.Rect);
                        double ToLeft = Get1616(TShapeOption.fillToLeft, 0);
                        double ToTop = Get1616(TShapeOption.fillToTop, 0);
 
                        GradientStops = GetBlending(Workbook, Blending, FillOpacity, FillBackOpacity, FillFocus, InvertColors, Color2, Color1);
                        break;
                     }
                default:
                    {
                        GradientDef = new TDrawingLinearGradient(FillAngle, false);
                        GradientStops = GetBlending(Workbook, Blending, FillOpacity, FillBackOpacity, FillFocus, InvertColors, Color2, Color1);
                        break;
                    }
            }

            return new TShapeFill(HasFill, new TGradientFill(null, RotateWithShape, FlipMode, GradientStops, GradientDef));
        }

        /// <summary>
        /// Sadly, we cannot change the focus on ColorBlended gradients with SetSigmaShape.
        /// </summary>
        /// <returns></returns>
        private static TDrawingGradientStop[] ChangeFocus(TDrawingGradientStop[] Blend, double FillFocus)
        {
            int k = Blend.Length;
            TDrawingGradientStop[] Result = new TDrawingGradientStop[k * 2 - 1];

            for (int i = 0; i < k; i++)
            {
                Result[i] = new TDrawingGradientStop((1 - Blend[k - i - 1].Position) * FillFocus, Blend[k - 1 - i].Color);
                Result[k * 2 - 2 - i] = new TDrawingGradientStop(FillFocus + Blend[k - 1 - i].Position * (1 - FillFocus), Blend[k - 1 - i].Color);
            }
            return Result;
        }

        private static void SwapColors(ref double FillOpacity, ref TDrawingColor Color1, ref double FillBackOpacity, ref TDrawingColor Color2)
        {
            TDrawingColor Tmp = Color1;
            Color1 = Color2;
            Color2 = Tmp;

            double t = FillBackOpacity;
            FillBackOpacity = FillOpacity;
            FillOpacity = t;
        }

        private TDrawingGradientStop[] GetBlending(ExcelFile Workbook, byte[] Values, double OpacityOrg, double OpacityDest, double FillFocus, bool InvertColors, TDrawingColor Color1, TDrawingColor Color2)
        {
            int BlendCount = 0;
            if (Values != null && Values.Length > 0)
            {
                BlendCount = (Values.Length - 6) / 8;
            }

            TDrawingGradientStop[] Result = new TDrawingGradientStop[BlendCount];

            int k = 6;
            for (int i = 0; i < BlendCount; i++)
            {
                TDrawingColor BlendCol = ColorFromLong(BitConverter.ToUInt32(Values, k), OpacityDest + (OpacityOrg - OpacityDest) * i / (BlendCount - 1f), Workbook, true, 0);
                double BlendPos = 1 - BitConverter.ToUInt16(Values, k + 5) / 256.0;
                Result[BlendCount - 1 - i] = new TDrawingGradientStop(BlendPos, BlendCol);
                k += 8;
            }

           EnsureMinimumAndMaximum(Color1, Color2, ref Result);

            if (InvertColors) //must be done before changefocus.
            {
                InvertColorBlend(Result);
            }

            if (FillFocus > 0)
            {
                Result = ChangeFocus(Result, FillFocus);
            }

            return Result;
        }

        internal static void EnsureMinimumAndMaximum(TDrawingColor Color1, TDrawingColor Color2, ref TDrawingGradientStop[] BlendedColors)
        {
            int Blends = BlendedColors.Length;
            bool NeedsZero = Blends <= 0 || BlendedColors[0].Position != 0;
            bool NeedsOne = Blends <= 0 || BlendedColors[Blends - 1].Position != 1;

            if (NeedsZero || NeedsOne)
            {
                if (NeedsZero) Blends++;
                if (NeedsOne) Blends++;
                TDrawingGradientStop[] Result = new TDrawingGradientStop[Blends];
                int k1 = 0;
                if (NeedsZero) { Result[0] = new TDrawingGradientStop(0, Color1); k1++; }
                for (int i = 0; i < BlendedColors.Length; i++)
                {
                    Result[k1 + i] = new TDrawingGradientStop(BlendedColors[i].Position, BlendedColors[i].Color);
                }
                if (NeedsOne) { Result[Blends - 1] = new TDrawingGradientStop(1, Color2); }
                BlendedColors = Result;
            }

        }

        internal static void InvertColorBlend(TDrawingGradientStop[] Result)
        {
            for (int i = 0; i < (Result.Length + 1) / 2; i++)
            {
                int n = Result.Length - 1 - i;

                TDrawingGradientStop Tmp = Result[n];

                Result[n] = new TDrawingGradientStop(1 - Result[i].Position, Result[i].Color);
                Result[i] = new TDrawingGradientStop(1 - Tmp.Position, Tmp.Color);
            }
        }

        private TShapeFill GetSolidBrush(ExcelFile xls)
        {
            double FillOpacity = Get1616(TShapeOption.fillOpacity, 1);
            TDrawingColor bkColor = ColorFromLong(TShapeOption.fillColor, 0xffffffff, FillOpacity, xls, true, 0);
            return new TShapeFill(HasFill, new TSolidFill(bkColor));
        }

        private TDrawingColor ColorFromLong(TShapeOption FillColor, UInt32 DefFill, double Opacity, ExcelFile Workbook, bool bkg, int RecursionLevel)
		{
            uint cl = GetUInt(FillColor, DefFill);
            return ColorFromLong(cl, Opacity, Workbook, bkg, RecursionLevel);
		}

        private TDrawingColor ColorFromLong(uint cl, double Opacity, ExcelFile Workbook, bool bkg, int RecursionLevel)
        {
            TDrawingColor Result = InternalColorFromLong(cl, Workbook, bkg, RecursionLevel);
            if (Opacity < 1)
            {
                Result = TDrawingColor.AddTransform(Result, new TColorTransform[] { new TColorTransform(TColorTransformType.Alpha, Opacity) });
            }

            return Result;
        }

		private TDrawingColor InternalColorFromLong(uint cl, ExcelFile Workbook, bool bkg, int RecursionLevel)
		{
            int ColorFlags = ((int)((cl & 0xFF000000) >> 24));

			if ((ColorFlags & 1) != 0)
				return TDrawingColor.FromRgb((byte)(cl & 0xFF), (byte)((cl & 0xFF00) >>8), (byte)((cl & 0xFF0000) >>16));

			if ((ColorFlags & 8) != 0) //Externally Indexed
			{
				int cp = (int)(cl & 0xFFFF) - 7;
                if (cp > 0 && cp <= Workbook.ColorPaletteCount)
                {
                    return Workbook.GetColorPalette(cp);
                }
                else if (cp == Workbook.ColorPaletteCount + 1 && bkg) //no fill
                {
                    return TDrawingColor.Transparent;
                }
                else  //auto
                    if (bkg)
                        return TDrawingColor.FromRgb(255, 255, 255);
                    else
                        return TDrawingColor.FromRgb(0, 0, 0);
			}

			if ((ColorFlags & 16) != 0) // SysIndex Color
			{
				return GetSysIndexColor(Workbook, cl, RecursionLevel);
			}

            return TDrawingColor.FromRgb((byte)(cl & 0xFF), (byte)((cl & 0xFF00) >> 8), (byte)((cl & 0xFF0000) >> 16));
        }

        private TDrawingColor GetSysIndexColor(ExcelFile Workbook, long cl, int RecursionLevel)
        {
            if (RecursionLevel > 20) return TDrawingColor.Transparent;
            if (cl <= 20) return GetSystemColor(cl);

            TDrawingColor Result = TDrawingColor.Transparent;

            if ((cl & 0x8000) != 0) Result = TDrawingColor.FromRgb(0x80, 0x80, 0x80);   // Make the color gray (before the above!)

            int p = (byte)(cl & 0xFF0000) >> 16; // Parameter used as above

            switch (cl & 0xFF)
            {
                case 0xF0:    // Use the fillColor property
                    Result = ColorFromLong(TShapeOption.fillColor, 0x0,
                        Get1616(TShapeOption.fillOpacity, 1), Workbook, true, RecursionLevel + 1);
                    break;

                case 0xF1: // Use the line color only if there is a line
                    if (!HasLine)
                    {
                        Result = ColorFromLong(TShapeOption.fillColor, 0x0,
                            Get1616(TShapeOption.fillOpacity, 1), Workbook, true, RecursionLevel + 1);
                    }
                    else
                    {
                        Result = ColorFromLong(TShapeOption.lineColor, 0x808080,
                            Get1616(TShapeOption.lineOpacity, 1), Workbook, true, RecursionLevel + 1);
                    }
                    break;

                case 0xF2: // Use the lineColor property
                    Result = ColorFromLong(TShapeOption.lineColor, 0x0,
                            Get1616(TShapeOption.lineOpacity, 1), Workbook, true, RecursionLevel + 1);
                    break;

                case 0xF3:    // Use the shadow color
                    Result = ColorFromLong(TShapeOption.shadowColor, 0x0,
                            Get1616(TShapeOption.shadowOpacity, 1), Workbook, true, RecursionLevel + 1);
                    break;

                case 0xF4:   // Use this color (only valid as described below)
                    break;
                case 0xF5:   // Use the fillBackColor property
                    Result = ColorFromLong(TShapeOption.fillBackColor, 0x0,
                            Get1616(TShapeOption.fillBackOpacity, 1), Workbook, true, RecursionLevel + 1);
                    break;

                case 0xF6:   // Use the lineBackColor property
                    Result = ColorFromLong(TShapeOption.lineBackColor, 0x0,
                        1, Workbook, true, RecursionLevel + 1);
                    break;
                case 0xF7:    // Use the fillColor unless no fill and line
                    Result = ColorFromLong(TShapeOption.fillColor, 0x0,
                            Get1616(TShapeOption.fillOpacity, 1), Workbook, true, RecursionLevel + 1);
                    break;

                case 0xFF: Result = TDrawingColor.FromRgb((byte)((cl & 0xFF00) >> 8), (byte)((cl & 0xFF0000) >> 16), (byte)((cl & 0xFF000000) >> 24)); break;  // Extract the color index
            }

            byte R, G, B;
            switch (cl & 0x0F00)   // function to apply
            {
                case 0x0100: Result = DarkColor(Workbook, Result, p / 255.0); break;   // Darken color by parameter/255
                case 0x0200: Result = LightColor(Workbook, Result, p / 255.0); break;   // Lighten color by parameter/255
                case 0x0300: Result = ChangeColor(Workbook, Result, p); break;        // Add grey level RGB(param,param,param)
                case 0x0400: Result = ChangeColor(Workbook, Result, -p); break;      // Subtract grey level RGB(p,p,p)
                case 0x0500:
                    Result.GetComponents(Workbook, out R, out G, out B);
                    Result = ChangeColor(Workbook, TDrawingColor.FromRgb((byte)p, (byte)p, (byte)p), R, G, B);  // Subtract from grey level RGB(p,p,p)
                    break; 

                // In the following "black" means maximum component value, white minimum.
                //   The operation is per component, to guarantee white combine with msocolorGray
                case 0x0600:
                    Result.GetComponents(Workbook, out R, out G, out B);
                    R = R < p ? (byte)0 : (byte)255;   // Black if < uParam, else white (>=)
                    G = G < p ? (byte)0 : (byte)255;   // Black if < uParam, else white (>=)
                    B = B < p ? (byte)0 : (byte)255;   // Black if < uParam, else white (>=)
                    Result = TDrawingColor.FromRgb(R, G, B); //no alpha yet
                    break;
            }

            if ((cl & 0x4000) != 0)
            {
                Result.GetComponents(Workbook, out R, out G, out B);
                unchecked
                {
                    Result = TDrawingColor.FromRgb((byte)(R ^ 128), (byte)(G ^ 128), (byte)(B ^ 128));   // Invert by toggling the top bit
                }
            }
            if ((cl & 0x2000) != 0)
            {
                Result.GetComponents(Workbook, out R, out G, out B);
                Result = TDrawingColor.FromRgb((byte)(255 - R), (byte)(255 - G), (byte)(255 - B));   // Invert color (at the *end*)
            }
           
            return Result; 
        }

        private static TDrawingColor GetSystemColor(long cl)
        {
            switch (cl)
            {
                case 0: return TDrawingColor.FromSystem(TSystemColor.BtnFace); //COLOR_BTNFACE
                case 1: return TDrawingColor.FromSystem(TSystemColor.WindowText);  // COLOR_WINDOWTEXT
                case 2: return TDrawingColor.FromSystem(TSystemColor.Menu); // COLOR_MENU
                case 3: return TDrawingColor.FromSystem(TSystemColor.Highlight);           // COLOR_HIGHLIGHT
                case 4: return TDrawingColor.FromSystem(TSystemColor.HighlightText);       // COLOR_HIGHLIGHTTEXT
                case 5: return TDrawingColor.FromSystem(TSystemColor.CaptionText);        // COLOR_CAPTIONTEXT
                case 6: return TDrawingColor.FromSystem(TSystemColor.ActiveCaption);       // COLOR_ACTIVECAPTION
                case 7: return TDrawingColor.FromSystem(TSystemColor.BtnHighlight);     // COLOR_BTNHIGHLIGHT
                case 8: return TDrawingColor.FromSystem(TSystemColor.BtnShadow);        // COLOR_BTNSHADOW
                case 9: return TDrawingColor.FromSystem(TSystemColor.BtnText);          // COLOR_BTNTEXT
                case 10: return TDrawingColor.FromSystem(TSystemColor.GrayText);            // COLOR_GRAYTEXT
                case 11: return TDrawingColor.FromSystem(TSystemColor.InactiveCaption);     // COLOR_INACTIVECAPTION
                case 12: return TDrawingColor.FromSystem(TSystemColor.InactiveCaptionText); // COLOR_INACTIVECAPTIONTEXT
                case 13: return TDrawingColor.FromSystem(TSystemColor.InfoBk);      // COLOR_INFOBK
                case 14: return TDrawingColor.FromSystem(TSystemColor.InfoText);            // COLOR_INFOTEXT
                case 15: return TDrawingColor.FromSystem(TSystemColor.MenuText);            // COLOR_MENUTEXT
                case 16: return TDrawingColor.FromSystem(TSystemColor.ScrollBar);           // COLOR_SCROLLBAR
                case 17: return TDrawingColor.FromSystem(TSystemColor.Window);              // COLOR_WINDOW
                case 18: return TDrawingColor.FromSystem(TSystemColor.WindowFrame);         // COLOR_WINDOWFRAME
                case 19: return TDrawingColor.FromSystem(TSystemColor.Light3d);             // COLOR_3DLIGHT
                case 20: return TDrawingColor.Transparent;                 // Count of system colors
            }

            return ColorUtil.Empty;
        }

        private static byte ChangeComponent(int R, int p)
        {
            int Result = R + p;
            if (Result < 0) Result = 0; if (Result > 255) Result = 255;
            return (byte)Result;
        }

        private static TDrawingColor ChangeColor(ExcelFile Workbook, TDrawingColor Clr, int p)
        {
            byte R, G, B;
            Clr.GetComponents(Workbook, out R, out G, out B);
            return TDrawingColor.FromRgb(ChangeComponent(R, p), ChangeComponent(G, p), ChangeComponent(B, p));
        }

        private static TDrawingColor ChangeColor(ExcelFile Workbook, TDrawingColor Clr, int p1, int p2, int p3)
        {
            byte R, G, B;
            Clr.GetComponents(Workbook, out R, out G, out B);
            return TDrawingColor.FromRgb(ChangeComponent(R, p1), ChangeComponent(G, p2), ChangeComponent(B, p3));
        }

        private static byte Light(int C, double p)
        {
            int Result = (int)(C * p);
            if (Result < 0) return 0;
            if (Result > 255) return 255;
            return (byte)Result;
        }
        
        private static TDrawingColor DarkColor(ExcelFile Workbook, TDrawingColor Clr, double p)
        {
            byte R, G, B;
            Clr.GetComponents(Workbook, out R, out G, out B);
            return TDrawingColor.FromRgb(Light(R, p), Light(G, p), Light(B, p)); //no alpha yet, if it had it we would have to add a transform.
        }

        private static TDrawingColor LightColor(ExcelFile Workbook, TDrawingColor Clr, double p)
        {
            byte R, G, B;
            Clr.GetComponents(Workbook, out R, out G, out B);
            return TDrawingColor.FromRgb((byte)(255 - Light(255 - R, p)), (byte)(255 - Light(255 - G, p)), (byte)(255 - Light(255 - B, p)));
        }
        #endregion

        #region Get Line
        public TShapeLine GetLine(ExcelFile xls)
        {
            return GetLine(xls, false, false, false, 0);
        }
        public TShapeLine GetLine(ExcelFile xls, bool Obscured, bool DefaultDrawLine, bool Shadow, int ShadowPass)
        {
            if (ShapeLine != null) return ShapeLine.Clone();
            uint lc = GetUInt(TShapeOption.lineColor, 0xff000000);

            if (!Obscured && !GetBool(TShapeOption.fNoLineDrawDash, (0x01 << 3), DefaultDrawLine)) return null; //fline

            int LineWidth = GetInt(TShapeOption.lineWidth, TLineStyle.DefaultWidth);
            
            TDrawingColor aColor = ColorFromLong(lc, 1, xls, false, 0);
            if (!Obscured &&  aColor == TDrawingColor.Transparent) return null;
            if (Shadow)
            {
                TShapeOption ShadowColor = ShadowPass <= 1 ? TShapeOption.shadowColor : TShapeOption.shadowHighlight;
                aColor = ColorFromLong(GetUInt(ShadowColor, 0x808080),
                    Get1616(TShapeOption.shadowOpacity, 1), xls, true, 0);
                if (aColor == TDrawingColor.Transparent) return null;
            }

            int Dashing = GetInt(TShapeOption.lineDashing, 0);
            TShapeLine Result = new TShapeLine(true, new TLineStyle(new TSolidFill(aColor), LineWidth, null, null, null, (TLineDashing) Dashing, null, null, null));

            return Result;
        }
        #endregion

        internal void SetShapeGeom(TShapeGeom aShapeGeom)
        {
            ShapeGeom = aShapeGeom;
        }

        internal void SetShapeFont(TShapeFont aShapeFont)
        {
            ShapeFont = aShapeFont;
        }

        internal void SetShapeEffects(TShapeEffects aShapeEffects)
        {
            ShapeEffects = aShapeEffects;
        }

        internal void SetEffectProps(TEffectProperties aEffectProps)
        {
            EffectProps = aEffectProps;
        }

        internal void SetTextExt(TDrawingRichString aTextExt)
        {
            TextExt = aTextExt;
        }

        internal void SetHlinkClick(TDrawingHyperlink aHlink)
        {
            HLinkClick = aHlink;
        }

        internal void SetHlinkHover(TDrawingHyperlink aHlink)
        {
            HLinkHover = aHlink;
        }

        internal void SetBodyPr(TBodyPr aBodyPr)
        {
            BodyPr = aBodyPr;
            if (BodyPr == null) return;

            SetIntProperty(TShapeOption.dxTextLeft, (int)BodyPr.l.Emu, 91440);
            SetIntProperty(TShapeOption.dyTextTop, (int)BodyPr.t.Emu, 91440 / 2);
            SetIntProperty(TShapeOption.dxTextRight, (int)BodyPr.r.Emu, 91440);
            SetIntProperty(TShapeOption.dyTextBottom, (int)BodyPr.b.Emu, 91440 / 2);

            SetBoolProperty(TShapeOption.fFitTextToShape, ~(uint)0x0808, true, false); //Clear ignore text margins.

        }

        internal void SetLstStyle(string aLstStyle)
        {
            LstStyle = aLstStyle;
        }

        internal TShapeGeom GetFinalShapeGeom()
        {
#if (!FRAMEWORK30 || COMPACTFRAMEWORK)
            return null;
#else
            if (ShapeGeom == null) return null; //only xlsx.

            string sp;
            if (!TDrawingPresetGeom.FromShapeType.TryGetValue(ShapeType, out sp)) return ShapeGeom.Clone();

            TShapeGeom DefaultGeom;
            if (!TShapePresets.ShapeList.TryGetValue(sp, out DefaultGeom)) return ShapeGeom.Clone();

            TShapeGeom NewGeom = DefaultGeom.Clone(ShapeGeom.AvList);

            return NewGeom;
#endif
        }


#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
        internal TShapeGeom GetGeom()
        {
            if (ShapeGeom != null) return ShapeGeom;
            string sp;
            if (TDrawingPresetGeom.FromShapeType.TryGetValue(ShapeType, out sp)) return CreateGeom(sp);
            return null;

        }

        private TShapeGeom CreateGeom(string sp)
        {
            TShapeGeom Result = new TShapeGeom(sp);
            TShapeGeom DefaultGeom;
            if (!TShapePresets.ShapeList.TryGetValue(sp, out DefaultGeom)) return Result;

            object o;
            for (int i = 0; i < DefaultGeom.AvList.Count; i++)
            {
                string AdjName = DefaultGeom.AvList[i].Name;
                TShapeOption so;
                if (!TShapePresets.GetBiff8Adjust(AdjName, out so)) continue;
                if (!Records.TryGetValue(so, out o)) continue;
                Result.AvList.Add(new TShapeGuide(AdjName, new TShapeVal(TShapePresets.ConvertAdjustFromBiff8(Convert.ToInt32(o), sp))));
            }
            return Result;
        }
#endif

        internal TShapeFont GetRawShapeFont()
        {
            return ShapeFont;
        }

        internal TShapeEffects GetRawShapeEffects()
        {
            return ShapeEffects;
        }

        internal TEffectProperties GetRawEffectProps()
        {
            return EffectProps;
        }

        internal TDrawingRichString GetRawTextExt()
        {
            return TextExt;
        }

        internal TBodyPr GetRawBodyPr()
        {
            return BodyPr;
        }

        internal string GetRawLstStyle()
        {
            return LstStyle;
        }

        internal TDrawingHyperlink GetRawHLinkClick()
        {
            return HLinkClick;
        }

        internal TDrawingHyperlink GetRawHLinkHover()
        {
            return HLinkHover;
        }

        internal TShapeFill GetRawShapeFill()
        {
            return ShapeFill;
        }

        internal TShapeLine GetRawShapeLine()
        {
            return ShapeLine;
        }

        internal bool UseTextMargins()
        {
            return !GetBool(TShapeOption.fFitTextToShape, 0x08, false); 
        }
    }
}

