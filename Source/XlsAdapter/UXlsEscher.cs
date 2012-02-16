using System;
using System.Diagnostics;
using System.IO;
using FlexCel.Core;
using System.Globalization;
using System.Collections.Generic;

#if (WPF)
using System.Windows.Media;
#else
using System.Drawing;
#endif

namespace FlexCel.XlsAdapter
{
    /// <summary>
    /// Base for Excel drawing records. It has 2 main descendants: TDrawingGroupRecord for the globals, and TDrawingRecord for the sheets.
    /// </summary>
    internal class TXlsEscherRecord: TxBaseRecord
    {
        internal TXlsEscherRecord(int aId, byte[] aData): base(aId, aData){}
    }

    /// <summary>
    /// Base for Drawing Group record and HeaderImg.
    /// </summary>
    internal class TBaseDrawingGroupRecord: TXlsEscherRecord
    {
        internal TBaseDrawingGroupRecord(int aId, byte[] aData): base(aId, aData){}
    }

    /// <summary>
    /// Drawing Group record. There is only one of them on a xls file, and it is on the global section.
    /// </summary>
    internal class TDrawingGroupRecord: TBaseDrawingGroupRecord
    {
        internal TDrawingGroupRecord(int aId, byte[] aData): base(aId, aData){}

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.DrawingGroup.LoadFromStream(WorkbookLoader.RecordLoader, this, false);
        }
    }

    /// <summary>
    /// Base for Drawing record and HeaderImg.
    /// </summary>
    internal class TBaseDrawingRecord: TXlsEscherRecord
    {
        internal TBaseDrawingRecord(int aId, byte[] aData): base(aId, aData){}
    }

    /// <summary>
    /// Drawing record. It holds a MSDrawing inside and can spawn many records.
    /// </summary>
    internal class TDrawingRecord: TBaseDrawingRecord
    {
        internal TDrawingRecord(int aId, byte[] aData): base(aId, aData){}

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            ws.Drawing.LoadFromStream(RecordLoader, ws.FWorkbookGlobals, this, ws.SheetType == TSheetType.Chart); 
        }

    }

    /// <summary>
    /// A parsed Drawing Group. As the on-file structure of drawings is so horrible, we won't attempt
    /// to keep it. We will parse to a better mem struct, and then save it back to the serialized file.
    /// </summary>
    internal class TDrawingGroup
    {
        private TEscherContainerRecord FDggContainer;
        private TEscherDwgGroupCache FRecordCache;
        private xlr MainRecordId;  //This might be a MSODRAWINGGROUP or a HEADERIMG 
        private byte[] ExtraData;

        internal TDrawingGroup(xlr aMainRecordId, int ExtraDataSize)
        {
            FRecordCache = new TEscherDwgGroupCache();
            MainRecordId = aMainRecordId;
            ExtraData = new byte[ExtraDataSize];
            if (ExtraData.Length>12)
            {
                ExtraData[0]=0x66; ExtraData[1]=0x08; ExtraData[12]=0x02;
            }
        }
    
        internal TEscherDwgGroupCache RecordCache { get {return FRecordCache;}}

        internal void Clear() 
        { 
            FDggContainer = null;
            FRecordCache= new TEscherDwgGroupCache();
        }

        internal void ClearIfEmpty()
        {
            if (FDggContainer != null && FRecordCache.Dgg != null && FRecordCache.Dgg.IsEmpty()) Clear();
        }

        internal bool HeaderImage{get{return ExtraData.Length>0;}}

        internal void LoadFromStream(TBaseRecordLoader RecordLoader, TBaseDrawingGroupRecord First, bool ChartCoords)
        {
            TEscherDwgCache DwgCache = new TEscherDwgCache(0, null, null, null, null, null, null, null, null, null);
            
            if (FDggContainer!=null) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
            int aPos=0;
            TxBaseRecord MyRecord = First; 
            TxBaseRecord CurrentRecord = First;

            byte[] ExtraData2 = new byte[ExtraData.Length];
            BitOps.ReadMem(ref MyRecord, ref aPos, ExtraData2);

            TEscherRecordHeader EscherHeader= new TEscherRecordHeader();
            BitOps.ReadMem(ref MyRecord, ref aPos, EscherHeader.Data);

            FDggContainer= new TEscherContainerRecord(EscherHeader, FRecordCache, DwgCache ,null);
            while (!FDggContainer.Loaded())
            {
                if ((MyRecord.Continue==null) && (aPos==MyRecord.DataSize))
                {
                    CurrentRecord=RecordLoader.LoadRecord(true) as TxBaseRecord;
                    if (MainRecordId == xlr.HEADERIMG)
                        CurrentRecord = CurrentRecord as THeaderImageGroupRecord;
                    else
                        CurrentRecord = CurrentRecord as TDrawingGroupRecord;

                    if (MyRecord==null) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
                    MyRecord=CurrentRecord;
                    aPos=ExtraData.Length;
                }

                FDggContainer.Load(ref MyRecord, ref aPos, HeaderImage, ChartCoords);

            } //while

        }

        internal void SaveToStream (IDataStream DataStream, TSaveData SaveData)
        {
            if (FDggContainer==null) return;
            TBreakList BreakList= new TBreakList(DataStream.Position, ExtraData);
            int NextPos=0;
            int RealSize=ExtraData.Length;
            int NewDwg= (int)MainRecordId;
            int Rid = MainRecordId == xlr.HEADERIMG? (int)xlr.HEADERIMG: (int)xlr.CONTINUE;
            FDggContainer.SplitRecords(ref NextPos, ref RealSize, ref NewDwg, BreakList, Rid, ExtraData.Length);
            BreakList.Add(0, NextPos);
            FDggContainer.SaveToStream(DataStream, SaveData, BreakList); 
        }

        internal long TotalSize()
        {
            if (FDggContainer==null) return 0;
            int NextPos=0;
            int RealSize=ExtraData.Length;
            int NewDwg= (int)MainRecordId;
            FDggContainer.SplitRecords(ref NextPos, ref RealSize, ref NewDwg, null, (int) xlr.CONTINUE, ExtraData.Length);
            return RealSize;
        }

        internal void AddDwg()
        {
            //not neded now. if (FRecordCache.Dgg!=null) FRecordCache.Dgg.DwgSaved++;
        }

        /// <summary>
        /// This function will test that there is a dwggroup, and if it doesn't, will create it.
        /// Should be called when adding new drawings on empty files.
        /// </summary>
        internal void EnsureDwgGroup()
        {
            TEscherDwgCache DwgCache = new TEscherDwgCache(0, null, null, null, null, null, null, null, null, null);
            if (FDggContainer == null)  // there is already a DwgGroup
            {
                //DggContainer
                TEscherRecordHeader EscherHeader = new TEscherRecordHeader();
                EscherHeader.Pre=0xF;
                EscherHeader.Id=(int)Msofbt.DggContainer;
                EscherHeader.Size=0;
                FDggContainer=new TEscherContainerRecord(EscherHeader, FRecordCache, DwgCache ,null);
                FDggContainer.LoadedDataSize=(int)EscherHeader.Size;
            }

            if (FDggContainer.FindRec<TEscherDggRecord>()==null)
            {
                //Dgg
                TEscherDggRecord FDgg=new TEscherDggRecord(FRecordCache, DwgCache ,FDggContainer);
                FDggContainer.ContainedRecords.Add(FDgg);
            }

            if (FDggContainer.FindRec<TEscherBStoreRecord>()==null)
            {
                // BStoreContainer
                TEscherRecordHeader EscherHeader = new TEscherRecordHeader();
                EscherHeader.Pre=0x2F;
                EscherHeader.Id=(int)Msofbt.BstoreContainer;
                EscherHeader.Size=0;
                TEscherBStoreRecord BStoreContainer= new TEscherBStoreRecord(EscherHeader, FRecordCache, DwgCache ,FDggContainer);
                BStoreContainer.LoadedDataSize=(int)EscherHeader.Size;
                FDggContainer.ContainedRecords.Add(BStoreContainer);
            }

            if (FDggContainer.FindRec<TEscherOPTRecord>()==null)
            {
                //OPT
                TEscherOPTRecord OPTRec= TEscherOPTRecord.CreateFromDataGlobalGroup(FRecordCache, DwgCache, FDggContainer);  //groupcreatefromdata
                FDggContainer.ContainedRecords.Add(OPTRec);
            }

            if (FDggContainer.FindRec<TEscherSplitMenuRecord>()==null)
            {
                //SplitMenuColors
                TEscherSplitMenuRecord SplitMenu= new TEscherSplitMenuRecord(FRecordCache, DwgCache, FDggContainer);
                FDggContainer.ContainedRecords.Add(SplitMenu);
            }

        }
    }

    /// <summary>
    /// A parsed Drawing. Again, as the on-file structure of drawings is so horrible, we won't attempt
    /// to keep it. We will parse to a better mem struct, and then save it back to the serialized file.
    /// </summary>
    internal class TDrawing
    {
        private TEscherContainerRecord FDgContainer;
        private TEscherDwgCache FRecordCache;
        private TDrawingGroup FDrawingGroup;
        private byte[] ExtraData;
        private xlr MainRecordId;

        private readonly byte[] EmptyBmp= 
            {
                0x28, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 
                0x00, 0x01, 0x00, 0x18, 0x00, 0x00, 0x00, 0x00, 0x00, 0x04, 0x00, 
                0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 
                0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 
                0x00, 0x00, 0xFF, 0xFF, 0xFF, 0x00
            };

        internal TDrawing(TDrawingGroup aDrawingGroup, xlr aMainRecordId, int ExtraDataSize)
        {
            FDrawingGroup=aDrawingGroup;
            FRecordCache= new TEscherDwgCache();
            MainRecordId = aMainRecordId;
            ExtraData = new byte[ExtraDataSize];
            if (ExtraData.Length>12)
            {
                ExtraData[0]=0x66; ExtraData[1]=0x08; ExtraData[12]=0x01;
            }
        }

        internal void Destroy()
        {
            Clear();
        }

        internal void Clear() 
        { 
            if (FDgContainer!=null) {FDgContainer.Destroy(); FDgContainer=null;}
            if (FDrawingGroup != null) FDrawingGroup.ClearIfEmpty(); 
            //Order is important... Cache should be freed after DgContainer
            FRecordCache.AnchorList=null;
            FRecordCache.Obj=null;
            FRecordCache.Shape=null;
            FRecordCache.Blip=null;
            FRecordCache.RadioButtons = null;
            FRecordCache= new TEscherDwgCache();
            FRecordCache.OptByName = new TEscherOptByNameCache();
        }

        internal bool HeaderImage{get{return ExtraData.Length>0;}}

        internal void CopyFrom(int RowOfs, int ColOfs, TDrawing aDrawing, TSheetInfo SheetInfo)
        {
            Clear();
            FRecordCache.MaxObjId = 0;
            FRecordCache.Dg=null; FRecordCache.Patriarch=null;

            if (aDrawing.FRecordCache.AnchorList!=null) 
            {
                FRecordCache.AnchorList= new TEscherAnchorCache();
                FRecordCache.Obj= new TEscherObjCache();
                FRecordCache.Shape= new TEscherShapeCache();
                FRecordCache.Blip= new TEscherOPTCache();
                FRecordCache.RadioButtons = new TRadioButtonCache();
                FRecordCache.OptByName = new TEscherOptByNameCache();
            }

            if (aDrawing.FDgContainer==null) {if (FDgContainer!=null) {FDgContainer.Destroy();FDgContainer=null;}} 
            else
            {
                SheetInfo.IncCopiedGen();
                FDgContainer= (TEscherContainerRecord) TEscherContainerRecord.Clone(aDrawing.FDgContainer, RowOfs, ColOfs, FRecordCache, FDrawingGroup.RecordCache, SheetInfo);
                FRecordCache.Shape.Sort(); // only here the values are loaded...
                if (FRecordCache.Solver!=null) FRecordCache.Solver.CheckMax(aDrawing.FRecordCache.Solver.MaxRuleId);

                FDrawingGroup.AddDwg();
            }
            //MADE: change cache

            Array.Copy(aDrawing.ExtraData, 0, ExtraData, 0, ExtraData.Length);
        }

        internal bool HasExternRefs()
        {
            if (FRecordCache==null || FRecordCache.Obj==null) return false;
            int Count = FRecordCache.Obj.Count;

            for (int i=0; i< Count;i++)
                if (FRecordCache.Obj[i].HasExternRefs()) return true;
            return false;
        }

        internal void ArrangeCopyRange(int RowOfs, int ColOfs, TSheetInfo SheetInfo)
        {
            if (FRecordCache==null || FRecordCache.Obj==null) return;
            int Count = FRecordCache.Obj.Count;

            for (int i=0; i< Count;i++)
                FRecordCache.Obj[i].ArrangeCopyRange(RowOfs, ColOfs, SheetInfo);
        }

        internal void CopyObjectsFrom(bool IncludeDontMoveAndResize, int RowOfs, int ColOfs, TDrawing aDrawing, TXlsCellRange CopyRange, TSheetInfo SheetInfo)
        {
            if (aDrawing==null || aDrawing.FRecordCache==null || aDrawing.FRecordCache.AnchorList==null) return;
            
            FDrawingGroup.EnsureDwgGroup();

            if ((FDgContainer==null) || (FRecordCache.AnchorList== null))  //no drawings on this sheet
                CreateBasicDrawingInfo();
            if (FRecordCache.Patriarch==null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);

            SheetInfo.IncCopiedGen();
            int aCount = aDrawing.FRecordCache.AnchorList.Count;
            for (int i = 0; i < aCount; i++)
            {
                TEscherBaseClientAnchorRecord al = aDrawing.FRecordCache.AnchorList[i];
                bool IsInRange;
                if (al.AllowCopy(IncludeDontMoveAndResize, CopyRange.Top, CopyRange.Bottom, CopyRange.Left, CopyRange.Right, out IsInRange))
                {
                    TEscherRecord Img = al.FindRoot();
                    if (Img != null) // && (aDrawing == this || !Img.HasExternRefs()))
                    {
                        Img = TEscherRecord.Clone(Img, RowOfs, ColOfs, FRecordCache, FDrawingGroup.RecordCache, SheetInfo);
                        Img.Parent = FRecordCache.Patriarch;
                        FRecordCache.Patriarch.FContainedRecords.Add(Img);
                    }
                }

            }
            FRecordCache.Shape.Sort(); // only here the values are loaded...
            if (FRecordCache.Solver!=null) FRecordCache.Solver.ArrangeCopyRange(RowOfs, ColOfs, SheetInfo);
        }

        internal void LoadFromStream(TBaseRecordLoader RecordLoader, TWorkbookGlobals Globals, TBaseDrawingRecord First, bool ChartCoords)
        {
            Debug.Assert (FDrawingGroup!=null,"DrawingGroup can't be null");
            if (FDgContainer!=null) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);

            FRecordCache.MaxObjId = 0;
            FRecordCache.Dg=null; FRecordCache.Patriarch=null; FRecordCache.Solver=null;
            FRecordCache.AnchorList= new TEscherAnchorCache();
            FRecordCache.Obj= new TEscherObjCache();
            FRecordCache.Shape= new TEscherShapeCache();
            FRecordCache.Blip= new TEscherOPTCache();
            FRecordCache.RadioButtons = new TRadioButtonCache();
            FRecordCache.OptByName = new TEscherOptByNameCache();
            
            int aPos=0;
            TxBaseRecord MyRecord = First; 
            TxBaseRecord CurrentRecord = First;

            byte[] ExtraData2 = new byte[ExtraData.Length];
            BitOps.ReadMem(ref MyRecord, ref aPos, ExtraData2);

            TEscherRecordHeader EscherHeader= new TEscherRecordHeader();
            BitOps.ReadMem(ref MyRecord, ref aPos, EscherHeader.Data);

            FDgContainer= new TEscherContainerRecord(EscherHeader, FDrawingGroup.RecordCache, FRecordCache ,null);
            TClientType ClientType=TClientType.Null;
            while ((!FDgContainer.Loaded()) || FDgContainer.WaitingClientData(ref ClientType))
            {
                if (!(FDgContainer.WaitingClientData(ref ClientType)))
                {
                    if ((MyRecord.Continue==null) && (aPos==MyRecord.DataSize))
                    {
                        CurrentRecord=RecordLoader.LoadRecord(false) as TxBaseRecord;
                        if (MainRecordId == xlr.HEADERIMG)
                            CurrentRecord = CurrentRecord as THeaderImageRecord;
                        else
                            CurrentRecord = CurrentRecord as TDrawingRecord;

                        if (MyRecord ==null) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
                        MyRecord=CurrentRecord;
                        aPos=0;
                    }

                    FDgContainer.Load(ref MyRecord, ref aPos, HeaderImage, ChartCoords);
                }
                else
                {
                    if (! ((MyRecord.Continue==null) && (aPos==MyRecord.DataSize))) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
                    TBaseRecord R = RecordLoader.LoadRecord(false);
                    if (R==null) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
                    
                    //Not very clean, (as in the delphi version using virtual constructors) but it works and avoids using Activator. (which needs more permissions to run)
                    TBaseClientData FClientData=null;
                    if ((R is TObjRecord)  && (ClientType == TClientType.TMsObj)) FClientData = new TMsObj();
                    if ((R is TTXORecord)  && (ClientType == TClientType.TTXO)) FClientData=new TTXO(RecordLoader.FontList);
                    if (FClientData != null) 
                    {
                        FClientData.LoadFromStream(RecordLoader, Globals, R);
                        FDgContainer.AssignClientData(FClientData, false);
                        if (FClientData.RemainingData!=null)
                        {
                            TxBaseRecord CdRecord=(TxBaseRecord) FClientData.RemainingData; //we dont have to free this
                            int CdPos=0;
                            FDgContainer.Load(ref CdRecord, ref CdPos, HeaderImage, ChartCoords);
                        }
                    } 
                    else XlsMessages.ThrowException(XlsErr.ErrInvalidDrawing);
                }
            } //while

            FRecordCache.Shape.Sort(); // only here the values are loaded...
            if (FRecordCache.Solver != null) FRecordCache.Solver.FixPointers();
        }

        internal void SaveToStream (IDataStream DataStream, TSaveData SaveData)
        {
            if (FDgContainer==null) return;
            TBreakList BreakList= new TBreakList(DataStream.Position, ExtraData);
            int NextPos=0;
            int RealSize=ExtraData.Length;
            int NewDwg= (int)MainRecordId;
            int Rid = MainRecordId == xlr.HEADERIMG? (int)xlr.HEADERIMG: (int)xlr.CONTINUE;
            FDgContainer.SplitRecords(ref NextPos, ref RealSize, ref NewDwg, BreakList, Rid, ExtraData.Length);
            BreakList.Add(0, NextPos);
            FDgContainer.SaveToStream(DataStream, SaveData, BreakList);
        }

        internal long TotalSize()
        {
            if (FDgContainer==null) return 0;
            int NextPos=0;
            int RealSize=ExtraData.Length;
            int NewDwg= (int)MainRecordId;
            FDgContainer.SplitRecords(ref NextPos, ref RealSize, ref NewDwg, null, (int) xlr.CONTINUE, ExtraData.Length);
            return RealSize;
        }

        internal void ArrangeInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
        {
            if ((FRecordCache.AnchorList != null) && (SheetInfo.SourceFormulaSheet == SheetInfo.InsSheet))
                FRecordCache.AnchorList.ArrangeInsertRange(CellRange, aRowCount, aColCount, SheetInfo, false);
            if (FRecordCache.Obj != null)
                FRecordCache.Obj.ArrangeInsertRange(CellRange, aRowCount, aColCount, SheetInfo, false);
        }

        internal void ArrangeMoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
        {
            if (FRecordCache.Obj != null)
                FRecordCache.Obj.ArrangeMoveRange(CellRange, NewRow, NewCol, SheetInfo);
        }

        internal void ArrangeCopySheet(TSheetInfo SheetInfo)
        {
            if (FRecordCache.Obj != null) 
                FRecordCache.Obj.ArrangeCopySheet(SheetInfo);
        }

        internal void InsertAndCopyRange(bool IncludeDontMoveAndResize, TXlsCellRange SourceRange, TFlxInsertMode InsertMode, int DestRow, int DestCol, int aRowCount, int aColCount, TRangeCopyMode CopyMode, 
            TSheetInfo SheetInfo)
        {
            if ((FDgContainer==null) || (FRecordCache.AnchorList== null)) return;  //no drawings on this sheet

            int myFirstRow=0;int myLastRow=0;
            if (DestRow>SourceRange.Top)
            {
                myFirstRow=SourceRange.Top; myLastRow=SourceRange.Bottom;
            } 
            else
            {
                myFirstRow=SourceRange.Top+aRowCount*SourceRange.RowCount;
                myLastRow=SourceRange.Bottom+aRowCount*SourceRange.RowCount;
            }

            int myFirstCol=0;int myLastCol=0;
            if (DestCol>SourceRange.Left)
            {
                myFirstCol=SourceRange.Left; myLastCol=SourceRange.Right;
            } 
            else
            {
                myFirstCol=SourceRange.Left+aColCount*SourceRange.ColCount;
                myLastCol=SourceRange.Right+aColCount*SourceRange.ColCount;
            }

            int RowOfs = DestRow-SourceRange.Top; 
            int ColOfs = DestCol-SourceRange.Left;

            //Insert cells
            ArrangeInsertRange(SourceRange.OffsetForIns(DestRow, DestCol, InsertMode), aRowCount, aColCount, SheetInfo);

            //Copy the images

            if (CopyMode != TRangeCopyMode.None && CopyMode != TRangeCopyMode.OnlyFormulasAndNoObjects)
            {
                UInt32List ShapesToCopy = new UInt32List();
                if (aRowCount > 0 || aColCount > 0)  //cache the shapes we want to copy
                {
                    GetObjectsInRange(IncludeDontMoveAndResize, myFirstRow, myLastRow, myFirstCol, myLastCol, SheetInfo.ObjectsInRange, ShapesToCopy);
                }

                if (ShapesToCopy.Count == 0) return;

                int aCount = ShapesToCopy.Count;

                //First the rows...
                int myDestRow = DestRow;
                for (int k = 0; k < aRowCount; k++) //Order must be this one. If not reports would fail.
                {
                    SheetInfo.IncCopiedGen();
                    for (int i = aCount - 1; i >= 0; i--)
                    {
                        TEscherRecord R = (TEscherRecord)FRecordCache.AnchorList[(int)ShapesToCopy[i]].CopyDwg(myDestRow - myFirstRow, DestCol - SourceRange.Left, SheetInfo);
                        AddCopyToCache(SheetInfo, k, R);
                    }
                    myDestRow += SourceRange.RowCount;
                    if (FRecordCache.Solver != null) FRecordCache.Solver.ArrangeCopyRange(RowOfs, ColOfs, SheetInfo);
                }

                //Now the columns... as we already copied the rows, now we will should make an array of images. But we currently support only copying columns xor rows
                int myDestCol = DestCol;
                for (int k = 0; k < aColCount; k++)
                {
                    SheetInfo.IncCopiedGen();
                    for (int i = aCount - 1; i >= 0; i--)
                    {
                        TEscherRecord R = (TEscherRecord)FRecordCache.AnchorList[(int)ShapesToCopy[i]].CopyDwg(DestRow - SourceRange.Top, myDestCol - myFirstCol, SheetInfo);
                        AddCopyToCache(SheetInfo, k, R);
                    }
                    myDestCol += SourceRange.ColCount;
                    if (FRecordCache.Solver != null) FRecordCache.Solver.ArrangeCopyRange(RowOfs, ColOfs, SheetInfo);
                }
            }
        }

        private static void AddCopyToCache(TSheetInfo SheetInfo, int k, TEscherRecord R)
        {
            if (SheetInfo.ObjectsInRange != null && SheetInfo.ObjectsInRange.IncludeCopies)
            {
                SheetInfo.ObjectsInRange.AddCopy(k, (R.Parent.FindRec<TEscherSpRecord>() as TEscherSpRecord).ShapeId);
            }
        }

        internal void GetObjectsInRange(bool IncludeDontMoveAndResize, int myFirstRow, int myLastRow, int myFirstCol, int myLastCol, TExcelObjectList ObjectsInRange, UInt32List ShapesToCopy)
        {
            if (FRecordCache.AnchorList == null) return;
            for (int i = FRecordCache.AnchorList.Count - 1; i >= 0; i--)
            {
                bool IsInRange;
                if (FRecordCache.AnchorList[i].AllowCopy(IncludeDontMoveAndResize, myFirstRow, myLastRow, myFirstCol, myLastCol, out IsInRange))
                {
                   if (ShapesToCopy != null) ShapesToCopy.Add((uint)i);
                }

                long ShpId;
                TEscherSpRecord sp = FRecordCache.AnchorList[i].Parent.FindRec<TEscherSpRecord>() as TEscherSpRecord;
                if (sp == null) FlxMessages.ThrowException(FlxErr.ErrInternal);
                ShpId = sp.ShapeId;
                if (IsInRange && ObjectsInRange != null) ObjectsInRange.Add(i + 1, ShpId); //id is 1 based
            }

            if (ObjectsInRange != null) ObjectsInRange.Reverse(); //we want the list in forward order.
        }
        
        internal void DeleteRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
        {
            if (FRecordCache.AnchorList==null) return;
            for (int i=FRecordCache.AnchorList.Count-1;i>=0;i--)
                if (FRecordCache.AnchorList[i].AllowDelete(CellRange.Top, CellRange.Bottom, CellRange.Left, CellRange.Right))
                {
                    if (FRecordCache.Patriarch==null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
                    FRecordCache.Patriarch.ContainedRecords.Remove(FRecordCache.AnchorList[i].FindRoot());
                }

            ArrangeInsertRange(CellRange, -aRowCount, -aColCount, SheetInfo);
        }        
 
        internal void MoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
        {
            if (FRecordCache.AnchorList==null) return;
            for (int i=FRecordCache.AnchorList.Count-1;i>=0;i--)
            {
                bool IsInRange;
                if (FRecordCache.AnchorList[i].AllowCopy(false, CellRange.Top, CellRange.Bottom, CellRange.Left, CellRange.Right, out IsInRange))
                {
                    if (FRecordCache.Patriarch==null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
                    ((TEscherImageAnchorRecord)FRecordCache.AnchorList[i]).ArrangeMoveRange(NewRow - CellRange.Top, NewCol - CellRange.Left, SheetInfo);
                }
                else  //only delete source if it is not moving inside the range.
                    if (FRecordCache.AnchorList[i].AllowDelete(NewRow, NewRow + CellRange.RowCount - 1, NewCol, NewCol + CellRange.ColCount - 1))
                {
                    if (FRecordCache.Patriarch==null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
                    FRecordCache.Patriarch.ContainedRecords.Remove(FRecordCache.AnchorList[i].FindRoot());
                }

            }

            ArrangeMoveRange(CellRange, NewRow, NewCol, SheetInfo);
        }        
 
        internal void ClearRange(TXlsCellRange CellRange)
        {
            if (FRecordCache.AnchorList==null) return;
            for (int i=FRecordCache.AnchorList.Count-1;i>=0;i--)
                if (FRecordCache.AnchorList[i].AllowDelete(CellRange.Top, CellRange.Bottom, CellRange.Left, CellRange.Right))
                {
                    if (FRecordCache.Patriarch==null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
                    FRecordCache.Patriarch.ContainedRecords.Remove(FRecordCache.AnchorList[i].FindRoot());
                }
        }   
     
        #region Named Ranges
        internal void UpdateDeletedRanges(TDeletedRanges DeletedRanges)
        {
            if (FRecordCache.Obj != null)
                FRecordCache.Obj.UpdateDeletedRanges(DeletedRanges);
        }
        #endregion

        internal int FindShapeIdIndex(long ShapeId)
        {
            if (FRecordCache.AnchorList == null) return -2;

            for (int i = 0; i < FRecordCache.AnchorList.Count; i++)
            {
                TEscherSpRecord cd = FRecordCache.AnchorList.VisibleItem(i).Parent.FindRec<TEscherSpRecord>() as TEscherSpRecord;
                if (cd.ShapeId == ShapeId) return i;
            }
            
            return -2;
        }

        internal TShapeType ShapeType(TEscherOPTRecord opt)
        {
            return opt.ShapeType;
        }


        internal TEscherClientDataRecord FindObjId(int ObjId)
        {
            if (FRecordCache.Obj == null) return null;
            for (int i = 0; i < FRecordCache.Obj.Count; i++)
                if (FRecordCache.Obj[i].ObjId == ObjId)
                    return FRecordCache.Obj[i];
            return null;
        }



        internal int DrawingCount
        {
            get
            {
                if (FRecordCache.Blip!=null) return FRecordCache.Blip.Count; else return 0;
            }
        }

        internal long ReferencesCount(int Index)
        {
            return FRecordCache.Blip[Index].ReferencesCount();
        }

        internal void AssignDrawing(int Index, byte[] Data, TXlsImgType DataType, bool IsObjIndex, string ObjectPath)
        {
            TEscherOPTRecord Opt = null;
            if (IsObjIndex)
            {
                TEscherContainerRecord R = TEscherContainerRecord.Navigate(FRecordCache, Index, ObjectPath); 
                if (R == null) return; Opt = (TEscherOPTRecord)R.FindRec<TEscherOPTRecord>();
            }
            else
            {
                Opt = FRecordCache.Blip[Index];
            }

            AssignDrawing(Opt, Data, DataType);
        }

        internal void AssignDrawing(TEscherOPTRecord Opt, byte[] Data, TXlsImgType DataType)
        {
            if (Opt == null) return;
            if (Data.Length == 0) ClearImage(Opt);  //XP crashes with a 0 byte image.
            else Opt.ReplaceImg(Data, DataType);
        }

        internal void AssignHeaderOrFooterDrawing(THeaderAndFooterKind Kind, THeaderAndFooterPos Section,
            byte[] Data, TXlsImgType DataType, THeaderOrFooterImageProperties Properties)
        {
            string SectionStr = GetSectionString(Kind, Section);
            if (FRecordCache.Blip==null)
            {
                AddImage(null, Data, DataType, Properties, true, SectionStr, null, 0, null, false);
                return;
            }

            for (int i=FRecordCache.Blip.Count-1; i>=0;i--)
            {
                TEscherOPTRecord blip = FRecordCache.Blip[i];
                if (String.Equals(blip.ShapeName, SectionStr, StringComparison.InvariantCultureIgnoreCase))
                {
                    if (Data.Length==0) ClearImage(i);  //XP crashes with a 0 byte image.
                    else blip.ReplaceImg(Data, DataType);
                    blip.SetAnchor(Properties.Anchor);
                    return;
                }
            }

            AddImage(null, Data, DataType, Properties, true, SectionStr, null, 0, null, false);
        }

        internal void GetDrawingFromStream(int Index, string ObjectPath, Stream Data, ref TXlsImgType DataType, bool UsesObjectIndex)
        {
            if (FRecordCache.Patriarch==null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
            Debug.Assert(UsesObjectIndex || Index < FRecordCache.Blip.Count, "Index out of range");

            TEscherOPTRecord Opt;
            if (UsesObjectIndex)
            {
                TEscherContainerRecord R = TEscherContainerRecord.Navigate(FRecordCache, Index, ObjectPath);
                if (R == null) return;
                Opt = (TEscherOPTRecord)R.FindRec<TEscherOPTRecord>();
            }
            else
            {
                Opt = FRecordCache.Blip[Index];
            }

            if (Opt==null) return;

            Opt.GetImageFromStream(Data, ref DataType);
        }

        private static string GetSString(THeaderAndFooterPos Section)
        {
            switch (Section)
            {
                case THeaderAndFooterPos.HeaderLeft: return "LH";
                case THeaderAndFooterPos.HeaderCenter: return "CH";
                case THeaderAndFooterPos.HeaderRight: return "RH";
                case THeaderAndFooterPos.FooterLeft: return "LF";
                case THeaderAndFooterPos.FooterCenter: return "CF";
                case THeaderAndFooterPos.FooterRight: return "RF";
            }
            return string.Empty;
        }

        private static string GetSectionString(THeaderAndFooterKind Kind, THeaderAndFooterPos Section)
        {
            switch (Kind)
            {
                case THeaderAndFooterKind.Default:
                    return GetSString(Section);
                case THeaderAndFooterKind.FirstPage:
                    return GetSString(Section) + "FIRST";
                case THeaderAndFooterKind.EvenPages:
                    return GetSString(Section) + "EVEN";
            }

            return GetSString(Section);
        }

        internal void GetDrawingFromStream(THeaderAndFooterKind Kind, THeaderAndFooterPos Section, Stream Data, ref TXlsImgType DataType)
        {
            string SectionStr = GetSectionString(Kind, Section);
            if (FRecordCache == null || FRecordCache.Blip == null) return;
            for (int i=FRecordCache.Blip.Count-1; i>=0;i--)
            {
                TEscherOPTRecord blip = FRecordCache.Blip[i];
                if (String.Equals(blip.ShapeName, SectionStr, StringComparison.InvariantCultureIgnoreCase))
                {
                    blip.GetImageFromStream(Data, ref DataType);
                    return;
                }
            }
        }


        internal TImageProperties GetImageProperties(int Index)
        {
            TEscherOPTRecord blip = FRecordCache.Blip[Index];

            int[] ParentCoords = null;

            TMsObj o = blip.GetObj();

            return new TImageProperties(blip.GetAnchor(ref ParentCoords), blip.FileName, blip.ShapeName, blip.CropArea, blip.TransparentColor, blip.Brightness, blip.Contrast, blip.Gamma,
                 o.IsLocked, o.IsPrintable, o.IsPublished, o.IsDisabled, o.IsDefaultSize, o.IsAutoFill, o.IsAutoLine, blip.AltText, null,
                 blip.PreferRelativeSize, blip.LockAspectRatio, blip.BiLevel, blip.Grayscale,
                 true);
        }

        internal THeaderOrFooterImageProperties GetHeaderOrFooterImageProperties(THeaderAndFooterKind Kind, THeaderAndFooterPos Section)
        {
            string SectionStr = GetSectionString(Kind, Section);
            if (FRecordCache == null || FRecordCache.Blip == null) return null;
            for (int i = FRecordCache.Blip.Count - 1; i >= 0; i--)
            {
                TEscherOPTRecord blip = FRecordCache.Blip[i];
                if (String.Equals(blip.ShapeName, SectionStr, StringComparison.InvariantCultureIgnoreCase))
                {
                    //Header and footers don't have MsObj
                    return HeaderOrFooterProps(blip);
                }
            }
            return null;
        }

        internal TEscherOPTRecord GetBlip(int index)
        {
            if (FRecordCache == null || FRecordCache.Blip == null) return null;
            return FRecordCache.Blip[index]; 
        }

        private static THeaderOrFooterImageProperties HeaderOrFooterProps(TEscherOPTRecord blip)
        {
            return new THeaderOrFooterImageProperties(blip.GetHeaderAnchor(), blip.FileName, blip.CropArea, blip.TransparentColor, blip.Brightness, blip.Contrast, blip.Gamma,
                true, true, false, false, false, false, false, null, 
                blip.PreferRelativeSize, blip.LockAspectRatio, blip.BiLevel, blip.Grayscale,
                true);
        }

        internal void SetImageProperties(int Index, TImageProperties Properties, TSheet sSheet)
        {
            FRecordCache.Blip[Index].SetAnchor(Properties.Anchor, sSheet);
            FRecordCache.Blip[Index].ShapeName = Properties.ShapeName;
            FRecordCache.Blip[Index].SetStringProperty(TShapeOption.wzDescription, Properties.AltText);
        }

        internal string DrawingName(int Index)
        {
            Debug.Assert(Index<FRecordCache.Blip.Count,"Index out of range");
            return FRecordCache.Blip[Index].ShapeName;
        }

        internal int GetDrawingRow(int Index)
        {
            Debug.Assert(Index<FRecordCache.Blip.Count,"Index out of range");
            return FRecordCache.Blip[Index].Row;
        }


        internal void DeleteImage(int Index)
        {
            if (FRecordCache.Blip==null) return;
            if (FRecordCache.Patriarch==null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
            FRecordCache.Patriarch.ContainedRecords.Remove(FRecordCache.Blip[Index].FindRoot());
        }

        internal void RemoveAutoFilter()
        {
            if (FRecordCache.Obj == null) return;
            if (FRecordCache.Patriarch == null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
            for (int i = FRecordCache.Obj.Count - 1; i >= 0; i--) 
            {
                TEscherClientDataRecord obj0 = FRecordCache.Obj[i];
                if (obj0.ClientData == null) continue;
                TMsObj ClientData = obj0.ClientData as TMsObj;
                if (ClientData == null) continue;
                if (ClientData.IsAutoFilter) 
                {
                    FRecordCache.Patriarch.ContainedRecords.Remove(obj0.FindRoot());
                }
            }
        }


        internal void DeleteHeaderOrFooterImage(THeaderAndFooterKind Kind, THeaderAndFooterPos Section)
        {
            if (FRecordCache.Blip == null) return;
            if (FRecordCache.Patriarch == null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
            string SectionStr = GetSectionString(Kind, Section);
            for (int i = FRecordCache.Blip.Count - 1; i >= 0; i--)
            {
                TEscherOPTRecord blip = FRecordCache.Blip[i];
                if (String.Equals(blip.ShapeName, SectionStr, StringComparison.InvariantCultureIgnoreCase))
                {
                    DeleteImage(i);
                    return;
                }
            }
        }

        internal void ClearImage(int ImageIndex)
        {
            FRecordCache.Blip[ImageIndex].ReplaceImg(EmptyBmp, TXlsImgType.Bmp);
        }

        internal void ClearImage(TEscherOPTRecord Opt)
        {
            Opt.ReplaceImg(EmptyBmp, TXlsImgType.Bmp);
        }

        #region z order
        internal void SendToBack(int Index)
        {
            int VisibleIndex=FRecordCache.AnchorList.VisibleIndex(Index);
            //swap visible and not visible stuff.
            for (int i= VisibleIndex; i>0;i--)
                SwapObjects(i,i-1);                                                          
        }

        internal void BringToFront(int Index)
        {
            int VisibleIndex=FRecordCache.AnchorList.VisibleIndex(Index);
            //swap visible and not visible stuff.
            for (int i= VisibleIndex; i<FRecordCache.AnchorList.Count-1;i++)
                SwapObjects(i,i+1);                                                          
        }

        private void SwapObjects(int ObjectIndex1, int ObjectIndex2)
        {
            //We won't Change BLIP cache, so the image will remain at the same pos. We will only change its position on the big list.
            if (ObjectIndex1<0 || ObjectIndex2<0) return;
            if (ObjectIndex1>=FRecordCache.AnchorList.Count || ObjectIndex2>=FRecordCache.AnchorList.Count) return;
            TEscherRecord o1=FRecordCache.AnchorList[ObjectIndex1].FindRoot();
            TEscherRecord o2=FRecordCache.AnchorList[ObjectIndex2].FindRoot();
            if (o1==null || o2==null) return;
            FRecordCache.AnchorList.Swap(ObjectIndex2, ObjectIndex1);
            FRecordCache.Patriarch.ContainedRecords.Swap(FRecordCache.Patriarch.ContainedRecords.IndexOf(o1), FRecordCache.Patriarch.ContainedRecords.IndexOf(o2));
        }
        
        internal void SendBack(int Index)
        {
            int VisibleIndex1=FRecordCache.AnchorList.VisibleIndex(Index);
            int VisibleIndex2=FRecordCache.AnchorList.VisibleIndex(Index-1);
            SwapObjects(VisibleIndex1, VisibleIndex2);
        }
        
        internal void SendForward(int Index)
        {
            int VisibleIndex1=FRecordCache.AnchorList.VisibleIndex(Index);
            int VisibleIndex2=FRecordCache.AnchorList.VisibleIndex(Index+1);
            SwapObjects(VisibleIndex1, VisibleIndex2);
        }
        #endregion

        internal int ObjectCount
        {
            get
            {
                if (FRecordCache==null || FRecordCache.AnchorList==null) return 0;
                return FRecordCache.AnchorList.VisibleCount();
            }
        }

        internal int LegacyCount
        {
            get
            {
                int Result = 0;
                if (FRecordCache == null || FRecordCache.AnchorList == null) return 0;
                for (int k = 0; k < FRecordCache.AnchorList.VisibleCount(); k++)
                {
                    TEscherContainerRecord Root = FRecordCache.AnchorList.VisibleItem(k).Parent;
                    for (int i = 0; i < Root.ContainedRecords.Count; i++)
                    {
                        TEscherClientDataRecord Cd = Root.ContainedRecords[i] as TEscherClientDataRecord;
                        if (Cd != null)
                        {
                            TMsObj MsObj = Cd.ClientData as TMsObj;
                            if (MsObj != null)
                            {
                                TObjectType ot = (TObjectType)MsObj.ObjType;
                                if (ot == TObjectType.CheckBox || ot == TObjectType.Button
                                    || ot == TObjectType.OptionButton || ot == TObjectType.GroupBox || IsComboButNotSpecial(MsObj, ot)
                                    || ot == TObjectType.ListBox
                                    || ot == TObjectType.Label
                                    || ot == TObjectType.Spinner
                                    || ot == TObjectType.ScrollBar) Result++;
                                break;
                            }
                        }
                    }
                }
                return Result;
            }

        }

        private bool IsComboButNotSpecial(TMsObj MsObj, TObjectType ot)
        {
            if (MsObj.IsSpecialDropdown) return false;
            return ot == TObjectType.ComboBox;
        }


        internal int ObjectIndexToImageIndex(int ObjectIndex)
        {
            TEscherRecord Root = FRecordCache.AnchorList.VisibleItem(ObjectIndex).FindRoot();
            for (int i = 0; i < FRecordCache.Blip.Count; i++)
                if (FRecordCache.Blip[i].FindRoot() == Root)
                    return i;
            return -1;
        }

        internal int ImageIndexToObjectIndex(int ImageIndex)
        {
            TEscherRecord Root=FRecordCache.Blip[ImageIndex].FindRoot();
            int aCount = ObjectCount;
            for (int i=0; i< aCount;i++)
                if (FRecordCache.AnchorList.VisibleItem(i).FindRoot()==Root)
                    return i;
            return -1;
        }

        internal string ObjectName(int Index)
        {
            TEscherOPTRecord Opt = GetOPT(Index);
            if (Opt==null) return String.Empty; else return Opt.ShapeName;
        }

        internal long ShapeId(int Index)
        {
            TEscherOPTRecord Opt = GetOPT(Index);
            if (Opt == null) return -1; else return Opt.ShapeId();
        }

        internal TEscherOPTRecord GetOPT(int Index)
        {
            if (FRecordCache.Patriarch == null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
            Debug.Assert(Index < ObjectCount, "Index out of range");
            return (TEscherOPTRecord)FRecordCache.AnchorList.VisibleItem(Index).Parent.FindRec<TEscherOPTRecord>();
        }

        internal TEscherClientDataRecord GetDwg(int Index)
        {
            if (FRecordCache.Patriarch == null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
            Debug.Assert(Index < ObjectCount, "Index out of range");
            return (TEscherClientDataRecord)FRecordCache.AnchorList.VisibleItem(Index).FindRoot().FindRec<TEscherClientDataRecord>();
        }

        internal TEscherClientTextBoxRecord GetClientTextBox(int Index)
        {
            if (FRecordCache.Patriarch == null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
            Debug.Assert(Index < ObjectCount, "Index out of range");
            return (TEscherClientTextBoxRecord)FRecordCache.AnchorList.VisibleItem(Index).FindRoot().FindRec<TEscherClientTextBoxRecord>();
        }


        internal TClientAnchor GetObjectAnchor(int Index)
        {
            if (FRecordCache.Patriarch==null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
            Debug.Assert(Index<ObjectCount,"Index out of range");
            TEscherOPTRecord Opt = (TEscherOPTRecord)FRecordCache.AnchorList.VisibleItem(Index).Parent.FindRec<TEscherOPTRecord>();
            int[] ParentCoords = null;
            if (Opt == null) return new TClientAnchor(); else return Opt.GetAnchor(ref ParentCoords);
        }

        internal static TClientAnchor GetObjectAnchorAbsolute(TEscherOPTRecord Opt,  IRowColSize Workbook)
        {
            if (Opt == null) return new TClientAnchor();

            List<TEscherOPTRecord> OptChain = new List<TEscherOPTRecord>();
            OptChain.Add(Opt);
            TEscherSpContainerRecord OriginalParentSp = Opt.Parent as TEscherSpContainerRecord;
            TEscherSpgrContainerRecord ParentSpgr = OriginalParentSp.Parent as TEscherSpgrContainerRecord;
            while (ParentSpgr != null)
            {
                TEscherSpContainerRecord ParentSp = ParentSpgr.FindRec<TEscherSpContainerRecord>() as TEscherSpContainerRecord;
                if (ParentSp.Opt != null) OptChain.Add(ParentSp.Opt);
                ParentSpgr = ParentSpgr.Parent as TEscherSpgrContainerRecord;
            }

            int[] ParentCoords = null;
            TClientAnchor GrpAnchor = OptChain[OptChain.Count - 1].GetAnchor(ref ParentCoords);


            if (OptChain.Count <= 1) return GrpAnchor; //most common case, not grouped.

            double h = 0; double w = 0;
            double x1 = 0; double y1 = 0;
            GrpAnchor.CalcImageCoords(ref h, ref w, Workbook);
            for(int i = OptChain.Count - 2; i >= 0; i--)
            {
                TClientAnchor TmpGrpAnchor = OptChain[i].GetAnchor(ref ParentCoords);
                if (TmpGrpAnchor.ChildAnchor != null)
                {
                    x1 = x1 + w * TmpGrpAnchor.ChildAnchor.Dx1;
                    y1 = y1 + h * TmpGrpAnchor.ChildAnchor.Dy1;
                    w = w * (TmpGrpAnchor.ChildAnchor.Dx2 - TmpGrpAnchor.ChildAnchor.Dx1);
                    h = h * (TmpGrpAnchor.ChildAnchor.Dy2 - TmpGrpAnchor.ChildAnchor.Dy1);
                }
            }
            return new TClientAnchor(GrpAnchor.AnchorType, GrpAnchor.Row1, GrpAnchor.Dy1Pix(Workbook) + (int)y1, GrpAnchor.Col1, GrpAnchor.Dx1Pix(Workbook) + (int)x1, (int)h, (int)w, Workbook);
        }



        internal void SetObjectAnchor(int Index, string ObjectPath, TClientAnchor Anchor, TSheet sSheet)
        {
            if (FRecordCache.Patriarch==null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
            TEscherContainerRecord Root = TEscherContainerRecord.Navigate(FRecordCache, Index, ObjectPath);
            if (Root == null)
            {
                XlsMessages.ThrowException(XlsErr.ErrObjectNotFound, ObjectPath);
            }

            TEscherOPTRecord Opt = (TEscherOPTRecord)Root.FindRec<TEscherOPTRecord>();
            if (Opt!=null) Opt.SetAnchor(Anchor, sSheet);
        }

        private TShapeProperties FindProps(TEscherContainerRecord Root, string ObjectPathAbsolute, bool GetOptions, int[] ParentSpgrCoords, out int[] NewParentSpgrCoords)
        {
            NewParentSpgrCoords = ParentSpgrCoords;
            TShapeProperties Result = new TShapeProperties();
            Result.ObjectPathAbsolute = ObjectPathAbsolute;
            if (Root != null)
            {
                for (int i = 0; i < Root.ContainedRecords.Count; i++)
                {
                    TEscherOPTRecord Opt = Root.ContainedRecords[i] as TEscherOPTRecord;
                    if (Opt != null)
                    {
                        Result.Anchor = Opt.GetAnchor(ref ParentSpgrCoords);
                        NewParentSpgrCoords = ParentSpgrCoords;
                        Result.ShapeName = Opt.ShapeName;
                        Result.Visible = Opt.Visible;
                        Result.ShapeGeom = Opt.GetFinalShapeGeom();
                        Result.ShapeThemeFont = Opt.GetRawShapeFont();
                        if (GetOptions) Result.ShapeOptions = Opt.ShapeOptions();
                        continue;
                    }

                    TEscherSpRecord Sp = Root.ContainedRecords[i] as TEscherSpRecord;
                    if (Sp != null)
                    {
                        Result.ShapeId = Sp.ShapeId;
                        Result.ShapeType = Sp.ShapeType;
                        Result.FlipH = (Sp.Flags & 0x40) != 0;
                        Result.FlipV = (Sp.Flags & 0x80) != 0;
                        continue;
                    }

                    TEscherClientDataRecord Cd = Root.ContainedRecords[i] as TEscherClientDataRecord;
                    if (Cd != null)
                    {
                        TTXO Txo = Cd.ClientData as TTXO;
                        if (Txo != null)
                        {
                            Result.Text = Txo.GetText();
                            Result.TextFlags = Txo.OptionFlags;
                            Result.TextRotation = Txo.Rotation;
                        }

                        TMsObj MsObj = Cd.ClientData as TMsObj;
                        if (MsObj != null)
                        {
                            Result.ObjectType = (TObjectType)MsObj.ObjType;
                            if (!MsObj.IsPrintable) Result.Print = false;  //Printable can also be a shape OPT.
                            Result.FIsActiveX = MsObj.HasPictFormula;
                            Result.FIsInternal = MsObj.IsAutoFilter || Result.ObjectType == TObjectType.Comment;
                        }
                        continue;
                    }

                    TEscherContainerRecord SpCr = Root.ContainedRecords[i] as TEscherContainerRecord;
                    if (SpCr != null)
                    {
                        int[] NewParentCoords;
                        Result.AddChild(FindProps(SpCr, ObjectPathAbsolute + i + FlxConsts.ObjectPathSeparator, GetOptions, ParentSpgrCoords, out NewParentCoords));
                        ParentSpgrCoords = NewParentCoords;
                        continue;
                    }

                    TEscherContainerRecord Cr = Root.ContainedRecords[i] as TEscherContainerRecord;
                    if (Cr != null)
                    {
                        int[] NewParentCoords; //We don't want to modify the original coords if it is not the first SP record.
                        Result.AddChild(FindProps(Cr, ObjectPathAbsolute + i + FlxConsts.ObjectPathSeparator, GetOptions, ParentSpgrCoords, out NewParentCoords));
                        continue;
                    }
                }
            }
            return Result;
        }

        internal TShapeProperties GetObjectProperties(int Index, bool GetOptions)
        {
            if (FRecordCache.Patriarch==null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
            Debug.Assert(Index < ObjectCount, "Index out of range");
            TEscherContainerRecord Root = FRecordCache.AnchorList.VisibleItem(Index).FindRoot() as TEscherContainerRecord;

#if (IMAGEDEBUG)
            DumpProps(Root, "");
#endif
            int[] RootCoords;
            return FindProps(Root, 
                FlxConsts.ObjectPathAbsolute + (Index + 1).ToString(CultureInfo.InvariantCulture) + FlxConsts.ObjectPathSeparator,
                GetOptions, null, out RootCoords);
        }

#if (IMAGEDEBUG)
        internal void DumpProps(TEscherContainerRecord r, string Indent)
        {
            Console.WriteLine(Indent + r.ToString());
            for(int i = 0; i < r.ContainedRecords.Count;i++)
            {
                TEscherRecord z = r.ContainedRecords[i];
                if (z is TEscherContainerRecord) DumpProps(z as TEscherContainerRecord, Indent + "       ");
                else
                {
                    Console.Write(Indent + "   *" + z.ToString());
                    if (z is TEscherSpgrRecord)
                    {
                        Console.Write("  -  " + (z as TEscherSpgrRecord).Bounds[0].ToString()+", "+(z as TEscherSpgrRecord).Bounds[1].ToString()+", "+(z as TEscherSpgrRecord).Bounds[2].ToString()+", "+(z as TEscherSpgrRecord).Bounds[3].ToString());
                    }

                    if (z is TEscherChildAnchorRecord)
                    {
                        Console.Write("  -  " + (z as TEscherChildAnchorRecord).Dx1.ToString() + ", " +
                            (z as TEscherChildAnchorRecord).Dy1.ToString() + ", " +(z as TEscherChildAnchorRecord).Dx2.ToString() +
                            ", " +(z as TEscherChildAnchorRecord).Dy2.ToString());
                    }
                    Console.WriteLine();
                }
            }
        }
#endif

        internal TShapeProperties GetObjectPropertiesByShapeId(long ShapeId, bool GetOptions)
        {
            if (FRecordCache.Patriarch == null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
            TEscherOPTRecord Opt = FRecordCache.OptByName.FindObjByShapeId(ShapeId);
            if (Opt == null) return null;
            TEscherContainerRecord Root = Opt.FindRoot() as TEscherContainerRecord;

            int[] RootCoords;
            return FindProps(Root,
                FlxConsts.ObjectPathAbsolute + (0).ToString(CultureInfo.InvariantCulture) + FlxConsts.ObjectPathSeparator,
                GetOptions, null, out RootCoords);
        }

        internal void SetObjectName(int Index, string ObjectPath, string Name)
        {
            if (FRecordCache.Patriarch == null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);

            TEscherContainerRecord R = TEscherContainerRecord.Navigate(FRecordCache, Index, ObjectPath);
            if (R == null) return;
            TEscherOPTRecord Opt = R.FindRec<TEscherOPTRecord>() as TEscherOPTRecord;
            if (Opt == null) return;
            Opt.ShapeName = Name;
        }

        internal void SetObjectText(int Index, string ObjectPath, TRichString Text, IFlexCelFontList aFontList, TObjectTextProperties TextProperties)
        {
            if (FRecordCache.Patriarch == null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);

            TEscherContainerRecord R = TEscherContainerRecord.Navigate(FRecordCache, Index, ObjectPath);
            SetObjectText(R, Text, aFontList, TextProperties);
        }

        internal void SetObjectTextExt(TEscherOPTRecord opt, TDrawingRichString Text, ExcelFile xls, TShapeProperties ShProps, TObjectTextProperties TextProperties)
        {
            TShapeFont shf = null;
            if (ShProps != null) shf = ShProps.ShapeThemeFont;
            SetObjectText(opt.Parent, Text.ToRichString(xls, shf), xls, TextProperties);
            opt.SetTextExt(Text); //Order is important, the line above will clear TextExt. 

        }

        internal void SetObjectText(TEscherContainerRecord R, TRichString Text, IFlexCelFontList aFontList, TObjectTextProperties TextProperties)
        {
            if (R == null) return;
            for (int i = 0; i < R.ContainedRecords.Count; i++)
            {
                Text = CheckObjectFont(Text, aFontList, R.ContainedRecords[i] as TEscherClientDataRecord);

                TEscherOPTRecord Opt1 = R.ContainedRecords[i] as TEscherOPTRecord;
                if (Opt1 != null)
                {
                    Opt1.SetTextExt(null);
                }

                TEscherClientTextBoxRecord TXORec = R.ContainedRecords[i] as TEscherClientTextBoxRecord;
                if (TXORec != null)
                {
                    TTXO Txo = TXORec.ClientData as TTXO;
                    if (Txo != null)
                    {
                        if (Text == null)
                        {
                            TEscherOPTRecord Opt = (TEscherOPTRecord)R.FindRec<TEscherOPTRecord>();
                            if (Opt != null) Opt.RemoveProperties((TShapeOption)128, (TShapeOption)255); //Text properties. Failing to remove them will crash excel.
                            R.ContainedRecords.Delete(i);
                            return;
                        }
                        Txo.SetText(Text);
                        SetObjectTextProps(TextProperties, TXORec);
                        return;
                    }
                }
            }
            if (Text==null) return;

            TEscherSpContainerRecord SpRec = R as TEscherSpContainerRecord;
            if (SpRec==null) return;

            TEscherClientTextBoxRecord TXORec2 = new TEscherClientTextBoxRecord(FDrawingGroup.RecordCache, FRecordCache, SpRec);  //.CreateFromData
            TTXO aTXO=new TTXO(aFontList, 0);  //.CreateFromData;
            aTXO.SetText(Text);

            TXORec2.AssignClientData(aTXO, true);
            SetObjectTextProps(TextProperties, TXORec2);
            SpRec.ContainedRecords.Add(TXORec2);
        }

        private static TRichString CheckObjectFont(TRichString Text, IFlexCelFontList aFontList, TEscherClientDataRecord obj)
        {
            if (Text != null && obj != null)
            {
                TMsObj msobj = obj.ClientData as TMsObj;
                if (msobj != null)
                {
                    TObjectType ot = (TObjectType)msobj.ObjType;
                    if (ot == TObjectType.CheckBox || ot == TObjectType.GroupBox || ot == TObjectType.OptionButton)
                    {
                        List<TRTFRun> Runs = new List<TRTFRun>(Text.GetRuns());
                        if (Runs.Count == 0 || Runs[0].FirstChar != 0)
                        {
                            TRTFRun rtf = new TRTFRun();
                            rtf.FirstChar = 0;
                            TFlxFont cbfont = new TFlxFont();
                            cbfont.Name = "Tahoma";
                            cbfont.Size20 = 160;
                            rtf.FontIndex = aFontList.AddFont(cbfont);
                            Runs.Insert(0, rtf);
                            Text = new TRichString(Text.Value, Runs, aFontList);
                        }
                    }
                }

            }
            return Text;
        }

        private static void SetObjectTextProps(TObjectTextProperties TextProperties, TEscherClientTextBoxRecord TXORec)
        {
            if (TextProperties != null)
            {
                TXORec.LockText = TextProperties.LockText;
                TXORec.HAlign = TextProperties.HAlignment;
                TXORec.VAlign = TextProperties.VAlignment;
                TXORec.TextRotation = TextProperties.TextRotation;
            }
        }
        

        internal void SetObjectProperty(int Index, string ObjectPath, TShapeOption Id, long Value)
        {
            if (FRecordCache.Patriarch==null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
            TEscherContainerRecord R = TEscherContainerRecord.Navigate(FRecordCache, Index, ObjectPath);
            if (R == null) return;
            TEscherOPTRecord Opt =(TEscherOPTRecord) R.FindRec<TEscherOPTRecord>();
            if (Opt!=null) 
            {
                Opt.SetLongProperty(Id, Value);
            }
        }

        internal void SetObjectProperty(int Index, string ObjectPath, TShapeOption Id, int PositionInSet, bool Value)
        {
            if (FRecordCache.Patriarch==null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
            TEscherContainerRecord R = TEscherContainerRecord.Navigate(FRecordCache, Index, ObjectPath);
            if (R == null) return;
            TEscherOPTRecord Opt =(TEscherOPTRecord) R.FindRec<TEscherOPTRecord>();
            if (Opt!=null) 
            {
                UInt32 lValue = 0;
                lValue = lValue | (UInt32)(1<<PositionInSet);
                if (Value)
                {
                    lValue |= (UInt32)(1<<(16 + PositionInSet)); //Property is set.
                    Opt.SetBoolProperty(Id, lValue, false, true);
                }
                else
                {
                    lValue = ~lValue;
                    Opt.SetBoolProperty(Id, lValue, true, false);

                    lValue = (UInt32)(1<<(16 + PositionInSet)); //Property is set.
                    Opt.SetBoolProperty(Id, lValue, false, true);
                }

            }
        }

        internal void SetObjectProperty(int Index, string ObjectPath, TShapeOption Id, string Value)
        {
            if (FRecordCache.Patriarch==null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
            TEscherContainerRecord R = TEscherContainerRecord.Navigate(FRecordCache, Index, ObjectPath);
            if (R==null) return;
            TEscherOPTRecord Opt =(TEscherOPTRecord) R.FindRec<TEscherOPTRecord>();
            if (Opt!=null) 
            {
                Opt.SetStringProperty(Id, Value);
            }
        }

        internal void SetObjectProperty(int Index, string ObjectPath, TShapeOption Id, THyperLink HLink)
        {
            if (FRecordCache.Patriarch==null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
            TEscherContainerRecord R = TEscherContainerRecord.Navigate(FRecordCache, Index, ObjectPath);
            if (R == null) return;
            TEscherOPTRecord Opt =(TEscherOPTRecord) R.FindRec<TEscherOPTRecord>();
            if (Opt!=null) 
            {
                Opt.SetHLinkProperty(Id, HLink);
            }
        }

        internal void SaveObjectCoords(TSheet sSheet)
        {
            if (FRecordCache.Patriarch==null) return;
            for (int i = FRecordCache.AnchorList.Count - 1; i >=0; i--)
            {
                FRecordCache.AnchorList[i].SaveObjectCoords(sSheet);
            }
        }

        internal void RestoreObjectCoords(TSheet dSheet)
        {
            if (FRecordCache.Patriarch==null) return;
            for (int i = FRecordCache.AnchorList.Count - 1; i >=0; i--)
            {
                FRecordCache.AnchorList[i].RestoreObjectCoords(dSheet);
            }
        }

        internal TFlxChart GetChart(int Index, string ObjectPath)
        {
            if (FRecordCache.Patriarch==null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
            TEscherContainerRecord R = TEscherContainerRecord.Navigate(FRecordCache, Index, ObjectPath);
            if (R == null) return null;
            if (R.ContainedRecords == null) return null;
            for (int i = R.ContainedRecords.Count - 1; i >= 0; i--)
            {
                TEscherClientDataRecord Cd = R.ContainedRecords[i] as TEscherClientDataRecord;
                if (Cd != null) return Cd.ClientData.Chart();
            }
            return null;
        }
        
        internal void DeleteObject(int Index)
        {
            if (FRecordCache.Patriarch==null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
            FRecordCache.Patriarch.ContainedRecords.Remove(FRecordCache.AnchorList.VisibleItem(Index).FindRoot());
        }

        internal void DeleteObject(int Index, string ObjectPath)
        {
            if (FRecordCache.Patriarch == null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
            TEscherContainerRecord Root = TEscherContainerRecord.Navigate(FRecordCache, Index, ObjectPath);
            if (Root == null)
            {
                XlsMessages.ThrowException(XlsErr.ErrObjectNotFound, ObjectPath);
            }
            TEscherContainerRecord RootParent = Root.Parent as TEscherContainerRecord;
            if (RootParent == null)
            {
                XlsMessages.ThrowException(XlsErr.ErrObjectNotFound, ObjectPath);
            }

            RootParent.ContainedRecords.Remove(Root);
        }

        private void CreateBasicDrawingInfo()
        {
            Debug.Assert (FDrawingGroup!=null,"DrawingGroup can't be null");
            FRecordCache.MaxObjId = 0;
            FRecordCache.Dg = null; FRecordCache.Patriarch = null; FRecordCache.Solver = null;

            FRecordCache.AnchorList=new TEscherAnchorCache();
            FRecordCache.Obj= new TEscherObjCache();
            FRecordCache.Shape= new TEscherShapeCache();
            FRecordCache.Blip= new TEscherOPTCache();
            FRecordCache.RadioButtons = new TRadioButtonCache();
            FRecordCache.OptByName = new TEscherOptByNameCache(); 

            TEscherRecordHeader EscherHeader=new TEscherRecordHeader();
            EscherHeader.Pre=0x0F;
            EscherHeader.Id=(int)Msofbt.DgContainer;
            EscherHeader.Size=0;
            FDgContainer= new TEscherContainerRecord(EscherHeader, FDrawingGroup.RecordCache, FRecordCache ,null);
            FDrawingGroup.AddDwg();

            //Add required records...
            int DgId; long FirstId;
            FDrawingGroup.RecordCache.Dgg.GetNewDgIdAndCluster(out DgId, out FirstId);
            TEscherDgRecord Dg= new TEscherDgRecord(0, DgId, FirstId, FDrawingGroup.RecordCache, FRecordCache, FDgContainer); //CreateFromData
            FDgContainer.ContainedRecords.Add(Dg);

            EscherHeader.Pre=0x0F;
            EscherHeader.Id=(int)Msofbt.SpgrContainer;
            EscherHeader.Size=0;
            FRecordCache.Patriarch = new TEscherSpgrContainerRecord(EscherHeader, FDrawingGroup.RecordCache, FRecordCache, FDgContainer);
            FDgContainer.ContainedRecords.Add(FRecordCache.Patriarch);

            EscherHeader.Id=(int)Msofbt.SpContainer;
            EscherHeader.Pre=0x0F;
            EscherHeader.Size=0; //Size for a container is calculated later
            TEscherSpContainerRecord SPRec= new TEscherSpContainerRecord(EscherHeader, FDrawingGroup.RecordCache, FRecordCache, FRecordCache.Patriarch);
            SPRec.LoadedDataSize=(int)EscherHeader.Size;
            FRecordCache.Patriarch.ContainedRecords.Add(SPRec);

            EscherHeader.Id=(int)Msofbt.Spgr;
            EscherHeader.Pre=0x01;
            EscherHeader.Size=16;
            TEscherDataRecord SPgrRec=new TEscherDataRecord(EscherHeader, FDrawingGroup.RecordCache, FRecordCache, FRecordCache.Patriarch);
            SPgrRec.LoadedDataSize=(int)EscherHeader.Size;
            SPgrRec.ClearData();
            SPRec.ContainedRecords.Add(SPgrRec);

            TEscherSpRecord SP = new TEscherSpRecord(0x02, FRecordCache.Dg.IncMaxShapeId(), FDrawingGroup.RecordCache, FRecordCache, SPRec,
                true, false, true, false, false); //CreateFromData
            SPRec.ContainedRecords.Add(SP);
        }

        internal TEscherClientDataRecord AddNewObject(TEscherSpContainerRecord SPRec, TShapeType ShapeType, TClientAnchor Anchor, TSheet sSheet, TEscherOPTRecord aOPTRec, TMsObj aMsObj)
        {
            return AddNewObject(null, null, SPRec, ShapeType, Anchor, null, sSheet, aOPTRec, aMsObj, new TDrawingPoint(), new Size(), 
                false);
        }

        internal TEscherClientDataRecord AddNewObject(TEscherContainerRecord GroupParent, TEscherContainerRecord Parent, TEscherSpContainerRecord SPRec, TShapeType ShapeType, TClientAnchor Anchor,
            THeaderOrFooterAnchor HFAnchor, TSheet sSheet, TEscherOPTRecord aOPTRec, TMsObj aMsObj, TDrawingPoint Offs, Size Ext, 
            bool IsGroup)
        {
            TEscherSpRecord SP = new TEscherSpRecord((((int)ShapeType) << 4) + 2, FRecordCache.Dg.IncMaxShapeId(), 
                FDrawingGroup.RecordCache, FRecordCache, SPRec,
                IsGroup, GroupParent != null && GroupParent != FDgContainer.Patriarch(), false, ShapeType != TShapeType.NotPrimitive, true); 
            SPRec.ContainedRecords.Add(SP);

            SPRec.ContainedRecords.Add(aOPTRec);

            if (HFAnchor != null)
            {
                TEscherRecordHeader RecordHeader = new TEscherRecordHeader();
                RecordHeader.Id = (int)Msofbt.ClientAnchor;
                RecordHeader.Pre = 0;
                THeaderOrFooterAnchor HeadClientAnchor = (THeaderOrFooterAnchor)HFAnchor.Clone();
                RecordHeader.Size = (int)HeadClientAnchor.Length;

                TEscherHeaderAnchorRecord HeadAnchorRec = new TEscherHeaderAnchorRecord(HeadClientAnchor, RecordHeader, FDrawingGroup.RecordCache, FRecordCache, SPRec);  //CreateFromData
                SPRec.ContainedRecords.Add(HeadAnchorRec);
            }
            else
            {
                if (Anchor != null)
                {
                    TEscherRecordHeader RecordHeader = new TEscherRecordHeader();
                    RecordHeader.Id = (int)Msofbt.ClientAnchor;
                    RecordHeader.Pre = 0;
                    RecordHeader.Size = TClientAnchor.Biff8Length;
                    TClientAnchor ClientAnchor = (TClientAnchor)Anchor.Clone();
                    TEscherImageAnchorRecord AnchorRec = TEscherImageAnchorRecord.CreateFromData(ClientAnchor, RecordHeader, FDrawingGroup.RecordCache, FRecordCache, SPRec, sSheet);
                    SPRec.ContainedRecords.Add(AnchorRec);
                }
                else
                {
                    TEscherChildAnchorRecord AnchorRec = TEscherChildAnchorRecord.CreateFromData(FDrawingGroup.RecordCache, FRecordCache,
                        SPRec, Offs, Ext.Height, Ext.Width);
                    SPRec.ContainedRecords.Add(AnchorRec);
                }
            }


            TEscherClientDataRecord ClientData = null;

            if (aMsObj != null)
            {
                ClientData = new TEscherClientDataRecord(FDrawingGroup.RecordCache, FRecordCache, SPRec);
                ClientData.AssignClientData(aMsObj, true);
                SPRec.ContainedRecords.Add(ClientData);
            }

            FRecordCache.OptByName.AddShapeId(aOPTRec);
            if (Parent == null) Parent = FRecordCache.Patriarch;
            Parent.ContainedRecords.Add(SPRec);
            return ClientData;
            
        }

        private TEscherSpContainerRecord CreateSPContainer()
        {
            return CreateSPContainer(null);
        }

        private TEscherSpContainerRecord CreateSPContainer(TEscherContainerRecord ObjParent)
        {
            FDrawingGroup.EnsureDwgGroup();

            if ((FDgContainer == null) || (FRecordCache.AnchorList == null))  //no drawings on this sheet
                CreateBasicDrawingInfo();

            TEscherRecordHeader RecordHeader = new TEscherRecordHeader();
            RecordHeader.Id = (int)Msofbt.SpContainer;
            RecordHeader.Pre = 0x0F;
            RecordHeader.Size = 0; //Size for a container is calculated later

            if (ObjParent == null) ObjParent = FRecordCache.Patriarch;
            TEscherSpContainerRecord SPRec = new TEscherSpContainerRecord(RecordHeader, FDrawingGroup.RecordCache, FRecordCache, ObjParent);
            SPRec.LoadedDataSize = (int)RecordHeader.Size;

            return SPRec;
        }


        private TEscherSpgrContainerRecord CreateSPGrContainer(TEscherContainerRecord ObjParent)
        {
            FDrawingGroup.EnsureDwgGroup();

            if ((FDgContainer == null) || (FRecordCache.AnchorList == null))  //no drawings on this sheet
                CreateBasicDrawingInfo();

            TEscherRecordHeader RecordHeader = new TEscherRecordHeader();
            RecordHeader.Id = (int)Msofbt.SpgrContainer;
            RecordHeader.Pre = 0x0F;
            RecordHeader.Size = 0; //Size for a container is calculated later

            if (ObjParent == null) ObjParent = FRecordCache.Patriarch;
            TEscherSpgrContainerRecord SPGrRec = new TEscherSpgrContainerRecord(RecordHeader, FDrawingGroup.RecordCache, FRecordCache, ObjParent);
            SPGrRec.LoadedDataSize = (int)RecordHeader.Size;

            return SPGrRec;
        }

        internal void AddImage(ExcelFile xls, byte[] Data, TXlsImgType DataType, TBaseImageProperties Properties, bool IsHeaderImage,
            string ShapeName, TSheet Sheet, int WorkingSheet, TEscherContainerRecord ObjParent, bool ReadingXlsx)
        {
            if (Data == null || Data.Length==0) 
            {
                Data=EmptyBmp;
                DataType=TXlsImgType.Bmp;
            }

            ImageUtils.CheckImgValid(ref Data, ref DataType, true);

            TEscherSpContainerRecord SPRec = CreateSPContainer(ObjParent);
            TEscherOPTRecord OPTRec=TEscherOPTRecord.CreateFromDataImg(Data, DataType, Properties, ShapeName, FDrawingGroup.RecordCache, FRecordCache, SPRec);  //CreateFromDataImg
            TClientAnchor ClientAnchor = null;
            THeaderOrFooterAnchor HFAnchor = null;
            TMsObj aMsObj = null;
            TImageProperties ImgProps = Properties as TImageProperties;
            if (IsHeaderImage)
            {
                HFAnchor = ((THeaderOrFooterImageProperties)Properties).Anchor;
            }
            else
            {
                ClientAnchor = (TClientAnchor)(ImgProps).Anchor;
                aMsObj = TMsObj.CreateEmptyImg(ref FRecordCache.MaxObjId, Properties);  
            }

            TDrawingPoint Offs = new TDrawingPoint();
            Size Ext = new Size();
            TObjectProperties ObjProps = ImgProps as TObjectProperties;
#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
            if (ObjProps != null)
            {
                Offs = ObjProps.Offs;
                Ext = ObjProps.Ext;
            }
#endif

            AddNewObject(ObjParent, ObjParent, SPRec, TShapeType.PictureFrame, ClientAnchor, HFAnchor, Sheet, OPTRec, aMsObj, Offs, Ext, false);
            if (ImgProps != null)
            {
                SetShapeProps(ImgProps.ShapeOptions, OPTRec, SPRec);
                AssignMacro(xls, WorkingSheet, ObjProps, aMsObj, ReadingXlsx);
            }
            SetExtendedProps(xls, ImgProps, OPTRec, TCommentProperties.DefaultLineColorSystem, TCommentProperties.DefaultFillColorSystem);

        }

        private static void AssignMacro(ExcelFile xls, int WorkingSheet, TObjectProperties ObjProps, TMsObj aMsObj, bool ReadingXlsx)
        {
            if (ObjProps == null || String.IsNullOrEmpty(ObjProps.Macro)) return;
            string Macro = ReadingXlsx ? TFormulaMessages.TokenString(TFormulaToken.fmStartFormula) + ObjProps.Macro : ObjProps.Macro;
            TFormulaConvertTextToInternal Converter = new TFormulaConvertTextToInternal(xls, WorkingSheet, false, Macro, true, false, false, null, TFmReturnType.Ref, false);
            if (ReadingXlsx)
            {
                Converter.SetReadingXlsx();
            }
            Converter.Parse();

            aMsObj.SetObjFormulaMacro(xls, Converter.GetTokens());
        }

#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
        internal void AddImage(ExcelFile xls, TSheet Sheet, int WorkingSheet, TImageProperties ImgProps, TEscherContainerRecord ObjParent, bool ReadingXlsx)
        {
            if (ImgProps.BlipFill == null || ImgProps.BlipFill.Blip == null) return;

            ImgProps.FileName = ImgProps.BlipFill.Blip.ImageFileName;
            ImgProps.SetCropArea(GetCropArea(ImgProps.BlipFill.SourceRect));
            AddImage(xls, ImgProps.BlipFill.Blip.PictureData, GetImageType(ImgProps.BlipFill.Blip.ContentType), ImgProps, false,
                ImgProps.ShapeName , Sheet, WorkingSheet, ObjParent, ReadingXlsx);
        }

        private static TCropArea GetCropArea(TDrawingRelativeRect? SrcRect)
        {
            if (!SrcRect.HasValue) return new TCropArea();
            return new TCropArea(
                (int)Math.Round(SrcRect.Value.Top * 65536),
                (int)Math.Round(SrcRect.Value.Bottom * 65536),
                (int)Math.Round(SrcRect.Value.Left * 65536),
                (int)Math.Round(SrcRect.Value.Right * 65536));
        }

        internal static TXlsImgType GetImageType(string mime)
        {
            switch (mime)
            {
                case "image/gif": return TXlsImgType.Gif;
                case "image/png": return TXlsImgType.Png;
                case "image/tiff": return TXlsImgType.Tiff;
                case "image/jpeg": return TXlsImgType.Jpeg;
                case "image/pict": return TXlsImgType.Pict;
                case "image/x-wmf": return TXlsImgType.Wmf;
                case "image/x-emf": return TXlsImgType.Emf;
                case "image/bmp": return TXlsImgType.Bmp;
            }
            return TXlsImgType.Unknown;
        }

        internal void AddShape(ExcelFile xls, TSheet Sheet, int WorkingSheet, TObjectProperties ObjProps, TEscherContainerRecord ObjParent, bool ReadingXlsx)
        {
            if (ObjProps.ShapeOptions == null) return;
            TEscherSpContainerRecord SPRec = CreateSPContainer(ObjParent);
            TEscherOPTRecord OPTRec = TEscherOPTRecord.CreateFromDataShape(FDrawingGroup.RecordCache, FRecordCache, SPRec); 

            TClientAnchor ClientAnchor = (TClientAnchor)ObjProps.Anchor;
            TMsObj aMsObj = TMsObj.CreateEmptyShape(ref FRecordCache.MaxObjId, ObjProps, true);
           
            AddNewObject(ObjParent, ObjParent, SPRec, ObjProps.ShapeOptions.ShapeType, ClientAnchor, null, Sheet, OPTRec, aMsObj, ObjProps.Offs, 
                ObjProps.Ext, false);
            if (ObjProps != null)
            {
                SetShapeProps(ObjProps.ShapeOptions, OPTRec, SPRec);
                if (ObjProps.FTextExt != null) SetObjectTextExt(OPTRec, ObjProps.FTextExt, xls, ObjProps.ShapeOptions, null);
                else
                {
                    if (ObjProps.FText != null) SetObjectText(OPTRec.Parent, ObjProps.FText, xls, null);
                }
                AssignMacro(xls, WorkingSheet, ObjProps, aMsObj, ReadingXlsx);
            }
            SetExtendedProps(xls, ObjProps, OPTRec, TCommentProperties.DefaultLineColorSystem, TCommentProperties.DefaultFillColorSystem);
        }

        private TEscherSpgrContainerRecord AddGroup(ExcelFile xls, TSheet Sheet, int WorkingSheet, TObjectProperties ObjProps, TEscherContainerRecord ObjParent, bool ReadingXlsx)
        {
            TEscherSpgrContainerRecord SPGr = CreateSPGrContainer(ObjParent);
            TEscherSpContainerRecord SPc = CreateSPContainer(SPGr);
            
            TEscherSpgrRecord SpgrRec = new TEscherSpgrRecord(FDrawingGroup.RecordCache, FRecordCache, SPc, 
                ObjProps.GetChOffs(), ObjProps.GetChExt().Height, ObjProps.GetChExt().Width);
            SPc.ContainedRecords.Add(SpgrRec);

            TEscherOPTRecord OPTRec = TEscherOPTRecord.CreateFromDataGroup(FDrawingGroup.RecordCache, FRecordCache, SPc);
            
            TClientAnchor ClientAnchor = (TClientAnchor)ObjProps.Anchor;
            TMsObj aMsObj = TMsObj.CreateEmptyGroup(ref FRecordCache.MaxObjId, ObjProps, true);

            AddNewObject(ObjParent, SPGr, SPc, ObjProps.ShapeOptions.ShapeType, ClientAnchor, null, Sheet, OPTRec, aMsObj, ObjProps.Offs, 
                ObjProps.Ext, true);
            if (ObjProps != null)
            {
                SetShapeProps(ObjProps.ShapeOptions, OPTRec, SPc);
                AssignMacro(xls, WorkingSheet, ObjProps, aMsObj, ReadingXlsx);
            }
            //Groups don't have colors or other stuff.
            //SetExtendedProps(xls, ObjProps, OPTRec, TCommentProperties.DefaultLineColorSystem, TCommentProperties.DefaultFillColorSystem);
            OPTRec.ShapeName = ObjProps.ShapeName;

            //SPGr.ContainedRecords.Add(SPc); Already added by addnewobject
            if (ObjParent == null) ObjParent = FDgContainer.Patriarch();
            ObjParent.ContainedRecords.Add(SPGr);
            return SPGr; 
        }


        internal void AddObject(ExcelFile xls, TSheet Sheet, int WorkingSheet, TObjectProperties ObjProps, TEscherContainerRecord ObjParent, bool ReadingXlsx)
        {
            if (ObjProps.ShapeOptions.ObjectType == TObjectType.Group)
            {
                if (ObjProps.FGroupedShapes.Count == 0) return;

                TEscherSpgrContainerRecord Group = AddGroup(xls, Sheet, WorkingSheet, ObjProps, ObjParent, ReadingXlsx);
                for (int i = 0; i < ObjProps.FGroupedShapes.Count; i++)
                {
                    AddObject(xls, Sheet, WorkingSheet, ObjProps.FGroupedShapes[i], Group, ReadingXlsx);
                }
                return;
            }

            if (ObjProps.BlipFill != null)
            {
                AddImage(xls, Sheet, WorkingSheet, ObjProps, ObjParent, ReadingXlsx);
                return;
            }
            AddShape(xls, Sheet, WorkingSheet, ObjProps, ObjParent, ReadingXlsx);

        }


#endif

        internal void AddAutoFilter(int Row, int[][] Cols, TSheet sSheet)
        {
            for (int i = 0; i < Cols.Length; i++)
            {
                AddAutoFilter(Row, Cols[i], sSheet);
            }
        }

        internal void AddAutoFilter(int Row, int[] Col, TSheet sSheet)
        {
            TEscherSpContainerRecord SPRec = CreateSPContainer();
            TClientAnchor ClientAnchor = new TClientAnchor((TFlxAnchorType)1, Row, 0, Col[0], 0, Row + 1, 0, Col[1] + 1, 0);//AutoFilters have a "1" as flag. This is not documented.
            TEscherOPTRecord OPTRec = TEscherOPTRecord.CreateFromDataAutoFilter(FDrawingGroup.RecordCache, FRecordCache, SPRec);
            TMsObj aMsObj = TMsObj.CreateEmptyAutoFilter(ref FRecordCache.MaxObjId);

            AddNewObject(SPRec,TShapeType.HostControl, ClientAnchor, sSheet, OPTRec, aMsObj);
        }

        internal TEscherClientDataRecord AddNewComment(ExcelFile xls, TImageProperties Properties, TSheet sSheet, bool ReadFromXlsx)
        {
            TEscherSpContainerRecord SPRec = CreateSPContainer();
            TEscherOPTRecord OPTRec = TEscherOPTRecord.CreateFromDataNote(FDrawingGroup.RecordCache, FRecordCache, SPRec, 0); //.CreateFromDataNote
            TMsObj aMsObj = TMsObj.CreateEmptyNote(ref FRecordCache.MaxObjId, Properties, ReadFromXlsx);

            TEscherClientDataRecord Result = AddNewObject(SPRec, TShapeType.TextBox, Properties.Anchor, sSheet, OPTRec, aMsObj);
            if (Properties != null) SetShapeProps(Properties.ShapeOptions, OPTRec, SPRec);

            TObjectProperties ObjProps = Properties as TObjectProperties;
            TObjectTextProperties TextProps = ObjProps == null ? null : ObjProps.FTextProperties;
            SetObjectText(FRecordCache.AnchorList.Count - 1, null, new TRichString(String.Empty), xls, TextProps);
            SetExtendedProps(xls, ObjProps, OPTRec, TCommentProperties.DefaultLineColorSystem, TCommentProperties.DefaultFillColorSystem);

            return Result;
        }

        private static void SetExtendedProps(IFlexCelPalette aPalette, TImageProperties ImgProps, TEscherOPTRecord OPTRec, TSystemColor DefaultSysColorFg, TSystemColor DefaultSysColorBg)
        {
            if (ImgProps == null) return;

            OPTRec.ShapeName = ImgProps.ShapeName;
            if (!String.IsNullOrEmpty(ImgProps.AltText)) OPTRec.SetStringProperty(TShapeOption.wzDescription, ImgProps.AltText);

            TObjectProperties ObjProps = ImgProps as TObjectProperties;
            if (ObjProps != null)
            {
                if (ObjProps.FAutoSize) OPTRec.SetBoolProperty(TShapeOption.fFitTextToShape, 0x20002, false, true);
                if (ObjProps.LockAspectRatio) OPTRec.SetBoolProperty(TShapeOption.fLockAgainstGrouping, 0x800080, false, true);

                if (ObjProps.FTextProperties != null)
                {
                    switch (ObjProps.FTextProperties.TextRotation)
                    {
                        case TTextRotation.Normal:
                            break;

                        case TTextRotation.Rotated90Degrees:
                            OPTRec.SetBoolProperty(TShapeOption.txflTextFlow, 0x02, false, true);
                            break;

                        case TTextRotation.RotatedMinus90Degrees:
                            OPTRec.SetBoolProperty(TShapeOption.txflTextFlow, 0x03, false, true);
                            break;

                        case TTextRotation.Vertical:
                            OPTRec.SetBoolProperty(TShapeOption.txflTextFlow, 0x05, false, true);
                            break;

                        default:
                            break;
                    }
                }

                if (aPalette != null)
                {
                    OPTRec.SetFillColor(aPalette, ObjProps.ShapeFill, DefaultSysColorBg);
                    OPTRec.SetLineStyle(aPalette, ObjProps.ShapeBorder, DefaultSysColorFg);
                }

#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
                if (ObjProps.ShapeOptions != null)
                {
                    OPTRec.SetShapeGeom(ObjProps.ShapeOptions.ShapeGeom);
                    OPTRec.SetShapeFont(ObjProps.ShapeOptions.ShapeThemeFont);
                }
                OPTRec.SetEffectProps(ObjProps.FEffectProperties);
                OPTRec.SetShapeEffects(ObjProps.FShapeEffects);
                OPTRec.SetHlinkClick(ObjProps.HLinkClick);
                OPTRec.SetHlinkHover(ObjProps.HLinkHover);
                OPTRec.SetBodyPr(ObjProps.BodyPr);
                OPTRec.SetLstStyle(ObjProps.LstStyle);
#endif

                OPTRec.Visible = !ObjProps.FHidden;
            }

        }

        internal bool GetRadioButton(int Index, string ObjectPath)
        {
            TCheckboxState st = GetCheckboxOrRadioButton(Index, ObjectPath, TObjectType.OptionButton);
            return st == TCheckboxState.Checked;
        }

        internal TCheckboxState GetCheckbox(int Index, string ObjectPath)
        {
            return GetCheckboxOrRadioButton(Index, ObjectPath, TObjectType.CheckBox);
        }

        private TCheckboxState GetCheckboxOrRadioButton(int Index, string ObjectPath, TObjectType ObjType)
        {
            if (FRecordCache.Patriarch == null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
            TEscherContainerRecord R = TEscherContainerRecord.Navigate(FRecordCache, Index, ObjectPath);
            if (R == null) return TCheckboxState.Indeterminate;
            for (int i = 0; i < R.ContainedRecords.Count; i++)
            {
                TEscherClientDataRecord Cd = R.ContainedRecords[i] as TEscherClientDataRecord;
                if (Cd != null)
                {
                    TMsObj obj = Cd.ClientData as TMsObj;
                    if (obj != null && (TObjectType)obj.ObjType == ObjType)
                    {
                        return obj.GetCheckbox();
                    }
                }
            }

            return TCheckboxState.Indeterminate;
        }

        internal void SetRadioButton(int Index, string ObjectPath, bool Selected, out int RbPosition, IRowColSize Workbook, out bool Changed)
        {
            Changed = false;
            RbPosition = 0;
            TCheckboxState Value = Selected ? TCheckboxState.Checked : TCheckboxState.Unchecked;

            
            TEscherContainerRecord R;
            bool RbChanged = SetCheckOrRadio(Index, ObjectPath, Value, TObjectType.OptionButton, out R);
            if (R == null || !RbChanged) return; //no need to change the other checkboxes.
            Changed = true;
            if (!Selected) return; //While checkbox changed, others weren't affected.

            UncheckOthersInGroup(R as TEscherSpContainerRecord, Workbook, ref RbPosition);
        }

        private void UncheckOthersInGroup(TEscherSpContainerRecord R, IRowColSize Workbook, ref int RbPosition)
        {
            if (R == null) return;
            if (FRecordCache == null || FRecordCache.Obj == null) return;

            List<TEscherClientDataRecord> BtnGroup;
            if (!FRecordCache.RadioButtons.Find(R, out BtnGroup, Workbook)) return;

            for (int i = 0; i < BtnGroup.Count; i++)
            {
                if (BtnGroup[i].Parent == R) RbPosition = i + 1;
                else
                {
                    TMsObj ms = BtnGroup[i].ClientData as TMsObj;
                    ms.SetCheckbox(TCheckboxState.Unchecked);
                }
            }
        }

        internal int GetRbPosition(int Index, string ObjectPath, IRowColSize Workbook)
        {
            if (FRecordCache.Patriarch == null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
            TEscherContainerRecord R = TEscherContainerRecord.Navigate(FRecordCache, Index, ObjectPath);
            if (R == null) return 0;

            List<TEscherClientDataRecord> BtnGroup;
            if (!FRecordCache.RadioButtons.Find(R as TEscherSpContainerRecord, out BtnGroup, Workbook)) return 0;
            for (int i = 0; i < BtnGroup.Count; i++)
            {
                if (BtnGroup[i].Parent == R) return i + 1;
            }

            return 0;
        }

        internal void SetCheckbox(int Index, string ObjectPath, TCheckboxState Value)
        {
            TEscherContainerRecord R;
            SetCheckOrRadio(Index, ObjectPath, Value, TObjectType.CheckBox, out R);
        }

        private bool SetCheckOrRadio(int Index, string ObjectPath, TCheckboxState Value, TObjectType ObjType, out TEscherContainerRecord R)
        {
            R = null;
            if (FRecordCache.Patriarch == null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
            R = TEscherContainerRecord.Navigate(FRecordCache, Index, ObjectPath);
            if (R == null) return false;
            for (int i = 0; i < R.ContainedRecords.Count; i++)
            {
                TEscherClientDataRecord Cd = R.ContainedRecords[i] as TEscherClientDataRecord;
                if (Cd != null)
                {
                    TMsObj obj = Cd.ClientData as TMsObj;
                    if (obj != null && (TObjectType)obj.ObjType == ObjType)
                    {
                        return obj.SetCheckbox(Value);
                    }
                }
            }

            return false;
        }

        internal int GetObjectSelection(int Index, string ObjectPath)
        {
            if (FRecordCache.Patriarch == null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
            TEscherContainerRecord R = TEscherContainerRecord.Navigate(FRecordCache, Index, ObjectPath);
            if (R == null) return 0;
            TEscherClientDataRecord Cd = R.FindRec<TEscherClientDataRecord>() as TEscherClientDataRecord;
            if (Cd != null)
            {
                TMsObj obj = Cd.ClientData as TMsObj;
                if (obj != null)
                {
                    return obj.GetObjectSelection();
                }
            }
            return 0;
        }

        internal void SetObjectSelection(int Index, string ObjectPath, int Value)
        {
            TEscherContainerRecord R = null;
            if (FRecordCache.Patriarch == null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
            R = TEscherContainerRecord.Navigate(FRecordCache, Index, ObjectPath);
            if (R == null) return;
            TEscherClientDataRecord Cd = R.FindRec<TEscherClientDataRecord>() as TEscherClientDataRecord;
            if (Cd != null)
            {
                TMsObj obj = Cd.ClientData as TMsObj;
                if (obj != null)
                {
                    obj.SetObjectSelection(Value);
                }
            }
        }

        internal int GetObjectSpinValue(int Index, string ObjectPath)
        {
            if (FRecordCache.Patriarch == null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
            TEscherContainerRecord R = TEscherContainerRecord.Navigate(FRecordCache, Index, ObjectPath);
            if (R == null) return 0;
            TEscherClientDataRecord Cd = R.FindRec<TEscherClientDataRecord>() as TEscherClientDataRecord;
            if (Cd != null)
            {
                TMsObj obj = Cd.ClientData as TMsObj;
                if (obj != null)
                {
                    return obj.GetObjectSpinValue(false, -1);
                }
            }
            return 0;
        }

        internal void SetObjectSpinValue(int Index, string ObjectPath, int Value)
        {
            TEscherContainerRecord R = null;
            if (FRecordCache.Patriarch == null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
            R = TEscherContainerRecord.Navigate(FRecordCache, Index, ObjectPath);
            if (R == null) return;
            TEscherClientDataRecord Cd = R.FindRec<TEscherClientDataRecord>() as TEscherClientDataRecord;
            if (Cd != null)
            {
                TMsObj obj = Cd.ClientData as TMsObj;
                if (obj != null)
                {
                    obj.SetObjectSpinValue(Value);
                }
            }
        }

        internal TCellAddress GetObjectLink(int Index, string ObjectPath, ExcelFile xls)
        {
            if (FRecordCache.Patriarch == null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
            TEscherContainerRecord R = TEscherContainerRecord.Navigate(FRecordCache, Index, ObjectPath);
            if (R == null) return null;

            TEscherClientDataRecord Cd = R.FindRec<TEscherClientDataRecord>() as TEscherClientDataRecord;
            if (Cd == null) return null;

            TMsObj obj = Cd.ClientData as TMsObj;
            if (obj == null) return null;

            if (obj.ObjType == TObjectType.OptionButton) //linked cell here is the first of the group.
            {
                List<TEscherClientDataRecord> BtnGrp;
                if (FRecordCache.RadioButtons.Find(R as TEscherSpContainerRecord, out BtnGrp, xls))
                {
                    if (BtnGrp.Count > 0) R = BtnGrp[0].Parent;
                }

                Cd = R.FindRec<TEscherClientDataRecord>() as TEscherClientDataRecord;
                if (Cd == null) return null;
                obj = Cd.ClientData as TMsObj;
                if (obj == null) return null;
            }

            return obj.GetObjectLink(xls);
        }


        internal void SetObjectLink(int Index, string ObjectPath, TCellAddress LinkedCellAddress, TParsedTokenList LinkedCellFmla, ExcelFile xls, bool ReadingXlsx)
        {
            if (FRecordCache.Patriarch == null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
            TEscherContainerRecord R = TEscherContainerRecord.Navigate(FRecordCache, Index, ObjectPath);
            if (R == null) return;
            TEscherClientDataRecord Cd = (TEscherClientDataRecord)R.FindRec<TEscherClientDataRecord>();
            if (Cd == null) return;

            TMsObj obj = Cd.ClientData as TMsObj;
            if (obj != null && CanHandleObjLink((TObjectType)obj.ObjType) )
            {
                if ((TObjectType)obj.ObjType == TObjectType.OptionButton)
                {
                    FRecordCache.RadioButtons.ChangeFmla(xls, Cd, LinkedCellAddress, LinkedCellFmla, ReadingXlsx);
                }
                else
                {
                    obj.SetObjFormulaLink(xls, LinkedCellAddress, LinkedCellFmla);
                }
            }
        }

        private bool CanHandleObjLink(TObjectType ot)
        {
            switch (ot)
            {
                case TObjectType.CheckBox:
                case TObjectType.OptionButton:
                case TObjectType.ComboBox:
                case TObjectType.ListBox:
                case TObjectType.Spinner:
                case TObjectType.ScrollBar:
                    return true;

                default:
                    return false;
            }
        }

        internal TCellAddressRange GetObjectInputRange(int Index, string ObjectPath, ExcelFile xls)
        {
            if (FRecordCache.Patriarch == null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
            TEscherContainerRecord R = TEscherContainerRecord.Navigate(FRecordCache, Index, ObjectPath);
            if (R == null) return null;

            TEscherClientDataRecord Cd = R.FindRec<TEscherClientDataRecord>() as TEscherClientDataRecord;
            if (Cd == null) return null;

            TMsObj obj = Cd.ClientData as TMsObj;
            if (obj == null) return null;

            return obj.GetObjectInputRange(xls);
        }

        internal void SetFormulaRange(int Index, string ObjectPath, TCellAddressRange InputRange, TParsedTokenList LinkedCellFmla, ExcelFile xls, bool ReadingXlsx)
        {
            if (FRecordCache.Patriarch == null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
            TEscherContainerRecord R = TEscherContainerRecord.Navigate(FRecordCache, Index, ObjectPath);
            if (R == null) return;
            TEscherClientDataRecord Cd = (TEscherClientDataRecord)R.FindRec<TEscherClientDataRecord>();
            if (Cd == null) return;

            TMsObj obj = Cd.ClientData as TMsObj;
            if (obj != null && CanHandleFmlaRange(obj.ObjType))
            {
                obj.SetObjFormulaRange(xls, InputRange, LinkedCellFmla);
            }
        }

        private bool CanHandleFmlaRange(TObjectType ot)
        {
            switch (ot)
            {
                case TObjectType.ListBox:
                case TObjectType.ComboBox:
                    return true;

                default:
                    return false;
            }
        }

        internal string GetObjectMacro(int Index, string ObjectPath, TCellList CellList)
        {
            if (FRecordCache.Patriarch == null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
            TEscherContainerRecord R = TEscherContainerRecord.Navigate(FRecordCache, Index, ObjectPath);
            if (R == null) return null;

            TEscherClientDataRecord Cd = R.FindRec<TEscherClientDataRecord>() as TEscherClientDataRecord;
            if (Cd == null) return null;

            TMsObj obj = Cd.ClientData as TMsObj;
            if (obj == null) return null;

            return obj.GetObjectMacro(CellList);
        }

        internal TSpinProperties GetObjectSpinProperties(int Index, string ObjectPath)
        {
            if (FRecordCache.Patriarch == null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
            TEscherContainerRecord R = TEscherContainerRecord.Navigate(FRecordCache, Index, ObjectPath);
            if (R == null) return null;

            TEscherClientDataRecord Cd = R.FindRec<TEscherClientDataRecord>() as TEscherClientDataRecord;
            if (Cd == null) return null;

            TMsObj obj = Cd.ClientData as TMsObj;
            if (obj == null) return null;

            return obj.GetSpinProps();
        }

        internal void SetObjectSpinProperties(int Index, string ObjectPath, TSpinProperties SpinProps)
        {
            if (FRecordCache.Patriarch == null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
            TEscherContainerRecord R = TEscherContainerRecord.Navigate(FRecordCache, Index, ObjectPath);
            if (R == null) return;
            TEscherClientDataRecord Cd = (TEscherClientDataRecord)R.FindRec<TEscherClientDataRecord>();
            if (Cd == null) return;

            TMsObj obj = Cd.ClientData as TMsObj;
            if (obj != null && CanHandleSpinner(obj.ObjType))
            {
                obj.SetSpinProps(SpinProps);
            }
        }

        private bool CanHandleSpinner(TObjectType ot)
        {
            switch (ot)
            {
                case TObjectType.ListBox:
                case TObjectType.ComboBox:
                case TObjectType.Spinner:
                case TObjectType.ScrollBar:
                    return true;

                default:
                    return false;
            }
        }


        internal int AddCheckbox(TClientAnchor Anchor, TRichString Text, TCheckboxState value, 
            ExcelFile xls, TSheet sSheet, TImageProperties Props, string name)
        {
            TEscherSpContainerRecord SPRec = CreateSPContainer();
            TEscherOPTRecord OPTRec = TEscherOPTRecord.CreateFromDataRadioOrCheckbox(FDrawingGroup.RecordCache, FRecordCache, SPRec);
            TMsObj aMsObj = TMsObj.CreateEmptyCheckbox(ref FRecordCache.MaxObjId, Props);
            if (value != TCheckboxState.Unchecked) aMsObj.SetCheckbox(value);
            return AddFormsObject(Anchor, Text, xls, sSheet, Props, SPRec, OPTRec, aMsObj, TVFlxAlignment.center, name);

        }


        internal int AddRadioButton(TClientAnchor Anchor, TRichString Text, TCheckboxState value, ExcelFile xls, 
            TSheet sSheet, TImageProperties Props, string name)
        {
            TEscherSpContainerRecord SPRec = CreateSPContainer();
            TEscherOPTRecord OPTRec = TEscherOPTRecord.CreateFromDataRadioOrCheckbox(FDrawingGroup.RecordCache, FRecordCache, SPRec);
            TMsObj aMsObj = TMsObj.CreateEmptyRadioButton(ref FRecordCache.MaxObjId, Props);
            if (value != TCheckboxState.Unchecked) aMsObj.SetCheckbox(value);
            return AddFormsObject(Anchor, Text, xls, sSheet, Props, SPRec, OPTRec, aMsObj, TVFlxAlignment.center, name);
        }

        internal int AddGroupBox(TClientAnchor Anchor, TRichString Text, 
            ExcelFile xls, TSheet sSheet, TImageProperties Props, string name)
        {
            TEscherSpContainerRecord SPRec = CreateSPContainer();
            TEscherOPTRecord OPTRec = TEscherOPTRecord.CreateFromDataGroupBox(FDrawingGroup.RecordCache, FRecordCache, SPRec);
            TMsObj aMsObj = TMsObj.CreateEmptyGroupBox(ref FRecordCache.MaxObjId, Props);
            return AddFormsObject(Anchor, Text, xls, sSheet, Props, SPRec, OPTRec, aMsObj, TVFlxAlignment.top, name);
        }

        internal int AddComboBox(TClientAnchor Anchor, int value,
            ExcelFile xls, TSheet sSheet, TImageProperties Props, string name)
        {
            TEscherSpContainerRecord SPRec = CreateSPContainer();
            TEscherOPTRecord OPTRec = TEscherOPTRecord.CreateFromDataComboBox(FDrawingGroup.RecordCache, FRecordCache, SPRec);
            TMsObj aMsObj = TMsObj.CreateEmptyComboBox(ref FRecordCache.MaxObjId, Props);
            return AddFormsObject(Anchor, null, xls, sSheet, Props, SPRec, OPTRec, aMsObj, TVFlxAlignment.center, name);
        }

        internal int AddListBox(TClientAnchor Anchor, int value,
            ExcelFile xls, TSheet sSheet, TImageProperties Props, string name, TListBoxSelectionType SelectionType)
        {
            TEscherSpContainerRecord SPRec = CreateSPContainer();
            TEscherOPTRecord OPTRec = TEscherOPTRecord.CreateFromDataListBox(FDrawingGroup.RecordCache, FRecordCache, SPRec);
            TMsObj aMsObj = TMsObj.CreateEmptyListBox(ref FRecordCache.MaxObjId, Props);
            aMsObj.SelectionType = SelectionType;
            return AddFormsObject(Anchor, null, xls, sSheet, Props, SPRec, OPTRec, aMsObj, TVFlxAlignment.center, name);
        }

        internal int AddLabel(TClientAnchor Anchor, TRichString Text, 
            ExcelFile xls, TSheet sSheet, TImageProperties Props, string name)
        {
            TEscherSpContainerRecord SPRec = CreateSPContainer();
            TEscherOPTRecord OPTRec = TEscherOPTRecord.CreateFromDataLabel(FDrawingGroup.RecordCache, FRecordCache, SPRec);
            TMsObj aMsObj = TMsObj.CreateEmptyLabel(ref FRecordCache.MaxObjId, Props);
            return AddFormsObject(Anchor, Text, xls, sSheet, Props, SPRec, OPTRec, aMsObj, TVFlxAlignment.center, name);
        }

        internal int AddSpinner(TClientAnchor Anchor,
            ExcelFile xls, TSheet sSheet, TImageProperties Props, string name)
        {
            TEscherSpContainerRecord SPRec = CreateSPContainer();
            TEscherOPTRecord OPTRec = TEscherOPTRecord.CreateFromDataSpinScroll(FDrawingGroup.RecordCache, FRecordCache, SPRec);
            TMsObj aMsObj = TMsObj.CreateEmptySpinner(ref FRecordCache.MaxObjId, Props);
            return AddFormsObject(Anchor, null, xls, sSheet, Props, SPRec, OPTRec, aMsObj, TVFlxAlignment.center, name);
        }

        internal int AddScrollBar(TClientAnchor Anchor,
            ExcelFile xls, TSheet sSheet, TImageProperties Props, string name)
        {
            TEscherSpContainerRecord SPRec = CreateSPContainer();
            TEscherOPTRecord OPTRec = TEscherOPTRecord.CreateFromDataSpinScroll(FDrawingGroup.RecordCache, FRecordCache, SPRec);
            TMsObj aMsObj = TMsObj.CreateEmptyScrollBar(ref FRecordCache.MaxObjId, Props);
            return AddFormsObject(Anchor, null, xls, sSheet, Props, SPRec, OPTRec, aMsObj, TVFlxAlignment.center, name);
        }

        private int AddFormsObject(TClientAnchor Anchor, TRichString Text, ExcelFile xls, 
            TSheet sSheet, TImageProperties Props, TEscherSpContainerRecord SPRec, TEscherOPTRecord OPTRec, TMsObj aMsObj,
            TVFlxAlignment DefaultVAlign, string Name)
        {
            AddNewObject(SPRec, TShapeType.HostControl, Anchor, sSheet, OPTRec, aMsObj);
            if (Props != null) SetShapeProps(Props.ShapeOptions, OPTRec, SPRec);

            int Result = FRecordCache.AnchorList.Count - 1;
            TObjectProperties ObjProps = Props as TObjectProperties;
            if (ObjProps == null) aMsObj.SetObject3D(false); else aMsObj.SetObject3D(ObjProps.FIs3D);

            TObjectTextProperties ObjTextProps = ObjProps == null ? null : ObjProps.FTextProperties;
            if (Text != null)
            {
                if (ObjTextProps == null) ObjTextProps = new TObjectTextProperties(true, THFlxAlignment.left, DefaultVAlign, TTextRotation.Normal);
                SetObjectText(Result, null, Text, xls, ObjTextProps);
            }

            if (ObjProps != null)
            {
#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
                aMsObj.SetComboProps(ObjProps.FComboBoxProperties);
                aMsObj.SetSpinProps(ObjProps.FSpinProperties);
                if (ObjProps.FSpinProperties != null) aMsObj.SetObjectSpinValue(ObjProps.FSpinProperties.FVal);
#endif
            }

            if (Name != null) OPTRec.ShapeName = Name; //name will also be set if Props is not null, this is just a shorcut for null props.
            SetExtendedProps(xls, ObjProps, OPTRec, TSystemColor.None, TSystemColor.Window);

            return Result;
        }

        private void SetShapeProps(TShapeProperties ShProps, TEscherOPTRecord OptRec, TEscherSpContainerRecord SpC)
        {
            if (ShProps == null) return;
            if (ShProps.FlipH) SpC.SP.Flags |= 0x40; else SpC.SP.Flags &= ~0x40;
            if (ShProps.FlipV) SpC.SP.Flags |= 0x80; else SpC.SP.Flags &= ~0x80;

            foreach (TShapeOption so in ShProps.ShapeOptions.Keys)
            {
                object obj = ShProps.ShapeOptions[so];
                if (obj is long) OptRec.SetLongProperty(so, (long)obj);
                
		 
            }
        }



        internal int AddButton(TClientAnchor Anchor, ExcelFile xls, TSheet sSheet, TObjectProperties Props)
        {
            TEscherSpContainerRecord SPRec = CreateSPContainer();
            TEscherOPTRecord OPTRec = TEscherOPTRecord.CreateFromDataButton(FDrawingGroup.RecordCache, FRecordCache, SPRec);
            TMsObj aMsObj = TMsObj.CreateEmptyButton(ref FRecordCache.MaxObjId, Props);
            AddNewObject(SPRec, TShapeType.HostControl, Anchor, sSheet, OPTRec, aMsObj);
            if (Props != null) SetShapeProps(Props.ShapeOptions, OPTRec, SPRec);

            int Result = FRecordCache.AnchorList.Count - 1;
            
            TObjectTextProperties ObjTextProps = Props.FTextProperties;
            if (Props.FText != null)
            {
                SetObjectText(Result, null, Props.FText, xls, ObjTextProps);
            }

            SetExtendedProps(xls, Props, OPTRec, TSystemColor.None, TSystemColor.Window);

            return Result;
        }

        internal void SetButtonMacro(int Index, string ObjectPath, TParsedTokenList Macro, ExcelFile xls)
        {
            if (FRecordCache.Patriarch == null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
            TEscherContainerRecord R = TEscherContainerRecord.Navigate(FRecordCache, Index, ObjectPath);
            if (R == null) return;
            for (int i = 0; i < R.ContainedRecords.Count; i++)
            {
                TEscherClientDataRecord Cd = R.ContainedRecords[i] as TEscherClientDataRecord;
                if (Cd != null)
                {
                    TMsObj obj = Cd.ClientData as TMsObj;
                    if (obj != null)
                    {
                        obj.SetObjFormulaMacro(xls, Macro);
                        return;
                    }
                }
            }
        }


        internal void FixRadioButtons(ExcelFile Workbook, int ActiveSheet)
        {
            if (FRecordCache.RadioButtons != null) FRecordCache.RadioButtons.FixLinks(Workbook, ActiveSheet);
        }

        internal string FindPath(string objectName)
        {
            if (FRecordCache.OptByName == null) return null;
            return FRecordCache.OptByName.Find(objectName);
        }


        internal bool IsSpecialDropdown(int i)
        {
            TEscherClientDataRecord obj0 = FRecordCache.Obj[i];
            if (obj0.ClientData == null) return false;
            TMsObj ClientData = obj0.ClientData as TMsObj;
            if (ClientData == null) return false;
            return ClientData.IsSpecialDropdown;
        }
    }


}
