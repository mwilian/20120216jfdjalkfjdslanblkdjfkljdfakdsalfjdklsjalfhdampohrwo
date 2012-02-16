using System;
using System.IO;
using FlexCel.Core;
using System.Collections.Generic;

namespace FlexCel.XlsAdapter
{


    internal class TEscherRecordList
    {

        protected List<TEscherRecord> FList;

        protected bool FSorted;
        internal TEscherRecordList()
        {
            FSorted=false;
            FList = new List<TEscherRecord>();
        }

        internal void Destroy()
        {
            for (int i = 0; i < Count; i++) if (this[i] != null) this[i].Destroy();
        }

        #region Generics
        internal void Add (TEscherRecord a)
        {
            FList.Add(a);
            FSorted=false;  //When we add the list gets unsorted
        }
        internal void Insert (int index, TEscherRecord a)
        {
            // We assume that when we insert, we respect the order, so we dont set Sorted=false
            FList.Insert(index, a);
        }

        protected void SetThis(TEscherRecord value, int index)
        {
            if (FList[index]!=null) FList[index].Destroy();
            FList[index]=value;
            FSorted=false;  //When we add the list gets unsorted
        }

        internal TEscherRecord this[int index] 
        {
            
            get {if (index < 0) index = 0;return FList[index];} 
            set {SetThis(value, index);}
        }

        internal void Swap(int Index1, int Index2)
        {
            TEscherRecord tmp=FList[Index1];
            FList[Index1]=FList[Index2];
            FList[Index2]=tmp;
            FSorted=false;
        }

        #endregion

        internal int Count
        {
            get {return FList.Count;}
        }

        internal void Sort()
        {
            FList.Sort();
            FSorted=true;
        }

        internal bool Find(TEscherRecord aRecord, ref int Index)
        {
            if (!FSorted) Sort();
            Index = FList.BinarySearch(0, FList.Count, aRecord, null); //Only BinarySearch compatible with CF.
            bool Result=Index>=0;
            if (Index<0) Index=~Index;
            return Result;
        }

        internal void SaveToStream(IDataStream DataStream, TSaveData SaveData, TBreakList BreakList)
        {
            for (int i=0; i< Count;i++)
                if (FList[i]!=null) FList[i].SaveToStream(DataStream, SaveData, BreakList);
        }

        internal void CopyFrom(int RowOfs, int ColOfs, TEscherRecordList aEscherRecordList, TEscherDwgCache NewDwgCache, TEscherDwgGroupCache NewDwgGroupCache, TSheetInfo SheetInfo)
        {
            if (aEscherRecordList == null) return;
            if (aEscherRecordList.FList == FList) XlsMessages.ThrowException(XlsErr.ErrInternal); //Should be different objects
            for (int i = 0; i < aEscherRecordList.Count; i++)
                Add(TEscherRecord.Clone(aEscherRecordList[i], RowOfs, ColOfs, NewDwgCache, NewDwgGroupCache, SheetInfo));

        }
        internal long TotalSizeNoSplit()
        {
            long Result=0;
            for (int i=0;i< Count;i++)
                Result+=FList[i].TotalSizeNoSplit();
            return Result;
        }

        internal void Delete(int Index)
        {
            TEscherRecord Obj = FList[Index];
            FList.RemoveAt(Index);
            Obj.Destroy();
        }

        internal void Remove(TEscherRecord Obj)
        {
            FList.Remove(Obj);
            Obj.Destroy();
        }

        internal int IndexOf(TEscherRecord Obj)
        {
            return FList.IndexOf(Obj);
        }

    }


    /// <summary>
    /// Base list of Escher records caches.
    /// All cache lists don't own the objects. They shouldn't call destroy.
    /// </summary>
    internal class TEscherRecordCache<T> where T:TEscherRecord
    {
        
        protected List<T> FList;
        protected IComparer<T> CacheComparer=null;

        protected bool FSorted;
        protected bool ListModified;
        internal TEscherRecordCache()
        {
            FSorted=false;
            FList = new List<T>();
        }

        #region Generics
        internal void Add (T a)
        {
            FList.Add(a);
            FSorted=false;  //When we add the list gets unsorted
            ListModified = true;
        }
        internal void Insert (int index, T a)
        {
            // We assume that when we insert, we respect the order, so we dont set Sorted=false
            FList.Insert(index, a);
            ListModified = true;            
        }

        protected void SetThis(T value, int index)
        {
            FList[index]=value;
            ListModified = true;
        }

        internal T this[int index] 
        {
            get {return FList[index];} 
            set {SetThis(value, index);}
        }

        #endregion

        internal int Count
        {
            get {return FList.Count;}
        }

        internal void ArrangeInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo, bool Forced)
        {
            for (int i=0; i< Count;i++)
                this[i].ArrangeInsertRange(CellRange, aRowCount, aColCount, SheetInfo, Forced);
        }

        internal void ArrangeMoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
        {
            for (int i=0; i< Count;i++)
                this[i].ArrangeMoveRange(CellRange, NewRow, NewCol, SheetInfo);
        }

        internal void Sort()
        {
            FList.Sort(CacheComparer);
            FSorted=true;
            ListModified = true;
        }

        internal void Delete(int Index)
        {
            //We DONT call destroy here, the records aren't owned
            FList.RemoveAt(Index);
            ListModified = true;
        }

        internal void Remove(T Obj)
        {
            //We DONT call destroy here, the records aren't owned
            FList.Remove(Obj);
        }


    } //TEscherRecordCache


    internal class TEscherObjCache : TEscherRecordCache<TEscherClientDataRecord>
    {
        internal TEscherObjCache()
        {
        }

        internal void ArrangeCopySheet(TSheetInfo SheetInfo)
        {
            for (int i=0; i< Count;i++)
                FList[i].ClientData.ArrangeCopySheet(SheetInfo);
        }

        #region Named Ranges
        internal void UpdateDeletedRanges(TDeletedRanges DeletedRanges)
        {
            for (int i=0; i< Count;i++)
                FList[i].ClientData.UpdateDeletedRanges(DeletedRanges);
        }
        #endregion
    }

    internal struct TRadioGroupCache
    {
        public TEscherClientDataRecord Group;
        public List<TEscherClientDataRecord> Buttons;

        public TRadioGroupCache(TEscherClientDataRecord grp)
        {
            Group = grp;
            Buttons = new List<TEscherClientDataRecord>();
        }
    }


    internal class TRadioButtonCache
    {
        Dictionary<int, TRadioGroupCache> RadioButtonsById;
        List<TRadioGroupCache> Tree;
        List<TEscherClientDataRecord> NewEntries;
        List<TEscherClientDataRecord> CopiedEntries;

        public TRadioButtonCache()
        {
            RadioButtonsById = new Dictionary<int, TRadioGroupCache>();
            Tree = new List<TRadioGroupCache>();
            Tree.Add(new TRadioGroupCache(null)); //first is no group.
            NewEntries = new List<TEscherClientDataRecord>();
            CopiedEntries = new List<TEscherClientDataRecord>();
        }

        internal void Remove(TEscherClientDataRecord ms)
        {
            TMsObj obj = ms.ClientData as TMsObj;
            if (obj == null) return;
            switch ((TObjectType)(obj.ObjType))
            {
                case TObjectType.GroupBox:
                    //remove from RadioButtonsById; and send all of those to copied.
                    for (int i = Tree.Count - 1; i >= 1; i--)
                    {
                        if (Tree[i].Group == ms)
                        {
                            for (int k = 0; k < Tree[i].Buttons.Count; k++)
                            {
                                RadioButtonsById.Remove(Tree[i].Buttons[k].ObjId);
                                CopiedEntries.Add(Tree[i].Buttons[k]);
                            }
                            Tree.RemoveAt(i);
                            return;
                        }
                    }
                    return;

                case TObjectType.OptionButton:
                    for (int i = NewEntries.Count - 1; i >= 0; i--)
                    {
                        if (NewEntries[i] == ms)
                        {
                            NewEntries.RemoveAt(i);
                            return;
                        }
                    }
                    for (int i = CopiedEntries.Count - 1; i >= 0; i--)
                    {
                        if (CopiedEntries[i] == ms)
                        {
                            CopiedEntries.RemoveAt(i);
                            return;
                        }
                    }

                    TRadioGroupCache rb;
                    if (RadioButtonsById.TryGetValue(ms.ObjId, out rb))
                    {
                        RadioButtonsById.Remove(ms.ObjId);
                        for (int i = rb.Buttons.Count - 1; i >= 0; i--)
                        {
                            if (rb.Buttons[i] == ms)
                            {
                                rb.Buttons.RemoveAt(i);
                                break;
                            }
                        }
                    }
                    return;
            }

        }

        internal void Add(TEscherClientDataRecord ms, bool Copied)
        {
            switch ((TObjectType)((ms.ClientData as TMsObj).ObjType))
            {
                case TObjectType.GroupBox:
                    TRadioGroupCache rg = new TRadioGroupCache(ms);
                    Tree.Add(rg);
                    return;

                case TObjectType.OptionButton:
                    List<TEscherClientDataRecord> Entries = Copied ? CopiedEntries : NewEntries;
                    Entries.Add(ms);
                    return;
            }
        }

        internal void BuildTree(IRowColSize Workbook)
        {
            ProcessNewEntries(Workbook);
            ProcessCopiedEntries(NewEntries, Workbook); //just in case it was not defined.
            ProcessCopiedEntries(CopiedEntries, Workbook);
        }

        private void ProcessCopiedEntries(List<TEscherClientDataRecord> Entries, IRowColSize Workbook)
        {
            for (int i = 0; i < Entries.Count; i++)
            {
                TEscherClientDataRecord obj = Entries[i];
                AddToTree(obj, true, Workbook);
            }
            Entries.Clear();
        }

        private void ProcessNewEntries(IRowColSize Workbook)
        {
            Dictionary<int, TEscherClientDataRecord> OtherRb = new Dictionary<int, TEscherClientDataRecord>();

            for (int i = NewEntries.Count - 1; i >= 0; i--)
            {
                TEscherClientDataRecord obj = NewEntries[i];
                if (obj.IsFirstRadioButton)
                {
                    if (AddToTree(obj, false, Workbook))
                    {
                        NewEntries.RemoveAt(i);
                    }
                }
                else
                {
                    OtherRb.Add(obj.ObjId, obj);
                }
            }

            for (int i = 0; i < Tree.Count; i++) //0 is root
            {
                if (Tree[i].Buttons.Count == 1)
                {
                    TEscherClientDataRecord Current = Tree[i].Buttons[0];
                    do
                    {
                        int NextId = Current.NextRbId;
                        if (NextId == 0) break; //if not careful, we could have an infinite loop here. for example is NextId = NextNextId, or maybe NextId = NextNextNextId

                        if (!OtherRb.TryGetValue(NextId, out Current)) break;
                        OtherRb.Remove(NextId); //to avoid infinite loop commented above.
                        NewEntries.Remove(Current);
                        Tree[i].Buttons.Add(Current);
                        RadioButtonsById[NextId] = Tree[i];
                    } while (true);
                }
            }
        }

        private bool AddToTree(TEscherClientDataRecord obj, bool MightNotBeFirst, IRowColSize Workbook)
        {
            TEscherOPTRecord CbObj = (obj.Parent as TEscherSpContainerRecord).FindRec<TEscherOPTRecord>() as TEscherOPTRecord;
            TClientAnchor CbAnchor = TDrawing.GetObjectAnchorAbsolute(CbObj, Workbook);

            for (int i = 1; i < Tree.Count; i++) // 0 is desktop
            {
                TEscherOPTRecord GrpObj = (Tree[i].Group.Parent as TEscherSpContainerRecord).FindRec<TEscherOPTRecord>() as TEscherOPTRecord;
                TClientAnchor GAnchor = TDrawing.GetObjectAnchorAbsolute(GrpObj, Workbook);

                if ((MightNotBeFirst || Tree[i].Buttons.Count == 0) && GAnchor.Contains(CbAnchor))
                {
                    AddRadioButton(obj, i);
                    return true;
                }
            }

            if (MightNotBeFirst || Tree[0].Buttons.Count == 0)
            {
                AddRadioButton(obj, 0);
                return true;
            }

            return false;
        }

        private void AddRadioButton(TEscherClientDataRecord obj, int i)
        {
            obj.IsFirstRadioButton = Tree[i].Buttons.Count == 0;
            if (Tree[i].Buttons.Count > 0)
            {
                Tree[i].Buttons[Tree[i].Buttons.Count - 1].NextRbId = obj.ObjId;
                (obj.ClientData as TMsObj).RemoveObjectFmlas(ft.CblsFmla, Tree[i].Buttons[0].ClientData as TMsObj); //only the first in the group might have a formula.

            }

            Tree[i].Buttons.Add(obj);
            obj.NextRbId = Tree[i].Buttons[0].ObjId; //docs say 0 can be here too, but excel writes a circular list. Both seem to work.
            RadioButtonsById[obj.ObjId] = Tree[i];
        }

        internal bool Find(TEscherSpContainerRecord R, out List<TEscherClientDataRecord> BtnGroup, IRowColSize Workbook)
        {
            BtnGroup = null;
            if (R == null) return false;
            TEscherClientDataRecord Cd = R.FindRec<TEscherClientDataRecord>() as TEscherClientDataRecord;
            if (Cd == null) return false;
            BuildTree(Workbook);

            TRadioGroupCache rgc;
            if (!RadioButtonsById.TryGetValue(Cd.ObjId, out rgc)) return false;

            BtnGroup = rgc.Buttons;
            return true;
        }

        internal void FixLinks(ExcelFile Workbook, int ActiveSheet)
        {
            BuildTree(Workbook);
            for (int i = 0; i < Tree.Count; i++)
            {
                if (Tree[i].Buttons.Count > 0)
                {
                    TMsObj ms = Tree[i].Buttons[0].ClientData as TMsObj;
                    if (ms != null)
                    {
                        TCellAddress LinkedCell = ms.GetObjectLink(Workbook);
                        if (LinkedCell != null)
                        {
                            CheckRadioButtons(Workbook, ActiveSheet, i, LinkedCell);
                        }
                    }
                }
            }
        }

        private void CheckRadioButtons(ExcelFile Workbook, int ActiveSheet, int i, TCellAddress LinkedCell)
        {
            int Sheet = ActiveSheet + 1;
            if (!string.IsNullOrEmpty(LinkedCell.Sheet)) Sheet = Workbook.GetSheetIndex(LinkedCell.Sheet, false);
            if (Sheet > 0)
            {
                int XF = 1;
                object r = Workbook.GetCellValue(Sheet, LinkedCell.Row, LinkedCell.Col, ref XF);
                if (r == null) { CheckGroup(i, 0); return; }
                if (r is TFlxFormulaErrorValue)
                {
                    if ((TFlxFormulaErrorValue)r == TFlxFormulaErrorValue.ErrNA) { CheckGroup(i, 0); return; }
                    //if it isn't n/a, its value doesn't matter. rb stays as is.
                }
                else
                {
                    double pd;
                    if (TBaseParsedToken.ExtToDouble(r, out pd)) //something like a string doesn't matter
                    {
                        if (pd <= 0 || pd >= int.MaxValue) { CheckGroup(i, 0); return; } // in this case it does matter, all cbs are unselected. 
                        int p = (int)pd;
                        CheckGroup(i, p);
                    }
                }
            }
        }

        private void CheckGroup(int TreePos, int cbPos)
        {
            for (int i = 0; i < Tree[TreePos].Buttons.Count; i++)
            {
                (Tree[TreePos].Buttons[i].ClientData as TMsObj).SetCheckbox(i + 1 == cbPos ? TCheckboxState.Checked : TCheckboxState.Unchecked);
            }
        }

        internal void ChangeFmla(ExcelFile Workbook, TEscherClientDataRecord Cd, TCellAddress LinkedCellAddress, TParsedTokenList LinkedCellFmla, bool ReadingXlsx)
        {
            //The line below would interfere with xlsx loading since it would build the tree and maybe remove links
            //before group boxes have been created.
            if (!ReadingXlsx) BuildTree(Workbook); //This avoids that 2 calls to different objects could have no secuential effects.
            TRadioGroupCache rgc;
            if (!RadioButtonsById.TryGetValue(Cd.ObjId, out rgc))
            {
                //It is not in the built part.
                (Cd.ClientData as TMsObj).SetObjFormulaLink(Workbook, LinkedCellAddress, LinkedCellFmla);
                return; 
            }
            (rgc.Buttons[0].ClientData as TMsObj).SetObjFormulaLink(Workbook, LinkedCellAddress, LinkedCellFmla);
        }
    }

    class TLinkedOpts
    {
        public TEscherOPTRecord Opt;
        public TLinkedOpts Next;
    }

    internal class TEscherOptByNameCache
    {
        Dictionary<string, TLinkedOpts> OptsByName;
        Dictionary<long, TEscherOPTRecord> OptsByShapeId;
        List<TEscherOPTRecord> NotLoadedOpt;

        public TEscherOptByNameCache()
        {
            OptsByName = new Dictionary<string, TLinkedOpts>(StringComparer.InvariantCultureIgnoreCase);
            OptsByShapeId = new Dictionary<long, TEscherOPTRecord>();
            NotLoadedOpt = new List<TEscherOPTRecord>();
        }

        internal void AddShapeId(TEscherOPTRecord Opt)
        {
            long ShapeId = Opt.ShapeId();
            if (ShapeId < 0) NotLoadedOpt.Add(Opt); else OptsByShapeId[ShapeId] = Opt;
        }

        internal void AddShapeName(TEscherOPTRecord Opt)
        {
            if (string.IsNullOrEmpty(Opt.ShapeName)) return;

            TLinkedOpts NeOpt = new TLinkedOpts();
            NeOpt.Opt = Opt;
            TLinkedOpts FirstOpt;
            if (OptsByName.TryGetValue(Opt.ShapeName, out FirstOpt))
            {
                NeOpt.Next = FirstOpt.Next;
                FirstOpt.Next = NeOpt;
            }
            else
            {
                NeOpt.Next = null;
                OptsByName.Add(Opt.ShapeName, NeOpt);
            }
        }

        internal void Remove(TEscherOPTRecord Opt, bool AlsoRemoveShapeId)
        {
            FillNotLoaded();
           if (AlsoRemoveShapeId)  OptsByShapeId.Remove(Opt.ShapeId());

            if (Opt == null || String.IsNullOrEmpty(Opt.ShapeName)) return;
            TLinkedOpts FirstOpt;
            if (!OptsByName.TryGetValue(Opt.ShapeName, out FirstOpt)) return;

            TLinkedOpts current = FirstOpt;
            TLinkedOpts previous = null;
            while (current != null)
            {
                if (current.Opt == Opt) break;
                previous = current;
                current = current.Next;
            }

            if (current == null) return;
            if (previous != null)
            {
                previous.Next = current.Next;
            }
            else
            {
                TLinkedOpts next = current.Next;
                OptsByName.Remove(Opt.ShapeName);
                if (next != null) OptsByName.Add(Opt.ShapeName, next);
            }
        }

        private void FillNotLoaded()
        {
            foreach (TEscherOPTRecord R in NotLoadedOpt)
            {
                long ShapeId = R.ShapeId();
                if (ShapeId == -1) FlxMessages.ThrowException(FlxErr.ErrInternal); 
                OptsByShapeId[ShapeId] = R;
            }

            NotLoadedOpt.Clear();
        }

        internal string Find(string objectName)
        {
            if (objectName == null) return null;
            TLinkedOpts FirstOp;
            if (!OptsByName.TryGetValue(objectName, out FirstOp)) return null;
            if (FirstOp.Next != null) return null;
            return FlxConsts.ObjectPathObjName + objectName;

        }

        internal TEscherOPTRecord FindObj(string objectName)
        {
            TLinkedOpts FirstOp;
            if (!OptsByName.TryGetValue(objectName, out FirstOp)) return null;
            if (FirstOp.Next != null) return null;
            return FirstOp.Opt;
        }

        internal TEscherOPTRecord FindObjByShapeId(string shapeId)
        {
            long SpId;
            FillNotLoaded();
            if (!TCompactFramework.TryParse(shapeId, out SpId)) return null;
            TEscherOPTRecord Result;
            if (!OptsByShapeId.TryGetValue(SpId, out Result)) return null;
            return Result;
        }

        internal TEscherOPTRecord FindObjByShapeId(long shapeId)
        {
            FillNotLoaded();
            TEscherOPTRecord Result;
            if (!OptsByShapeId.TryGetValue(shapeId, out Result)) return null;
            return Result;
        }
    }

    class ShapeComparer : IComparer<TEscherSpRecord>
    {
        #region IComparer<TEscherSpRecord> Members

        public int Compare(TEscherSpRecord x, TEscherSpRecord y)
        {
            if (x.ShapeId < y.ShapeId) return -1;
            else if (x.ShapeId > y.ShapeId) return 1; else return 0;
        }

        #endregion
    }


    internal class TEscherAnchorCache : TEscherRecordCache<TEscherBaseClientAnchorRecord>
    {
        internal TEscherAnchorCache()
        {
        }

        internal void Swap(int Index1, int Index2)
        {
            TEscherBaseClientAnchorRecord tmp = FList[Index1];
            FList[Index1]=FList[Index2];
            FList[Index2]=tmp;
            FSorted=false;
        }


        internal int VisibleCount()
        {
            return FList.Count;
        /*	int aCount=FList.Count;
            int Result = FList.Count;
            for (int i=0; i<aCount;i++)
            {
                TEscherClientAnchorRecord It = (TEscherClientAnchorRecord) FList[i];
                if (It.HasComment)
                {
                    Result--;
                    continue;
                }
                TEscherContainerRecord cr= It.Parent as TEscherContainerRecord;
                if (cr!=null)
                {
                    for (int k=0; k<cr.FContainedRecords.Count; k++)
                    {
                        TEscherClientDataRecord cd = cr.FContainedRecords[k] as TEscherClientDataRecord;
                        if (cd!=null && cd.ClientData!=null)
                        {
                            TMsObj msobj = cd.ClientData as TMsObj;
                            if (msobj!=null)
                            {
                                if (msobj.ObjType == 0x19) //Comment
                                {
                                    Result--;
                                    It.HasComment=true;
                                    break;
                                }
                            }
                        }
                    }
                }
            }
            return Result;*/
        }

        internal TEscherBaseClientAnchorRecord VisibleItem(int index)
        {
            return this[index];
/*			int VisibleIndex=0;
            int aCount = Count;
            for (int i=0; i<aCount;i++)
            {
                TEscherClientAnchorRecord It = (TEscherClientAnchorRecord) FList[i];
                if (It.HasComment)
                    continue;

                if (index==VisibleIndex)
                    return It;
                VisibleIndex++;
            }        
            FlxMessages.ThrowException(FlxErr.ErrInvalidValue, FlxParam.ObjectIndex);
            return null; //keep compiler happy.
            */
        }

        internal int VisibleIndex(int index)
        {
            return index;
/*			int VisibleIndex=0;
            int aCount = Count;
            for (int i=0; i<aCount;i++)
            {
                TEscherClientAnchorRecord It = (TEscherClientAnchorRecord) FList[i];
                if (It.HasComment)
                    continue;

                if (index==VisibleIndex)
                    return VisibleIndex;
                VisibleIndex++;
            }        
            FlxMessages.ThrowException(FlxErr.ErrInvalidValue, FlxParam.ObjectIndex);
            return -1; //keep compiler happy.
*/		}
    }

    internal class TEscherShapeCache : TEscherRecordCache<TEscherSpRecord>
    {
        private TEscherSpRecord SearchRecord = new TEscherSpRecord(0);
        private static readonly ShapeComparer StaticShapeComparer = new ShapeComparer(); //STATIC*

        internal TEscherShapeCache()
        {
            CacheComparer = StaticShapeComparer;
        }

        internal bool Find(long aShapeId, ref int Index)
        {
            if (!FSorted) Sort();
            SearchRecord.ShapeId = aShapeId;
            Index = FList.BinarySearch(0, FList.Count, SearchRecord, CacheComparer); //Only BinarySearch compatible with CF.
            bool Result = Index >= 0;
            if (Index < 0) Index = ~Index;
            return Result;
        }
    }



    internal class OPTComparer : IComparer<TEscherOPTRecord>
    {
        #region IComparer<TEscherOPTRecord> Members

        public int Compare(TEscherOPTRecord x, TEscherOPTRecord y)
        {
            return x.Row.CompareTo(y.Row);
        }

        #endregion
    }

    internal class TEscherOPTCache : TEscherRecordCache<TEscherOPTRecord>
    {
        private static readonly OPTComparer StaticOPTComparer= new OPTComparer(); //STATIC*
        internal TEscherOPTCache()
        {
            CacheComparer=StaticOPTComparer;
        }
        
        internal bool Find(int aRow, ref int Index)
        {
            if (!FSorted) Sort();
            Index=FList.BinarySearch(0, FList.Count, new TEscherOPTRecord(aRow), CacheComparer); //Only BinarySearch compatible with CF.
            bool Result=Index>=0;
            if (Index<0) Index=~Index;
            return Result;
        }

        internal void FixOPTPositions()
        {
            if (!ListModified) return;
            for (int i = 0; i < Count; i++) FList[i].PosInList = i + 1;
            ListModified = false;
        }
    }

    internal class TRBreak
    {
        public int Id;
        public int Size;
        public int AcumSize;
    }
    
    internal class TBreakList
    {
        protected List<TRBreak> FList;
        private int CurrentPos;
        private long ZeroPos;
        private byte[] FExtraData;
        private byte[] FExtraData2;

        internal TBreakList(long aZeroPos, byte[] aExtraData)
        {
            FList = new List<TRBreak>();
            ZeroPos=aZeroPos;
            CurrentPos=0;
            FExtraData = aExtraData;
            FExtraData2 = new byte[FExtraData.Length];
            Array.Copy(FExtraData,0, FExtraData2, 0, FExtraData2.Length);
            if (FExtraData2.Length>12) FExtraData2[12]=0x06;
        }

        #region Generics
        internal void Add(TRBreak a)
        {
            FList.Add(a);
        }
        internal void Insert(int index, TRBreak a)
        {
            FList.Insert(index, a);
        }

        protected void SetThis(TRBreak value, int index)
        {
            FList[index] = value;
        }

        internal TRBreak this[int index] 
        {
            get {return FList[index];} 
            set {SetThis(value, index);}
        }

        #endregion

        internal int Count
        {
            get {return FList.Count;}
        }


        internal void Add(int aId, int aSize)
        {
            TRBreak RBreak = new TRBreak();
            RBreak.Id=aId;
            RBreak.Size=aSize;
            if (Count>0) RBreak.Size+=FExtraData.Length;
            RBreak.AcumSize=RBreak.Size;
            if (Count>0) RBreak.AcumSize+=this[Count-1].AcumSize;
            Add(RBreak);
        }

        internal long AcumSize()
        {
            if ((CurrentPos>= Count) || (CurrentPos<0)) XlsMessages.ThrowException(XlsErr.ErrInternal);
            return this[CurrentPos].AcumSize + ZeroPos + XlsConsts.SizeOfTRecordHeader*(CurrentPos);
        }

        internal void AddToZeroPos(int Delta)
        {
            ZeroPos+= Delta;
        }

        internal int CurrentId
        {
            get
            {
                if ((CurrentPos>= Count) || (CurrentPos<0)) XlsMessages.ThrowException(XlsErr.ErrInternal);
                return this[CurrentPos].Id;
            }
        }

        internal int CurrentSize
        {
            get
            {
                if ((CurrentPos+1>= Count) || (CurrentPos+1<0)) XlsMessages.ThrowException(XlsErr.ErrInternal);
                return this[CurrentPos+1].Size;
            }
        }

        internal void IncCurrent()
        {
            CurrentPos++;
        }

        internal int ExtraDataLen()
        {
            return FExtraData.Length;
        }

        internal byte[] GetExtraData
        {
            get
            {
                if (CurrentPos==0) return FExtraData;
                return FExtraData2;
            }
        }
    }


}

