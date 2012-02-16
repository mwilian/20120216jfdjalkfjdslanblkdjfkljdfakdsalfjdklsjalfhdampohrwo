using System;
using System.Text;
using System.Diagnostics;
using FlexCel.Core;
using System.Collections.Generic;

namespace FlexCel.XlsAdapter
{
    internal abstract class TObjSubrecord
    {
        internal abstract void SaveToStream(IDataStream Workbook, TSaveData SaveData, int pData);
        internal abstract int TotalSize();
        internal abstract TObjSubrecord Clone();

        public abstract ft SubrecordType { get;}

        internal abstract bool HasExternRefs();

        internal static byte[] CloneArray(byte[] arr)
        {
            if (arr == null) return null;
            byte[] Result = new byte[arr.Length];
            Array.Copy(arr, 0, Result, 0, arr.Length);
            return Result;
        }

        internal virtual void ArrangeCopyRange(int RowOfs, int ColOfs, TSheetInfo SheetInfo)
        {
        }

        internal virtual void ArrangeInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
        {
        }

        internal virtual void ArrangeMoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
        {
        }

        internal virtual void ArrangeCopySheet(TSheetInfo SheetInfo)
        {
        }

        internal virtual void UpdateDeletedRanges(TDeletedRanges DeletedRanges)
        {
        }
    }

    internal class TxObjSubrecord: TObjSubrecord
    {
        internal byte[] Data;

        internal TxObjSubrecord(byte[] aData, int aPos, int aLen)
        {
            Data = new byte[aLen];
            Array.Copy(aData, aPos, Data, 0, aLen);
        }

        internal override TObjSubrecord Clone()
        {
            TxObjSubrecord Result = new TxObjSubrecord(Data, 0, Data.Length);
            return Result;
        }

        internal override int TotalSize()
        {
            if (Data == null) return 0;
            return Data.Length;
        }

        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int pData)
        {
            FixValues(SaveData.ObjectHeight, SaveData.FixPage);
            Workbook.Write(Data, Data.Length);
        }

        internal void FixValues(double ObjectHeight, bool FixPage)
        {
            if (SubrecordType == ft.Sbs)
            {
                int Min = GetWord(10);
                if (Min < 0 || Min > Int16.MaxValue) { Min = 0; SetWord(10, 0); }
                if (GetWord(12) < Min) SetWord(12, Min);
                if (GetWord(12) > Int16.MaxValue) SetWord(12, Int16.MaxValue);

                if (GetWord(8) < Min) SetWord(8, Min);
                if (GetWord(8) > GetWord(12)) SetWord(8, GetWord(12));

                if (FixPage && ObjectHeight > 0)
                {
                    int dPage = Convert.ToInt32(Math.Floor(ObjectHeight / 193.0));
                    SetWord(16, dPage);
                }
            }
        }
        public override ft SubrecordType
        {
            get
            {
                if (Data == null) return 0;
                return (ft)BitOps.GetWord(Data, 0);
            }
        }

        internal override bool HasExternRefs()
        {
            return false;
        }

        internal int GetWord(int p)
        {
            return BitOps.GetWord(Data, p);
        }

        internal void SetWord(int p, int v)
        {
            BitOps.SetWord(Data, p, v);
        }

        internal void SetBoolWord(int p, int mask, bool value)
        {
            if (value) BitOps.SetWord(Data, p, BitOps.GetWord(Data, p) | mask);
            else BitOps.SetWord(Data, p, BitOps.GetWord(Data, p) & ~mask);
        }
    }

    internal class TObjFormula
    {
        private TParsedTokenList FTokens;
        private byte[] EmbedInfo;
        
        public TObjFormula(byte[] aData, ref int p, TNameRecordList Names)
        {
            int ObjFmlaLen = BitOps.GetWord(aData, p);
            p += 2;

            if (ObjFmlaLen == 0) //No cbfmla at all.
            {
                FTokens = new TParsedTokenList(new TBaseParsedToken[0]);
            }
            else
            {
                int FmlaLen = BitOps.GetWord(aData, p) & 0x7FFF;
                p += 2 + 4;
                FTokens = TTokenManipulator.CreateObjFmlaFromBiff8(Names, 0, 0, aData, p, FmlaLen, false);
                p += FmlaLen;

                if (ObjFmlaLen > FmlaLen + 2 + 4)
                {
                    EmbedInfo = new byte[ObjFmlaLen - (FmlaLen + 2 + 4)];
                    Array.Copy(aData, p, EmbedInfo, 0, EmbedInfo.Length);
                    p += EmbedInfo.Length;
                }
            }
        }

        public TObjFormula(TParsedTokenList aTokens)
        {
            FTokens = aTokens;
            EmbedInfo = null;
        }


        public TParsedTokenList Tokens { get { return FTokens; } }


        internal void SaveToStream(IDataStream Workbook, TSaveData SaveData)
        {
            int FmlaNoArrayLen;
            byte[] ParsedData = TFormulaConvertInternalToBiff8.GetTokenData(SaveData.Globals.Names, Tokens, TFormulaType.Normal, out FmlaNoArrayLen);
            
            int s = TotalSize(ParsedData, false);
            Workbook.Write16((UInt16)(s + s % 2));
            if (s > 0)
            {
                Workbook.Write16((UInt16)ParsedData.Length);
                Workbook.Write32(0);
                Workbook.Write(ParsedData, ParsedData.Length);
                if (EmbedInfo != null) Workbook.Write(EmbedInfo, EmbedInfo.Length);
                if (s % 2 != 0) Workbook.Write(new byte[1], 1); //Padding.
            }
        }

        internal int TotalSize()
        {
            int FmlaNoArrayLen;
            byte[] ParsedData = TFormulaConvertInternalToBiff8.GetTokenData(null, Tokens, TFormulaType.Normal, out FmlaNoArrayLen);
            return TotalSize(ParsedData, true) + 2;
        }

        private int TotalSize(byte[] ParsedData, bool Pad)
        {
            int Result = ParsedData.Length;
            if (Result > 0) Result += 2 + 4; //Add fmla len and empty bytes.
            if (EmbedInfo != null) Result += EmbedInfo.Length;
            if (Pad) Result += Result % 2;

            return Result;
        }

        internal TObjFormula Clone()
        {
            TObjFormula Result = new TObjFormula(Tokens.Clone());
            Result.EmbedInfo = TFmlaObjSubrecord.CloneArray(EmbedInfo);
            return Result;
        }

        internal void SetTokens(TParsedTokenList aTokens)
        {
            FTokens = aTokens;
            //If we ever set a picfmla here, we should create/invalidate EmbedInfo depending in the tokens.
        }
    }

    internal class TFmlaObjSubrecord : TObjSubrecord
    {
        protected ft Ft;
        TObjFormula OFmla;
        byte[] Before;
        byte[] After;

        public TFmlaObjSubrecord(ft aFt, TObjFormula aOFmla, byte[] aBefore, byte[] aAfter)
        {
            Ft = aFt;
            Before = aBefore;
            After = aAfter;
            OFmla = aOFmla;
        }

        internal TParsedTokenList Tokens { get { return OFmla.Tokens; } }

        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int pData)
        {
            Workbook.Write16((UInt16)Ft);
            CalcBefOffs(pData);
            if (Before != null) Workbook.Write(Before, Before.Length);
            OFmla.SaveToStream(Workbook, SaveData);
            if (After != null) Workbook.Write(After, After.Length);

        }

 

        internal override int TotalSize()
        {
            int bef = Before == null ? 0 : Before.Length;
            int aft = After == null ? 0 : After.Length;
            return 2 + bef + aft + OFmla.TotalSize();
        }

        internal override TObjSubrecord Clone()
        {
            return new TFmlaObjSubrecord(Ft, OFmla.Clone(), CloneArray(Before), CloneArray(After));
        }

        public override ft SubrecordType
        {
            get { return Ft; }
        }

        internal override bool HasExternRefs()
        {
            return TTokenManipulator.HasExternRefs(Tokens);
        }

        internal bool IsAutoFilter()
        {
            if (Ft != ft.LbsData) return false;
            if (After == null) return false;

            int FlagPos = 4;
            bool fUseCb = (After[FlagPos] & 0x01) != 0;

            int lct = After[FlagPos + 1];
            return fUseCb && lct == 0x03;
        }

        internal bool IsSpecialDropdown()
        {
            if (Ft != ft.LbsData) return false;
            if (After == null) return false;

            int FlagPos = 4;
            bool fUseCb = (After[FlagPos] & 0x01) != 0;

            int lct = After[FlagPos + 1];
            return fUseCb && lct != 0x00;
        }
        internal override void ArrangeCopyRange(int RowOfs, int ColOfs, TSheetInfo SheetInfo)
        {
            TTokenManipulator.ArrangeInsertAndCopyRange(Tokens, new TXlsCellRange(0, 0, -1, -1), -1, -1, 0, 0,
                RowOfs, ColOfs, SheetInfo, true, null); //Sheet info doesn't have meaning on copy, except to create the Bounds cache.        
        }

        internal override void ArrangeInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
        {
            TTokenManipulator.ArrangeInsertAndCopyRange(Tokens, CellRange, -1, -1, 
                aRowCount, aColCount, 0, 0, SheetInfo, true, null);
        }

        internal override void ArrangeMoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
        {
            TTokenManipulator.ArrangeMoveRange(Tokens, CellRange, -1, -1, NewRow, NewCol, SheetInfo, null);
        }

        internal override void ArrangeCopySheet(TSheetInfo SheetInfo)
        {
            TTokenManipulator.ArrangeInsertSheets(Tokens, SheetInfo);
        }

        internal override void UpdateDeletedRanges(TDeletedRanges DeletedRanges)
        {
            TTokenManipulator.UpdateDeletedRanges(Tokens, DeletedRanges);
        }

        internal int GetAfterWord(int p)
        {
            return BitOps.GetWord(After, p);
        }

        internal void SetAfterWord(int p, int v)
        {
            BitOps.SetWord(After, p, v);
        }

        internal void SetAfterBoolWord(int p, int mask, bool value)
        {
            if (value) BitOps.SetWord(After, p, BitOps.GetWord(After, p) | mask);
            else BitOps.SetWord(After, p, BitOps.GetWord(After, p) & ~mask);
        }

        internal TCellAddress Address(ExcelFile xls)
        {
            TCellAddressRange Result = Range(xls);
            if (Result == null) return null;

            return Result.TopLeft;
        }

        internal TCellAddressRange Range(ExcelFile xls)
        {
            Tokens.ResetPositionToLast();
            TBaseParsedToken token = Tokens.LightPop();

            TRef3dToken r3d = token as TRef3dToken;
            if (r3d != null)
            {
                string ebook; int s1; int s2;
                r3d.ParseExternSheet(xls, out ebook, out s1, out s2);
                if (s1 > 0 && s1 <= xls.SheetCount)
                {
                    string sname = xls.GetSheetName(s1);
                    return new TCellAddressRange(
                        new TCellAddress(sname, r3d.Row + 1, r3d.Col + 1, r3d.RowAbs, r3d.ColAbs),
                        new TCellAddress(sname, r3d.Row + 1, r3d.Col + 1, r3d.RowAbs, r3d.ColAbs)
                    );
                }
                return null;
            }

            TArea3dToken a3d = token as TArea3dToken;
            if (a3d != null)
            {
                string ebook; int s1; int s2;
                a3d.ParseExternSheet(xls, out ebook, out s1, out s2);
                if (s1 > 0 && s1 <= xls.SheetCount)
                {
                    string sname = xls.GetSheetName(s1);
                    return new TCellAddressRange(
                        new TCellAddress(sname, a3d.Row1 + 1, a3d.Col1 + 1, a3d.RowAbs1, a3d.ColAbs1),
                        new TCellAddress(sname, a3d.Row2 + 1, a3d.Col2 + 1, a3d.RowAbs2, a3d.ColAbs2)
                    );
                }
                return null;
            }

            TRefToken r = token as TRefToken;
            if (r != null)
            {
                return new TCellAddressRange(
                    new TCellAddress(r.Row + 1, r.Col + 1, r.RowAbs, r.ColAbs),
                    new TCellAddress(r.Row + 1, r.Col + 1, r.RowAbs, r.ColAbs)
                );
            }

            TAreaToken a = token as TAreaToken;
            if (a != null)
            {
                return new TCellAddressRange(
                    new TCellAddress(a.Row1 + 1, a.Col1 + 1, a.RowAbs1, a.ColAbs1),
                    new TCellAddress(a.Row2 + 1, a.Col2 + 1, a.RowAbs2, a.ColAbs2)
                );
            }
            return null;
        }

        internal void SetFormula(TParsedTokenList aTokens)
        {
            OFmla.SetTokens(aTokens);
        }

        private void CalcBefOffs(int pData)
        {
            if (Ft == ft.PictFmla)
            {
                BitOps.SetWord(Before, 0, OFmla.TotalSize() + After.Length);
            }
            if (Ft == ft.LbsData) //This is not really documented, but it looks like the number here is MaxRecordSize - 4 - pData  = 0x201c - PData
            {
                BitOps.SetWord(Before, 0, 0x201C - pData);
            }

        }

        internal bool CanBeFullyRemoved()
        {
            return After == null && Before == null;
        }

        internal void RemoveTokens()
        {
            TParsedTokenList aTokens = new TParsedTokenList(new TBaseParsedToken[0]);
            SetFormula(aTokens);
        }
    }

    /// <summary>
    /// Object Record. This record interacts with MSODRAWING ones.
    /// </summary>
    internal class TObjRecord : TBaseRecord
    {
        List<TObjSubrecord> SubRecords;
        Dictionary<ft, int> SubRecordPos;
        static Dictionary<ft, int> ObjPos = CreateObjPos();

        internal TContinueRecord Continue; //It can handle part of the next MSODrawing record.

        internal TObjRecord(int aId, byte[] aData, TNameRecordList Names): this()
        { 
            int p = 0;
            while (p < aData.Length)
            {
                ft srId = (ft)BitOps.GetWord(aData, p);

                if (srId == 0) //ftend
                {
                    Add(new TxObjSubrecord(aData, p, aData.Length - p));
                    p = aData.Length;
                    break;
                }

                int cb = BitOps.GetWord(aData, p + 2);

                switch (srId)
                {
                    case ft.SbsFmla:
                    case ft.CblsFmla:
                    case ft.Macro:
                        int z = p + 2;
                        Add(new TFmlaObjSubrecord(srId, new TObjFormula(aData, ref z, Names), null, null));
                        break;

                    case ft.PictFmla:
                        int zPic = p + 4;
                        TObjFormula OFmla = new TObjFormula(aData, ref zPic, Names);
                        Add(new TFmlaObjSubrecord(srId, OFmla, new byte[] { aData[p + 2], aData[p + 3] }, GetPictRemains(aData, zPic, p + 4 + cb)));
                        break;

                    case ft.LbsData:
                        int zLbs = p + 4;
                        TObjFormula OFmla2 = new TObjFormula(aData, ref zLbs, Names);
                        Add(new TFmlaObjSubrecord(srId, OFmla2, new byte[] { aData[p + 2], aData[p + 3] }, GetLbsRemains(aData, ref zLbs, ObjType == TObjectType.ComboBox)));
                        p = zLbs;
                        cb = -4;
                        break;

                    case ft.Cmo:
                        for (int i = 6; i < cb; i++)
                            aData[p + 4 + i] = 0;  //if we dont clear those reserved bits excel might crash. (for example when deleting the text form an autoshape)
                        Add(new TxObjSubrecord(aData, p, cb + 4));
                        break;

                    default:
                        Add(new TxObjSubrecord(aData, p, cb + 4));
                        break;
                }

                p += cb + 4;
            }

        }

        private byte[] GetLbsRemains(byte[] aData, ref int zLbs, bool HasDropData)
        {
            int p = aData.Length;
            byte[] Result = new byte[p - zLbs];
            Array.Copy(aData, zLbs, Result, 0, Result.Length);
            zLbs = p;
            return Result;

            /* This should be the correct way to do it, but it doesn't work. The count reported can be wrong.
             * So we will just use the fact that lbsdata is the last record when it is present.
            int p = zLbs + 8;
            if (HasDropData)
            {
                p += 6;
                int sl = (int)StrOps.GetStrLen(true, aData, p, false, 0);

                sl += sl % 2;
                p += sl;
            }

            bool fValidPlex = (BitOps.GetWord(aData, zLbs + 4) & 0x01) != 0;
            int cLines = BitOps.GetWord(aData, zLbs) & 0x7FFF;

            if (fValidPlex)
            {
                for (int i = 0; i < cLines; i++)
                {
                    int sl = (int)StrOps.GetStrLen(true, aData, p, false, 0);
                    p += sl;
                }
            }

            int wListType = BitOps.GetWord(aData, zLbs + 4) & 0x30;
            bool Hasbsels = wListType != 0;
            if (Hasbsels)
            {
                p += cLines * 1;
            }



            byte[] Result = new byte[p - zLbs];
            Array.Copy(aData, zLbs, Result, 0, Result.Length);
            zLbs = p;
            return Result;*/
        }

        private byte[] GetPictRemains(byte[] aData, int StartPos, int FinalPos)
        {
            int sz = FinalPos - StartPos;
            byte[] Result = new byte[sz];
            Array.Copy(aData, StartPos, Result, 0, sz);

            return Result;
        }

        public TObjRecord()
        {
            SubRecords = new List<TObjSubrecord>();
            SubRecordPos = new Dictionary<ft, int>();
        }

        TxObjSubrecord cmo
        {
            get
            {
                int index;
                if (!SubRecordPos.TryGetValue(ft.Cmo, out index)) return null;
                return (TxObjSubrecord)SubRecords[index];
            }
        }

        TObjSubrecord rec(ft aFt)
        {
            int index;
            if (!SubRecordPos.TryGetValue(aFt, out index)) return null;
            return SubRecords[index];
        }

        internal TxObjSubrecord recx(ft aFt)
        {
            return rec(aFt) as TxObjSubrecord;
        }
        
        internal TFmlaObjSubrecord recf(ft aFt)
        {
            return rec(aFt) as TFmlaObjSubrecord;
        }
 
        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
        {
            Workbook.WriteHeader((int)xlr.OBJ,(UInt16)TotalSizeNoHeaders());
            long pStart = Workbook.Position;
            foreach (TObjSubrecord sr in SubRecords)
            {
                sr.SaveToStream(Workbook, SaveData, (int)(Workbook.Position - pStart));
            }
        }

        internal override int TotalSize()
        {
            return TotalSizeNoHeaders() + XlsConsts.SizeOfTRecordHeader;
        }

        internal override int TotalSizeNoHeaders()
        {
            int Size = 0;
            foreach (TObjSubrecord sr in SubRecords)
            {
                Size += sr.TotalSize();
            }
            return Size;
        }

        internal bool IsAutoFilter()
        {
            int LbsPos;
            if (!SubRecordPos.TryGetValue(ft.LbsData, out LbsPos)) return false;

            return ((TFmlaObjSubrecord)SubRecords[LbsPos]).IsAutoFilter();
        }

        internal bool IsSpecialDropdown()
        {
            int LbsPos;
            if (!SubRecordPos.TryGetValue(ft.LbsData, out LbsPos)) return false;

            return ((TFmlaObjSubrecord)SubRecords[LbsPos]).IsSpecialDropdown();
        }

        public int ObjId 
        { 
            get
            {
                return BitOps.GetWord(cmo.Data, 6);
            }
            set
            {
                BitOps.SetWord(cmo.Data, 6, value);
            }
        }

        public int ObjFlags
        {
            get
            {
                return BitOps.GetWord(cmo.Data, 8);
            }
            set
            {
                BitOps.SetWord(cmo.Data, 8, value);
            }
        }

        public TObjectType ObjType
        {
            get
            {
                return (TObjectType)BitOps.GetWord(cmo.Data, 4);
            }
        }


        public bool HasPictFormula
        {
            get
            {
                return SubRecordPos.ContainsKey(ft.PictFmla);
            }
        }

        internal override int GetId
        {
            get { return (int)xlr.OBJ; }
        }

        internal override void AddContinue(TContinueRecord aContinue)
        {
            if (Continue != null) XlsMessages.ThrowException(XlsErr.ErrInvalidContinue);
            Continue = aContinue;
        }

        protected override TBaseRecord DoCopyTo(TSheetInfo SheetInfo)
        {
            TObjRecord Result = new TObjRecord();
            Result.SubRecords = new List<TObjSubrecord>();
            Result.SubRecordPos = new Dictionary<ft, int>();
            foreach (TObjSubrecord sr in SubRecords)
            {
                Result.SubRecordPos.Add(sr.SubrecordType, Result.SubRecords.Count);
                Result.SubRecords.Add(sr.Clone());
            }

            return Result;
        }

        internal bool ObjsHaveExternRefs()
        {
            for (int i = 0; i < SubRecords.Count; i++)
            {
                if (SubRecords[i].HasExternRefs()) return true;
            }
            return false;
        }

        #region Insert And copy
        internal void ArrangeCopyRange(int RowOfs, int ColOfs, TSheetInfo SheetInfo)
        {
            foreach (TObjSubrecord sr in SubRecords)
            {
                sr.ArrangeCopyRange(RowOfs, ColOfs, SheetInfo);
            }
        }

        internal void ArrangeInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
        {
            foreach (TObjSubrecord sr in SubRecords)
            {
                sr.ArrangeInsertRange(CellRange, aRowCount, aColCount, SheetInfo);
            }
        }

        internal void ArrangeMoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
        {
            foreach (TObjSubrecord sr in SubRecords)
            {
                sr.ArrangeMoveRange(CellRange, NewRow, NewCol, SheetInfo);
            }
        }

        internal void ArrangeCopySheet(TSheetInfo SheetInfo)
        {
            foreach (TObjSubrecord sr in SubRecords)
            {
                sr.ArrangeCopySheet(SheetInfo);
            }
        }

        internal void UpdateDeletedRanges(TDeletedRanges DeletedRanges)
        {
            foreach (TObjSubrecord sr in SubRecords)
            {
                sr.UpdateDeletedRanges(DeletedRanges);
            }
        }
        #endregion

        #region Obj Manipulation
        internal bool SetCheckbox(TCheckboxState Value)
        {
            TxObjSubrecord FtCblsData = recx(ft.CblsData);
            if (FtCblsData == null) return false;

            int v;
            switch (Value)
            {
                case TCheckboxState.Unchecked:
                    v = 0;
                    break;
                case TCheckboxState.Checked:
                    v = 1;
                    break;
                default:
                    v = 2;
                    break;
            }

            bool Changed = FtCblsData.GetWord(4) != v;
            FtCblsData.SetWord(4, v);
            return Changed;
        }

        internal TCheckboxState GetCheckbox()
        {
            TxObjSubrecord FtCblsData = recx(ft.CblsData);
            if (FtCblsData == null) return TCheckboxState.Indeterminate;
            switch (FtCblsData.GetWord(4))
            {
                case 0: return TCheckboxState.Unchecked;
                case 1: return TCheckboxState.Checked;
                default: return TCheckboxState.Indeterminate;
            }
        }

        public bool Obj3D 
        { 
            get
            {
                TxObjSubrecord FtCblsData = recx(ft.CblsData);
                if (FtCblsData != null)
                {
                    return (FtCblsData.GetWord(10) & 1) == 0;
                }

                TxObjSubrecord FtGrpBox = recx(ft.GboData);
                if (FtGrpBox != null)
                {
                    return (FtGrpBox.GetWord(8) & 1) == 0;
                }

                TFmlaObjSubrecord FtLbls = recf(ft.LbsData);
                if (FtLbls != null)
                {
                    return (FtLbls.GetAfterWord(4) & 8) == 0;
                }

                return true;
            }
            set
            {
                TxObjSubrecord FtCblsData = recx(ft.CblsData);
                if (FtCblsData != null)
                {
                    FtCblsData.SetBoolWord(10, 1, !value);
                }

                TxObjSubrecord FtGrpBox = recx(ft.GboData);
                if (FtGrpBox != null)
                {
                    FtGrpBox.SetBoolWord(8, 1, !value);
                }

                TFmlaObjSubrecord FtLbls = recf(ft.LbsData);
                if (FtLbls != null)
                {
                    FtLbls.SetAfterBoolWord(4, 8, !value);
                }

            }
        }

        #endregion

        #region Add


        private static Dictionary<ft, int> CreateObjPos()
        {
            Dictionary<ft, int> Result = new Dictionary<ft, int>();
            Result.Add(ft.Cmo, 0); 
            Result.Add(ft.Gmo, 1); 
            Result.Add(ft.Cf, 2); //picFormat
            Result.Add(ft.PioGrbit, 3); //picFlags
            Result.Add(ft.Cbls, 4); //cbls
            Result.Add(ft.Rbo, 5); //rbo
            Result.Add(ft.Sbs, 6); //sbs
            Result.Add(ft.Nts, 7); //nts
            Result.Add(ft.Macro, 8); //macro
            Result.Add(ft.PictFmla, 9); //PicFmla
            Result.Add(ft.CblsFmla, 10); //objLinkFmla
            Result.Add(ft.SbsFmla, 11); //objLinkFmla 2
            Result.Add(ft.CblsData, 12); //cbls

            Result.Add(ft.RboData, 14); //rboData
            Result.Add(ft.EdoData, 15); //edoData
            Result.Add(ft.LbsData, 16); //lbsData
            Result.Add(ft.GboData, 17); //gboData

            return Result;
        }

        private bool ObjIdGoesAfter(ft ObjId, ft BaseId)
        {
            if (ObjId == 0) return true;
            if (BaseId == 0) return false;
            int ObjIdPos; int FtIdPos;
            if (!ObjPos.TryGetValue(ObjId, out ObjIdPos)) return false;
            if (!ObjPos.TryGetValue(BaseId, out FtIdPos)) return true; //it shouldn't really happen
            return ObjIdPos > FtIdPos;

        }
        private void Add(TObjSubrecord sr)
        {
            Debug.Assert(!SubRecordPos.ContainsKey(sr.SubrecordType));
            int i = 0; 
            while (i < SubRecords.Count &&  ObjIdGoesAfter(sr.SubrecordType, SubRecords[i].SubrecordType))
            {
                i++;   
            }
            SubRecords.Insert(i, sr);
            SubRecordPos[sr.SubrecordType] = i;
            for (int k = i + 1; k < SubRecords.Count; k++)
            {
                SubRecordPos[SubRecords[k].SubrecordType] = k;
            }
        }

        #endregion
        #region Obj Formula

        internal void SetObjectFormula(ft aFt, TParsedTokenList aTokens)
        {
            TFmlaObjSubrecord f = recf(aFt);
            if (f == null) Add(new TFmlaObjSubrecord(aFt, new TObjFormula(aTokens), CalcBefore(aFt), CalcAfter(aFt))); 
            else f.SetFormula(aTokens);
        }

        private byte[] CalcAfter(ft aFt)
        {
            switch (aFt)
            {
                case ft.Macro: return null;
                case ft.PictFmla: return null;
                
                case ft.SbsFmla: 
                case ft.CblsFmla:
                    return null;

                case ft.LbsData:
                    return new byte[]{0x00, 0x00,   
                                      0x00, 0x00, 0x08, 0x00, 
                                      0x00, 0x00, 0x00, 0x00,   
                                      0x08, 0x00, 0x00, 0x00};

                default:
                    return null;
            }
        }

        private byte[] CalcBefore(ft aFt)
        {
            switch (aFt)
            {
                case ft.Macro: return null;
                case ft.PictFmla: return new byte[2]; //must have the full len.
                
                case ft.SbsFmla:
                case ft.CblsFmla:
                    return null;
                
                case ft.LbsData:
                    return new byte[] { 0xEE, 0x1F };

                default:
                    return null;
            }
        }

        internal void RemoveFmla(ft aFt)
        {
            TFmlaObjSubrecord f = recf(aFt);
            if (f == null) return;
            if (f.CanBeFullyRemoved())
            {
                int index = SubRecordPos[aFt];
                SubRecords.RemoveAt(index);
                SubRecordPos.Remove(aFt);

                for (int i = 0; i < SubRecords.Count; i++)
                {
                    ft srt = SubRecords[i].SubrecordType;
                    if (SubRecordPos[srt] > index) SubRecordPos[srt]--;
                }
            }
            else
            {
                f.RemoveTokens();
            }

        }

        #endregion
    }

    /// <summary>
    /// Generic Object.
    /// </summary>
    internal class TMsObj: TBaseClientData
    {
        private TObjRecord FObjRecord;
        private TFlxChart FChart;
        private TxBaseRecord FImData;

        #region Empty objects
        private static readonly byte[] EmptyMsObjImg = 
            {
                0x15, 0x00, 0x12, 0x00,   0x08, 0x00, 0x01, 0x00,   0x11, 0x60, 0x00, 0x00, //Note that here ObjId=1. This has to be changed later on code
                0x00, 0x00, 0x00, 0x00,   0x00, 0x00, 0x00, 0x00,   0x00, 0x00, // ftCmo
                0x07, 0x00, 0x02, 0x00,   0xFF, 0xFF, //ftCf
                0x08, 0x00, 0x02 ,0x00,   0x01, 0x00, //ftPioGrbit
                0x00, 0x00, 0x00, 0x00  //ftEnd
            };

            private static readonly byte[] EmptyMsObjShape = 
            {
               0x15, 0x00, 0x12, 0x00,    0x1E, 0x00, 0x01, 0x00,   0x11, 0x60, 0x00, 0x00, 
               0x00, 0x00, 0x00, 0x00,    0x00, 0x00, 0x00, 0x00,   0x00, 0x00, //ftCmo. 0x1e might be diff in a rectangle or line
        
               0x00, 0x00, 0x00, 0x00  //ftEnd
            };

            private static readonly byte[] EmptyMsObjGroup = 
            {
               0x15, 0x00, 0x12, 0x00,    0x00, 0x00, 0x01, 0x00,   0x11, 0x60, 0x00, 0x00, 
               0x00, 0x00, 0x00, 0x00,    0x00, 0x00, 0x00, 0x00,   0x00, 0x00, //ftCmo.

               0x06, 0x00, 0x02, 0x00,    0x00, 0x00, //ftGmo
        
               0x00, 0x00, 0x00, 0x00  //ftEnd
            };

        private static readonly byte[] EmptyMsObjNote =
            {
                0x15, 0x00, 0x12, 0x00,   0x19, 0x00, 0x01, 0x00,   0x11, 0x40, 0x00, 0x00, //Note that here ObjId=1. This has to be changed later on code
                0x00, 0x00, 0x00, 0x00,   0x00, 0x00, 0x00, 0x00,   0x00, 0x00, // ftCmo
                
                0x0D, 0x00, 0x16, 0x00,
                0x00, 0x12, 0x23, 0x8C,   0x6C, 0x50, 0x8B, 0x4D,   0xA8, 0xC5, 0x7B, 0x64, 
                0xFF, 0xA8, 0xC5, 0xA3,   0x00, 0x00, 0x10, 0x00,   0x00, 0x00,  //ftNts
                
                0x00, 0x00, 0x00, 0x00 //ftEnd
            };

        private static readonly byte[] EmptyMsObjAutoFilter =
            {
                0x15, 0x00, 0x12, 0x00,   0x14, 0x00, 0x01, 0x00,   0x01, 0x21, 0x00, 0x00, //Note that here ObjId=1. This has to be changed later on code
                0x00, 0x00, 0x00, 0x00,   0x00, 0x00, 0x00, 0x00,   0x00, 0x00, // ftCmo
                
                0x0C, 0x00, 0x14, 0x00,   0x00, 0x00, 0x00, 0x00,   0x00, 0x00, 0x00, 0x00,
                0x64, 0x00, 0x01, 0x00,   0x0A, 0x00, 0x00, 0x00,   0x10, 0x00, 0x01, 0x00, //ftSbs  Scroll bar data
                
                0x13, 0x00, 0xEE, 0x1F,   0x00, 0x00, 0x00, 0x00,   0x04, 0x00, 0x01, 0x03, 
                0x00, 0x00, 0x02, 0x00,   0x08, 0x00, 0x6C, 0x00,   //List box data. this record has a negative length, but Excel writes it this way.
                0x00, 0x00, 0x00, 0x00 //no ftEnd in list?
            };

        private static readonly byte[] EmptyMsObjComboBox =
            {
                0x15, 0x00, 0x12, 0x00,   0x14, 0x00, 0x01, 0x00,   0x11, 0x00, 0x00, 0x00, //changes "Print" with autofilter.
                0x00, 0x00, 0x00, 0x00,   0x00, 0x00, 0x00, 0x00,   0x00, 0x00, // ftCmo
                
                0x0C, 0x00, 0x14, 0x00,   0x00, 0x00, 0x00, 0x00,   0x00, 0x00, 0x00, 0x00, 
                0x00, 0x00, 0x01, 0x00,   0x01, 0x00, 0x00, 0x00,   0x10, 0x00, 0x01, 0x00, //ftSbs  Scroll bar data. Changes iMax and dPage with autofilter.

                0x13, 0x00, 0xEE, 0x1F,   0x00, 0x00, 0x00, 0x00,   0x00, 0x00, 0x08, 0x00, 
                0x00, 0x00, 0x00, 0x00,   0x08, 0x00, 0x00, 0x00,   //List box data. this record has a negative length, but Excel writes it this way.

                0x00, 0x00, 0x00, 0x00 // no ftEnd in list, this is part of listboxdata.

            };

        private static readonly byte[] EmptyMsObjListBox =
            {
                0x15, 0x00, 0x12, 0x00,   0x12, 0x00, 0x01, 0x00,   0x11, 0x00, 0x00, 0x00, //changes "Print" with autofilter.
                0x00, 0x00, 0x00, 0x00,   0x00, 0x00, 0x00, 0x00,   0x00, 0x00, // ftCmo
                
                0x0C, 0x00, 0x14, 0x00,   0x00, 0x00, 0x00, 0x00,   0x00, 0x00, 0x00, 0x00, 
                0x00, 0x00, 0x01, 0x00,   0x01, 0x00, 0x00, 0x00,   0x10, 0x00, 0x01, 0x00, //ftSbs  Scroll bar data. Changes iMax and dPage with autofilter.

                0x13, 0x00, 0xEE, 0x1F,   0x00, 0x00, 0x00, 0x00,   0x00, 0x00, 0x08, 0x00, 
                0x00, 0x00 // no ftEnd in list, this is part of listboxdata.

            };


        private static readonly byte[] EmptyMsObjCheckbox =
            {
                0x15, 0x00, 0x12, 0x00,   0x0B, 0x00, 0x01, 0x00,   0x11, 0x00, 0x00, 0x00, //Note that here ObjId=1. This has to be changed later on code
                0x00, 0x00, 0x00, 0x00,   0x00, 0x00, 0x00, 0x00,   0x00, 0x00, // ftCmo
                
                0x0A, 0x00, 0x0C, 0x00,   0x00, 0x00, 0x00, 0x00,   0x00, 0x00, 0x00, 0x00,   0x00, 0x00, 0x00, 0x00,  //ftcbls ignored
                0x12, 0x00, 0x08, 0x00,   0x00, 0x00, 0x00, 0x00,   0x00, 0x00, 0x03, 0x00,  // FtCblsData
                0x00, 0x00, 0x00, 0x00 //ftEnd

            };

        private static readonly byte[] EmptyMsObjRadioButton =
            {
                0x15, 0x00, 0x12, 0x00,   0x0C, 0x00, 0x01, 0x00,   0x11, 0x00, 0x00, 0x00, //Note that here ObjId=1. This has to be changed later on code
                0x00, 0x00, 0x00, 0x00,   0x00, 0x00, 0x00, 0x00,   0x00, 0x00, // ftCmo

                0x0A, 0x00, 0x0C, 0x00,   0x00, 0x00, 0x00, 0x00,   0x00, 0x00, 0x00, 0x00,   0x00, 0x00, 0x00, 0x00,  //ftcbls ignored
                0x0B, 0x00, 0x06, 0x00,   0x00, 0x00, 0x00, 0x00,   0x00, 0x00, //ftRbo
                0x12, 0x00, 0x08, 0x00,   0x00, 0x00, 0x00, 0x00,   0x00, 0x00, 0x03, 0x00, // FtCblsData
                0x11, 0x00, 0x04, 0x00,   0x00, 0x00, 0x00, 0x00,   0x00, 0x00, 0x00, 0x00, //ftRboData 

                0x00, 0x00, 0x00, 0x00 //ftEnd

            };

        private static readonly byte[] EmptyMsObjGroupBox =
            {
                0x15, 0x00, 0x12, 0x00,   0x13, 0x00, 0x01, 0x00,   0x11, 0x00, 0x00, 0x00, //Note that here ObjId=1. This has to be changed later on code
                0x00, 0x00, 0x00, 0x00,   0x00, 0x00, 0x00, 0x00,   0x00, 0x00, // ftCmo

                0x0F, 0x00, 0x06, 0x00,   0x00, 0x00, 0x00, 0x00,   0x00, 0x00, //FtGroup
                0x00, 0x00, 0x00, 0x00 //ftEnd

            };


        private static readonly byte[] EmptyMsObjButton =
            {
                0x15, 0x00, 0x12, 0x00,   0x07, 0x00, 0x01, 0x00,   0x01, 0x40, 0x00, 0x00, //Note that here ObjId=1. This has to be changed later on code 
                0x00, 0x00, 0x00, 0x00,   0x00, 0x00, 0x00, 0x00,   0x00, 0x00, // ftCmo                                 
                0x00, 0x00, 0x00, 0x00 //ftEnd

            };

        private static readonly byte[] EmptyMsObjLabel =
            {
                0x15, 0x00, 0x12, 0x00,   0x0E, 0x00, 0x01, 0x00,   0x11, 0x00, 0x00, 0x00, //Note that here ObjId=1. This has to be changed later on code
                0x00, 0x00, 0x00, 0x00,   0x00, 0x00, 0x00, 0x00,   0x00, 0x00, // ftCmo
                
                0x00, 0x00, 0x00, 0x00 //ftEnd

            };

        private static readonly byte[] EmptyMsObjSpinner =
            {
                0x15, 0x00, 0x12, 0x00,   0x10, 0x00, 0x01, 0x00,   0x11, 0x00, 0x00, 0x00, //Note that here ObjId=1. This has to be changed later on code
                0x00, 0x00, 0x00, 0x00,   0x00, 0x00, 0x00, 0x00,   0x00, 0x00, // ftCmo

                0x0C, 0x00, 0x14, 0x00,   0x00, 0x00, 0x00, 0x00,   0x00, 0x00, 0x00, 0x00, 
                0x64, 0x00, 0x01, 0x00,   0x01, 0x00, 0x00, 0x00,   0x10, 0x00, 0x01, 0x00, //ftSbs  Scroll bar data. 

                
                0x00, 0x00, 0x00, 0x00 //ftEnd

            };

        private static readonly byte[] EmptyMsObjScrollBar =
            {
                0x15, 0x00, 0x12, 0x00,   0x11, 0x00, 0x01, 0x00,   0x11, 0x00, 0x00, 0x00, //Note that here ObjId=1. This has to be changed later on code
                0x00, 0x00, 0x00, 0x00,   0x00, 0x00, 0x00, 0x00,   0x00, 0x00, // ftCmo

                0x0C, 0x00, 0x14, 0x00,   0x00, 0x00, 0x00, 0x00,   0x00, 0x00, 0x00, 0x00, 
                0x64, 0x00, 0x01, 0x00,   0x01, 0x00, 0x00, 0x00,   0x10, 0x00, 0x01, 0x00, //ftSbs  Scroll bar data. 

                
                0x00, 0x00, 0x00, 0x00 //ftEnd

            };

        #endregion

        internal TMsObj(){}

        #region Create Empty objects
        internal static TMsObj CreateEmptyImg(ref int MaxId, TBaseImageProperties Props)
        {
            return CreateEmpty(ref MaxId, EmptyMsObjImg, Props, false); //image   we don't read autofill/autoline from xlsx, so we give false here as ReadFromXlsx
        }

        internal static TMsObj CreateEmptyShape(ref int MaxId, TBaseImageProperties Props, bool ReadFromXlsx)
        {
            return CreateEmpty(ref MaxId, EmptyMsObjShape, Props, ReadFromXlsx); 
        }

        internal static TMsObj CreateEmptyGroup(ref int MaxId, TBaseImageProperties Props, bool ReadFromXlsx)
        {
            return CreateEmpty(ref MaxId, EmptyMsObjGroup, Props, ReadFromXlsx);
        }

        internal static TMsObj CreateEmptyNote(ref int MaxId, TImageProperties Props, bool ReadFromXlsx)
        {
            return CreateEmpty(ref MaxId, EmptyMsObjNote, Props, ReadFromXlsx); //comment
        }

        internal static TMsObj CreateEmptyAutoFilter(ref int MaxId)
        {
            return CreateEmpty(ref MaxId, EmptyMsObjAutoFilter, null, false); //combo box
        }

        internal static TMsObj CreateEmptyCheckbox(ref int MaxId, TBaseImageProperties Props)
        {
            return CreateEmpty(ref MaxId, EmptyMsObjCheckbox, Props, false); //checkbox
        }

        internal static TMsObj CreateEmptyRadioButton(ref int MaxId, TBaseImageProperties Props)
        {
            return CreateEmpty(ref MaxId, EmptyMsObjRadioButton, Props, false); //radio button
        }
        
        internal static TMsObj CreateEmptyGroupBox(ref int MaxId, TBaseImageProperties Props)
        {
            return CreateEmpty(ref MaxId, EmptyMsObjGroupBox, Props, false); //GroupBox
        }

        internal static TMsObj CreateEmptyButton(ref int MaxId, TBaseImageProperties Props)
        {
            return CreateEmpty(ref MaxId, EmptyMsObjButton, Props, false); //false
        }

        internal static TMsObj CreateEmptyComboBox(ref int MaxId, TBaseImageProperties Props)
        {
            return CreateEmpty(ref MaxId, EmptyMsObjComboBox, Props, false); //ComboBox
        }

        internal static TMsObj CreateEmptyListBox(ref int MaxId, TBaseImageProperties Props)
        {
            return CreateEmpty(ref MaxId, EmptyMsObjListBox, Props, false); 
        }

        internal static TMsObj CreateEmptyLabel(ref int MaxId, TBaseImageProperties Props)
        {
            return CreateEmpty(ref MaxId, EmptyMsObjLabel, Props, false); 
        }

        internal static TMsObj CreateEmptySpinner(ref int MaxId, TBaseImageProperties Props)
        {
            return CreateEmpty(ref MaxId, EmptyMsObjSpinner, Props, false);
        }

        internal static TMsObj CreateEmptyScrollBar(ref int MaxId, TBaseImageProperties Props)
        {
            return CreateEmpty(ref MaxId, EmptyMsObjScrollBar, Props, false);
        }

        private static TMsObj CreateEmpty(ref int MaxId, byte[] Data, TBaseImageProperties Props, bool ReadFromXlsx)
        {
            TMsObj Result = new TMsObj();

            int MyDataSize = Data.Length;
            byte[] MyData = new byte[MyDataSize];
            Data.CopyTo(MyData, 0);

            if (Props != null)
            {
                unchecked
                {
                    if (Props.Lock) MyData[8] |= 0x01; else MyData[8] &= (byte)~0x01;
                    if (Props.DefaultSize) MyData[8] |= 0x04; else MyData[8] &= (byte)~0x04;
                    if (Props.Published) MyData[8] |= 0x08; else MyData[8] &= (byte)~0x08;
                    if (Props.Print) MyData[8] |= 0x10; else MyData[8] &= (byte)~0x10;
                    if (Props.Disabled) MyData[8] |= 0x80; else MyData[8] &= (byte)~0x80;

                    if (ReadFromXlsx)
                    {
                        if (Props.AutoFill) MyData[9] |= 0x20; else MyData[9] &= (byte)~0x20;
                        if (Props.AutoLine) MyData[9] |= 0x40; else MyData[9] &= (byte)~0x40;
                    }

                }
            }

            Result.FObjRecord = new TObjRecord((int)xlr.OBJ, MyData, null);
            Result.ArrangeId(ref MaxId);

            return Result;
        }
        #endregion

        internal bool IsAutoFilter
        {
            get
            {
                //An AutoFilter object seems to be composed of 3 records: ftcmo, scrollbar data and listbox data.
                //Now, normal comboboxes also have the same 3 records. The bytes that "bind" a listbox into an AutoFilter are 
                //bytes 56 and 57, and they should be 0x01 and 0x03. 
                return FObjRecord.IsAutoFilter();
            }
        }

        /// <summary>
        /// Autofilter, datavalidation
        /// </summary>
        internal bool IsSpecialDropdown
        {
            get
            {
                return FObjRecord.IsSpecialDropdown();
            }
        }

        protected override int GetId() {  if (FObjRecord!=null) return FObjRecord.ObjId; else return 0;}
        protected override void SetId(int Value){if (FObjRecord!=null) FObjRecord.ObjId = Value;}

        public TObjectType ObjType 
        {
            get
            {
                if (FObjRecord == null) return 0;
                return FObjRecord.ObjType;
            }
        }

        int ObjFlags
        {
            get
            {
                if (FObjRecord == null) return 0;
                return FObjRecord.ObjFlags;
            }
        }

        internal bool IsLocked { get { return (ObjFlags & 0x01) != 0; } }
        internal bool IsDefaultSize { get { return (ObjFlags & 0x04) != 0; } }
        internal bool IsPublished { get { return (ObjFlags & 0x08) != 0; } }
        internal bool IsPrintable { get { return (ObjFlags & 0x10) != 0; } }
        internal bool IsDisabled { get { return (ObjFlags & 0x80) != 0; } }

        internal bool IsAutoFill { get { return (ObjFlags & 0x2000) != 0; } }
        internal bool IsAutoLine { get { return (ObjFlags & 0x4000) != 0; } }

        internal override void ArrangeId(ref int MaxId)
        {
            base.ArrangeId(ref MaxId);
            MaxId++;
            Id = MaxId;
        }

        internal override void Clear()
        {
            FObjRecord=null;
            FChart=null;
            FImData=null;
            RemainingData=null;
        }

        public bool HasPictFormula
        {
            get
            {
                if (FObjRecord == null) return false;
                return FObjRecord.HasPictFormula; 
            }
        }

        protected override TBaseClientData DoCopyTo(TSheetInfo SheetInfo)
        {
            if (HasPictFormula) XlsMessages.ThrowException(XlsErr.ErrCantCopyPictFmla);
            TMsObj Result= new TMsObj();
            Result.FObjRecord= (TObjRecord) TObjRecord.Clone(FObjRecord, SheetInfo);
            if (FChart != null)
            {
                TCopiedGen SaveCopiedGen = SheetInfo.CopiedGen;
                SheetInfo.PushCopiedGen();
                Result.FChart = (TFlxChart)TFlxChart.Clone(FChart, SheetInfo);
                SheetInfo.PopCopiedGen(SaveCopiedGen);
            }
            if (FImData!=null) Result.FImData= (TxBaseRecord)TxBaseRecord.Clone(FImData, SheetInfo);
            return Result;
        }

        internal override bool HasExternRefs()
        {
            if (FChart!=null && FChart.HasExternRefs()) return true;
            if (FObjRecord != null && FObjRecord.ObjsHaveExternRefs()) return true;
            return false;
        }

        #region Save
        internal override void LoadFromStream(TBaseRecordLoader RecordLoader, TWorkbookGlobals Globals, TBaseRecord First)
        {
            Clear();
            TObjRecord rFirst = (TObjRecord)First;
            if (rFirst.ObjType == TObjectType.Chart)
            {
                TBOFRecord R=RecordLoader.LoadRecord(false) as TBOFRecord;
                if (R == null) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
                FChart= new TFlxChart(Globals, true);
                try
                {
                    FChart.LoadFromStream(RecordLoader, R);
                }
                catch (Exception)
                {
                    FChart=null;
                    throw;
                } //except
            } 
            else
                if (rFirst.ObjType == TObjectType.Picture)
            {
                if (RecordLoader.RecordHeader.Id==(int)xlr.IMDATA) 
                    FImData=(TxBaseRecord)RecordLoader.LoadRecord(false);
            }

            RemainingData = rFirst.Continue;
            rFirst.Continue = null;

            if (FChart != null && RemainingData == null)
            {
                RemainingData = FChart.RemainingData;
                FChart.RemainingData = null;
            }

            //this must be the last statement, so if there is an exception, we dont take First
            FObjRecord= rFirst;
        }

        internal override void SaveToStream(IDataStream DataStream, TSaveData SaveData)
        {
            if (FObjRecord==null) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
      
            FObjRecord.SaveToStream(DataStream, SaveData, 0);
            if (FImData!=null) FImData.SaveToStream(DataStream, SaveData, 0);
            if (FChart!=null)  FChart.Save(DataStream, SaveData, null);
        }

        internal override long TotalSize()
        {
            if (FObjRecord == null) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
            long Result = FObjRecord.TotalSize();
            if (FChart != null) Result += FChart.TotalSize(null, false);
            if (FImData != null) Result += FImData.TotalSize();
            return Result;
        }
        #endregion

        #region InsertAndCopy
        internal override void ArrangeCopyRange(int RowOfs, int ColOfs, TSheetInfo SheetInfo)
        {
            if (FChart!=null) FChart.Chart.ArrangeCopyRange(RowOfs, ColOfs, SheetInfo);
            if (FObjRecord != null) FObjRecord.ArrangeCopyRange(RowOfs, ColOfs, SheetInfo);
        }

        internal override void ArrangeInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
        {
            if (FChart!=null) FChart.ArrangeInsertRange(CellRange, aRowCount, aColCount, SheetInfo);
            if (FObjRecord != null) FObjRecord.ArrangeInsertRange(CellRange, aRowCount, aColCount, SheetInfo); 
        }

        internal override void ArrangeMoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
        {
            if (FChart!=null) FChart.ArrangeMoveRange(CellRange, NewRow, NewCol, SheetInfo);
            if (FObjRecord != null) FObjRecord.ArrangeMoveRange(CellRange, NewRow, NewCol, SheetInfo);  
        }

        internal override void ArrangeCopySheet(TSheetInfo SheetInfo)
        {
            if (FChart!=null) FChart.ArrangeCopySheet(SheetInfo);
            if (FObjRecord != null) FObjRecord.ArrangeCopySheet(SheetInfo);  
        }

        internal override void UpdateDeletedRanges(TDeletedRanges DeletedRanges)
        {
            if (FChart!=null) FChart.UpdateDeletedRanges(DeletedRanges);
            if (FObjRecord != null) FObjRecord.UpdateDeletedRanges(DeletedRanges);
        }
        #endregion

        internal override TClientType ObjRecord()
        {
            return TClientType.TMsObj;
        }

        internal override TFlxChart Chart()
        {
            return FChart;
        }

        #region Checkboxes / radio buttons

        internal bool SetCheckbox(TCheckboxState Value)
        {
            return FObjRecord.SetCheckbox(Value);
        }

        internal TCheckboxState GetCheckbox()
        {
            return FObjRecord.GetCheckbox();
        }

        internal void SetObject3D(bool Value)
        {
            FObjRecord.Obj3D = Value;
        }

        internal bool GetObject3D()
        {
            return FObjRecord.Obj3D;
        }

        internal bool GetRbFirstInGroup()
        {
            TxObjSubrecord rboData = FObjRecord.recx(ft.RboData);
            if (rboData == null) return false;
            return (rboData.GetWord(6) & 0x01) != 0;
        }

        internal void SetRbFirstInGroup(bool Value)
        {
            TxObjSubrecord rboData = FObjRecord.recx(ft.RboData);
            if (rboData == null) return;
            rboData.SetBoolWord(6, 0x01, Value);
        }

        internal int GetRbNextId()
        {
            TxObjSubrecord rboData = FObjRecord.recx(ft.RboData);
            if (rboData == null) return 0;
            return rboData.GetWord(4);
        }

        internal void SetRbNextId(int Value)
        {
            TxObjSubrecord rboData = FObjRecord.recx(ft.RboData);
            if (rboData == null) return;
            rboData.SetWord(4, Value);
        }

        internal TComboBoxProperties GetComboProps()
        {
            TFmlaObjSubrecord lbsData = FObjRecord.recf(ft.LbsData);
            if (lbsData == null) return null;

            TComboBoxProperties Result = new TComboBoxProperties();
            Result.SelectionType = SelectionType;

            if (ObjType == TObjectType.ComboBox)
            {
                Result.DropLines = lbsData.GetAfterWord(10);
            }

            return Result;
        }

        internal bool HasSpinProps()
        {
            return FObjRecord.recx(ft.Sbs) != null;
        }

        internal TSpinProperties GetSpinProps()
        {
            TxObjSubrecord sbs = FObjRecord.recx(ft.Sbs);
            if (sbs == null) return null;
            TSpinProperties Result = new TSpinProperties();

            Result.Min = sbs.GetWord(10);
            Result.Max = sbs.GetWord(12);
            Result.Incr = sbs.GetWord(14);
            Result.Page = sbs.GetWord(16);
            Result.Dx = sbs.GetWord(20);

            return Result;
        }

        internal int GetObjectSelection()
        {
            TFmlaObjSubrecord lbsData = FObjRecord.recf(ft.LbsData);
            if (lbsData == null) return 0;
            return lbsData.GetAfterWord(2);
        }

        internal void SetObjectSelection(int v)
        {
            TFmlaObjSubrecord lbsData = FObjRecord.recf(ft.LbsData);
            if (lbsData == null) return;
            lbsData.SetAfterWord(2, v);
        }

        internal int GetObjectSpinValue(bool fix, double ObjectHeight)
        {
            TxObjSubrecord sbsData = FObjRecord.recx(ft.Sbs);
            if (sbsData == null) return 0;
            if (fix) sbsData.FixValues(ObjectHeight, ObjType == TObjectType.ListBox);
            return sbsData.GetWord(8);
        }

        internal void SetObjectSpinValue(int v)
        {
            TxObjSubrecord sbsData = FObjRecord.recx(ft.Sbs);
            if (sbsData == null) return;
            sbsData.SetWord(8, v);
        }

        internal TListBoxSelectionType SelectionType
        {
            get
            {
                TFmlaObjSubrecord lbsData = FObjRecord.recf(ft.LbsData);
                if (lbsData == null) return 0;
                return (TListBoxSelectionType)((lbsData.GetAfterWord(4) >> 4) & 0x3);
            }
            set
            {
                TFmlaObjSubrecord lbsData = FObjRecord.recf(ft.LbsData);
                if (lbsData == null) return;
                bool v1 = ((int)value & 0x1) != 0;
                bool v2 = ((int)value & 0x2) != 0;
                lbsData.SetAfterBoolWord(4, 0x10, v1);
                lbsData.SetAfterBoolWord(4, 0x20, v2);

            }
        }

        internal void SetComboProps(TComboBoxProperties cbProps)
        {
            if (cbProps == null) return;
            TFmlaObjSubrecord lbsData = FObjRecord.recf(ft.LbsData);
            if (lbsData != null)
            {
                if (ObjType == TObjectType.ComboBox)
                {
                    lbsData.SetAfterWord(10, cbProps.DropLines);
                }

                SelectionType = cbProps.SelectionType;
            }

            SetObjectSelection(cbProps.Sel);
        }

        internal void SetSpinProps(TSpinProperties spinProps)
        {
            if (spinProps == null) return;
            TxObjSubrecord sbs = FObjRecord.recx(ft.Sbs);
            if (sbs == null) return;

            sbs.SetWord(14, spinProps.Incr);
            sbs.SetWord(16, spinProps.Page);
            sbs.SetWord(20, spinProps.Dx);

            if (ObjType == TObjectType.ScrollBar || ObjType == TObjectType.Spinner)
            {
                sbs.SetWord(10, spinProps.Min);
                sbs.SetWord(12, spinProps.Max);
            }
        }


        ft GetFmlaLinkId()
        {
            switch (ObjType)
            {
                case TObjectType.Spinner:
                case TObjectType.ScrollBar:
                case TObjectType.ListBox:
                case TObjectType.ComboBox:
                    return ft.SbsFmla;
               
            }

            return ft.CblsFmla;
        }

        internal TCellAddress GetObjectLink(ExcelFile xls)
        {
            TFmlaObjSubrecord obj = FObjRecord.recf(GetFmlaLinkId());
            if (obj == null || obj.Tokens.Count == 0) return null;
            return obj.Address(xls);
        }

        internal TCellAddressRange GetObjectInputRange(ExcelFile xls)
        {
            TFmlaObjSubrecord obj = FObjRecord.recf(ft.LbsData);
            if (obj == null || obj.Tokens.Count == 0) return null;
            return obj.Range(xls);
        }

        internal string GetObjectMacro(TCellList CellList)
        {
            TFmlaObjSubrecord obj = FObjRecord.recf(ft.Macro);
            if (obj == null || obj.Tokens.Count == 0) return null;
            return TFormulaConvertInternalToText.AsString(obj.Tokens, 0, 0, CellList, 
                CellList.Globals, FlxConsts.Max_FormulaStringConstant, false, true);
        }

        private string GetObjFmlaXlsx(TCellList CellList, ft aFt)
        {
            TFmlaObjSubrecord obj = FObjRecord.recf(aFt);
            if (obj == null || obj.Tokens.Count == 0) return null;
            return TFormulaConvertInternalToText.AsString(obj.Tokens, 0, 0, CellList, 
                CellList.Globals, FlxConsts.Max_FormulaStringConstant, true);
        }

        internal string GetFmlaLinkXlsx(TCellList CellList)
        {
            return GetObjFmlaXlsx(CellList, GetFmlaLinkId());
        }

        internal string GetFmlaRangeXlsx(TCellList CellList)
        {
            return GetObjFmlaXlsx(CellList, ft.LbsData);
        }

        internal string GetFmlaMacroXlsx(TCellList CellList)
        {
            return GetObjFmlaXlsx(CellList, ft.Macro);
        }

        private void SetObjFormula(ft aFt, TParsedTokenList TokenList)
        {
            if (TokenList == null)
            {
                RemoveObjectFmlas(aFt, null);
                return;
            }

            FObjRecord.SetObjectFormula(aFt, TokenList);
        }

        internal void SetObjFormulaLink(ExcelFile xls, TCellAddress LinkedCellAddress, TParsedTokenList LinkedCellFmla)
        {
            TParsedTokenList TokenList = null;
            if (LinkedCellFmla != null) TokenList = LinkedCellFmla;
            else if (LinkedCellAddress != null) TokenList = GetTokensFromAddress(xls, LinkedCellAddress);

            SetObjFormula(GetFmlaLinkId(), TokenList);
        }

        internal void SetObjFormulaRange(ExcelFile xls, TCellAddressRange InputRange, TParsedTokenList LinkedCellFmla)
        {
            TParsedTokenList TokenList = null;
            if (LinkedCellFmla != null) TokenList = LinkedCellFmla;
            else if (InputRange != null) TokenList = GetTokensFromRange(xls, InputRange);

            SetObjFormula(ft.LbsData, TokenList);
        }

        internal void SetObjFormulaMacro(ExcelFile xls, TParsedTokenList TokenList)
        {
            SetObjFormula(ft.Macro, TokenList);
        }

        private static TParsedTokenList GetTokensFromAddress(ExcelFile xls, TCellAddress LinkedCellAddress)
        {
            TRefToken r;
            if (string.IsNullOrEmpty(LinkedCellAddress.Sheet))
            {
                r = new TRefToken(ptg.Ref, LinkedCellAddress.Row - 1, LinkedCellAddress.Col - 1, LinkedCellAddress.RowAbsolute, LinkedCellAddress.ColAbsolute);

            }
            else
            {
                r = new TRef3dToken(ptg.Ref3d, xls.GetExternSheet(LinkedCellAddress.Sheet, false), LinkedCellAddress.Row - 1, LinkedCellAddress.Col - 1,
                    LinkedCellAddress.RowAbsolute, LinkedCellAddress.ColAbsolute);
            }

            TParsedTokenList TokenList = new TParsedTokenList(new TBaseParsedToken[] { r });
            return TokenList;
        }

        private static TParsedTokenList GetTokensFromRange(ExcelFile xls, TCellAddressRange InputRange)
        {
            TAreaToken r;
            if (InputRange.TopLeft.Sheet != InputRange.BottomRight.Sheet)
            {
                XlsMessages.ThrowException(XlsErr.ErrRangeMustHaveSameSheet, InputRange.TopLeft.Sheet, InputRange.BottomRight.Sheet);
            }
            if (string.IsNullOrEmpty(InputRange.TopLeft.Sheet))
            {
                r = new TAreaToken(ptg.Area, InputRange.TopLeft.Row - 1, InputRange.TopLeft.Col - 1, InputRange.TopLeft.RowAbsolute, InputRange.TopLeft.ColAbsolute,
                    InputRange.BottomRight.Row - 1, InputRange.BottomRight.Col - 1, InputRange.BottomRight.RowAbsolute, InputRange.BottomRight.ColAbsolute);
            }
            else
            {
                r = new TArea3dToken(ptg.Area3d, xls.GetExternSheet(InputRange.TopLeft.Sheet, false), 
                    InputRange.TopLeft.Row - 1, InputRange.TopLeft.Col - 1, InputRange.TopLeft.RowAbsolute, InputRange.TopLeft.ColAbsolute,
                    InputRange.BottomRight.Row - 1, InputRange.BottomRight.Col - 1, InputRange.BottomRight.RowAbsolute, InputRange.BottomRight.ColAbsolute);
            }

            TParsedTokenList TokenList = new TParsedTokenList(new TBaseParsedToken[] { r });
            return TokenList;
        }


        internal void RemoveObjectFmlas(ft aFt, TMsObj MoveToObj)
        {
            if (MoveToObj != null)
            {
                 TFmlaObjSubrecord obj = FObjRecord.recf(aFt);
                 if (obj != null)
                 {
                     MoveToObj.SetObjFormula(aFt, obj.Tokens);
                 }
            }
            FObjRecord.RemoveFmla(aFt);
        }

        #endregion

    }

    /// <summary>
    /// TextBox Object.
    /// </summary>
    internal class TTXO: TBaseClientData
    {
        private TTXORecord FTXO;
        private IFlexCelFontList FFontList;

        internal TTXO(IFlexCelFontList aFontList)
        {
            FFontList=aFontList;
        }

        /// <summary>
        /// Create from data
        /// </summary>
        /// <param name="aFontList">Pointer to the FontList. It should be a pointer, so it can be modified.</param>
        /// <param name="Dummy"></param>
        internal TTXO(IFlexCelFontList aFontList, int Dummy)
            : this(aFontList)
        {
            FTXO=new TTXORecord();
            //No need for continue recods
        }

        protected void ScanRecord(ref string Value, ref TRTFRun[] RTFRuns)
        {
            Value = String.Empty; RTFRuns = new TRTFRun[0];
            if (FTXO.Continue == null) return;
            int Len = BitOps.GetWord(FTXO.Data, 10);

            if (Len == 0)
            {
                RemainingData = FTXO.Continue;
                FTXO.Continue = null;
                return;
            }
            TxBaseRecord TxtRec = FTXO.Continue;
            int aPos = 1;
            if (TxtRec != null)
            {
                byte OptionFlags = (byte)(TxtRec.Data[0] & 1); byte ActualOptionFlags = OptionFlags;
                StringBuilder s = new StringBuilder();
                StrOps.ReadStr(ref TxtRec, ref aPos, s, OptionFlags, ref ActualOptionFlags, Len);
                Value = s.ToString();
            }
            else Value = String.Empty;

            Len = BitOps.GetWord(FTXO.Data, 12) / 8;
            RTFRuns = new TRTFRun[Len]; //last is TXOFinalRun
            if (TxtRec != null) TxtRec = TxtRec.Continue;
            aPos = 0;
            byte[] r = new byte[8];
            for (int i = 0; i < Len; i++) 
            {
                BitOps.ReadMem(ref TxtRec, ref aPos, r);
                RTFRuns[i].FirstChar = BitOps.GetWord(r, 0);
                RTFRuns[i].FontIndex = BitOps.GetWord(r, 2) & 0x3FF;
            }
            if ((TxtRec != null) && (TxtRec.Continue != null))
            {
                RemainingData = TxtRec.Continue;
                TxtRec.Continue = null;
            }
            else RemainingData = null;
        }

        internal override void Clear()
        {
            FTXO=null;
            RemainingData=null;
        }

        protected override TBaseClientData DoCopyTo(TSheetInfo SheetInfo)
        {
            IFlexCelFontList aFontList = FFontList;
            if (SheetInfo != null && SheetInfo.DestGlobals != null && SheetInfo.DestGlobals.Fonts != null) aFontList = SheetInfo.DestGlobals.Workbook;
            TTXO Result= new TTXO(aFontList);
            if (FTXO!=null) Result.FTXO = (TTXORecord) TTXORecord.Clone(FTXO, SheetInfo);

            if (FTXO.Continue != null && FTXO.Continue.Continue != null)
                TChartBaseRecord.AdaptRTF(0, 8, SheetInfo.SourceGlobals, SheetInfo.DestGlobals, FTXO.Continue.Continue.Data, ref Result.FTXO.Continue.Continue.Data);
            return Result;
        }

        internal override void LoadFromStream(TBaseRecordLoader RecordLoader, TWorkbookGlobals Globals, TBaseRecord First)
        {
            FTXO=First as TTXORecord;
            //We have to search for any Continues that do not apply here.
            GetText();
        }

        internal override void SaveToStream(IDataStream DataStream, TSaveData SaveData)
        {
            if (FTXO!=null) FTXO.SaveToStream(DataStream, SaveData, 0);
        }

        internal override long TotalSize()
        {
            return FTXO.TotalSize();
        }

        internal override void ArrangeInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
        {
            //Nothing
        }

        internal override void ArrangeMoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
        {
            //nothing
        }


        internal override void ArrangeCopySheet(TSheetInfo SheetInfo)
        {
            //Nothing
        }

        internal override void UpdateDeletedRanges(TDeletedRanges DeletedRanges)
        {
            //Nothing
        }


        internal override TClientType ObjRecord()
        {
            return TClientType.TTXO;
        }

        internal TFlxFont GetFontList(int fontIndex)
        {
            return FFontList.GetFont(fontIndex);
        }

        internal TFlxFont[] CreateFontList()
        {
            TFlxFont[] Result= new TFlxFont[FFontList.FontCount+1];
            int p=0;
            for (int i=0; i<FFontList.FontCount;i++)
            {
                if (i==4) 
                {
                    p++;
                    Result[4]=GetFontList(4);
                }
                Result[i+p]=GetFontList(i+p);
            }
            return Result;
        }

        internal int OptionFlags
        {
            get
            {
                return BitOps.GetWord(FTXO.Data, 0);
            }
            set
            {
                BitOps.SetWord(FTXO.Data, 0, value);
            }
        }

        internal int Rotation
        {
            get
            {
                return BitOps.GetWord(FTXO.Data, 2);
            }
            set
            {
                BitOps.SetWord(FTXO.Data, 2, value);
            }
        }

            internal TRichString GetText()
            {
                string Result=String.Empty;
                TRTFRun[] RTFRuns=null;
                ScanRecord(ref Result, ref RTFRuns);
                return new TRichString(Result, RTFRuns, CreateFontList());  //(We create it here, because it can dinamically change)
            }

            internal void SetText(TRichString value)
            {
                string sValue = value.Value;
                if (sValue == null) sValue = String.Empty;
                int Len = sValue.Length;
                int RTFCount = value.RTFRunCount;

                while (RTFCount > 0 && value.RTFRun(RTFCount - 1).FirstChar > Len) RTFCount--; //Characters longer than the string len should not be copied.
                TRTFRun[] RTFRun = new TRTFRun[RTFCount];
                for (int i = 0; i < RTFCount; i++) RTFRun[i] = value.RTFRun(i);

                if (RTFRun.Length <= 0 || RTFRun[0].FirstChar != 0) RTFCount++;
                if (RTFRun.Length <= 0 || RTFRun[RTFRun.Length - 1].FirstChar != Len) RTFCount++;
                //We need always at least 2 RTFRuns.

                BitOps.SetWord(FTXO.Data, 10, Len); //length of text
                if (Len > 0) BitOps.SetWord(FTXO.Data, 12, 8 * RTFCount); else BitOps.SetWord(FTXO.Data, 12, 0); //length of formatting runs
                FTXO.Continue = null;
                if (Len > 0)
                {
                    if (StrOps.IsWide(sValue))
                    {
                        byte[] Dat = new byte[Len * 2 + 1];
                        Dat[0] = 1;
                        Encoding.Unicode.GetBytes(sValue).CopyTo(Dat, 1);
                        FTXO.Continue = new TContinueRecord((int)xlr.CONTINUE, Dat);
                    }
                    else
                    {
                        byte[] Dat = new byte[Len + 1];
                        Dat[0] = 0;
                        bool b = StrOps.CompressUnicode(sValue, Dat, 1);
                        Debug.Assert(b);
                        FTXO.Continue = new TContinueRecord((int)xlr.CONTINUE, Dat);
                    }

                    int RTFLen = RTFCount * 8;
                    byte[] Dat2 = new byte[RTFLen];
                    int pos = 0;
                    if (RTFRun.Length <= 0 || RTFRun[0].FirstChar != 0)
                    {
                        //first run will have the XF = 0.
                        pos++;
                    }
                    if (RTFRun.Length <= 0 || RTFRun[RTFRun.Length - 1].FirstChar != Len)
                    {
                        BitOps.SetWord(Dat2, Dat2.Length - 8, Len);
                    }

                    for (int i = 0; i < RTFRun.Length; i++)
                    {
                        BitOps.SetWord(Dat2, (i + pos) * 8, RTFRun[i].FirstChar);
                        int FontIndex = RTFRun[i].FontIndex; if (FontIndex < 0 || FontIndex >= value.MaxFontIndex) FontIndex = 0;
                        BitOps.SetWord(Dat2, (i + pos) * 8 + 2, FFontList.AddFont(value.GetFont(FontIndex)));
                    }

                    FTXO.Continue.Continue = new TContinueRecord((int)xlr.CONTINUE, Dat2);

                }
            }
    }
}
