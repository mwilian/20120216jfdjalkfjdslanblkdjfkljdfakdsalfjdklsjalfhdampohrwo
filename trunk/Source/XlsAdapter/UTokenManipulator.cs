using System;
using System.Diagnostics;
using System.Text;
using System.Globalization;

using FlexCel.Core;

namespace FlexCel.XlsAdapter
{
    internal sealed class TTokenManipulator
    {
        private TTokenManipulator(){}

        public static TParsedTokenList CreateFromBiff8(TNameRecordList Names, int aRow, int aCol, byte[] aData, int aStart, int aLen, bool aRelative3dRanges, out bool HasSubtotal, out bool HasAggregate) 
        {
            TFormulaConvertBiff8ToInternal rpn = new TFormulaConvertBiff8ToInternal();
            return rpn.ParseRPN(Names, aRow, aCol, aData, aStart, aLen, aRelative3dRanges, out HasSubtotal, out HasAggregate, false);
        }

        internal static TParsedTokenList CreateObjFmlaFromBiff8(TNameRecordList Names, int aRow, int aCol, byte[] aData, int aStart, int aLen, bool aRelative3dRanges) 
        {
            TFormulaConvertBiff8ToInternal rpn = new TFormulaConvertBiff8ToInternal();
            bool HasSubtotal; bool HasAggregate;
            return rpn.ParseRPN(Names, aRow, aCol, aData, aStart, aLen, aRelative3dRanges, out HasSubtotal, out HasAggregate, true);
        }



        #region External References
        internal static bool HasExternRefs(TParsedTokenList Data)
        {
            Data.ResetPositionToLast();
            while (!Data.Bof())
            {
                TBaseParsedToken r = Data.LightPop();
                ptg id = r.GetBaseId;
                // This check is a little simplistic because an Area3d or Ref3d might in fact refer to the same sheet. But then, the externsheet is not copied so  
                // the reference will be invalid anyway. The "right" thing to do would be to convert external refs to the same sheet to external refs on the new sheet.
                if (id == ptg.Area3d || id == ptg.Ref3d || id == ptg.NameX) return true;
            }

            return false;
        }

        internal static bool HasExternLinks(TParsedTokenList Data, TReferences References)
        {
            Data.ResetPositionToLast();
            while (!Data.Bof())
            {
                TBaseParsedToken r = Data.LightPop();
                ptg id = r.GetBaseId;

                if (id == ptg.Area3d || id == ptg.Ref3d || id == ptg.NameX)
                {
                    if (References != null)
                    {
                        int ExternSheet = r.ExternSheet;
                        if (!References.IsLocalSheet(ExternSheet)) return true;
                    }
                }
            }

            return false;
        }
        #endregion

        #region InsertAndMove
        internal static void ArrangeInsertAndCopyRange(TParsedTokenList Data, TXlsCellRange CellRange, int FmlaRow, int FmlaCol, int aRowCount, int aColCount, int CopyRowOffset, int CopyColOffset, TSheetInfo SheetInfo, bool aAllowedAbsolute, TFormulaBounds Bounds)
        {
            new TInsertTokens(CellRange, aRowCount, aColCount, CopyRowOffset, CopyColOffset, 
                SheetInfo, false, aAllowedAbsolute, FmlaRow, FmlaCol, Bounds).Go(Data, SheetInfo.SourceFormulaSheet == SheetInfo.InsSheet);
        }

        internal static void ArrangeMoveRange(TParsedTokenList Data, TXlsCellRange CellRange, int FmlaRow, int FmlaCol, int NewRow, int NewCol, TSheetInfo SheetInfo, TFormulaBounds Bounds)
        {
            new TMoveTokens(CellRange, FmlaRow, FmlaCol, NewRow, NewCol, SheetInfo, Bounds).Go(Data, SheetInfo.SourceFormulaSheet == SheetInfo.InsSheet);
        }

        internal static void ArrangeInsertSheets(TParsedTokenList Data, TSheetInfo SheetInfo)
        {
            new TInsertTokens(new TXlsCellRange(0,0,-1,-1), 0, 0, 0, 0, 
                SheetInfo, true,  true, 0, 0, null).Go(Data, SheetInfo.SourceFormulaSheet == SheetInfo.InsSheet);
        }
        #endregion

        #region Misc
        internal static void UpdateDeletedRanges(TParsedTokenList Data, TDeletedRanges DeletedRanges)
        {
            Data.ResetPositionToLast();
            while (!Data.Bof())
            {
                TBaseParsedToken tk = Data.LightPop();
                TBaseFunctionToken ft = tk as TBaseFunctionToken;
                if (ft != null)
                {
                    //we need to ensure we don't delete the used _xlfn. ranges. Used def fn don't need to check, because they use the name in the tokenlist.
                    if (ft.GetFunctionData().FutureInXls)
                    {
                        int NameId = DeletedRanges.Names.GetNamePos(-1, ft.GetFunctionData().FutureName);
                        if (NameId >= 0) DeletedRanges.Reference(NameId); //No need for recursion here, this name can't use anything else. Also, we don't need to update refs to this range.
                    }

                    continue;
                }

                TNameToken r = tk as TNameToken; //this includes namex
                if (r == null) continue;
                if (r.GetBaseId == ptg.NameX && !DeletedRanges.References.IsLocalSheet(r.ExternSheet)) return;  //This name does not point to a name in the NAME table.

                if (DeletedRanges.Update)
                {
                    UpdateRange(r, DeletedRanges);
                }
                else
                {
                    FindReferences(r, DeletedRanges);
                }
            }
        }

        private static void UpdateRange(TNameToken r, TDeletedRanges DeletedRanges)
        {
            int NameId = r.NameIndex - 1;
            if (NameId < 0 || NameId >= DeletedRanges.Count) return;

            Debug.Assert(DeletedRanges.Referenced(NameId), "Can't delete ranges that have references. Excel does not do it, and doesn't provide a way to create invalid references for ranges.");

            int ofs = DeletedRanges.Offset(NameId);
            if (ofs == 0) return;

            NameId -= ofs;
            Debug.Assert(NameId >= 0);

            r.NameIndex = NameId + 1;

        }

        private static void FindReferences(TNameToken r, TDeletedRanges DeletedRanges)
        {
            int NameId = r.NameIndex - 1;
            if (NameId < 0 || NameId >= DeletedRanges.Count) return;
            if (DeletedRanges.Referenced(NameId)) return; //Avoid infinite loop if one range refers to other and vice versa.
            DeletedRanges.Reference(NameId);

            //We also need to recursively find all other names referenced by this name.
            TNameRecord Name = DeletedRanges.Names[NameId];
            Name.UpdateDeletedRanges(DeletedRanges);
        }

        internal static void ArrangeSharedFormulas(TParsedTokenList Data, int Row, int Col, bool FromBiff8)
        {
            Data.ResetPositionToLast();
            while (!Data.Bof())
            {
                TBaseParsedToken r = Data.LightPop();
                
                switch (r.GetBaseId)
                {
                    case ptg.RefN:
                    case ptg.AreaN:
                        Data.UnShare(Row, Col, FromBiff8);
                        break;
                }
            }
        }

        internal static int TotalSizeWithArray(TParsedTokenList Data, TFormulaType FmlaType)
        {
            if (Data.TextLenght > FlxConsts.Max_FormulaLen97_2003) FlxMessages.ThrowException(FlxErr.ErrFormulaTooLong, TFormulaConvertInternalToText.AsString(Data, 0, 0,null));
            int DataLenNoArray;
            byte[] bData = TFormulaConvertInternalToBiff8.GetTokenData(null, Data, FmlaType, out DataLenNoArray);
            return bData.Length;
        }

        internal static void SaveToStream(TNameRecordList Names, IDataStream Workbook, TFormulaType FmlaType, TParsedTokenList Data, bool WriteLen)
        {
            if (Data.TextLenght > FlxConsts.Max_FormulaLen97_2003) FlxMessages.ThrowException(FlxErr.ErrFormulaTooLong, TFormulaConvertInternalToText.AsString(Data, 0, 0, null));
            int FmlaNoArrayLen;
            byte[] bData = TFormulaConvertInternalToBiff8.GetTokenData(Names, Data, FmlaType, out FmlaNoArrayLen);
            if (WriteLen) Workbook.Write16((UInt16)FmlaNoArrayLen);
            Workbook.Write(bData, bData.Length);
        }

        #endregion
    }
    
    internal abstract class TInsertOrMovetokens
    {
        protected abstract void Do3D(TBaseParsedToken tk, out bool RefIsInInsertingSheet);
        protected abstract void DoName(TNameToken aName);
        protected abstract void DoNameX(TNameXToken aNamex);
        protected abstract void DoRef(TRefToken reft, bool RefIsInInsertingSheet);
        protected abstract void DoArea(TAreaToken areat, bool RefIsInInsertingSheet);
        protected abstract void DoMemArea(TParsedTokenList Tokens, TMemAreaToken areat, bool RefIsInInsertingSheet);

        protected abstract void DoTable(TTableToken table);
        protected abstract void DoArrayFmla(TExp_Token exp);

        internal void Go(TParsedTokenList Data, bool RefIsInInsertingSheet)
        {
            Data.ResetPositionToStart();
            while (!Data.Eof())
            {
                TBaseParsedToken r = Data.ForwardPop();
                ptg id = r.GetBaseId;
                switch (id)
                {
                    case ptg.Exp:
                        DoArrayFmla((TExp_Token)r);
                        break;
                    case ptg.Tbl:
                        TTableToken tbl = r as TTableToken;
                        if (tbl != null) //might be also TTableObjToken
                        {
                            DoTable(tbl);
                        }
                        break;
                    case ptg.Name:
                        DoName((TNameToken)r);
                        break;

                    case ptg.RefErr:
                    case ptg.Ref:
                    case ptg.RefN:
                        DoRef((TRefToken)r, RefIsInInsertingSheet);
                        break;

                    case ptg.AreaErr:
                    case ptg.Area:
                    case ptg.AreaN:
                        DoArea((TAreaToken)r, RefIsInInsertingSheet);
                        break;

                    case ptg.MemArea:
                        DoMemArea(Data, (TMemAreaToken)r, RefIsInInsertingSheet);
                        break;

                    case ptg.NameX:
                        {
                            //We'll do Do3d only for local names.
                            //bool ThisRefIsInInsertingSheet;
                            //Do3D(r, out ThisRefIsInInsertingSheet);
                            DoNameX((TNameXToken)r);
                        }
                        break;

                    case ptg.Ref3d:
                    case ptg.Ref3dErr:
                        {
                            bool ThisRefIsInInsertingSheet;
                            Do3D(r, out ThisRefIsInInsertingSheet);
                            DoRef((TRef3dToken)r, ThisRefIsInInsertingSheet);
                            break;
                        }

                    case ptg.Area3dErr:
                    case ptg.Area3d:
                        {
                            bool ThisRefIsInInsertingSheet;
                            Do3D(r, out ThisRefIsInInsertingSheet);
                            DoArea((TArea3dToken)r, ThisRefIsInInsertingSheet);
                        }
                        break;
                }
            }
        }


        protected static void IncRowColForTableArray(ref int RowCol, int InsPos, int Offset, int Max, bool CheckInside)
        {
            long w = RowCol;
            //Handle deletes...
            if (CheckInside && (Offset < 0) && (InsPos >= 0) && (w >= InsPos) && (w < InsPos - Offset))
            {
                XlsMessages.ThrowException(XlsErr.ErrCantMovePartOfArrayFormula); 
                //we should never have an error when updating tables or exp tokens. The error should show when deleting the cell, but that would be to much costly in performance.
                return;
            }

            w += Offset;

            if ((w < 0) || (w > Max))
            {
                XlsMessages.ThrowException(XlsErr.ErrCantMovePartOfArrayFormula); //we should never have an error when updating tables or exp tokens.
                return;
            }

            RowCol = (int)w;
        }

        protected static int MoveRef(TBaseRefToken reft, ref int RowCol, int AbsRowCol, int InsPos, int Offs, int FmlaRowCol, bool RowRel, bool RefInside, bool FmlaInside, int MaxRowCol, bool CheckInside)       
        {
            if (RowRel)
            {
                if (FmlaRowCol < 0) return AbsRowCol; //this is a name or similar, in this case we don't update refs.
                if (RefInside ^ FmlaInside)
                {
                    if (CheckInside) //Here only matters if RefInside is true. If FmlaInside is true, the ref will be destroyed anyway.
                    {
                        if (InsPos >= 0 && RowCol >= InsPos && RowCol < InsPos - Offs) reft.CreateInvalidRef(); //deleted relative ref
                    }

                    int rofs = FmlaInside ? -Offs : Offs;
                    int a = TBaseRefToken.WrapSigned(RowCol + rofs, MaxRowCol);
                    RowCol = a;
                    AbsRowCol = TBaseRefToken.WrapSigned(FmlaRowCol + a, MaxRowCol);
                }
            }
            else
            {
                if (RefInside)
                {
                    reft.IncRowCol(ref RowCol, InsPos, Offs, MaxRowCol, CheckInside);
                    AbsRowCol = RowCol;
                }
            }
            return AbsRowCol;
        }

    }

    internal class TInsertTokens : TInsertOrMovetokens
    {
        #region Variables
        TXlsCellRange CellRange;
        int FmlaRow;
        int FmlaCol;
        int RowCount;
        int ColCount;
        int CopyColOffset;
        int CopyRowOffset;
        TSheetInfo SheetInfo;
        bool InsertingSheet;
        bool AllowedAbsolute;
        TFormulaBounds Bounds;
        bool SemiAbsoluteMode; //When copying in semi-absolute, absolute formulas inside the range will move.
        #endregion

        #region Constructor
        internal TInsertTokens(TXlsCellRange aCellRange, int aRowCount, int aColCount, int aCopyRowOffset, int aCopyColOffset,
            TSheetInfo aSheetInfo, bool aInsertingSheet, bool aAllowedAbsolute, int aFmlaRow, int aFmlaCol, TFormulaBounds aBounds)
        {
            CellRange = aCellRange;
            FmlaRow = aFmlaRow;
            FmlaCol = aFmlaCol;
            RowCount = aRowCount;
            ColCount = aColCount;
            CopyRowOffset = aCopyRowOffset;
            CopyColOffset = aCopyColOffset;
            SheetInfo = aSheetInfo;
            InsertingSheet = aInsertingSheet;
            AllowedAbsolute = aAllowedAbsolute;
            SemiAbsoluteMode = aSheetInfo.SemiAbsoluteMode;
            Bounds = aBounds;
        }
        #endregion

        protected override void Do3D(TBaseParsedToken tk, out bool RefIsInInsertingSheet)
        {
            RefIsInInsertingSheet = false;

            if (SheetInfo.SourceReferences != null)
            {

                if (SheetInfo.DestReferences != SheetInfo.SourceReferences) //Copy the external refs to the new file.
                {
                    int refpos = SheetInfo.DestReferences.CopySheet(tk.ExternSheet, SheetInfo);
                    if (refpos < 0)
                    {
                        //CreateInvalidRef(ref Data[tkPos]);
                        //SetWord(Data, tPos, SheetInfo.DestReferences.SetSheet(SheetInfo.DestFormulaSheet)); //Ensure we have a valid externsheet. A reference to a non existing externsheet will raise an error in Excel, even if the reference is not valid.
                        tk.ExternSheet = SheetInfo.DestReferences.AddSheet(SheetInfo.DestGlobals.SheetCount, 0xFFFF); //A reference to a deleted sheet.
                    }
                    else
                    {
                        tk.ExternSheet = refpos;        //this copies external refs to the old sheet to the new sheet
                        if (Bounds != null) Bounds.AddSheets(SheetInfo.DestReferences.GetAllSheets(refpos));
                    }
                }
                else
                {
                    int SingleLocalSheet = SheetInfo.SourceReferences.GetJustOneSheet(tk.ExternSheet);
                    RefIsInInsertingSheet = SingleLocalSheet == SheetInfo.InsSheet;

                    if (InsertingSheet && (SingleLocalSheet == SheetInfo.SourceFormulaSheet)) //We will only convert those names that reference a single sheet.
                        tk.ExternSheet = SheetInfo.SourceReferences.AddSheet(SheetInfo.SourceGlobals.SheetCount, SheetInfo.InsSheet);        //this copies external refs to the old sheet to the new sheet

                    if (Bounds != null) Bounds.AddSheets(SheetInfo.DestReferences.GetAllSheets(tk.ExternSheet));
                }
            }
            else
            {
                Debug.Assert(Bounds == null);  //we would have problems if references is null and we try to use Bounds.
            }
        }

        private bool FindDestNameInLocalNames(int DestSheet, TNameToken aName, TNameRecord Name)
        {
            for (int i = SheetInfo.DestNames.Count - 1; i >= 0; i--)
            {
                TNameRecord NewName = SheetInfo.DestNames[i];
                if (NewName.RangeSheet == DestSheet &&
                    String.Equals(NewName.Name, Name.Name, StringComparison.CurrentCultureIgnoreCase))
                {
                    aName.NameIndex = i + 1;
                    return true;
                }
            }
            return false;
        }

        protected override void DoName(TNameToken aName)
        {
            if (SheetInfo.SourceReferences != null && SheetInfo.SourceNames != null)
            {
                TNameRecord Name = SheetInfo.SourceNames[aName.NameIndex - 1];
                int NameSheet = Name.RangeSheet;


                if (SheetInfo.SourceReferences != SheetInfo.DestReferences) //Copy the names to the new file.
                {
                    //Copy the name to the new reference.
                    if (!FindDestNameInLocalNames(SheetInfo.DestFormulaSheet, aName, Name))  //Search in local names
                    {
                        // If not found in local names, we will add a new one, even if it is defined as global. So if
                        // we copy to sheets with each one a ref to a name in the sheet, both reference different things.

                        //This would cause infinite recursion if the name is "A = A" (or A=B and B=A)
                        //TNameRecord NewName = Name.CopyTo(SheetInfo.DestFormulaSheet, CopyRowOffset, CopyColOffset, SheetInfo);

                        int DestSheet = SheetInfo.DestFormulaSheet;
                        if (Name.RangeSheet < 0 && (Name.IsAddin || Name.Data.Count == 0)) DestSheet = Name.RangeSheet; //Macros and stuff should be copied in the global sheet.
                        TNameRecord NewName = TNameRecord.CreateTempName(Name.Name, DestSheet);
                        int NamePos;
                        bool Added = SheetInfo.DestNames.AddNameIfNotExists(NewName, out NamePos);
                        if (Added)
                        {
                            SheetInfo.DestNames[NamePos] = Name.CopyTo(DestSheet, CopyRowOffset, CopyColOffset, SheetInfo);
                        }
                        aName.NameIndex = NamePos + 1;// + 1;

                    }

                }
                else
                    if (InsertingSheet && (NameSheet == SheetInfo.SourceFormulaSheet || NameSheet < 0))
                    {
                        FindDestNameInLocalNames(SheetInfo.InsSheet, aName, Name);//If not found do nothing, it is a ref to a global name in the same sheet, so we just keep it.
                    }
            }
        }

        private void DoExternalName(TNameXToken aNamex)
        {
            if (SheetInfo.SourceReferences != null)
            {
                if (SheetInfo.SourceReferences != SheetInfo.DestReferences) //Copy the names to the new file.
                {
                    int ExternNameIndex = aNamex.NameIndex;
                    aNamex.ExternSheet = SheetInfo.DestReferences.CopyExternName(aNamex.ExternSheet, ref ExternNameIndex, SheetInfo);
                    aNamex.NameIndex = ExternNameIndex;
                }

                //Normal copy from one sheet to other does not need anything. The reference is kept.
            }
        }

        protected override void DoNameX(TNameXToken aNamex)
        {
            int ExternSheet = aNamex.ExternSheet;

            if (SheetInfo.SourceReferences != null)
            {
                if (SheetInfo.SourceReferences.IsLocalSheet(ExternSheet))
                {
                    DoName(aNamex);
                    bool ThisRefIsInInsertingSheet;
                    Do3D(aNamex, out ThisRefIsInInsertingSheet);

                }
                else
                {
                    DoExternalName(aNamex);
                }
            }
        }

        private void OffsetCopy(TBaseRefToken reft, ref int RowCol, int CopyOffset, int MaxRowCol, bool RowColAbs, bool ForgetAbsolute)
        {
            if (reft.CanHaveRelativeOffsets && !RowColAbs) return;  //Offsets never have to be adapted when copying
            
            bool AbsoluteRef = AllowedAbsolute && RowColAbs;
            if (!AbsoluteRef || ForgetAbsolute) reft.IncRowCol(ref RowCol, -1, CopyOffset, MaxRowCol, true);  //Fix the copy.
        }

        protected override void DoRef(TRefToken reft, bool RefIsInInsertingSheet)
        {
            int r = reft.Row;
            int c = reft.Col;

            if (RefIsInInsertingSheet)
            {
                bool RowIsOffset = reft.CanHaveRelativeOffsets && !reft.RowAbs;  //CanHaveRelativeOffsets is true for RefN tokens.
                bool ColIsOffset = reft.CanHaveRelativeOffsets && !reft.ColAbs;

                if (RowIsOffset) r = TBaseRefToken.WrapRow(FmlaRow + r, false);
                if (ColIsOffset) c = TBaseRefToken.WrapColumn(FmlaCol + c, false);

                bool FmlaInsideDown = (FmlaRow >= CellRange.Top) && CellRange.HasCol(FmlaCol);
                bool FmlaInsideRight = (FmlaCol >= CellRange.Left) && CellRange.HasRow(FmlaRow);
                bool RefInsideDown = (r >= CellRange.Top) && (CellRange.HasCol(c));
                bool RefInsideRight = (c >= CellRange.Left) && (CellRange.HasRow(r));

                r = MoveRef(reft, ref reft.Row, r, CellRange.Top, RowCount * CellRange.RowCount, FmlaRow, RowIsOffset, RefInsideDown, FmlaInsideDown, FlxConsts.Max_Rows, true);
                c = MoveRef(reft, ref reft.Col, c, CellRange.Left, ColCount * CellRange.ColCount, FmlaCol, ColIsOffset, RefInsideRight, FmlaInsideRight, FlxConsts.Max_Columns, true);
            }

            bool ForgetAbsolute = SemiAbsoluteMode && CellRange.HasRow(r) && CellRange.HasCol(c);

            OffsetCopy(reft, ref reft.Row, CopyRowOffset, FlxConsts.Max_Rows, reft.RowAbs, ForgetAbsolute);
            OffsetCopy(reft, ref reft.Col, CopyColOffset, FlxConsts.Max_Columns, reft.ColAbs, ForgetAbsolute);

            if (Bounds != null)
            {
                Bounds.AddRow(reft.Row);
                Bounds.AddCol(reft.Col);
            }
        }

        protected override void DoArea(TAreaToken areat, bool RefIsInInsertingSheet)
        {
            int r1 = areat.Row1;
            int r2 = areat.Row2;
            int c1 = areat.Col1;
            int c2 = areat.Col2;

            if (RefIsInInsertingSheet)
            {
                bool RowRel1 = areat.CanHaveRelativeOffsets && !areat.RowAbs1;  //CanHaveRelativeOffsets is true for RefN tokens.
                bool ColRel1 = areat.CanHaveRelativeOffsets && !areat.ColAbs1;
                bool RowRel2 = areat.CanHaveRelativeOffsets && !areat.RowAbs2;
                bool ColRel2 = areat.CanHaveRelativeOffsets && !areat.ColAbs2;

                if (RowRel1) r1 = TBaseRefToken.WrapRow(FmlaRow + r1, false);
                if (ColRel1) c1 = TBaseRefToken.WrapColumn(FmlaCol + c1, false);
                if (RowRel2) r2 = TBaseRefToken.WrapRow(FmlaRow + r2, false);
                if (ColRel2) c2 = TBaseRefToken.WrapColumn(FmlaCol + c2, false);

                bool InColRange = CellRange.HasCol(c1) && CellRange.HasCol(c2);
                if (InColRange)
                {
                    if (RowCount < 0)
                        areat.DeleteRowsArea(FmlaRow, RowCount, CellRange);  //Handles the complexities of deleting ranges.
                }

                if (RowCount > 0)
                {
                    bool FmlaInsideDown = (FmlaRow >= CellRange.Top) && CellRange.HasCol(FmlaCol);
                    bool RefInsideDown1 = r1 >= CellRange.Top && InColRange;
                    bool RefInsideDown2 = r2 >= CellRange.Top && r2 != FlxConsts.Max_Rows && InColRange;
                    r1 = MoveRef(areat, ref areat.Row1, r1, CellRange.Top, RowCount * CellRange.RowCount, FmlaRow, RowRel1, RefInsideDown1, FmlaInsideDown, FlxConsts.Max_Rows, true);
                    r2 = MoveRef(areat, ref areat.Row2, r2, CellRange.Top, RowCount * CellRange.RowCount, FmlaRow, RowRel2, RefInsideDown2, FmlaInsideDown, FlxConsts.Max_Rows, true);
                }

                bool InRowRange = CellRange.HasRow(r1) && CellRange.HasRow(r2);
                if (InRowRange)
                {
                    if (ColCount < 0)
                        areat.DeleteColsArea(FmlaCol, ColCount, CellRange);  //Handles the complexities of deleting ranges.
                }

                if (ColCount > 0)
                {
                    bool FmlaInsideRight = (FmlaCol >= CellRange.Left) && CellRange.HasRow(FmlaRow);
                    bool RefInsideRight1 = c1 >= CellRange.Left && InRowRange;
                    bool RefInsideRight2 = c2 >= CellRange.Left && c2 != FlxConsts.Max_Columns && InRowRange;
                    c1 = MoveRef(areat, ref areat.Col1, c1, CellRange.Top, ColCount * CellRange.ColCount, FmlaCol, ColRel1, RefInsideRight1, FmlaInsideRight, FlxConsts.Max_Columns, true);
                    c2 = MoveRef(areat, ref areat.Col2, c2, CellRange.Top, ColCount * CellRange.ColCount, FmlaCol, ColRel2, RefInsideRight2, FmlaInsideRight, FlxConsts.Max_Columns, true);
                }
            }


            bool ForgetAbsolute1 = SemiAbsoluteMode && CellRange.HasRow(r1) && CellRange.HasCol(c1);
            bool ForgetAbsolute2 = SemiAbsoluteMode && CellRange.HasRow(r2) && CellRange.HasCol(c2);

            OffsetCopy(areat, ref areat.Row1, CopyRowOffset, FlxConsts.Max_Rows, areat.RowAbs1, ForgetAbsolute1);
            OffsetCopy(areat, ref areat.Col1, CopyColOffset, FlxConsts.Max_Columns, areat.ColAbs1, ForgetAbsolute1);
            OffsetCopy(areat, ref areat.Row2, CopyRowOffset, FlxConsts.Max_Rows, areat.RowAbs2, ForgetAbsolute2);
            OffsetCopy(areat, ref areat.Col2, CopyColOffset, FlxConsts.Max_Columns, areat.ColAbs2, ForgetAbsolute2);

            if (Bounds != null)
            {
                Bounds.AddRow(areat.Row1);
                Bounds.AddRow(areat.Row2);
                Bounds.AddCol(areat.Col1);
                Bounds.AddCol(areat.Col2);
            }
        }

        protected override void DoMemArea(TParsedTokenList Tokens, TMemAreaToken areat, bool RefIsInInsertingSheet)
        {
            //Instead of trying to keep this list of areas synchronized, merging ranges that are together, etc, we will just remove this token. It isn't really needed anyway.
            Tokens.RemoveToken();
            Tokens.MoveBack(); //We can't call LightPop here to go back, since it would raise an Exception if at first token.
        }

        protected override void DoTable(TTableToken table)
        {
            ArrangeTokenTableAndArray(ref table.Row, ref table.Col);
        }

        protected override void DoArrayFmla(TExp_Token exp)
        {
            ArrangeTokenTableAndArray(ref exp.Row, ref exp.Col);
        }

        private void ArrangeTokenTableAndArray(ref int Row, ref int Col)
        {
            if (Bounds != null) { Bounds.Sheet1 = -2; }
            if ((SheetInfo.SourceFormulaSheet != SheetInfo.InsSheet) || InsertingSheet) return;

            if (Row >= CellRange.Top && CellRange.HasCol(Col)) IncRowColForTableArray(ref Row, CellRange.Top, RowCount * CellRange.RowCount, FlxConsts.Max_Rows, true);
            IncRowColForTableArray(ref Row, -1, CopyRowOffset, FlxConsts.Max_Rows, true);
            if (Col >= CellRange.Left && CellRange.HasRow(Row)) IncRowColForTableArray(ref Col, CellRange.Left, ColCount * CellRange.ColCount, FlxConsts.Max_Columns, true);
            IncRowColForTableArray(ref Col, -1, CopyColOffset, FlxConsts.Max_Columns, true);

        }
    }

    internal class TMoveTokens : TInsertOrMovetokens
    {
        #region Variables
        TXlsCellRange CellRange;
        int FmlaRow; int FmlaCol;
        int NewRow;
        int NewCol; 
        TSheetInfo SheetInfo; 
        TFormulaBounds Bounds;
        #endregion

        #region Constructor
        internal TMoveTokens(TXlsCellRange aCellRange, int aFmlaRow, int aFmlaCol, int aNewRow, int aNewCol, TSheetInfo aSheetInfo, TFormulaBounds aBounds)
        {
            CellRange = aCellRange;
            FmlaRow = aFmlaRow;
            FmlaCol = aFmlaCol;
            NewRow = aNewRow;
            NewCol = aNewCol;
            SheetInfo = aSheetInfo;
            Bounds = aBounds;
        }
        #endregion

        protected override void Do3D(TBaseParsedToken tk, out bool RefIsInInsertingSheet)
        {
            RefIsInInsertingSheet = false;

            if (SheetInfo.SourceReferences != null)
            {
                RefIsInInsertingSheet = SheetInfo.SourceReferences.GetJustOneSheet(tk.ExternSheet) == SheetInfo.InsSheet;
                if (Bounds != null) Bounds.AddSheets(SheetInfo.SourceReferences.GetAllSheets(tk.ExternSheet));
            }
            else
            {
                Debug.Assert(Bounds == null);  //we would have problems if references is null and we try to use Bounds.
            }
        }

        protected override void DoName(TNameToken aName)
        {
            //No changes when moving things.
        }

        protected override void DoNameX(TNameXToken aNamex)
        {
            //No changes when moving things.
        }

        protected override void DoRef(TRefToken reft, bool RefIsInInsertingSheet)
        {
            int r = reft.Row; int c = reft.Col;

            if (RefIsInInsertingSheet)
            {
                TXlsCellRange NewRange = CellRange.Offset(NewRow, NewCol);

                bool RowIsOffset = reft.CanHaveRelativeOffsets && !reft.RowAbs;  //CanHaveRelativeOffsets is true for RefN tokens.
                bool ColIsOffset = reft.CanHaveRelativeOffsets && !reft.ColAbs;

                if (RowIsOffset) r = TBaseRefToken.WrapRow(FmlaRow + r, false);
                if (ColIsOffset) c = TBaseRefToken.WrapColumn(FmlaCol + c, false);

                bool RefInside = CellRange.HasRow(r) && CellRange.HasCol(c);
                bool FmlaInside = CellRange.HasRow(FmlaRow) && CellRange.HasCol(FmlaCol);

                //only if the ref was *not* on the original range.
                if (!RefInside && NewRange.HasRow(r) && NewRange.HasCol(c))
                {
                    reft.CreateInvalidRef();
                }
                else
                {
                    r = MoveRef(reft, ref reft.Row, r, 0, NewRow - CellRange.Top, FmlaRow, RowIsOffset, RefInside, FmlaInside, FlxConsts.Max_Rows, false);
                    c = MoveRef(reft, ref reft.Col, c, 0, NewCol - CellRange.Left, FmlaCol, ColIsOffset, RefInside, FmlaInside, FlxConsts.Max_Columns, false);
                }
            }

            if (Bounds != null)
            {
                Bounds.AddRow(r);
                Bounds.AddCol(c);
            }
        }

        protected override void DoArea(TAreaToken areat, bool RefIsInInsertingSheet)
        {
            int r1 = areat.Row1;
            int r2 = areat.Row2;
            int c1 = areat.Col1;
            int c2 = areat.Col2;

            if (RefIsInInsertingSheet)
            {
                TXlsCellRange NewRange = CellRange.Offset(NewRow, NewCol);

                bool RowRel1 = areat.CanHaveRelativeOffsets && !areat.RowAbs1;  //CanHaveRelativeOffsets is true for RefN tokens.
                bool ColRel1 = areat.CanHaveRelativeOffsets && !areat.ColAbs1;
                bool RowRel2 = areat.CanHaveRelativeOffsets && !areat.RowAbs2;
                bool ColRel2 = areat.CanHaveRelativeOffsets && !areat.ColAbs2;

                if (RowRel1) r1 = TBaseRefToken.WrapRow(FmlaRow + r1, false);
                if (ColRel1) c1 = TBaseRefToken.WrapColumn(FmlaCol + c1, false);
                if (RowRel2) r2 = TBaseRefToken.WrapRow(FmlaRow + r2, false);
                if (ColRel2) c2 = TBaseRefToken.WrapColumn(FmlaCol + c2, false);

                bool RefInside = CellRange.HasRow(r1) && CellRange.HasCol(c1) && CellRange.HasRow(r2) && CellRange.HasCol(c2);
                bool FmlaInside = CellRange.HasRow(FmlaRow) && CellRange.HasCol(FmlaCol);

                //only if the ref was *not* on the original range.
                if (!RefInside && NewRange.HasRow(r1) && NewRange.HasCol(c1) && NewRange.HasRow(r2) && NewRange.HasCol(c2))
                {
                    areat.CreateInvalidRef();
                }
                else
                {
                    r1 = MoveRef(areat, ref areat.Row1, r1, 0, NewRow - CellRange.Top, FmlaRow, RowRel1, RefInside, FmlaInside, FlxConsts.Max_Rows, false);
                    c1 = MoveRef(areat, ref areat.Col1, c1, 0, NewCol - CellRange.Left, FmlaCol, ColRel1, RefInside, FmlaInside, FlxConsts.Max_Columns, false);
                    r2 = MoveRef(areat, ref areat.Row2, r2, 0, NewRow - CellRange.Top, FmlaRow, RowRel2, RefInside, FmlaInside, FlxConsts.Max_Rows, false);
                    c2 = MoveRef(areat, ref areat.Col2, c2, 0, NewCol - CellRange.Left, FmlaCol, ColRel2, RefInside, FmlaInside, FlxConsts.Max_Columns, false);
                }
            }
            if (Bounds != null)
            {
                Bounds.AddRow(r1);
                Bounds.AddRow(r2);
                Bounds.AddCol(c1);
                Bounds.AddCol(c2);
            }

        }

        protected override void DoMemArea(TParsedTokenList Tokens, TMemAreaToken areat, bool RefIsInInsertingSheet)
        {
            //Instead of trying to keep this list of areas synchronized, merging ranges that are together, etc, we will just remove this token. It isn't really needed anyway.
            Tokens.RemoveToken();
            Tokens.MoveBack(); //We can't call LightPop here to go back, since it would raise an Exception if at first token.
        }

        protected override void DoTable(TTableToken table)
        {
            ArrangeTokenTableAndArray(ref table.Row, ref table.Col);
        }

        protected override void DoArrayFmla(TExp_Token exp)
        {
            ArrangeTokenTableAndArray(ref exp.Row, ref exp.Col);
        }

        private void ArrangeTokenTableAndArray(ref int Row, ref int Col)
        {
            if (Bounds != null) { Bounds.Sheet1 = -2; }
            if (SheetInfo.SourceFormulaSheet != SheetInfo.InsSheet) return;

            TXlsCellRange NewRange = CellRange.Offset(NewRow, NewCol);

            if (CellRange.HasRow(Row) && CellRange.HasCol(Col))
            {
                IncRowColForTableArray(ref Row, -1, NewRow - CellRange.Top, FlxConsts.Max_Rows, false);
                IncRowColForTableArray(ref Col, -1, NewCol - CellRange.Left, FlxConsts.Max_Columns, false);
            }
            else
                if (NewRange.HasRow(Row) && NewRange.HasCol(Col))
                {
                    FlxMessages.ThrowException(FlxErr.ErrInternal); //we should never have an error when updating tables or exp tokens. The error should show when deleting the cell
                }

        }
    }

    internal static class TRangeManipulator
    {

        internal static void ArrangeInsertRange(TXlsCellRange Refe, TXlsCellRange CellRange, int aRowCount, int aColCount)
        {
            InsertFirst(ref Refe.Top, CellRange.Top, FlxConsts.Max_Rows, CellRange.RowCount, aRowCount, XlsErr.ErrTooManyRows);
            InsertLast(ref Refe.Bottom, CellRange.Bottom, FlxConsts.Max_Rows, CellRange.RowCount, aRowCount);
            InsertFirst(ref Refe.Left, CellRange.Left, FlxConsts.Max_Columns, CellRange.ColCount, aColCount, XlsErr.ErrTooManyColumns);
            InsertLast(ref Refe.Right, CellRange.Right, FlxConsts.Max_Columns, CellRange.ColCount, aColCount);
        }

        private static void InsertFirst(ref int First, int MaxRange, int MaxAll, int RangeCount, int aRowCount, XlsErr ErrorWhenTooMany)
        {
            if (First >= MaxRange)
            {
                First += RangeCount * aRowCount;
                if (First > MaxAll)
                {
                    XlsMessages.ThrowException(ErrorWhenTooMany, First + 1, MaxAll + 1);
                }
            }
        }

        private static void InsertLast(ref int Last, int MaxRange, int MaxAll, int RangeCount, int aRowCount)
        {
            if (Last >= MaxRange)
            {
                if (Last >= MaxAll && aRowCount < 0)
                {
                }
                else
                {
                    Last += RangeCount * aRowCount;
                }
                if (Last > MaxAll) Last = MaxAll;
            }
        }

        internal static void ArrangeMoveRange(TXlsCellRange Refe, TXlsCellRange CellRange, int NewRow, int NewCol)
        {
            if (Refe.Top < CellRange.Top || Refe.Bottom > CellRange.Bottom || Refe.Left < CellRange.Left || Refe.Right > CellRange.Right) return;

            int DeltaRow = NewRow - Refe.Top;
            int DeltaCol = NewCol - Refe.Left;

            MoveOne(ref Refe.Top, DeltaRow, FlxConsts.Max_Rows, XlsErr.ErrTooManyRows);
            MoveOne(ref Refe.Bottom, DeltaRow, FlxConsts.Max_Rows, XlsErr.ErrTooManyRows);
            MoveOne(ref Refe.Left, DeltaCol, FlxConsts.Max_Columns, XlsErr.ErrTooManyColumns);
            MoveOne(ref Refe.Right, DeltaCol, FlxConsts.Max_Columns, XlsErr.ErrTooManyColumns);
        }

        private static void MoveOne(ref int r, int Delta, int MaxAll, XlsErr ErrorWhenTooMany)
        {
            r += Delta;
            if (r > MaxAll)
            {
                XlsMessages.ThrowException(ErrorWhenTooMany, r + 1, MaxAll + 1);
            }
        }
    }
}
