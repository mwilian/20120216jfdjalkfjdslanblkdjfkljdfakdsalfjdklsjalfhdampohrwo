using System;
using FlexCel.Core;

using System.IO;
using System.Collections.Generic;
using System.Globalization;


namespace FlexCel.XlsAdapter
{
    /// <summary>
    /// A list of sheets.
    /// </summary>
    internal class TSheetList
    {
        protected List<TSheet> FList;

        internal TSheetList()
        {
            FList = new List<TSheet>();
        }
            
        #region Generics
        internal int Add (TSheet a)
        {
            FList.Add(a);
            return FList.Count - 1;
        }

        internal void Insert (int index, TSheet a)
        {
            FList.Insert(index, a);
        }

        protected void SetThis(TSheet value, int index)
        {
            FList[index]=value;
        }

        internal TSheet this[int index] 
        {
            get {return FList[index];} 
            set {SetThis(value, index);}
        }

        internal int Count
        {
            get {return FList.Count;}
        }

        //We shouldn't need to destroy anything here, as when we clear the sheetlist we are probably
        //clearing the whole thing. Going and destroying all globals would be a waste of time, as globals should be destroyed too.
        internal void Clear()
        {
            FList.Clear();
        }

        internal void DeleteSheets(int SheetIndex, int SheetCount)
        {
            for (int i=0; i<SheetCount;i++)
            {
                if (SheetIndex>= Count) return;
                TSheet Sh=this[SheetIndex];
                FList.RemoveAt(SheetIndex);
                Sh.Destroy(); //This will call destroy on child objects, to clear SST references and BSE drawing refs.
            }
        }
        #endregion

        internal int IndexOf(TSheet SearchRecord)
        {
            for (int i = Count - 1; i >= 0; i--)
            {
                if (FList[i] == SearchRecord) return i;
            }
            return -1;
        }

        internal void InsertAndCopyRange(TXlsCellRange SourceRange, int DestRow, int DestCol, int aRowCount, int aColCount, TRangeCopyMode CopyMode, TFlxInsertMode InsertMode, TSheetInfo SheetInfo)
        {
            TXlsCellRange NewCellRange = SourceRange.Offset(DestRow, DestCol);

            FList[SheetInfo.InsSheet].InsertAndCopyRange(SourceRange, DestRow, DestCol, aRowCount, aColCount, CopyMode, InsertMode, SheetInfo);
            for (int i = 0; i < Count; i++)
                if (i != SheetInfo.InsSheet)
                {
                    SheetInfo.SourceFormulaSheet = i;
                    SheetInfo.DestFormulaSheet = i;
                    FList[i].ArrangeInsertRange(NewCellRange, aRowCount, aColCount, SheetInfo);
                }
        }

        internal void DeleteRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
        {
            FList[SheetInfo.InsSheet].DeleteRange(CellRange, aRowCount, aColCount, SheetInfo);
            for (int i=0; i< Count;i++)
                if (i!=SheetInfo.InsSheet)
                {
                    SheetInfo.SourceFormulaSheet=i;
                    SheetInfo.DestFormulaSheet=i;
                    FList[i].ArrangeInsertRange(CellRange, -aRowCount, -aColCount, SheetInfo);
                }
        }

        internal void MoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
        {
            FList[SheetInfo.InsSheet].MoveRange(CellRange, NewRow, NewCol, SheetInfo);
            for (int i = 0; i < Count; i++)
                if (i != SheetInfo.InsSheet)
                {
                    SheetInfo.SourceFormulaSheet = i;
                    SheetInfo.DestFormulaSheet = i;
                    FList[i].ArrangeMoveRange(CellRange, NewRow, NewCol, SheetInfo);
                }
        }


        internal void SaveToStream(IDataStream DataStream, TSaveData SaveData)
        {
            for (int i = 0; i < Count; i++)
            {
                SaveData.SavingSheet = FList[i];
                FList[i].SaveToStream(DataStream, SaveData);
            }
        }

        internal void SaveToPxl(TPxlStream PxlStream, TPxlSaveData SaveData)
        {
            for (int i=0; i< Count;i++)
                FList[i].SaveToPxl(PxlStream, SaveData);
        }

        internal void MergeFromPxlSheets(TSheetList SourceSheets)
        {
            for (int i=0; i< Count;i++)
                FList[i].MergeFromPxlSheet(SourceSheets[i]);
        }


        internal void UpdateDeletedRanges(int FirstSheet, int SheetCount, TDeletedRanges DeletedRanges)
        {
            for (int i = Count - 1; i >= 0; i--)
            {
                if (i< FirstSheet || i >= FirstSheet + SheetCount)
                {
                    FList[i].UpdateDeletedRanges(DeletedRanges);
                }
            }
        }
    }


    /// <summary>
    /// This class encapsulates the complete Excel workbook. It is the main class on XlsAdapter, and all Requests come through it.
    /// </summary>
    internal class TWorkbook
    {
#if (FRAMEWORK30)
        private TFileProps FileProps;
#endif
        private TWorkbookGlobals FGlobals;
        private TSheetList FSheets;
        internal bool Loaded;

        internal TWorkbook(ExcelFile aXls, TFileProps aFileProps)
        {
            FGlobals= new TWorkbookGlobals(aXls);
            FSheets = new TSheetList();
#if (FRAMEWORK30)
            FileProps = aFileProps;
#endif
        }

        internal TWorkbookGlobals Globals{get{return FGlobals;}}
        internal TSheetList Sheets{get{return FSheets;}}

        internal bool IsWorkSheet(int Index)
        {
            return (Sheets[Index] is TWorkSheet);
        }

        internal TSheetType SheetType(int Index)
        {
            return Sheets[Index].SheetType;
        }

        internal TWorkSheet WorkSheets(int Index)
        {
            return (TWorkSheet)Sheets[Index];
        }

        internal int ActiveSheet 
        {
            get 
            {
                return Globals.ActiveSheet;
            } 
            set 
            { 
                //if ((Globals.ActiveSheet>=0) && (Globals.ActiveSheet<Sheets.Count))  //Active sheet might become invalid if we delete sheets.
                //    Sheets[Globals.ActiveSheet].Selected=false;
                //We have to loop on ALL sheets, because copying might copy selected sheets.
                int aCount=Sheets.Count;
                for (int i=0;i<aCount;i++) 
                    Sheets[i].Selected=false;

                Globals.SetFirstSheetVisible(0);
                Globals.ActiveSheet=value;
                Sheets[value].Selected=true;
            }
        }

        internal void FixRows()
        {
            for (int i = 0; i < Sheets.Count; i++)
            {
                Sheets[i].DeleteEmptyRowRecords();
            }
        }

        internal void FixNames()
        {
            FGlobals.Names.CleanUnusedNames(this);
        }

        private void FixTheme()
        {
            FGlobals.ThemeRecord.CalcData(this);
        }


        internal bool IsXltTemplate
        {
            get
            {
                return Globals.IsXltTemplate;
            }
            set
            {
                Globals.IsXltTemplate = value;
            }
        }

        #region Loading
        private void InitLoading(TBaseRecordLoader RecordLoader, TProtection Protection)
        {
            Loaded = false;
            Sheets.Clear();
            Globals.Clear();
        }

        private void LoadBinWorkbook(TBinRecordLoader RecordLoader)
        {
            RecordLoader.ReadHeader();
            int RecordId = RecordLoader.RecordHeader.Id;

            while (!RecordLoader.Eof && (RecordLoader.RecordHeader.Id != 0))
            {
                RecordId = RecordLoader.RecordHeader.Id;
                TBOFRecord RBOF = RecordLoader.LoadRecord(false) as TBOFRecord;
                if (RecordId == (int)xlr.BOF)
                {
                    RecordLoader.SwitchSheet();

                    switch (RBOF.BOFType)
                    {
                        case (int)xlb.Globals:
                            Globals.LoadFromStream(RecordLoader, RBOF);
                            if (RecordLoader.VirtualReader != null) RecordLoader.VirtualReader.StartReading();
                            break;
                        case (int)xlb.Worksheet:
                            FSheets[FSheets.Add(new TWorkSheet(Globals))].LoadFromStream(RecordLoader, RBOF);
                            break;
                        case (int)xlb.Chart:
                            FSheets[FSheets.Add(new TFlxChart(Globals, false))].LoadFromStream(RecordLoader, RBOF);
                            break;
                        case (int)xlb.Macro:
                            FSheets[FSheets.Add(new TMacroSheet(Globals))].LoadFromStream(RecordLoader, RBOF);
                            break;
                        default:
                            FSheets[FSheets.Add(new TFlxUnsupportedSheet(Globals))].LoadFromStream(RecordLoader, RBOF);
                            break;
                    } //case

                    if (RecordLoader.VirtualReader != null) RecordLoader.VirtualReader.Flush();

                }
                else
                    if (RecordId != (int)xlr.EOF)  //There can be 2 eof at the end of the file
                        XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);

                if (Globals.SheetCount > 0 && Globals.SheetCount <= FSheets.Count) break; //There shouldn't be any garbage here, but some weird non-created-by-excel files might have it, and Excel will load them fine. 

            } //while

        }

#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
        private void LoadXlsxWorkbook(TXlsxRecordLoader RecordLoader, out bool MacroEnabled)
        {
            RecordLoader.ReadTheme(Globals.ThemeRecord);
            List<string> ExternalRefs = new List<string>();
            List<string> NameDefinitions = new List<string>();
            RecordLoader.ReadWorkbook(Globals, ExternalRefs, NameDefinitions, out MacroEnabled);

            RecordLoader.ReadCustomFileProperties(FileProps);
            RecordLoader.ReadCustomXMLData(Globals.CustomXMLData);
            Globals.References.EnsureLocalSupBook(Globals.SheetCount); //Local Supbook is reference [0]
            foreach (string ExternalRef in ExternalRefs)
            {
                RecordLoader.ReadExternalLink(Globals, ExternalRef);
            }

            RecordLoader.LoadNameDefinitions(Globals, NameDefinitions); //shouldn't be done until all boundsheets have been read and external refs too.

            RecordLoader.ReadStyles(Globals);
            RecordLoader.ReadSST();
            RecordLoader.ReadConnections(Globals);

            Globals.EnsureRequiredRecords();
            Loaded = true;

            if (Globals.SheetCount <= 0) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
            if (RecordLoader.VirtualReader != null) RecordLoader.VirtualReader.StartReading();
            for (int i = 0; i < Globals.SheetCount; i++)
            {
                RecordLoader.SwitchSheets();
                RecordLoader.ReadSheet(Globals.SheetRelationshipId(i), FSheets, Globals);
                if (RecordLoader.VirtualReader != null) RecordLoader.VirtualReader.Flush();
            }

            RecordLoader.ReadMacros();
        }

#endif

        private void EndLoading(TProtection Protection, TEncryptionData Encryption, TSST SST)
        {
            // References from LABELSST to SST have been loaded, we can sort
            // Globals.SST.Sort(); Not needed here, we are using a hash and entries are automatically sorted.
            SST.ClearIndexData(); //All labelSST records have been loaded. We can go on...

            //now we can safely sort, all BSEs are pointers, no integers
            if (Globals.DrawingGroup.RecordCache.BStore != null) Globals.DrawingGroup.RecordCache.BStore.ContainedRecords.Sort();

            FixRows();
            if (Encryption != null) Protection.OpenPassword = Encryption.ReadPassword; //The read password might have been modified if the event was used.
            Loaded = true;

        }

        internal void LoadFromStream(TBinRecordLoader RecordLoader, TProtection Protection)
        {
            InitLoading(RecordLoader, Protection);
            Globals.Patterns.EnsureRequiredFills();
            LoadBinWorkbook(RecordLoader);
            EndLoading(Protection, RecordLoader.Encryption, RecordLoader.SST);
        }

#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
        internal void LoadFromStream(TXlsxRecordLoader RecordLoader, TProtection Protection, out bool MacroEnabled)
        {
            InitLoading(null, Protection);
            LoadXlsxWorkbook(RecordLoader, out MacroEnabled);
            Globals.Patterns.EnsureRequiredFills(); //after loading... some files might not have style 1.
            EndLoading(Protection, RecordLoader.Encryption, RecordLoader.SST);
        }
#endif
        #endregion

        /// <summary>
        /// Fixes the offset for the sheets on the global section. If offsets are wrong, Excel will crash, even when FlexCel
        /// will be able to read the file ok. (FlexCel does not use sheet offsets).
        /// </summary>
        private void FixBoundSheetsOffset(TEncryptionData Encryption, bool Repeatable)
        {
            Globals.SST.FixRefs();
            long TotalOfs=Globals.TotalSize(Encryption, Repeatable);  //Includes the EOF on workbook Globals
            if (Globals.SheetCount!= Sheets.Count) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);

            for (int i=0; i< Globals.SheetCount;i++)
            {
                Globals.SheetSetOffset(i,TotalOfs);
                TotalOfs+=(Sheets[i].TotalSize(Encryption, Repeatable));
            }
        }

        private void FixRangeBoundSheetsOffset(int SheetIndex, TXlsCellRange CellRange, TEncryptionData Encryption, bool Repeatable)
        {
            Globals.SST.FixRefs();
            long TotalOfs=Globals.TotalRangeSize(SheetIndex, CellRange, Encryption, Repeatable);  //Includes the EOF on workbook Globals
            if (Globals.SheetCount != Sheets.Count) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);

            Globals.SheetSetOffset(SheetIndex, TotalOfs);
        }

        private void FixNotes()
        {
            for (int i = FSheets.Count - 1; i >= 0; i--)
            {
                FSheets[i].FixNotes();
            }
        }

        private void FixCodeNames()
        {
            if (!FGlobals.HasMacro) return;
            List<string> Names = new List<string>();
            Names.Add(FGlobals.CodeName);
            for (int i = 0; i < FSheets.Count; i++)
            {
                string s = FSheets[i].CodeName;
                if (s.Length > 0) Names.Add(s.ToUpper(CultureInfo.InvariantCulture));
            }
            Names.Sort();

            for (int i = 0; i < FSheets.Count; i++)
            {
                if (FSheets[i].CodeName.Length == 0)
                {
                    string SheetName = FGlobals.GetSheetName(i);
                    int k = SheetName.Length - 1;
                    while ((k >= 0) && (Char.IsDigit(SheetName[k]))) k--;
                    string s = SheetName.Substring(k + 1);
                    SheetName = SheetName.Substring(0, k + 1);
                    string UpSheetName = SheetName.ToUpper(CultureInfo.CurrentCulture);
                    if (s.Length == 0) k = 0; else k = Convert.ToInt32(s);
                    while (Names.BinarySearch(0, Names.Count, UpSheetName + s, null) >= 0)  //Watch out CF!
                    {
                        k++;
                        s = k.ToString();
                    }

                    FSheets[i].CodeName = SheetName + s;
                    int ListIndex = Names.BinarySearch(0, Names.Count, UpSheetName + s, null);
                    Names.Insert(~ListIndex, UpSheetName + s);
                }
            }
        }


        private void FixSheetVisible()
        {
            int NewFirstSheetVisible=-1;
            
            //Verify we have not selected a hidden sheet. This will cause excel to crash.
            for (int i = FSheets.Count - 1; i >= 0; i--)
            {
                if (FGlobals.GetSheetVisible(i) == TXlsSheetVisible.Visible) NewFirstSheetVisible = i;
                else
                    if (FSheets[i].Selected) XlsMessages.ThrowException(XlsErr.ErrHiddenSheetSelected);
            }

            if (NewFirstSheetVisible==-1) XlsMessages.ThrowException(XlsErr.ErrNoSheetVisible);

            int sv = FGlobals.GetFirstSheetVisible();
            if (sv < 0 || sv >= FSheets.Count
                || FGlobals.GetSheetVisible(sv) != TXlsSheetVisible.Visible)
            {
                FGlobals.SetFirstSheetVisible(NewFirstSheetVisible);
            }
        }

        private void FixRadioButtons()
        {
            for (int i = FSheets.Count - 1; i >= 0; i--)
            {
                FSheets[i].FixRadioButtons(Globals.Workbook, i);
            }
        }

        private void FixPivotTables()
        {
#if(FRAMEWORK30 && !COMPACTFRAMEWORK)
            for (int i = FSheets.Count - 1; i >= 0; i--)
            {
                FSheets[i].XlsxPivotTables.MarkUsedCacheIds();
            }

            FGlobals.XlsxPivotCache.RemoveUnusedCaches();
#endif
        }

        internal void SaveToStream(IDataStream DataStream, TSaveData SaveData)
        {
            PrepareToSave();
            FixBoundSheetsOffset(DataStream.Encryption, SaveData.Repeatable);

            FGlobals.SaveToStream(DataStream, SaveData);
            FSheets.SaveToStream(DataStream, SaveData);
        }

        internal void PrepareToSave()
        {
            FixTheme();
            FixNames();
            FixCodeNames(); //before fixing offsets.
            FixRows(); //Should be before calculating the offset. it is needed because some non-excel generated files might not have all rows.
            FixNotes();
            FixRadioButtons();
            FixPivotTables();

            FixSheetVisible();
        }

        internal void SaveToPxl(TPxlStream PxlStream)
        {
            FixRows(); //Should be before calculating the offset. it is needed because some non-excel generated files might not have all rows.
            FixSheetVisible();

            TPxlSaveData SaveData = new TPxlSaveData(FGlobals);
            FGlobals.SaveToPxl(PxlStream, SaveData);
            FSheets.SaveToPxl(PxlStream, SaveData);
        }

        internal void SaveRangeToStream(IDataStream DataStream, TSaveData SaveData, int SheetIndex, TXlsCellRange CellRange)
        {
            FixNames();
            FixCodeNames(); //before fixing offsets.
            FixRows(); //Should be before calculating the offset. It is needed because some non-excel generated files might not have all rows.
            FixRangeBoundSheetsOffset(SheetIndex, CellRange, DataStream.Encryption, SaveData.Repeatable);
            FixNotes();
            FixRadioButtons();

            FGlobals.SaveRangeToStream(DataStream, SaveData, SheetIndex, CellRange);
            //we don't have to check SheetIndex is ok. this was done on FGlobals.SaveRangetoStream
            SaveData.SavingSheet = FSheets[SheetIndex];
            FSheets[SheetIndex].SaveRangeToStream(DataStream, SaveData, SheetIndex, CellRange);
        }

        internal void CleanFlags()
        {
            int aCount = Sheets.Count;
            for (int i = 0; i < aCount; i++) //We need to clean the flags in *all* the sheets before calculating.
            {
                Sheets[i].Cells.CellList.CleanFlags();
            }
        }

        internal void Recalc(ExcelFile aXls, TUnsupportedFormulaList Ufl)
        {
            aXls.SetUnsupportedFormulaList(Ufl);
            try
            {
                int aCount = Sheets.Count;
                for (int i = 0; i < aCount; i++)
                {
                    Sheets[i].Cells.CellList.Recalc(aXls, i + 1);
                }
            }
            finally
            {
                aXls.SetUnsupportedFormulaList(null);
            }
        }

        internal void ClearFormulaResults()
        {
            int aCount=Sheets.Count;
            for (int i=0;i < aCount; i++)
            {
                Sheets[i].Cells.CellList.ClearFormulaResult();
            }
        }

        internal void ConvertFormulasToValues(int SheetIndex, bool OnlyExternal)
        {
            TCellList Cells = Sheets[SheetIndex].Cells.CellList;

            int RowCount= Cells.Count;
            for (int r=0; r< RowCount;r++)
            {
                int ColCount=Cells[r].Count;
                for (int c=ColCount-1; c>=0; c--)
                {
                    TFormulaRecord fr= (Cells[r][c] as TFormulaRecord);
                    if (fr!=null && (!OnlyExternal || fr.HasExternRefs()))
                    {
                        Cells.SetValue(r, fr.Col, fr.FormulaValue, fr.XF);
                    }
                }
            }
        }

        internal void ConvertExternalNamesToRefErrors()
        {
            for (int i = FGlobals.Names.Count - 1; i >= 0; i--)
            {
                if (FGlobals.Names[i].HasExternRefs(FGlobals.References)) FGlobals.Names.DeleteName(i, this);
            }
        }


        internal void ForceAutoRecalc()
        {
            int aCount=Sheets.Count;
            for (int i=0;i < aCount; i++)
            {
                Sheets[i].Cells.CellList.ForceAutoRecalc();
            }
        }

#if (!COMPACTFRAMEWORK && !MONOTOUCH && !SILVERLIGHT)
        internal void RecalcRowHeights(ExcelFile Workbook, bool Forced, bool KeepAutoFit, float Adjustment, int AdjustmentFixed, int MinHeight, int MaxHeight, TAutofitMerged AutofitMerged)
        {
            int aCount=Sheets.Count;
            for (int i=0;i < aCount; i++)
            {
                Sheets[i].Cells.CellList.RecalcRowHeights(Workbook, 0, Sheets[i].Cells.Count - 1, Forced, KeepAutoFit, false, Adjustment, AdjustmentFixed, MinHeight, MaxHeight, AutofitMerged);
                Sheets[i].RestoreObjectCoords();
            }
        }

#endif

        internal bool R1C1
        {
            get
            {
                return !Globals.CalcOptions.A1RefMode;
            }
            set
            {
                Globals.CalcOptions.A1RefMode = !value;
            }
        }


        #region Manipulating Methods

        internal void InsertAndCopyRange(int SheetNo, TXlsCellRange SourceRange,
            int DestRow, int DestCol, int aRowCount, int aColCount, TRangeCopyMode CopyMode, TFlxInsertMode InsertMode, bool aSemiAbsoluteMode, TExcelObjectList ObjectsInRange)
        {
            //Some error handling
            if (
                (SourceRange.Top>SourceRange.Bottom) || (SourceRange.Top<0) || (DestRow> FlxConsts.Max_Rows) ||
                ((CopyMode!=TRangeCopyMode.None) && (aRowCount>0) && (SourceRange.Top<DestRow) && (DestRow<=SourceRange.Bottom)) 
                || (DestRow+(SourceRange.Bottom-SourceRange.Top+1)*aRowCount-1>FlxConsts.Max_Rows)
                || (DestRow<0)
                )
                XlsMessages.ThrowException(XlsErr.ErrBadCopyRows, FlxConsts.Max_Rows+1);
            
            if (
                (SourceRange.Left>SourceRange.Right) || (SourceRange.Left<0) || (DestCol> FlxConsts.Max_Columns) ||
                ((CopyMode!=TRangeCopyMode.None) && (aColCount>0) && (SourceRange.Left<DestCol) && (DestCol<=SourceRange.Right)) 
                || (DestCol+(SourceRange.Right-SourceRange.Left+1)*aColCount-1>FlxConsts.Max_Columns)
                || (DestCol<0)
                )
                XlsMessages.ThrowException(XlsErr.ErrBadCopyCols, FlxConsts.Max_Columns + 1);

            if ((SheetNo<0) || (SheetNo>= Sheets.Count)) XlsMessages.ThrowException(XlsErr.ErrInvalidSheetNo, SheetNo, 0, Sheets.Count-1);

            TSheetInfo SheetInfo= new TSheetInfo(SheetNo, SheetNo, SheetNo, Globals, Globals, Sheets, Sheets, aSemiAbsoluteMode);
            SheetInfo.ObjectsInRange = ObjectsInRange;

            FSheets.InsertAndCopyRange(
                SourceRange, DestRow, DestCol, aRowCount, aColCount, 
                CopyMode, InsertMode, SheetInfo);
            Globals.InsertAndCopyRange(
                SourceRange, InsertMode, DestRow, DestCol, aRowCount, aColCount,
                SheetInfo);

            //FSheets[SheetNo].FixAutoFilter(SheetNo);  Not needed now, we are going to replace it.
            if (aColCount > 0) FSheets[SheetNo].AddNewAutoFilters(SheetNo, SourceRange.Top, SourceRange.Bottom, DestCol, DestCol); 
        }
        
        internal void DeleteRange(int SheetNo, TXlsCellRange CellRange, TFlxInsertMode InsertMode)
        {
            if((SheetNo>= Sheets.Count)||(SheetNo<0)) XlsMessages.ThrowException(XlsErr.ErrInvalidSheetNo, SheetNo, 0, Sheets.Count-1);

            TSheetInfo SheetInfo= new TSheetInfo(SheetNo, SheetNo, SheetNo, Globals, Globals, Sheets, Sheets, false);

            if (InsertMode==TFlxInsertMode.NoneDown || InsertMode == TFlxInsertMode.NoneRight)
            {
                FSheets[SheetNo].ClearRange(CellRange);
                return;
            }

            int aRowCount=0;
            int aColCount=0;
            if ((InsertMode== TFlxInsertMode.ShiftRangeDown)|| (InsertMode== TFlxInsertMode.ShiftRowDown)) aRowCount=1; else aColCount=1;

            FSheets.DeleteRange(CellRange, aRowCount, aColCount, SheetInfo);
            Globals.DeleteRange(CellRange, aRowCount, aColCount, SheetInfo);
            
            //FSheets[SheetNo].FixAutoFilter(SheetNo);
            //we sadly have to reconstruct the whole autofilter, or the blue arrows will not stay correct. Just fixing AUTOFILTERINFO (with the line above) is not enough.
            if (aColCount > 0) FSheets[SheetNo].AddNewAutoFilters(SheetNo, CellRange.Top, CellRange.Bottom, CellRange.Left, CellRange.Right); 

        }

        internal void MoveRange(int SheetNo, TXlsCellRange CellRange, int NewRow, int NewCol)
        {
            //Some error handling
            if (
                (CellRange.Top>CellRange.Bottom) || (CellRange.Top<0) || (NewRow + CellRange.RowCount - 1> FlxConsts.Max_Rows)				
                )
                XlsMessages.ThrowException(XlsErr.ErrBadMoveCall);

            if (
                (CellRange.Left>CellRange.Right) || (CellRange.Left<0) || (NewCol + CellRange.ColCount - 1> FlxConsts.Max_Columns)
                )
                XlsMessages.ThrowException(XlsErr.ErrBadMoveCall);
            

            if((SheetNo>= Sheets.Count)||(SheetNo<0)) XlsMessages.ThrowException(XlsErr.ErrInvalidSheetNo, SheetNo, 0, Sheets.Count-1);

            TSheetInfo SheetInfo= new TSheetInfo(SheetNo, SheetNo, SheetNo, Globals, Globals, Sheets, Sheets, false);

            FSheets.MoveRange(CellRange, NewRow, NewCol, SheetInfo);
            Globals.MoveRange(CellRange, NewRow, NewCol, SheetInfo);
        }

        internal void DeleteSheets(int SheetPos, int SheetCount)
        {
            if  (SheetPos> Sheets.Count) XlsMessages.ThrowException(XlsErr.ErrInvalidSheetNo, SheetPos, 0, Sheets.Count);
            Globals.DeleteSheets(SheetPos, SheetCount, this);
            Sheets.DeleteSheets(SheetPos, SheetCount); 
        }

        internal void UpdateDeletedRanges(int SheetIndex, int SheetCount, TDeletedRanges DeletedRanges)
        {
            Globals.UpdateDeletedRanges(SheetIndex, SheetCount, DeletedRanges);
            Sheets.UpdateDeletedRanges(SheetIndex, SheetCount, DeletedRanges);
        }

        internal TDeletedRanges FindUnreferencedRanges(int SheetIndex, int SheetCount)
        {
            TDeletedRanges Result = new TDeletedRanges(Globals.Names.Count, Globals.References, Globals.Names);
            UpdateDeletedRanges(SheetIndex, SheetCount, Result);
            Result.Update = true;
            return Result;
        }


        internal void InsertSheets(int CopyFrom, int InsertBefore, int SheetCount, TWorkbook SourceWorkbook)
        {
            if (SourceWorkbook == null) SourceWorkbook = this;
            if (CopyFrom >= SourceWorkbook.Sheets.Count) XlsMessages.ThrowException(XlsErr.ErrInvalidSheetNo, CopyFrom, -1, SourceWorkbook.Sheets.Count - 1);
            if (InsertBefore > Sheets.Count) XlsMessages.ThrowException(XlsErr.ErrInvalidSheetNo, InsertBefore, 0, Sheets.Count);

            TSheet aSheet=null;
            int OptionFlags=0;

            if (CopyFrom >= 0)
            {
                aSheet = SourceWorkbook.Sheets[CopyFrom];
                OptionFlags = SourceWorkbook.Globals.SheetOptionFlags(CopyFrom);
            }

            int NewCopyFrom = CopyFrom;
            if (SourceWorkbook == this && CopyFrom >= InsertBefore) NewCopyFrom += SheetCount;
            Globals.InsertSheets(CopyFrom, InsertBefore, OptionFlags, XlsMessages.GetString(XlsErr.BaseSheetName), SheetCount, SourceWorkbook.Sheets);

            TSheetInfo SheetInfo = new TSheetInfo(-1, -1, -1, SourceWorkbook.Globals, Globals, aSheet, null, false);
            for (int i = 0; i < SheetCount; i++)
            {
                SheetInfo.InsSheet = InsertBefore + i;
                SheetInfo.SourceFormulaSheet = NewCopyFrom;
                SheetInfo.DestFormulaSheet = InsertBefore;
                SheetInfo.DestSheet = null; //keep it null, since the reference does not exist yet.

                if (aSheet == null)
                    Sheets.Insert(InsertBefore, TWorkSheet.CreateFromData(Globals, Globals.Workbook.XlsBiffVersion, Globals.Workbook.ExcelFileFormat));
                else
                {
                    SheetInfo.DestFormulaSheet = InsertBefore + i;
                    Globals.Names.InsertSheets(NewCopyFrom, InsertBefore + i, 1, SheetInfo, SourceWorkbook == this); //names must be inserted before the sheet is cloned, so formulas can refer to them.
                    Sheets.Insert(InsertBefore + i, TSheet.Clone(aSheet, SheetInfo));
                    SheetInfo.DestSheet = Sheets[InsertBefore + i];
                    Sheets[InsertBefore + i].ArrangeCopySheet(SheetInfo);
                }
            }
        }

        internal static TSheet CopySheetMisc(TSheetInfo SheetInfo)
        {
            return SheetInfo.SourceSheet.CopyMiscData(SheetInfo);
        }

        #endregion

        #region Merge
        internal void MergeFromPxlWorkbook(TWorkbook SourceWorkbook)
        {
            FGlobals.MergeFromPxlGlobals(SourceWorkbook.FGlobals);
            FSheets.MergeFromPxlSheets(SourceWorkbook.FSheets); 
        }
        #endregion


        internal bool FindSheet(string SheetName, out int SheetIndex)
        {
            SheetIndex = FGlobals.BoundSheets.SheetNames[SheetName];
            return SheetIndex >= 0;
        }
    }
}
