using System;
using FlexCel.Core;

namespace FlexCel.XlsAdapter
{
    /// <summary>
    /// Selection. As this record is difficult to modify (it might split), it will be saved on a different place.
    /// </summary>
    internal class TBiff8SelectionRecord: TxBaseRecord
    {
        internal TBiff8SelectionRecord(int aId, byte[] aData): base(aId, aData){}

        internal TPanePosition PanePosition
        {
            get
            {
                return (TPanePosition) Data[0];
            }
            set
            {
                if ((int)value<0 || (int)value>3) Data[0]=3; else Data[0]=(byte)value;
            }
        }

        internal int ActiveCellRow
        {
            get
            {
                return GetWord(1);
            }
            set
            {
                if (value>=0 && value <= FlxConsts.Max_Rows)
                    SetWord(1, value);
            }
        }

        internal int ActiveCellCol
        {
            get
            {
                return GetWord(3);
            }
            set
            {
                if (value>=0 && value <= FlxConsts.Max_Columns)
                    SetWord(3, value);
            }
        }

        internal int ActiveCellRef
        {
            get
            {
                return GetWord(5);
            }
        }

        internal TXlsCellRange[] SelectedCells
        {
            get
            {
                TXlsCellRange[] Result = new TXlsCellRange[GetWord(7)];
                for (int i=0;i< Result.Length;i++)
                    Result[i] = new TXlsCellRange(GetWord(9+i*6), Data[13+i*6], GetWord(11+i*6), Data[14+i*6]);

                return Result;
            }
        }

        internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
        {
            if (Loader.CustomView != null) Loader.CustomView.Selection.AddBiff8Record(this);
            else ws.Window.Selection.AddBiff8Record(this);
        }


    }

    /// <summary>
    /// Selection for one pane, when window is split. There might exist 4 different panes
    /// </summary>
    internal class TPaneSelection
    {
        int ActiveRow;
        int ActiveCol;
        int ActiveSel;

        TXlsCellRange[] SelectedCells;

        internal static TPaneSelection Clone(TPaneSelection Source)
        {
            if (Source ==null) return null;
            TPaneSelection Result = new TPaneSelection();
            Result.ActiveRow = Source.ActiveRow;
            Result.ActiveCol = Source.ActiveCol;
            if (Source.SelectedCells==null)
            {
                Result.SelectedCells=null;
            }
            else
            {
                Result.SelectedCells = new TXlsCellRange[Source.SelectedCells.Length];
                for (int i=0;i< Result.SelectedCells.Length;i++)
                {
                    Result.SelectedCells[i] = (TXlsCellRange)Source.SelectedCells[i].Clone();
                }
            }
            return Result;
        }

        UInt16 ActualRecordSize(ref int Pos)
        {
            if (Pos> SelectedCells.Length) FlxMessages.ThrowException(FlxErr.ErrInternal);
            int DeltaPos = ((XlsConsts.MaxRecordDataSize - XlsConsts.SizeOfTRecordHeader -9) / 6)-1;
            Pos+=DeltaPos;
            int IncludedCells = DeltaPos;
            if (Pos>SelectedCells.Length) 
            {
                IncludedCells= SelectedCells.Length-(Pos-DeltaPos);
                Pos=SelectedCells.Length;
            }
            return (UInt16)( XlsConsts.SizeOfTRecordHeader + 
                9 + //Fixed part
                IncludedCells*6);
        } 
        
        internal void SaveToStream(IDataStream DataStream, TSaveData SaveData, byte PanePos)
        {
            if (SelectedCells==null || SelectedCells.Length==0) return;

            if ((SaveData.ExcludedRecords & TExcludedRecords.CellSelection) != 0) return; //Note that this will invalidate the size, but it doesnt matter as this is not saved for real use. We could write blanks here if we wanted to keep the offsets right.

            byte[] ppos = new byte[1]; ppos[0]=PanePos;

            int CurrentPos=0;
            do
            {
                int OldPos = CurrentPos;
                UInt16 RecordSize = ActualRecordSize(ref CurrentPos);
                DataStream.WriteHeader((UInt16)xlr.SELECTION, (UInt16)( RecordSize - XlsConsts.SizeOfTRecordHeader));
                DataStream.Write(ppos, ppos.Length);

                DataStream.WriteRow(ActiveRow);
                DataStream.WriteCol(ActiveCol);

                DataStream.Write16((UInt16)ActiveSel); 

                int len = CurrentPos-OldPos;
                DataStream.Write16((UInt16)len);

                unchecked
                {
                    for (int i=0; i< len;i++)
                    {
                        DataStream.WriteRow(SelectedCells[i].Top);
                        DataStream.WriteRow(SelectedCells[i].Bottom);
                        DataStream.WriteColByte(SelectedCells[i].Left);
                        DataStream.WriteColByte(SelectedCells[i].Right);
                    }
                }
            }
            while (CurrentPos<SelectedCells.Length);

        }

        internal void SaveToPxl(TPxlStream PxlStream, TPxlSaveData SaveData)
        {
            if (ActiveRow > FlxConsts.Max_PxlRows) return;
            if (ActiveCol > FlxConsts.Max_PxlColumns) return;

            TXlsCellRange Selection = new TXlsCellRange(ActiveRow, ActiveCol, ActiveRow, ActiveCol);
            if (SelectedCells!=null && SelectedCells.Length>0) 
            {
                TXlsCellRange Selection2 = SelectedCells[0];
                if (Selection2.Left >= 0 && Selection2.Right <= FlxConsts.Max_PxlColumns && Selection2.Left <= Selection2.Right)
                {
                    if (Selection2.Top >= 0 && Selection2.Bottom <= FlxConsts.Max_PxlRows && Selection2.Top <= Selection2.Bottom)
                    {
                        Selection = Selection2;
                    }
                }
            }

            PxlStream.WriteByte((byte) pxl.SELECTION);
            PxlStream.Write16((UInt16)Selection.Top); 
            PxlStream.WriteByte((byte)Selection.Left);
            PxlStream.Write16((UInt16)Selection.Bottom); 
            PxlStream.WriteByte((byte)Selection.Right);

            PxlStream.Write16((UInt16)ActiveRow); 
            PxlStream.WriteByte((byte)ActiveCol);

        }

        internal long TotalSize()
        {
            if (SelectedCells==null || SelectedCells.Length==0) return 0;
            long Result=0;
            int CurrentPos=0;
            do
            {
                Result += ActualRecordSize(ref CurrentPos);
            }
            while (CurrentPos<SelectedCells.Length);

            return Result;
        }

        internal void AddRecord(TBiff8SelectionRecord aRecord)
        {
            ActiveRow = aRecord.ActiveCellRow;
            ActiveCol = aRecord.ActiveCellCol;

            TXlsCellRange[] Selection = aRecord.SelectedCells;
            if (SelectedCells==null) //99.999% of the cases. we need hundreds of separate cells to be selected in order to need more than one.
                SelectedCells=Selection;
            else  //Not very eficient, but it will almost never happen, and when it happens it doesn't matter.
            {
                TXlsCellRange[] FinalSelection = new TXlsCellRange[SelectedCells.Length + Selection.Length];
                Array.Copy(SelectedCells,0,FinalSelection,0,SelectedCells.Length);
                Array.Copy(Selection,0,FinalSelection,SelectedCells.Length,Selection.Length);
                SelectedCells = FinalSelection;
            }
            ActiveSel = aRecord.ActiveCellRef;
        }

        internal TXlsCellRange[] GetSelection()
        {
            return SelectedCells;
        }

        internal TCellAddress GetActiveCellBase1()
        { 
            return new TCellAddress(ActiveRow + 1, ActiveCol + 1);
        }

        internal int GetActiveCellId()
        {
            return ActiveSel;
        }
        
        internal void Select(TXlsCellRange[] CellRange,int aActiveRow, int aActiveCol, int aActiveCellId)
        {
            if (CellRange==null || CellRange.Length==0) return;
            if (aActiveRow < 0) ActiveRow = CellRange[0].Top; else ActiveRow = aActiveRow;
            if (aActiveCol < 0) ActiveCol = CellRange[0].Left; else ActiveCol = aActiveCol;
            ActiveSel = aActiveCellId;

            SelectedCells = CellRange;
        }

    }

    /// <summary>
    /// A whole selection for a sheet.
    /// </summary>
    internal class TSheetSelection
    {
        private TPaneSelection[] Panes;

        internal TSheetSelection()
        {
            Panes = new TPaneSelection[4];
        }

        internal static TSheetSelection Clone(TSheetSelection Source)
        {
            if (Source == null) return null;
            TSheetSelection Result = new TSheetSelection();
        {
            for (int i=0; i< Result.Panes.Length; i++)
            {
                Result.Panes[i] = TPaneSelection.Clone(Source.Panes[i]);
            }
        }
            return Result;
        }

        internal long TotalSize(TWindow win)
        {
            long Result =0;
            for (int i= Panes.Length-1; i>=0; i--)
            {
                if (MustSavePane(i, win)) Result+= Panes[i].TotalSize();
            }
            return Result;
        }

        internal bool MustSavePane(int i, TWindow win)
        {
            if (Panes[i] == null || win == null) return false;
            if (win.Window2.IsFrozen)
            {
                return i == (int)win.Pane.ActivePaneForSelection();
            }
            return true;
        }

        internal void AddBiff8Record(TBiff8SelectionRecord aRecord)
        {
            int PanePos = (int)aRecord.PanePosition;
            if (PanePos<0 || PanePos>3) return; //invalid pane.

            if (Panes[PanePos]==null) Panes[PanePos]= new TPaneSelection();
            Panes[PanePos].AddRecord(aRecord);
        }

        internal void Select(TPanePosition Pane, TXlsCellRange[] CellRange, int ActiveRow, int ActiveCol, int ActiveCellId)
        {
            int PanePos = (int)Pane;
            if (PanePos<0 || PanePos>3) return; //invalid pane.

            if (Panes[PanePos]==null) Panes[PanePos]= new TPaneSelection();
            Panes[PanePos].Select(CellRange, ActiveRow, ActiveCol, ActiveCellId);
        }


        private TPaneSelection CalcPane(TPanePosition Pane)
        {
            int PanePos = (int)Pane;
            if (PanePos < 0 || PanePos > 3) return null; //invalid pane.

            if (Panes[PanePos] == null) return null;
            return Panes[PanePos];
        }
        
        internal TXlsCellRange[] GetSelection(TPanePosition Pane)
        {
            TPaneSelection P = CalcPane(Pane);
            if (P == null) return null;
            return P.GetSelection();
        }

        internal TCellAddress GetActiveCellBase1(TPanePosition Pane)
        {
            TPaneSelection P = CalcPane(Pane);
            if (P == null) return null;
            return P.GetActiveCellBase1();
        }

        internal int ActiveCellId(TPanePosition Pane)
        {
            TPaneSelection P = CalcPane(Pane);
            if (P == null) return 0;
            return P.GetActiveCellId();
        }

        internal void SaveToStream(IDataStream DataStream, TSaveData SaveData, TWindow win)
        {
            byte[] RecordOrder = {3,1,2,0};

            for (int i=0; i< Panes.Length; i++)
            {
                if (MustSavePane(RecordOrder[i], win))
                {
                    Panes[RecordOrder[i]].SaveToStream(DataStream, SaveData, (byte)RecordOrder[i]);
                }
            }
        }

        internal void SaveToPxl(TPxlStream PxlStream, TPxlSaveData SaveData)
        {
            TPaneSelection Pane = Panes[3];
            if (Pane == null) return;
            Pane.SaveToPxl(PxlStream, SaveData);
        }
    }


}
