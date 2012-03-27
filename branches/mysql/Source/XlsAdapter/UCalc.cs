using System;
using FlexCel.Core;
using System.Diagnostics;
using System.Collections.Generic;

namespace FlexCel.XlsAdapter
{
    internal class TFormulaCache
    {
        private TFormulaRecord FirstFormula;
        private TFormulaRecord LastFormula;

        internal TFormulaCache()
        {
            FirstFormula = null;
            LastFormula = null;
        }

        public void Add(TFormulaRecord r)
        {
            r.Prev = LastFormula;
            r.Next = null;
            if (FirstFormula == null) FirstFormula = r;
            if (LastFormula != null) LastFormula.Next = r;
            LastFormula = r;
        }

        public void Delete(TFormulaRecord r)
        {
            if (r.Prev == null) //first record
            {
                Debug.Assert(r == FirstFormula);
                FirstFormula = r.Next;
            }
            else
            {
                r.Prev.Next = r.Next;
            }

            if (r.Next == null) //Last record
            {
                Debug.Assert(r == LastFormula);
                LastFormula = r.Prev;
            }
            else
            {
                r.Next.Prev = r.Prev;
            }

        }

        public void Clear()
        {
            FirstFormula = null;
            LastFormula = null;
        }

        internal void ArrangeInsertRangeRows(TXlsCellRange CellRange, int aRowCount, TSheetInfo SheetInfo)
        {
            TFormulaRecord r= FirstFormula;
            while (r != null)
            {
                r.ArrangeInsertRange(r.FRow, CellRange, aRowCount, 0, SheetInfo);
                r = r.Next;
            }
        }

        public TCellAddress[] GetTables()
        {
            List<TCellAddress> Cells = new List<TCellAddress>();
            TFormulaRecord r = FirstFormula;
            while (r != null)
            {
                if (r.TableRecord != null)
                {
                    Cells.Add(new TCellAddress(r.FRow + 1, r.Col + 1));
                }

                r = r.Next;
            }
            return Cells.ToArray();
        }

        internal void ArrangeInsertRangeCols(TXlsCellRange CellRange, int aColCount, TSheetInfo SheetInfo)
        {
            TFormulaRecord r = FirstFormula;
            while (r != null)
            {
                r.ArrangeInsertRange(r.FRow, CellRange, 0, aColCount, SheetInfo);
                r = r.Next;
            }
        }

		internal void ArrangeMoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
		{
            TFormulaRecord r = FirstFormula;
            while (r != null)
            {
                r.ArrangeMoveRange(CellRange, NewRow, NewCol, SheetInfo);
                r = r.Next;
			}
		}

        internal void ArrangeInsertSheet(TSheetInfo SheetInfo)
        {
            TFormulaRecord r = FirstFormula;
            while (r != null)
            {
                r.ArrangeInsertSheet(SheetInfo);
                r = r.Next;
            }
        }

        public void Recalc(TCellList CellList, ExcelFile aXls, int SheetIndexBase1)   
        {
			//Do this loop in normal order, not reversed, to help avoid stack overflows. Normally references grow down and to the right 
			//(for example cell A2 = "A1+1"). If we do the loop in reverse, we start by A900, and to calculate it we need A899 so we need A898...
			//And this can lead to an stack overflow error. Of course this is no guarantee, an user migh have A2= A3-1, etc, but that is less likely.

            TFormulaRecord r = FirstFormula;
            while (r != null)
            {
                if (aXls != null)
                {
                    aXls.SetUnsupportedFormulaCellAddress(new TCellAddress(aXls.GetSheetName(SheetIndexBase1), r.FRow + 1, r.Col + 1, false, false));
                }

                TCalcState CalcState = new TCalcState();
                TFormulaRecord rPrev = r.Prev; //r.prev won't be modified by Recalc, as it has already been recalculated
                r.Recalc(CellList, aXls, SheetIndexBase1, CalcState, new TCalcStack());
                
                if (rPrev == null) r = FirstFormula; else r = rPrev.Next; //Recalc might reorder the calc chain, so we want the next in the chain position, not the next to the modified r..

                r = r.Next;
            }
        }

        public void ClearResults()   
        {
            TFormulaRecord r = FirstFormula;
            while (r != null)
            {
                r.ClearResult();
                r = r.Next;
            }
        }
  
        public void ForceAutoRecalc()   
        {
            TFormulaRecord r = FirstFormula;
            while (r != null)
            {
                r.ForceAutoRecalc();
                r = r.Next;
            }
        }

        public void CleanFlags()
        {
            TFormulaRecord r = FirstFormula;
            while (r != null)
            {
                r.NotRecalculating();
                r.NotRecalculated();
                r = r.Next;
            }
        }

        public bool HasData()
        {
            return FirstFormula != null;
        }

		#region Named Ranges
        internal void UpdateDeletedRanges(TDeletedRanges DeletedRanges)
        {
            TFormulaRecord r = FirstFormula;
            while (r != null)
            {
                r.UpdateDeletedRanges(DeletedRanges);
                r = r.Next;
            }
        }
		
		#endregion


        #region Reorder Calc Chain

        public bool ReorderCalcChain(IFormulaRecord i1, IFormulaRecord i2)
        {
            TFormulaRecord r1 = i1 as TFormulaRecord;
            TFormulaRecord r2 = i2 as TFormulaRecord;
            
            if (r1 == null || r2 == null || r1 == r2) return false;
            if (r2.Prev == null) return false;
            if (r2.Prev == r1) { SwapAdjacent(r1, r2); return true; }
            if (r1.Prev == r2) { SwapAdjacent(r2, r1); return true; }

            SwapNeighbors(r1, r2);
            return true;
        }

        private void SwapNeighbors(TFormulaRecord r1, TFormulaRecord r2)
        {
            if (r1.Prev == null)
            {
                FirstFormula = r2;
            }
            else
            {
                r1.Prev.Next = r2;
            }

            if (r2.Next == null)
            {
                LastFormula = r2.Prev;
            }
            else
            {
                r2.Next.Prev = r2.Prev;
            }

            r2.Prev.Next = r2.Next;

            r2.Prev = r1.Prev;
            r1.Prev = r2;
            r2.Next = r1;
        }

        private void SwapAdjacent(TFormulaRecord r1, TFormulaRecord r2)
        {
            if (r1.Prev == null)
            {
                FirstFormula = r2;
            }
            else
            {
                r1.Prev.Next = r2;
            }

            if (r2.Next == null)
            {
                LastFormula = r1;
            }
            else
            {
                r2.Next.Prev = r1;
            }

            r1.Next = r2.Next;
            r2.Prev = r1.Prev;
            r2.Next = r1;
            r1.Prev = r2;
        }
        #endregion
    }
}
