using System;
using System.Diagnostics;
using System.Collections;
using System.Collections.Generic;

using FlexCel.Core;

namespace FlexCel.XlsAdapter
{
	/// <summary>
	/// A static utility class to sort a range of cells.
	/// </summary>
	internal sealed class TSortRange
	{
		private TSortRange()
		{
		}

		internal static void Sort (XlsFile xls, TXlsCellRange Range, bool ByRows, int[] Keys, TSortOrder[] SortOrder, IComparer Comparer)
		{
			xls.Recalc(); //make sure formulas are ok. Also, sets Date1904 to the correct value.
			TKeyList Items = new TKeyList();
			if (ByRows) FillRowRange(Items, xls, Range, Keys, SortOrder, Comparer);
			else FillColRange(Items, xls, Range, Keys, SortOrder, Comparer);

			Items.Sort();

			MoveRowRange(Items, xls, Range, ByRows);
		}

		private static void FillRowRange(TKeyList Items, XlsFile xls, TXlsCellRange Range, int[] Keys, TSortOrder[] SortOrder, IComparer Comparer)
		{
			for (int r = Range.Top; r <= Range.Bottom; r++)
			{
				object[] Values;
				if (Keys == null)
				{
					int Len = Range.ColCount;
					if (Len > 8) Len = 8; //avoid taking up too much memory.
					Values = new object[Len];
					for (int c = 0; c < Values.Length; c++)
					{
						Values[c] =  xls.GetCellValue(r, Range.Left + c);
					}

				}
				else
				{
					Values = new object[Keys.Length];
					for (int c = 0; c < Keys.Length; c++)
					{
						Values[c] =  xls.GetCellValue(r, Keys[c]);
					}

				}
				Items.Add(new TKeyItem(r, Values, SortOrder, Comparer));
			}
		}

		private static void FillColRange(TKeyList Items, XlsFile xls, TXlsCellRange Range, int[] Keys, TSortOrder[] SortOrder, IComparer Comparer)
		{
			for (int c = Range.Left; c <= Range.Right; c++)
			{
				object[] Values;
				if (Keys == null)
				{
					int Len = Range.RowCount;
					if (Len > 8) Len = 8; //avoid taking up too much memory.
					Values = new object[Len];
					for (int r = 0; r < Values.Length; r++)
					{
						Values[r] =  xls.GetCellValue(Range.Top + r, c);
					}

				}
				else
				{
					Values = new object[Keys.Length];
					for (int r = 0; r < Keys.Length; r++)
					{
						Values[r] =  xls.GetCellValue(Keys[r], c);
					}

				}
				Items.Add(new TKeyItem(c, Values, SortOrder, Comparer));
			}
		}
		
		private static void MoveRowRange(TKeyList Items, ExcelFile xls, TXlsCellRange Range, bool ByRows)
		{
			if (!ByRows) Range = Range.Transpose();
			//optimization if part (or all) of the range is already sorted.
			int FirstRow = Range.Top;
            FirstRow = AdvanceToNextUnsorted(Items, Range, FirstRow);
			if (FirstRow > Range.Bottom) return;
            int LastRow = Range.Bottom;
            int StagingRow = LastRow + 1;

			if (ByRows)
				xls.InsertAndCopyRange(new TXlsCellRange(StagingRow, Range.Left, StagingRow, Range.Right), StagingRow, Range.Left, 1, TFlxInsertMode.ShiftRangeDown, TRangeCopyMode.None); //add a row at the bottom to swap rows.
			else
				xls.InsertAndCopyRange(new TXlsCellRange(Range.Left, StagingRow, Range.Right, StagingRow), Range.Left, StagingRow, 1, TFlxInsertMode.ShiftRangeRight, TRangeCopyMode.None); //add a column at the right to swap columns.

            do
            {
                PartialSort(Items, xls, Range, ByRows, FirstRow, StagingRow);
                int fr = FirstRow;
                FirstRow = AdvanceToNextUnsorted(Items, Range, FirstRow);
                if (fr == FirstRow) FlxMessages.ThrowException(FlxErr.ErrInternal);

            } while (FirstRow < Range.Bottom);

			if (ByRows)
				xls.DeleteRange(new TXlsCellRange(StagingRow, Range.Left, StagingRow, Range.Right), TFlxInsertMode.ShiftRangeDown);
			else
				xls.DeleteRange(new TXlsCellRange(Range.Left, StagingRow, Range.Right, StagingRow), TFlxInsertMode.ShiftRangeRight);
		}

        private static int AdvanceToNextUnsorted(TKeyList Items, TXlsCellRange Range, int FirstRow)
        {
            while (FirstRow <= Range.Bottom)
            {
                if (Items[FirstRow - Range.Top].Position == FirstRow)
                {
                    FirstRow++;
                }
                else break;
            }
            return FirstRow;
        }

        private static void PartialSort(TKeyList Items, ExcelFile xls, TXlsCellRange Range, bool ByRows, int FirstRow, int StagingRow)
        {
            int Slot = FirstRow - Range.Top;
            int LastSlot = -1;
            int destRow = StagingRow;
            int Row = FirstRow;
            do
            {
                if (Row == destRow) FlxMessages.ThrowException(FlxErr.ErrInternal);
                MoveRows(xls, Range, ByRows, Row, destRow);
                if (LastSlot >= 0) Items[LastSlot].Position = destRow;

                int MovedFrom = Row;
                if (Slot == Items.Count) break; //we are moving from the staging area.
                Row = Items[Slot].Position;
                if (Row == FirstRow) Row = StagingRow;
                LastSlot = Slot;
                Slot = Row - Range.Top;
                destRow = MovedFrom;

            } while (true);
        }

        private static void MoveRows(ExcelFile xls, TXlsCellRange Range, bool ByRows, int Row, int DestRow)
        {
            if (ByRows)
                xls.MoveRange(new TXlsCellRange(Row, Range.Left, Row, Range.Right), DestRow, Range.Left, TFlxInsertMode.NoneDown);
            else
                xls.MoveRange(new TXlsCellRange(Range.Left, Row, Range.Right, Row), Range.Left, DestRow, TFlxInsertMode.NoneDown);
        }
	}



	internal class TKeyItem: IComparable
	{
		internal int Position;
		internal object[] Values;
		TSortOrder[] SortOrder;
		internal IComparer Comparer;

		internal TKeyItem(int aPosition, object[] aValues, TSortOrder[] aSortOrder, IComparer aComparer)
		{
			Position = aPosition;
			Values = aValues;
			SortOrder = aSortOrder;
			if (aComparer == null) Comparer = TCellComparer.Instance; else Comparer = aComparer;
		}

		#region IComparable Members

		public int CompareTo(object obj)
		{
			TKeyItem obj2 = obj as TKeyItem;
			if (obj2 == null) return -1;
			if (Values.Length > obj2.Values.Length) return 1;
			if (Values.Length < obj2.Values.Length) return -1;

			for (int i = 0; i < Values.Length; i++)
			{
				int Result = Comparer.Compare(Values[i], obj2.Values[i]);
				if (Values[i] != null && obj2.Values[i] != null)  //null values do not change with sort order.
				{
					if (SortOrder != null && i < SortOrder.Length && SortOrder[i] == TSortOrder.Descending) Result = -Result;
				}
				if (Result != 0) return Result;
			}
			return 0;
		}

		#endregion
	}

#if (FRAMEWORK20)
    internal class TKeyList : List<TKeyItem>
    {
#else
	internal class TKeyList: ArrayList
	{
		public new TKeyItem this[int index] {get{return (TKeyItem) base[index];}}
#endif
    }

	internal class TCellComparer: IComparer
	{
		internal static readonly TCellComparer Instance= new TCellComparer(); //STATIC*

		#region IComparer Members

		public int Compare(object x, object y)
		{
			x = TExcelTypes.ConvertToAllowedObject(x, TBaseParsedToken.Dates1904);
			y = TExcelTypes.ConvertToAllowedObject(y, TBaseParsedToken.Dates1904);
			//null values go always at the bottom. (in ascending or descending order)
			if (x == null)
			{
				if (y == null) return 0;
				return 1;
			}
			if (y == null) return -1;

			if (x is TFlxFormulaErrorValue)
			{
				if (y is TFlxFormulaErrorValue)
				{
					return ((int)x).CompareTo((int)y);
				}
				return 1;
			}
			if (y is TFlxFormulaErrorValue) return -1;


			object Result = TBaseParsedToken.CompareValues(x, y);
			if (Result is int) return (int)Result;
			return 0;
		}

		#endregion
	}

}
