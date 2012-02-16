using System;
using FlexCel.Core;
using System.Text;
using System.Globalization;

namespace FlexCel.XlsAdapter
{
	/// <summary>
	/// A class to perform search and /or replace on a sheet.
	/// </summary>
	internal sealed class TSearchAndReplace
	{
		private TSearchAndReplace()	{}

		public static void Search(ExcelFile xls, TXlsCellRange Range, TCellAddress Start, bool ByRows, TSearchOrReplace Action)
		{
			int StartRow = 1;
			int EndRow = 0;
			int StartCol = 1;
			int EndCol = 0;
			if (Range != null)
			{
				StartRow = Range.Top;
				EndRow = Range.Bottom;
				StartCol = Range.Left;
				EndCol = Range.Right;
			}
			else
			{
				EndRow = xls.RowCount;
				EndCol = xls.ColCount;
			}

			int FirstStartCol = StartCol;
			int FirstStartRow = StartRow;
			if (Start != null)
			{
				FirstStartRow = Start.Row;
				if (ByRows) FirstStartRow++;
				FirstStartCol = Start.Col;
				if (!ByRows) FirstStartCol++;

				if (FirstStartRow > EndRow) 
				{
					FirstStartRow = StartRow;
					FirstStartCol++;
				}
				
				if (FirstStartCol > EndCol)
				{
					FirstStartCol = StartCol;
					FirstStartRow++;
				}
			}

			if (ByRows)
				SearchByRows(xls, FirstStartRow, StartRow, EndRow, FirstStartCol, EndCol, Action);
			else
				SearchByCols(xls, FirstStartRow, EndRow, FirstStartCol, StartCol, EndCol, Action);


		}

		private static void SearchByRows(ExcelFile xls, int FirstStartRow, int StartRow, int EndRow, int StartCol, int EndCol, TSearchOrReplace Action)
		{
            int RealEndRow = Math.Min(xls.RowCount, EndRow);
            int RealEndCol = 0; //we won't use xls.ColCount as it would have to loop over all rows. We can do this here anyway.
            
			for (int col = StartCol; col <= EndCol; col++)
			{
				int sr;
                if (col == StartCol)
                {
                    sr = FirstStartRow;
                }
                else
                {
                    sr = StartRow;
                }

                if (col > StartCol + 1 && col > RealEndCol) break;
                
				for (int row = sr; row <= RealEndRow; row++)
				{
                    if (col <= StartCol + 1) //We might start searching in the middle of a range. So we need to have at least 2 columns checked to be sure all rows have been inspected for max col.
                    {
                        int ccInRow = xls.ColFromIndex(row, xls.ColCountInRow(row));
                        if (ccInRow > RealEndCol) RealEndCol = ccInRow;
                    }

					object OldVal = xls.GetCellValue(row, col);
					if (Action.Go(xls, OldVal, row, col)) return;
				}
			}
			Action.Clear();
		}

		private static void SearchByCols(ExcelFile xls, int StartRow, int EndRow, int FirstStartCol, int StartCol, int EndCol, TSearchOrReplace Action)
		{
            int RealEndRow = Math.Min(xls.RowCount, EndRow);
			for (int row = StartRow; row <= RealEndRow; row++)
			{
				int cIndex = row == StartRow? xls.ColToIndex(StartRow, FirstStartCol): xls.ColToIndex(StartRow, StartCol);
				int ColsInRow = xls.ColCountInRow(row);
				while (cIndex <= ColsInRow)
				{
					int XF = -1;
					int col = xls.ColFromIndex(row, cIndex);
					if (col < StartCol) 
					{
						cIndex++;
						continue;
					}
					if (col > EndCol) break;
					object OldVal = xls.GetCellValueIndexed(row, cIndex, ref XF);
					if (Action.Go(xls, OldVal, row, col)) return;
					cIndex++;
				}
			}
			Action.Clear();
		}

	}

	/// <summary>
	/// An abstract class that will be used by specialized search or replace actions.
	/// </summary>
	internal abstract class TSearchOrReplace
	{
		protected readonly bool FCaseInsensitive, FSearchInFormulas, FWholeCellContents;
		private StringComparison FStringComparison;
		object SearchItem;
		string SearchStr;
		protected string UpperSearchItem;
		protected string SearchToStr;

        internal TSearchOrReplace(bool aDates1904, object aSearchItem, bool aCaseInsensitive, bool aSearchInFormulas, bool aWholeCellContents)
		{
			SearchItem = TExcelTypes.ConvertToAllowedObject(aSearchItem, aDates1904);
			SearchStr = aSearchItem as String;
			FCaseInsensitive = aCaseInsensitive;
			if (aCaseInsensitive) FStringComparison = StringComparison.CurrentCultureIgnoreCase; else FStringComparison = StringComparison.CurrentCulture;

			FSearchInFormulas = aSearchInFormulas;
			FWholeCellContents = aWholeCellContents;

			if (SearchItem != null) 
			{
				SearchToStr = aSearchItem.ToString();
				UpperSearchItem = SearchToStr.ToUpper(CultureInfo.CurrentCulture);
			}
		}

		internal virtual bool Go(ExcelFile xls, object value, int row, int col)
		{
			object v = value;
			if (FSearchInFormulas)
			{
				TFormula fmla = value as TFormula;
				if (fmla != null)
				{
                    if (!fmla.Span.IsTopLeft) return false; //can't change part of a shared formula.
					v = fmla.Text;
				}
			}

			if (SearchItem == null)
			{
				if (v == null) 
				{
					return OnFound(xls, v, row, col);
				}
				return false;
			}

			if (v == null) return false;

			if (FWholeCellContents)
			{
				if (SearchStr != null)
				{
					string sv = v as string;
					if (sv != null)
					{
						if (String.Equals(sv, SearchStr, FStringComparison))
						{
							return OnFound(xls, v, row, col);
						}
					}
					return false;
				}

				if (SearchItem.Equals(v))
				{
					return OnFound(xls, v, row, col);
				}
				return false;
			}

			//partial searches
			string sv2 = v.ToString();
			if (FCaseInsensitive)
			{
				if (sv2.ToUpper(CultureInfo.CurrentCulture).IndexOf(UpperSearchItem) >=0) 
				{
					return OnFound(xls, v, row, col);
				}
			}
			else
			{
				if (sv2.IndexOf(SearchToStr) >=0) 
				{
					return OnFound(xls, v, row, col);
				}
			}


			return false;
		}

		internal abstract void Clear();

		internal abstract bool OnFound(ExcelFile xls, object oldval, int row, int col);
	}

	internal class TSearch: TSearchOrReplace
	{

		public TCellAddress Cell;

		internal TSearch(bool aDates1904, object aSearchItem, bool aCaseInsensitive, bool aSearchInFormulas, bool aWholeCellContents): base(aDates1904, aSearchItem, aCaseInsensitive, aSearchInFormulas, aWholeCellContents)
		{		
		}

		internal override bool Go(ExcelFile xls, object value, int row, int col)
		{
			Cell = null;
			return base.Go (xls, value, row, col);
		}


		internal override bool OnFound(ExcelFile xls, object oldval, int row, int col)
		{
			Cell = new TCellAddress(row, col);
			return true;
		}

		internal override void Clear()
		{
			Cell = null;
		}

	}


	internal class TReplace: TSearchOrReplace
	{
		TRichString NewValue;
		public int ReplaceCount;

		internal TReplace(bool aDates1904, object aSearchItem, object aNewValue, bool aCaseInsensitive, bool aSearchInFormulas, bool aWholeCellContents): base(aDates1904, aSearchItem, aCaseInsensitive, aSearchInFormulas, aWholeCellContents)
		{
			NewValue = aNewValue as TRichString;
			if (NewValue == null && aNewValue != null)
			{
				NewValue = new TRichString(aNewValue.ToString());
			}
			if (NewValue == null) NewValue = new TRichString();
		}

		internal override bool OnFound(ExcelFile xls, object oldval, int row, int col)
		{
            TFormula oldFmla = oldval as TFormula;

			if (!FSearchInFormulas && oldFmla != null) return false; //do not replace if it is a formula.
			if (oldval == null) return false;
			
			TRichString OldStr = oldval as TRichString;
			if (OldStr == null) OldStr = new TRichString(FlxConvert.ToString(oldval));
			TRichString newStr = OldStr.Replace(SearchToStr, NewValue.ToString(), FCaseInsensitive);

			if (newStr != null && newStr.Value != null && newStr.Length > 0 && newStr.Value.StartsWith(TFormulaMessages.TokenString(TFormulaToken.fmStartFormula)))
			{
                TFormulaSpan Span = oldFmla != null ? oldFmla.Span : new TFormulaSpan();
				xls.SetCellValue(row, col, new TFormula(newStr.Value, null, Span));
			}
			else
			{
                if (oldFmla != null && !oldFmla.Span.IsOneCell) return false; //can't replace a shared formula with simple text.
                xls.SetCellFromString(row, col, newStr);
			}

			ReplaceCount++;
			return false;
		}




		internal override void Clear()
		{
		}

	}


}
