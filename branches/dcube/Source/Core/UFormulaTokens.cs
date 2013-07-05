using System;
using System.Text;
using System.Globalization;
using System.Diagnostics;
using System.Runtime.Serialization;
using System.Collections.Generic;
using System.Collections;

#if (MONOTOUCH)
using Color = MonoTouch.UIKit.UIColor;
using System.Drawing;
#else
	#if (WPF)
	using System.Windows.Media;
	#else
	using Colors = System.Drawing.Color;
	using System.Drawing;
	#endif
#endif

namespace FlexCel.Core
{
	#region WorkbookInfo
	internal class TWorkbookInfo
	{
		internal ExcelFile Xls;
		internal int SheetIndexBase1;
		internal int Row;
		internal int Col;
		internal int RowCount;
		internal int ColCount;

		internal int RowOfs;
		internal int ColOfs;

		internal int SumProductCount;

		internal bool IsArrayFormula;

		internal TDebugStack DebugStack;
		internal int FullDataSetIndex; //Not the best place for it to be...

		internal TWorkbookInfo(ExcelFile aXls, int aSheetIndexBase1, int aRow, int aCol, int aRowCount, int aColCount, int aRowOfs, int aColOfs, bool aIsArrayFormula)
		{
			Xls = aXls;
			SheetIndexBase1 = aSheetIndexBase1;
			Row = aRow;
			Col = aCol;
			RowCount = aRowCount;
			ColCount = aColCount;

			RowOfs = aRowOfs;
			ColOfs = aColOfs;
			SumProductCount = 0;

			IsArrayFormula = aIsArrayFormula;
		}

		/// <summary>
		/// Will not do a deep copy, so the debugstack will be the same.
		/// </summary>
		/// <returns></returns>
		internal TWorkbookInfo ShallowClone()
		{
			return (TWorkbookInfo) MemberwiseClone();
		}

        internal void AddUnsupported(TUnsupportedFormulaErrorType ErrorType, string FuncName)
        {
            if (Xls != null) Xls.AddUnsupported(ErrorType, FuncName);
        }
    }
	#endregion

	#region DebugInfo
	internal class TDebugItem
	{
		internal string Tag;
		internal object Value;
		internal int Level;

		internal TDebugItem(string aTag, object aValue, int aLevel)
		{
			Tag = aTag;
			Value = aValue;
			Level = aLevel;
		}
	}

	internal class TDebugStack
	{
        private List<TDebugItem> FList;
		private int Level;

		internal TDebugStack()
		{
            FList = new List<TDebugItem>();
		}

		internal TDebugItem Add(string aTag, object aValue)
		{
			TDebugItem Result = new TDebugItem(aTag, aValue, Level);
			FList.Add(Result);
			return Result;
		}

		internal void IncLevel()
		{
			Level++;
		}

		internal void DecLevel()
		{
			Level--;
		}

		private static void AddPreamble(StringBuilder sb, int aLevel)
		{
			for (int i = 0; i < aLevel; i++)
			{
				sb.Append("   ");
			}
		}

		private void AddArray(ExcelFile Workbook, StringBuilder sb, TFlxFont BaseFont, TRTFList RTF, Array a)
		{
			int SecondDimension = a.Rank < 2 ? 1 : a.GetLength(1);
			for (int r = 0; r < SecondDimension; r++)
			{
				SetColor(Workbook, sb.Length, BaseFont, RTF, BaseFont.Color);
				sb.Append(TFormulaMessages.TokenChar(TFormulaToken.fmOpenArray));
				for (int i = 0; i < a.GetLength(0); i++)
				{
					if (a.Rank < 2) AddValue(Workbook, sb, BaseFont, RTF, a.GetValue(i)); else AddValue(Workbook, sb, BaseFont, RTF, a.GetValue(i, r));
					SetColor(Workbook, sb.Length, BaseFont, RTF, BaseFont.Color);
					if (i < a.GetLength(0) - 1) sb.Append(", ");
				}
				sb.Append(TFormulaMessages.TokenChar(TFormulaToken.fmCloseArray));
			}
		}

		private void SetColor(ExcelFile Workbook, int Position, TFlxFont BaseFont, TRTFList RTF, TExcelColor aColor)
		{
			TRTFRun run;
			run.FirstChar = Position;
			TFlxFont NewFont = (TFlxFont)BaseFont.Clone();
			NewFont.Color = aColor;
			run.FontIndex = Workbook.AddFont(NewFont);
			RTF.Add(run);
		}

		private void AddValue(ExcelFile Workbook, StringBuilder sb, TFlxFont BaseFont, TRTFList RTF, object Value)
		{
			if (Value == null) sb.Append("null");
			else
			{
				Array arr = Value as Array;
				if (arr != null) AddArray(Workbook, sb, BaseFont, RTF, arr);
				else 
				{
					if (Value is String || Value is TRichString)
					{
						SetColor(Workbook, sb.Length, BaseFont, RTF, TExcelColor.FromArgb(0x993300));
						sb.Append("\u201c" + Value.ToString() + "\u201d");
					}
					else
					{
						TExcelColor TextColor = TExcelColor.Automatic;
						if (Value is Boolean) TextColor = Colors.Blue;
						else if (Value is TFlxFormulaErrorValue) TextColor = Colors.Red;

                        if (!TextColor.IsAutomatic) SetColor(Workbook, sb.Length, BaseFont, RTF, TextColor); 
                        else
						{
							SetColor(Workbook, sb.Length, BaseFont, RTF, BaseFont.Color);
						}
						sb.Append(Value.ToString());
					}
				}
			}

		}

		public TRichString ToRichString(ExcelFile Workbook, int XF)
		{
			TRTFList RTF = new TRTFList();
			StringBuilder sb = new StringBuilder();
			TFlxFormat baseFmt = Workbook.GetFormat(XF);
			baseFmt.Font.Style = TFlxFontStyles.Bold;
			int fmtKey = Workbook.AddFont(baseFmt.Font);
			baseFmt.Font.Style = TFlxFontStyles.Italic;

            bool First = true;
			foreach (TDebugItem dbg in FList)
			{
                if (!First) sb.Append("\n");
                First = false;

				TRTFRun run;
				run.FirstChar = sb.Length;
				run.FontIndex = fmtKey;
				RTF.Add(run);
				AddPreamble(sb, dbg.Level);
				sb.Append(dbg.Tag);
				sb.Append(": ");

				AddValue(Workbook, sb, baseFmt.Font, RTF, dbg.Value);
			}

			return new TRichString(sb.ToString(), RTF, Workbook);
		}



	}

	#endregion

	#region Utility
	/// <summary>
	/// Used to summarize a range of cells.
	/// </summary>
	internal class TAddress
	{
		internal TWorkbookInfo wi;
		internal string BookName; //null if no external book.
		internal int Sheet;
		internal int Row;
		internal int Col;

		internal TAddress(TWorkbookInfo awi, int aSheet, int aRow, int aCol) : this(awi, null, aSheet, aRow, aCol)
		{
		}

		internal TAddress(TWorkbookInfo awi, string aBookName, int aSheet, int aRow, int aCol)
		{
			wi = awi;
			BookName = aBookName;
			Sheet = aSheet;
			Row = aRow;
			Col = aCol;
		}

	}

	internal class TAddressList
	{
		private TAddress[][] FList;

		internal TAddressList(TAddress[] adr)
		{
			FList = new TAddress[][] { adr };
		}

        internal TAddressList(List<TAddress[]> adr)
        {
            FList = adr.ToArray();
        }

		internal void Add(TAddressList ad2)
		{
			TAddress[][] tmp = new TAddress[FList.Length + ad2.FList.Length][];
			Array.Copy(FList, 0, tmp, 0, FList.Length);
			Array.Copy(ad2.FList, 0, tmp, FList.Length, ad2.FList.Length);
			FList = tmp;
		}

		internal int Count { get { return FList.Length; } }
		internal TAddress[] this[int index]
		{
			get
			{
				return FList[index];
			}
		}

		internal bool Has3dRef()
		{
			for (int i = 0; i < FList.Length; i++)
				if (FList[i][0].wi.Xls != FList[i][1].wi.Xls || FList[i][0].Sheet != FList[i][1].Sheet) return true;
			return false;
		}

	}


	internal struct TOneCellRef
	{
		internal int Row;
		internal int Col;

        internal TOneCellRef(int aRow, int aCol)
        {
            Row = aRow;
            Col = aCol;
        }

		internal bool IsCell(int aRow, int aCol)
		{
			return Row == aRow && Col == aCol;
		}

		internal bool IsEmpty()
		{
			return Row <= 0 || Col <= 0;
		}

        public override bool Equals(object obj)
        {
            return obj is TOneCellRef && this == (TOneCellRef) obj;
        }

        public static bool operator==(TOneCellRef o1, TOneCellRef o2)
        {
            return o1.Row == o2.Row && o1.Col == o2.Col;
        }

        public static bool operator!=(TOneCellRef o1, TOneCellRef o2)
        {
            return o1.Row != o2.Row && o1.Col != o2.Col;
        }

        public override int  GetHashCode()
        {
 	        return HashCoder.GetHash(Row, Col);
        }

	}

    internal struct TCalcStack //being a struct, it is copied each time in each recursive call. This is different from CalcState.
    {
        internal int Level;
        internal IFormulaRecord ParentFmla;
        internal int ParentSheetBase1;
        internal ExcelFile ParentXls;
    }

    internal interface IFormulaRecord
    {
    }


	internal class TCalcState
	{
		internal bool IgnoreHidden;
        internal bool IgnoreErrors;
		internal bool InSubTotal;

        internal bool Aborted;

		internal int WhatIfRow;  //1 based, 0 means none
		internal int WhatIfCol;  //1 based, 0 means none
		internal int WhatIfSheet;
		internal TOneCellRef TableRowCell;
		internal TOneCellRef TableColCell;
		internal object TableRowValue;
		internal object TableColValue;

		internal TCalcState Clone()
		{
			TCalcState Result = new TCalcState();
			Result.IgnoreHidden = IgnoreHidden;
            Result.IgnoreErrors = IgnoreErrors;
			Result.InSubTotal = InSubTotal;
            Result.Aborted = Aborted;

			Result.WhatIfRow = WhatIfRow;
			Result.WhatIfCol = WhatIfCol;
			Result.WhatIfSheet = WhatIfSheet;
			Result.TableRowCell = TableRowCell;
			Result.TableColCell = TableColCell;
			Result.TableRowValue = TableRowValue;
			Result.TableColValue = TableColValue;
			return Result;
		}
	}

	#endregion

	#region Aggregating objects
	internal abstract class TBaseAggregate
	{
		internal abstract object Agg(TWorkbookInfo wi, int Sheet1, int Sheet2, int Row1, int Col1, int Row2, int Col2, TCalcState CalcState, TCalcStack CalcStack);
		internal abstract object AggValues(object val1, object val2);
		internal abstract object AggArray(object[,] arr);
		internal virtual void AggAverages(TAverageValue val1, double val2)
		{
		}

		protected static void OrderRange(ref int Sheet1, ref int Sheet2, ref int Row1, ref int Col1, ref int Row2, ref int Col2)
		{
			TBaseParsedToken.OrderRange(ref Sheet1, ref Sheet2, ref Row1, ref Col1, ref Row2, ref Col2);
		}

		protected static object ConvertToAllowedObject(object o, bool Dates1904)
		{
			return TExcelTypes.ConvertToAllowedObject(o, Dates1904);
		}

		internal virtual bool PropagateOnEquality
		{
			get { return false; }
		}
	}

	internal abstract class TBaseErrAggregate : TBaseAggregate
	{
		internal override object Agg(TWorkbookInfo wi, int Sheet1, int Sheet2, int Row1, int Col1, int Row2, int Col2, TCalcState CalcState, TCalcStack CalcStack)
		{
			if (Sheet1 == Sheet2)
			{
				if (Row1 == Row2)
				{
					if (Col1 == Col2) return wi.Xls.GetCellValueAndRecalc(Sheet1, Row1, Col1, CalcState, CalcStack);
					int wiCol = wi.Col + 1 + wi.ColOfs;
					if (wiCol >= Col1 && wiCol <= Col2) return wi.Xls.GetCellValueAndRecalc(Sheet1, Row1, wiCol, CalcState, CalcStack);
				}
				if (Col1 == Col2)
				{
					int wiRow = wi.Row + 1 + wi.RowOfs;
					if (wiRow >= Row1 && wiRow <= Row2) return wi.Xls.GetCellValueAndRecalc(Sheet1, wiRow, Col1, CalcState, CalcStack);
				}
			}


			return TFlxFormulaErrorValue.ErrValue;
		}

		internal override object AggValues(object val1, object val2)
		{
			return TFlxFormulaErrorValue.ErrValue;
		}
	}

	internal class TErr2Aggregate : TBaseErrAggregate
	{
		internal static readonly TErr2Aggregate Instance = new TErr2Aggregate();

		internal override object AggArray(object[,] arr)
		{
			return arr;
		}

	}

	internal class TArrayAggregate : TBaseAggregate
	{
		internal static readonly TArrayAggregate Instance = new TArrayAggregate();

		internal override bool PropagateOnEquality
		{
			get
			{
				return true;
			}
		}


		internal override object Agg(TWorkbookInfo wi, int Sheet1, int Sheet2, int Row1, int Col1, int Row2, int Col2, TCalcState CalcState, TCalcStack CalcStack)
		{
			if (Sheet1 == Sheet2)
			{
				if (Row1 == Row2 && Col1 == Col2)
					return wi.Xls.GetCellValueAndRecalc(Sheet1, Row1, Col1, CalcState, CalcStack);

				OrderRange(ref Sheet1, ref Sheet2, ref Row1, ref Col1, ref Row2, ref Col2);
				object[,] Result = new object[Row2 - Row1 + 1, Col2 - Col1 + 1];

				int MaxRow = wi.Xls.GetRowCount(Sheet1);
				for (int r = Row1; r <= Row2; r++)
				{
					if (r > MaxRow) break;

					for (int cIndex = wi.Xls.ColToIndex(Sheet1, r, Col2); cIndex > 0; cIndex--)
					{
						int c = wi.Xls.ColFromIndex(Sheet1, r, cIndex);
						if (c > Col2 || c == 0) continue;
						if (c < Col1) break;

						Result[r - Row1, c - Col1] =
							ConvertToAllowedObject(wi.Xls.GetCellValueAndRecalc(Sheet1, r, c, CalcState, CalcStack), wi.Xls.OptionsDates1904);
					}
				}
				return Result;
			}
			return TFlxFormulaErrorValue.ErrValue;
		}

		private static object AggOneDimToArray(object[,] v2, object val1, int ofs)
		{
			if (v2.GetLength(1) == 1)
			{
				object[,] Result = new object[v2.GetLength(1) + 1, 1];
				for (int i = 0; i < Result.GetLength(0) - 1; i++)
				{
					Result[i + ofs, 0] = v2[i, 0];
				}
				if (ofs == 0)
					Result[Result.GetLength(0) - 1, 0] = val1;
				else
					Result[0, 0] = val1;

				return Result;
			}
			if (v2.GetLength(0) == 1)
			{
				object[,] Result = new object[1, v2.GetLength(1) + 1];
				for (int i = 0; i < Result.GetLength(1) - 1; i++)
				{
					Result[0, i + ofs] = v2[0, i];
				}

				if (ofs == 0)
					Result[0, Result.GetLength(1) - 1] = val1;
				else
					Result[0, 0] = val1;

				return Result;
			}
			return TFlxFormulaErrorValue.ErrValue;
		}

		private static void Add2Arrays(ref object[,] ResultArray, object[,] v1, object[,] v2, int ofs1, int ofs2)
		{
			int cofs1 = 1 - ofs1;
			int cofs2 = 1 - ofs2;

			for (int i = v2.GetLength(ofs2) - 1; i >= 0; i--)
			{
				int v21 = v2.GetLength(cofs2);
				for (int j = v21 - 1; j >= 0; j--)
				{
					int ri = (ofs2 == 0) ? i : j;
					int rj = (ofs2 == 0) ? j : i;

					object item = ofs2 == 0 ? v2[i, j] : v2[j, i];
					ResultArray[ri, rj] = item;
				}

				for (int j = v1.GetLength(cofs1) - 1; j >= 0; j--)
				{
					int ri = (ofs2 == 0) ? i : j + v21;
					int rj = (ofs2 == 0) ? j + v21 : i;

					object item = ofs1 == 0 ? v1[i, j] : v1[j, i];
					ResultArray[ri, rj] = item;
				}
			}

		}

		internal override object AggValues(object val1, object val2)
		{
			object[,] v1 = val1 as object[,];
			object[,] v2 = val2 as object[,];
			if (v1 == null)
			{
				if (v2 == null)
				{
					object[,] Result = new object[2, 1];
					Result[0, 0] = val2;
					Result[1, 0] = val1;
					return Result;
				}
				else return AggOneDimToArray(v2, val1, 0);
			}
			else
			{
				if (v2 == null) return AggOneDimToArray(v1, val2, 1);
			}

			//here v2 && v1 != null

			if (v2.GetLength(0) == v1.GetLength(0))
			{
				object[,] Result = new object[v1.GetLength(0), v1.GetLength(1) + v2.GetLength(1)];
				Add2Arrays(ref Result, v1, v2, 0, 0);
				return Result;

			}

			if (v2.GetLength(1) == v1.GetLength(1))
			{
				object[,] Result = new object[v1.GetLength(0) + v2.GetLength(0), v1.GetLength(1)];
				Add2Arrays(ref Result, v1, v2, 1, 1);
				return Result;
			}

			if (v2.GetLength(1) == v1.GetLength(0))
			{
				object[,] Result = new object[v1.GetLength(1) + v2.GetLength(0), v2.GetLength(1)];
				Add2Arrays(ref Result, v1, v2, 0, 1);
				return Result;
			}

			if (v2.GetLength(0) == v1.GetLength(1))
			{
				object[,] Result = new object[v2.GetLength(0), v1.GetLength(0) + v2.GetLength(1)];
				Add2Arrays(ref Result, v1, v2, 1, 0);
				return Result;
			}


			return TFlxFormulaErrorValue.ErrValue;
		}

		internal override object AggArray(object[,] arr)
		{
			return arr;
		}
	}

	internal class TUdfAggregate : TBaseAggregate
	{
		internal static readonly TUdfAggregate Instance = new TUdfAggregate();

		internal override object Agg(TWorkbookInfo wi, int Sheet1, int Sheet2, int Row1, int Col1, int Row2, int Col2, TCalcState CalcState, TCalcStack CalcStack)
		{
			/* We can't do this, since there would be no way to return a reference to a cell.
			if (Sheet1 == Sheet2 && Row1 == Row2 && Col1 == Col2)
			{
				return wi.Xls.GetCellValueAndRecalc(Sheet1, Row1, Col1, CalcState, CalcStack);
			}
			*/

			return new TXls3DRange(Sheet1, Sheet2, Row1, Col1, Row2, Col2);
		}

		internal override object AggValues(object val1, object val2)
		{
			return TFlxFormulaErrorValue.ErrNA;
		}


		internal override object AggArray(object[,] arr)
		{
			return arr;
		}


	}

	internal class TAverageAggregate : TBaseAggregate
	{
		private bool CountAnything;
		internal static readonly TAverageAggregate Instance0 = new TAverageAggregate(false);
		internal static readonly TAverageAggregate InstanceA = new TAverageAggregate(true);

		private TAverageAggregate(bool aCountAnything)
		{
			CountAnything = aCountAnything;
		}

		internal override object Agg(TWorkbookInfo wi, int Sheet1, int Sheet2, int Row1, int Col1, int Row2, int Col2, TCalcState CalcState, TCalcStack CalcStack)
		{
			OrderRange(ref Sheet1, ref Sheet2, ref Row1, ref Col1, ref Row2, ref Col2);
			double ResultValue = 0;
			long ValueCount = 0;
			for (int s = Sheet1; s <= Sheet2; s++)
			{
				int MaxRow = wi.Xls.GetRowCount(s);
				for (int r = Row1; r <= Row2; r++)
				{
					if (r > MaxRow) break;
					if (CalcState.IgnoreHidden && wi.Xls.GetRowHidden(s, r)) continue;

					for (int cIndex = wi.Xls.ColToIndex(s, r, Col2); cIndex > 0; cIndex--)
					{
						int c = wi.Xls.ColFromIndex(s, r, cIndex);
						if (c > Col2 || c == 0) continue;
						if (c < Col1) break;

						object v1 = ConvertToAllowedObject(wi.Xls.GetCellValueAndRecalc(s, r, c, CalcState, CalcStack), wi.Xls.OptionsDates1904);
                        if (v1 is TFlxFormulaErrorValue)
                        {
                            if (CalcState.IgnoreErrors) continue;
                            return v1;
                        }
                        if (v1 is double)
						{
							ResultValue += (double)v1;
							ValueCount++;
						}
						else
							if (CountAnything && v1 != null)
						{
							ValueCount++;
						}

					}
				}
			}
			return new TAverageValue(ResultValue, ValueCount);
		}

		internal override object AggValues(object val1, object val2)
		{
			if (val1 is TFlxFormulaErrorValue) return val1;
			if (val2 is TFlxFormulaErrorValue) return val2;
			TAverageValue a1 = (val1 as TAverageValue);
			TAverageValue a2 = (val2 as TAverageValue);
			if (a1 == null)
			{
				if (val1 == null) return a2;
				return TFlxFormulaErrorValue.ErrValue;
			}
			if (a2 == null)
			{
				if (val2 == null) return a1;
				return TFlxFormulaErrorValue.ErrValue;
			}

			return a1.Add(a2);
		}

		internal override object AggArray(object[,] arr)
		{
			TAverageValue Result = new TAverageValue(0, 0);
			for (int i = 0; i < arr.GetLength(0); i++)
				for (int k = 0; k < arr.GetLength(1); k++)
				{
					if (arr[i, k] is double)
					{
						Result.Sum += (double)arr[i, k];
						Result.ValueCount++;
					}
				}
			return Result;
		}

		internal override void AggAverages(TAverageValue val1, double val2)
		{
			val1.Sum += val2;
			val1.ValueCount++;
		}

	}

	internal class TSumAggregate : TBaseAggregate
	{
		internal static readonly TSumAggregate Instance = new TSumAggregate();

		internal override object Agg(TWorkbookInfo wi, int Sheet1, int Sheet2, int Row1, int Col1, int Row2, int Col2, TCalcState CalcState, TCalcStack CalcStack)
		{
			OrderRange(ref Sheet1, ref Sheet2, ref Row1, ref Col1, ref Row2, ref Col2);
			double Result = 0;
			for (int s = Sheet1; s <= Sheet2; s++)
			{
				int MaxRow = wi.Xls.GetRowCount(s);
				for (int r = Row1; r <= Row2; r++)
				{
					if (r > MaxRow) break;
					if (CalcState.IgnoreHidden && wi.Xls.GetRowHidden(s, r)) continue;

					for (int cIndex = wi.Xls.ColToIndex(s, r, Col2); cIndex > 0; cIndex--)
					{
						int c = wi.Xls.ColFromIndex(s, r, cIndex);
						if (c > Col2 || c == 0) continue;
						if (c < Col1) break;

						object v1 = ConvertToAllowedObject(wi.Xls.GetCellValueAndRecalc(s, r, c, CalcState, CalcStack), wi.Xls.OptionsDates1904);
                        if (v1 is TFlxFormulaErrorValue)
                        {
                            if (CalcState.IgnoreErrors) continue;
                            return v1;
                        }
                        if (v1 is double)
							Result += (double)v1;
					}
				}
			}
			return Result;
		}

		internal override object AggValues(object val1, object val2)
		{
			if (val1 is TFlxFormulaErrorValue) return val1;
			if (val2 is TFlxFormulaErrorValue) return val2;

			double r1 = 0;
			if (val1 != null)
			{
				if (!TBaseParsedToken.GetDouble(val1, out r1)) return TFlxFormulaErrorValue.ErrValue;
			}
			double r2 = 0;
			if (val2 != null)
			{
				if (!TBaseParsedToken.GetDouble(val2, out r2)) return TFlxFormulaErrorValue.ErrValue;
			}

			return r1 + r2;
		}

		internal override object AggArray(object[,] arr)
		{
			double Result = 0;
			for (int i = 0; i < arr.GetLength(0); i++)
				for (int k = 0; k < arr.GetLength(1); k++)
				{
					if (arr[i, k] is TFlxFormulaErrorValue) return arr[i, k];
					if (arr[i, k] is double)
						Result += (double)arr[i, k];
				}
			return Result;
		}

	}

	internal class TGeoMeanAggregate : TBaseAggregate
	{
		internal static readonly TGeoMeanAggregate Instance = new TGeoMeanAggregate();

		private TGeoMeanAggregate()
		{
		}

		internal override object Agg(TWorkbookInfo wi, int Sheet1, int Sheet2, int Row1, int Col1, int Row2, int Col2, TCalcState CalcState, TCalcStack CalcStack)
		{
			OrderRange(ref Sheet1, ref Sheet2, ref Row1, ref Col1, ref Row2, ref Col2);
			double ResultValue = 0;
			long ValueCount = 0;
			for (int s = Sheet1; s <= Sheet2; s++)
			{
				int MaxRow = wi.Xls.GetRowCount(s);
				for (int r = Row1; r <= Row2; r++)
				{
					if (r > MaxRow) break;
					if (CalcState.IgnoreHidden && wi.Xls.GetRowHidden(s, r)) continue;

					for (int cIndex = wi.Xls.ColToIndex(s, r, Col2); cIndex > 0; cIndex--)
					{
						int c = wi.Xls.ColFromIndex(s, r, cIndex);
						if (c > Col2 || c == 0) continue;
						if (c < Col1) break;

						object v1 = ConvertToAllowedObject(wi.Xls.GetCellValueAndRecalc(s, r, c, CalcState, CalcStack), wi.Xls.OptionsDates1904);
                        if (v1 is TFlxFormulaErrorValue)
                        {
                            if (CalcState.IgnoreErrors) continue;
                            return v1;
                        }
                        if (v1 is double)
						{
							if ((double)v1 <= 0) return TFlxFormulaErrorValue.ErrNum;
							if (ValueCount > 0)
							{
								ResultValue *= (double)v1;
							}
							else
							{
								ResultValue = (double)v1;
							}
							if (Double.IsNaN(ResultValue) || Double.IsInfinity(ResultValue)) return TFlxFormulaErrorValue.ErrNum;
							ValueCount++;
						}
					}
				}
			}
			return new TAverageValue(ResultValue, ValueCount);
		}

		internal override object AggValues(object val1, object val2)
		{
			if (val1 is TFlxFormulaErrorValue) return val1;
			if (val2 is TFlxFormulaErrorValue) return val2;
			TAverageValue a1 = (val1 as TAverageValue);
			TAverageValue a2 = (val2 as TAverageValue);
			if (a1 == null)
			{
				if (val1 == null) return a2;
				return TFlxFormulaErrorValue.ErrValue;
			}
			if (a2 == null)
			{
				if (val2 == null) return a1;
				return TFlxFormulaErrorValue.ErrValue;
			}

			return a1.Mult(a2);
		}

		internal override object AggArray(object[,] arr)
		{
			TAverageValue Result = new TAverageValue(0, 0);
			for (int i = 0; i < arr.GetLength(0); i++)
				for (int k = 0; k < arr.GetLength(1); k++)
				{
					if (arr[i, k] is double)
					{
						if ((double)arr[i, k] <= 0) return TFlxFormulaErrorValue.ErrNum;

						if (Result.ValueCount == 0)
							Result.Sum = (double)arr[i, k];
						else
							Result.Sum *= (double)arr[i, k];

						Result.ValueCount++;
					}
				}
			return Result;
		}

		internal override void AggAverages(TAverageValue val1, double val2)
		{
			if (val1.ValueCount == 0) val1.Sum = val2; else val1.Sum *= val2;
			val1.ValueCount++;
		}

	}

	internal class THarMeanAggregate : TBaseAggregate
	{
		internal static readonly THarMeanAggregate Instance = new THarMeanAggregate();

		private THarMeanAggregate()
		{
		}

		internal override object Agg(TWorkbookInfo wi, int Sheet1, int Sheet2, int Row1, int Col1, int Row2, int Col2, TCalcState CalcState, TCalcStack CalcStack)
		{
			OrderRange(ref Sheet1, ref Sheet2, ref Row1, ref Col1, ref Row2, ref Col2);
			double ResultValue = 0;
			long ValueCount = 0;
			for (int s = Sheet1; s <= Sheet2; s++)
			{
				int MaxRow = wi.Xls.GetRowCount(s);
				for (int r = Row1; r <= Row2; r++)
				{
					if (r > MaxRow) break;
					if (CalcState.IgnoreHidden && wi.Xls.GetRowHidden(s, r)) continue;

					for (int cIndex = wi.Xls.ColToIndex(s, r, Col2); cIndex > 0; cIndex--)
					{
						int c = wi.Xls.ColFromIndex(s, r, cIndex);
						if (c > Col2 || c == 0) continue;
						if (c < Col1) break;
						
						object v1 = ConvertToAllowedObject(wi.Xls.GetCellValueAndRecalc(s, r, c, CalcState, CalcStack), wi.Xls.OptionsDates1904);
                        if (v1 is TFlxFormulaErrorValue)
                        {
                            if (CalcState.IgnoreErrors) continue;
                            return v1;
                        }
                        if (v1 is double)
						{
							if ((double)v1 <= 0) return TFlxFormulaErrorValue.ErrNum;
							ResultValue += 1 / (double)v1;
							if (Double.IsNaN(ResultValue) || Double.IsInfinity(ResultValue)) return TFlxFormulaErrorValue.ErrNum;
							ValueCount++;
						}
					}
				}
			}
			return new TAverageValue(ResultValue, ValueCount);
		}

		internal override object AggValues(object val1, object val2)
		{
			if (val1 is TFlxFormulaErrorValue) return val1;
			if (val2 is TFlxFormulaErrorValue) return val2;
			TAverageValue a1 = (val1 as TAverageValue);
			TAverageValue a2 = (val2 as TAverageValue);
			if (a1 == null)
			{
				if (val1 == null) return a2;
				return TFlxFormulaErrorValue.ErrValue;
			}
			if (a2 == null)
			{
				if (val2 == null) return a1;
				return TFlxFormulaErrorValue.ErrValue;
			}

			return a1.Add(a2);
		}

		internal override object AggArray(object[,] arr)
		{
			TAverageValue Result = new TAverageValue(0, 0);
			for (int i = 0; i < arr.GetLength(0); i++)
				for (int k = 0; k < arr.GetLength(1); k++)
				{
					if (arr[i, k] is double)
					{
						if ((double)arr[i, k] <= 0) return TFlxFormulaErrorValue.ErrNum;

						Result.Sum += 1 / (double)arr[i, k];
						Result.ValueCount++;
					}
				}
			return Result;
		}

		internal override void AggAverages(TAverageValue val1, double val2)
		{
			val1.Sum += 1 / val2;
			val1.ValueCount++;
		}
	}

	internal class TSumSqAggregate : TBaseAggregate
	{
		internal static readonly TSumSqAggregate Instance = new TSumSqAggregate();

		internal override object Agg(TWorkbookInfo wi, int Sheet1, int Sheet2, int Row1, int Col1, int Row2, int Col2, TCalcState CalcState, TCalcStack CalcStack)
		{
			OrderRange(ref Sheet1, ref Sheet2, ref Row1, ref Col1, ref Row2, ref Col2);
			double ResultValue = 0;
			for (int s = Sheet1; s <= Sheet2; s++)
			{
				int MaxRow = wi.Xls.GetRowCount(s);
				for (int r = Row1; r <= Row2; r++)
				{
					if (r > MaxRow) break;
					if (CalcState.IgnoreHidden && wi.Xls.GetRowHidden(s, r)) continue;

					for (int cIndex = wi.Xls.ColToIndex(s, r, Col2); cIndex > 0; cIndex--)
					{
						int c = wi.Xls.ColFromIndex(s, r, cIndex);
						if (c > Col2 || c == 0) continue;
						if (c < Col1) break;
						
						object v1 = ConvertToAllowedObject(wi.Xls.GetCellValueAndRecalc(s, r, c, CalcState, CalcStack), wi.Xls.OptionsDates1904);
                        if (v1 is TFlxFormulaErrorValue)
                        {
                            if (CalcState.IgnoreErrors) continue;
                            return v1;
                        }
                        if (v1 is double)
							ResultValue += (double)v1 * (double)v1;
					}
				}
			}
			return Math.Sqrt(ResultValue);  //to avoid aggregating twice.
		}

		internal override object AggValues(object val1, object val2)
		{
			if (val1 is TFlxFormulaErrorValue) return val1;
			if (val2 is TFlxFormulaErrorValue) return val2;

			double r1 = 0;
			if (val1 != null)
			{
				if (!TBaseParsedToken.GetDouble(val1, out r1)) return TFlxFormulaErrorValue.ErrValue;
			}
			double r2 = 0;
			if (val2 != null)
			{
				if (!TBaseParsedToken.GetDouble(val2, out r2)) return TFlxFormulaErrorValue.ErrValue;
			}

			return Math.Sqrt(r1 * r1 + r2 * r2);
		}

		internal override object AggArray(object[,] arr)
		{
			double ResultValue = 0;
			for (int i = 0; i < arr.GetLength(0); i++)
				for (int k = 0; k < arr.GetLength(1); k++)
				{
					if (arr[i, k] is TFlxFormulaErrorValue) return arr[i, k];
					if (arr[i, k] is double)
						ResultValue += (double)arr[i, k] * (double)arr[i, k];
				}
			return Math.Sqrt(ResultValue);
		}

	}

	internal class TProductAggregate : TBaseAggregate
	{
		internal static readonly TProductAggregate Instance = new TProductAggregate();

		internal override object Agg(TWorkbookInfo wi, int Sheet1, int Sheet2, int Row1, int Col1, int Row2, int Col2, TCalcState CalcState, TCalcStack CalcStack)
		{
			OrderRange(ref Sheet1, ref Sheet2, ref Row1, ref Col1, ref Row2, ref Col2);
			object Result = null;
			for (int s = Sheet1; s <= Sheet2; s++)
			{
				int MaxRow = wi.Xls.GetRowCount(s);
				for (int r = Row1; r <= Row2; r++)
				{
					if (r > MaxRow) break;
					if (CalcState.IgnoreHidden && wi.Xls.GetRowHidden(s, r)) continue;

					for (int cIndex = wi.Xls.ColToIndex(s, r, Col2); cIndex > 0; cIndex--)
					{
						int c = wi.Xls.ColFromIndex(s, r, cIndex);
						if (c > Col2 || c == 0) continue;
						if (c < Col1) break;
						
						object v1 = ConvertToAllowedObject(wi.Xls.GetCellValueAndRecalc(s, r, c, CalcState, CalcStack), wi.Xls.OptionsDates1904);
                        if (v1 is TFlxFormulaErrorValue)
                        {
                            if (CalcState.IgnoreErrors) continue;
                            return v1;
                        }

						if (v1 is double)
						{
							if (Result != null) Result = (double)Result * (double)v1; else Result = (double)v1;
						}
					}
				}
			}
			return Result;
		}

		internal override object AggValues(object val1, object val2)
		{
			if (val1 is TFlxFormulaErrorValue) return val1;
			if (val2 is TFlxFormulaErrorValue) return val2;

			double r1 = 1;
			if (val1 != null)
			{
				if (!TBaseParsedToken.GetDouble(val1, out r1)) return TFlxFormulaErrorValue.ErrValue;
			}
			double r2 = 1;
			if (val2 != null)
			{
				if (!TBaseParsedToken.GetDouble(val2, out r2)) return TFlxFormulaErrorValue.ErrValue;
			}

			return r1 * r2;
		}

		internal override object AggArray(object[,] arr)
		{
			object Result = null;
			for (int i = 0; i < arr.GetLength(0); i++)
				for (int k = 0; k < arr.GetLength(1); k++)
				{
					if (arr[i, k] is TFlxFormulaErrorValue) return arr[i, k];
					if (arr[i, k] is double)
					{
						if (Result != null) Result = (double)Result * (double)arr[i, k]; else Result = (double)arr[i, k];
					}
				}
			return Result;
		}

	}

	internal class TCountAggregate : TBaseAggregate
	{
		internal static readonly TCountAggregate Instance = new TCountAggregate();

		internal override object Agg(TWorkbookInfo wi, int Sheet1, int Sheet2, int Row1, int Col1, int Row2, int Col2, TCalcState CalcState, TCalcStack CalcStack)
		{
			OrderRange(ref Sheet1, ref Sheet2, ref Row1, ref Col1, ref Row2, ref Col2);
			long ResultValue = 0;
			for (int s = Sheet1; s <= Sheet2; s++)
			{
				int MaxRow = wi.Xls.GetRowCount(s);
				for (int r = Row1; r <= Row2; r++)
				{
					if (r > MaxRow) break;
					if (CalcState.IgnoreHidden && wi.Xls.GetRowHidden(s, r)) continue;

					for (int cIndex = wi.Xls.ColToIndex(s, r, Col2); cIndex > 0; cIndex--)
					{
						int c = wi.Xls.ColFromIndex(s, r, cIndex);
						if (c > Col2 || c == 0) continue;
						if (c < Col1) break;
						
						object v1 = ConvertToAllowedObject(wi.Xls.GetCellValueAndRecalc(s, r, c, CalcState, CalcStack), wi.Xls.OptionsDates1904);
                        if (v1 is TFlxFormulaErrorValue)
                        {
                            if (CalcState.IgnoreErrors) continue;
                            //Count doesn't care about errors. if (v1 is TFlxFormulaErrorValue) return v1;
                        }
                        if (CalcState.Aborted) return TFlxFormulaErrorValue.ErrNA;
						if (v1 is double)
							ResultValue++;
					}
				}
			}
			return new TAverageValue(0, ResultValue);
		}

		internal override object AggValues(object val1, object val2)
		{
			if (val1 is TFlxFormulaErrorValue) return val1;
			if (val2 is TFlxFormulaErrorValue) return val2;
			TAverageValue a1 = (val1 as TAverageValue);
			TAverageValue a2 = (val2 as TAverageValue);
			if (a1 == null)
			{
				if (val1 == null) return a2;
				return TFlxFormulaErrorValue.ErrValue;
			}
			if (a2 == null)
			{
				if (val2 == null) return a1;
				return TFlxFormulaErrorValue.ErrValue;
			}

			return a1.Add(a2);
		}

		internal override object AggArray(object[,] arr)
		{
			TAverageValue Result = new TAverageValue(0, 0);
			for (int i = 0; i < arr.GetLength(0); i++)
				for (int k = 0; k < arr.GetLength(1); k++)
				{
					if (arr[i, k] is double)
					{
						Result.ValueCount++;
					}
				}
			return Result;
		}

	}

	internal class TCountAAggregate : TBaseAggregate
	{
		internal static readonly TCountAAggregate InstanceAll = new TCountAAggregate(false);
		internal static readonly TCountAAggregate InstanceBlank = new TCountAAggregate(true);

		bool CountBlanks;

		internal TCountAAggregate(bool aCountBlanks)
		{
			CountBlanks = aCountBlanks;
		}
		internal override object Agg(TWorkbookInfo wi, int Sheet1, int Sheet2, int Row1, int Col1, int Row2, int Col2, TCalcState CalcState, TCalcStack CalcStack)
		{
			OrderRange(ref Sheet1, ref Sheet2, ref Row1, ref Col1, ref Row2, ref Col2);
			long ResultValue = 0;
			long ValuesProcessed = 0;
			for (int s = Sheet1; s <= Sheet2; s++)
			{
				int MaxRow = wi.Xls.GetRowCount(s);
				for (int r = Row1; r <= Row2; r++)
				{
					if (r > MaxRow) break;
					if (CalcState.IgnoreHidden && wi.Xls.GetRowHidden(s, r)) continue;

					for (int cIndex = wi.Xls.ColToIndex(s, r, Col2); cIndex > 0; cIndex--)
					{
						int c = wi.Xls.ColFromIndex(s, r, cIndex);
						if (c > Col2 || c == 0) continue;
						if (c < Col1) break;
						ValuesProcessed++;
						
						object v1 = ConvertToAllowedObject(wi.Xls.GetCellValueAndRecalc(s, r, c, CalcState, CalcStack), wi.Xls.OptionsDates1904);
                        if (v1 is TFlxFormulaErrorValue)
                        {
                            if (CalcState.IgnoreErrors) continue;
                            //Count doesn't care about errors. if (v1 is TFlxFormulaErrorValue) return v1;
                        }

                        if (CalcState.Aborted) return TFlxFormulaErrorValue.ErrNA;
                        
                        if ((CountBlanks ^ (v1 != null)) || CountBlanks && v1 is String && ((string)v1).Length == 0)
							ResultValue++;
					}
				}
			}
			if (CountBlanks) ResultValue += ((long)Sheet2 - Sheet1 + 1) * (Row2 - Row1 + 1) * (Col2 - Col1 + 1) - ValuesProcessed;
			return new TAverageValue(0, ResultValue);
		}

		internal override object AggValues(object val1, object val2)
		{
			if (val1 is TFlxFormulaErrorValue) return val1;
			if (val2 is TFlxFormulaErrorValue) return val2;
			TAverageValue a1 = (val1 as TAverageValue);
			TAverageValue a2 = (val2 as TAverageValue);
			if (a1 == null)
			{
				if (val1 == null) return a2;
				return TFlxFormulaErrorValue.ErrValue;
			}
			if (a2 == null)
			{
				if (val2 == null) return a1;
				return TFlxFormulaErrorValue.ErrValue;
			}

			return a1.Add(a2);
		}

		internal override object AggArray(object[,] arr)
		{
			TAverageValue Result = new TAverageValue(0, 0);
			for (int i = 0; i < arr.GetLength(0); i++)
				for (int k = 0; k < arr.GetLength(1); k++)
				{
					if ((CountBlanks ^ (arr[i, k] != null)) || CountBlanks && arr[i, k] is String && ((string)arr[i, k]).Length == 0)
					{
						Result.ValueCount++;
					}
				}
			return Result;
		}

	}

	internal class TMinAggregate : TBaseAggregate
	{
		private bool CountAnything;
		internal static readonly TMinAggregate Instance0 = new TMinAggregate(false);
		internal static readonly TMinAggregate InstanceA = new TMinAggregate(true);

		private TMinAggregate(bool aCountAnything)
		{
			CountAnything = aCountAnything;
		}

		internal override object Agg(TWorkbookInfo wi, int Sheet1, int Sheet2, int Row1, int Col1, int Row2, int Col2, TCalcState CalcState, TCalcStack CalcStack)
		{
			OrderRange(ref Sheet1, ref Sheet2, ref Row1, ref Col1, ref Row2, ref Col2);
			object Result = null;
		
			for (int s = Sheet1; s <= Sheet2; s++)
			{
				int MaxRow = wi.Xls.GetRowCount(s);
				for (int r = Row1; r <= Row2; r++)
				{
					if (r > MaxRow) break;
					if (CalcState.IgnoreHidden && wi.Xls.GetRowHidden(s, r)) continue;

					for (int cIndex = wi.Xls.ColToIndex(s, r, Col2); cIndex > 0; cIndex--)
					{
						int c = wi.Xls.ColFromIndex(s, r, cIndex);
						if (c > Col2 || c == 0) continue;
						if (c < Col1) break;

						object v1 = ConvertToAllowedObject(wi.Xls.GetCellValueAndRecalc(s, r, c, CalcState, CalcStack), wi.Xls.OptionsDates1904);
                        if (v1 is TFlxFormulaErrorValue)
                        {
                            if (CalcState.IgnoreErrors) continue;
                            return v1;
                        }

						if (v1 is bool && CountAnything)
							if ((bool)v1) v1 = 1.0; else v1 = 0.0;
						if (v1 is string && CountAnything)
							v1 = 0.0;

						if (v1 is double)
							if (Result == null || (double)v1 < (double)Result)
							{
								Result = (double)v1;
							}
					}
				}
			}

			return Result;
		}

		internal override object AggValues(object val1, object val2)
		{
			if (val1 is TFlxFormulaErrorValue) return val1;
			if (val2 is TFlxFormulaErrorValue) return val2;

			double r1 = 0;
			if (val1 == null) return val2;
			if (!TBaseParsedToken.GetDouble(val1, out r1)) return TFlxFormulaErrorValue.ErrValue;
			double r2 = 0;
			if (val2 == null) return val1;
			if (!TBaseParsedToken.GetDouble(val2, out r2)) return TFlxFormulaErrorValue.ErrValue;

			return Math.Min(r1, r2);
		}

		internal override object AggArray(object[,] arr)
		{
			object Result = null;
			for (int i = 0; i < arr.GetLength(0); i++)
				for (int k = 0; k < arr.GetLength(1); k++)
				{
					if (arr[i, k] is double)
					{
						double v = (double)arr[i, k];
						if (Result == null || (double)Result > v) Result = v;
					}
				}
			return Result;
		}
	}

	internal class TMaxAggregate : TBaseAggregate
	{
		private bool CountAnything;

		internal static readonly TMaxAggregate Instance0 = new TMaxAggregate(false);
		internal static readonly TMaxAggregate InstanceA = new TMaxAggregate(true);

		private TMaxAggregate(bool aCountAnything)
		{
			CountAnything = aCountAnything;
		}

		internal override object Agg(TWorkbookInfo wi, int Sheet1, int Sheet2, int Row1, int Col1, int Row2, int Col2, TCalcState CalcState, TCalcStack CalcStack)
		{
			OrderRange(ref Sheet1, ref Sheet2, ref Row1, ref Col1, ref Row2, ref Col2);
			object Result = null;
			for (int s = Sheet1; s <= Sheet2; s++)
			{
				int MaxRow = wi.Xls.GetRowCount(s);
				for (int r = Row1; r <= Row2; r++)
				{
					if (r > MaxRow) break;
					if (CalcState.IgnoreHidden && wi.Xls.GetRowHidden(s, r)) continue;

					for (int cIndex = wi.Xls.ColToIndex(s, r, Col2); cIndex > 0; cIndex--)
					{
						int c = wi.Xls.ColFromIndex(s, r, cIndex);
						if (c > Col2 || c == 0) continue;
						if (c < Col1) break;

						object v1 = ConvertToAllowedObject(wi.Xls.GetCellValueAndRecalc(s, r, c, CalcState, CalcStack), wi.Xls.OptionsDates1904);
                        if (v1 is TFlxFormulaErrorValue)
                        {
                            if (CalcState.IgnoreErrors) continue;
                            return v1;
                        }

						if (v1 is bool && CountAnything)
							if ((bool)v1) v1 = 1.0; else v1 = 0.0;
						if (v1 is string && CountAnything)
							v1 = 0.0;

						if (v1 is double)
							if (Result == null || (double)v1 > (double)Result)
							{
								Result = (double)v1;
							}
					}
				}
			}

			return Result;
		}

		internal override object AggValues(object val1, object val2)
		{
			if (val1 is TFlxFormulaErrorValue) return val1;
			if (val2 is TFlxFormulaErrorValue) return val2;

			double r1 = 0;
			if (val1 == null) return val2;
			if (!TBaseParsedToken.GetDouble(val1, out r1)) return TFlxFormulaErrorValue.ErrValue;
			double r2 = 0;
			if (val2 == null) return val1;
			if (!TBaseParsedToken.GetDouble(val2, out r2)) return TFlxFormulaErrorValue.ErrValue;

			return Math.Max(r1, r2);
		}

		internal override object AggArray(object[,] arr)
		{
			object Result = null;
			for (int i = 0; i < arr.GetLength(0); i++)
				for (int k = 0; k < arr.GetLength(1); k++)
				{
					if (arr[i, k] is double)
					{
						double v = (double)arr[i, k];
						if (Result == null || (double)Result < v) Result = v;
					}
				}
			return Result;
		}

	}

	internal class TMaxMinKAggregate : TBaseAggregate
	{
		int kesim;
		bool DoMax;
		internal bool Used;

		internal TMaxMinKAggregate(int k, bool aDoMax)
		{
			kesim = k;
			DoMax = aDoMax;
			Used = false;
		}

		internal override object Agg(TWorkbookInfo wi, int Sheet1, int Sheet2, int Row1, int Col1, int Row2, int Col2, TCalcState CalcState, TCalcStack CalcStack)
		{
			Used = true;
			if (kesim <= 0) return TFlxFormulaErrorValue.ErrNum;
			double[] Values = new double[kesim];
			int ValuesUsed = 0;

			OrderRange(ref Sheet1, ref Sheet2, ref Row1, ref Col1, ref Row2, ref Col2);
			for (int s = Sheet1; s <= Sheet2; s++)
			{
				int MaxRow = wi.Xls.GetRowCount(s);
				for (int r = Row1; r <= Row2; r++)
				{
					if (r > MaxRow) break;
					if (CalcState.IgnoreHidden && wi.Xls.GetRowHidden(s, r)) continue;

					for (int cIndex = wi.Xls.ColToIndex(s, r, Col2); cIndex > 0; cIndex--)
					{
						int c = wi.Xls.ColFromIndex(s, r, cIndex);
						if (c > Col2 || c == 0) continue;
						if (c < Col1) break;
						
						object v1 = ConvertToAllowedObject(wi.Xls.GetCellValueAndRecalc(s, r, c, CalcState, CalcStack), wi.Xls.OptionsDates1904);
                        if (v1 is TFlxFormulaErrorValue)
                        {
                            if (CalcState.IgnoreErrors) continue;
                            return v1;
                        }

						if (v1 is double) AddDouble(Values, ref ValuesUsed, (double)v1);
					}
				}
			}

			if (ValuesUsed != Values.Length) return TFlxFormulaErrorValue.ErrNum;
			return (Values[Values.Length - 1]);
		}

		private void AddDouble(double[] Values, ref int ValuesUsed, double v)
		{
			if (ValuesUsed < Values.Length)
			{
				if ((ValuesUsed == 0 || IsLess(v, Values[ValuesUsed - 1])))
				{
					Values[ValuesUsed] = v;
				}
				else
				{
					InsertIntoValues(Values, ValuesUsed, v);
				}
				ValuesUsed++;
			}
			else
			{
				InsertIntoValues(Values, ValuesUsed, v);
			}

		}

		private void InsertIntoValues(double[] Values, int ValuesUsed, double v)
		{
			int k = ValuesUsed - 1;
			while (k >= 0)
			{
				if (IsLess(Values[k], v))
				{
					if (k + 1 < Values.Length) Values[k + 1] = Values[k];
				}
				else
					break;

				k--;
			}

			if (k + 1 < Values.Length) Values[k + 1] = v;
		}

		internal bool IsLess(double v1, double v2)
		{
			if (DoMax) return v1 < v2;
			return v1 > v2;
		}

		internal override object AggValues(object val1, object val2)
		{
			Used = true;
			return TFlxFormulaErrorValue.ErrNA;
		}

		internal override object AggArray(object[,] arr)
		{
			Used = true;
			if (kesim <= 0) return TFlxFormulaErrorValue.ErrNum;
			double[] Values = new double[kesim];
			int ValuesUsed = 0;

			for (int i = 0; i < arr.GetLength(0); i++)
				for (int k = 0; k < arr.GetLength(1); k++)
				{
					if (arr[i, k] is double)
					{
						double v = (double)arr[i, k];
						AddDouble(Values, ref ValuesUsed, v);
					}
				}
			if (ValuesUsed != Values.Length) return TFlxFormulaErrorValue.ErrNum;
			return (Values[Values.Length - 1]);
		}

	}

	internal class TSquaredDiffAggregate : TBaseAggregate
	{
		private double FAvg;
		private bool CountAnything;
		internal TSquaredDiffAggregate(double aAvg, bool aCountAnything)
		{
			FAvg = aAvg;
			CountAnything = aCountAnything;
		}

		internal override object Agg(TWorkbookInfo wi, int Sheet1, int Sheet2, int Row1, int Col1, int Row2, int Col2, TCalcState CalcState, TCalcStack CalcStack)
		{
			OrderRange(ref Sheet1, ref Sheet2, ref Row1, ref Col1, ref Row2, ref Col2);
			double ResultValue = 0;
			long ValueCount = 0;
			for (int s = Sheet1; s <= Sheet2; s++)
			{
				int MaxRow = wi.Xls.GetRowCount(s);
				for (int r = Row1; r <= Row2; r++)
				{
					if (r > MaxRow) break;
					if (CalcState.IgnoreHidden && wi.Xls.GetRowHidden(s, r)) continue;

					for (int cIndex = wi.Xls.ColToIndex(s, r, Col2); cIndex > 0; cIndex--)
					{
						int c = wi.Xls.ColFromIndex(s, r, cIndex);
						if (c > Col2 || c == 0) continue;
						if (c < Col1) break;

						object v1 = ConvertToAllowedObject(wi.Xls.GetCellValueAndRecalc(s, r, c, CalcState, CalcStack), wi.Xls.OptionsDates1904);
                        if (v1 is TFlxFormulaErrorValue)
                        {
                            if (CalcState.IgnoreErrors) continue;
                            return v1;
                        }

						if (CountAnything)
						{
							if (v1 is bool)
								if ((bool)v1) v1 = 1.0; else v1 = 0.0;
							if (v1 is string)
								v1 = 0.0;

						}

						if (v1 is double)
						{
							double m = ((double)v1 - FAvg);
							ResultValue += m * m;
							ValueCount++;
						}
					}
				}
			}

			return new TAverageValue(ResultValue, ValueCount);
		}

		internal override object AggValues(object val1, object val2)
		{
			if (val1 is TFlxFormulaErrorValue) return val1;
			if (val2 is TFlxFormulaErrorValue) return val2;
			TAverageValue a1 = (val1 as TAverageValue);
			TAverageValue a2 = (val2 as TAverageValue);
			if (a1 == null)
			{
				if (val1 == null) return a2;
				return TFlxFormulaErrorValue.ErrValue;
			}
			if (a2 == null)
			{
				if (val2 == null) return a1;
				return TFlxFormulaErrorValue.ErrValue;
			}

			return a1.AddSquaredDiff(a2, FAvg);
		}

		internal override object AggArray(object[,] arr)
		{
			TAverageValue Result = new TAverageValue(0, 0);
			for (int i = 0; i < arr.GetLength(0); i++)
				for (int k = 0; k < arr.GetLength(1); k++)
				{
					if (arr[i, k] is double)
					{
						double x = (double)arr[i, k] - FAvg;
						Result.Sum += x * x;
						Result.ValueCount++;
					}
				}
			return Result;
		}

		internal override void AggAverages(TAverageValue val1, double val2)
		{
			val1.Sum += (val2 - FAvg) * (val2 - FAvg);
			val1.ValueCount++;
		}

	}

	internal class TNSquaredDiffAggregate : TBaseAggregate
	{
		private double FAvg;
		private bool CountAnything;
		private int N;
		internal TNSquaredDiffAggregate(double aAvg, bool aCountAnything, int aN)
		{
			FAvg = aAvg;
			CountAnything = aCountAnything;
			N = aN;
		}

		internal override object Agg(TWorkbookInfo wi, int Sheet1, int Sheet2, int Row1, int Col1, int Row2, int Col2, TCalcState CalcState, TCalcStack CalcStack)
		{
			OrderRange(ref Sheet1, ref Sheet2, ref Row1, ref Col1, ref Row2, ref Col2);
			double ResultValue = 0;
			long ValueCount = 0;
			for (int s = Sheet1; s <= Sheet2; s++)
			{
				int MaxRow = wi.Xls.GetRowCount(s);
				for (int r = Row1; r <= Row2; r++)
				{
					if (r > MaxRow) break;
					if (CalcState.IgnoreHidden && wi.Xls.GetRowHidden(s, r)) continue;

					for (int cIndex = wi.Xls.ColToIndex(s, r, Col2); cIndex > 0; cIndex--)
					{
						int c = wi.Xls.ColFromIndex(s, r, cIndex);
						if (c > Col2 || c == 0) continue;
						if (c < Col1) break;
						
						object v1 = ConvertToAllowedObject(wi.Xls.GetCellValueAndRecalc(s, r, c, CalcState, CalcStack), wi.Xls.OptionsDates1904);
                        if (v1 is TFlxFormulaErrorValue)
                        {
                            if (CalcState.IgnoreErrors) continue;
                            return v1;
                        }


						if (CountAnything)
						{
							if (v1 is bool)
								if ((bool)v1) v1 = 1.0; else v1 = 0.0;
							if (v1 is string)
								v1 = 0.0;

						}

						if (v1 is double)
						{
							double m = ((double)v1 - FAvg);
							ResultValue += Math.Pow(m, N);
							ValueCount++;
						}
					}
				}
			}

			return new TAverageValue(ResultValue, ValueCount);
		}

		internal override object AggValues(object val1, object val2)
		{
			if (val1 is TFlxFormulaErrorValue) return val1;
			if (val2 is TFlxFormulaErrorValue) return val2;
			TAverageValue a1 = (val1 as TAverageValue);
			TAverageValue a2 = (val2 as TAverageValue);
			if (a1 == null)
			{
				if (val1 == null) return a2;
				return TFlxFormulaErrorValue.ErrValue;
			}
			if (a2 == null)
			{
				if (val2 == null) return a1;
				return TFlxFormulaErrorValue.ErrValue;
			}

			return a1.AddNSquaredDiff(a2, FAvg, N);
		}

		internal override object AggArray(object[,] arr)
		{
			TAverageValue Result = new TAverageValue(0, 0);
			for (int i = 0; i < arr.GetLength(0); i++)
				for (int k = 0; k < arr.GetLength(1); k++)
				{
					if (arr[i, k] is double)
					{
						double x = (double)arr[i, k] - FAvg;
						Result.Sum += Math.Pow(x, N);
						Result.ValueCount++;
					}
				}
			return Result;
		}

		internal override void AggAverages(TAverageValue val1, double val2)
		{
			val1.Sum += Math.Pow((val2 - FAvg), N);
			val1.ValueCount++;
		}

	}

	internal class TModDiffAggregate : TBaseAggregate
	{
		private double FAvg;
		private bool CountAnything;
		internal TModDiffAggregate(double aAvg, bool aCountAnything)
		{
			FAvg = aAvg;
			CountAnything = aCountAnything;
		}

		internal override object Agg(TWorkbookInfo wi, int Sheet1, int Sheet2, int Row1, int Col1, int Row2, int Col2, TCalcState CalcState, TCalcStack CalcStack)
		{
			OrderRange(ref Sheet1, ref Sheet2, ref Row1, ref Col1, ref Row2, ref Col2);
			double ResultValue = 0;
			long ValueCount = 0;
			for (int s = Sheet1; s <= Sheet2; s++)
			{
				int MaxRow = wi.Xls.GetRowCount(s);
				for (int r = Row1; r <= Row2; r++)
				{
					if (r > MaxRow) break;
					if (CalcState.IgnoreHidden && wi.Xls.GetRowHidden(s, r)) continue;

					for (int cIndex = wi.Xls.ColToIndex(s, r, Col2); cIndex > 0; cIndex--)
					{
						int c = wi.Xls.ColFromIndex(s, r, cIndex);
						if (c > Col2 || c == 0) continue;
						if (c < Col1) break;

						object v1 = ConvertToAllowedObject(wi.Xls.GetCellValueAndRecalc(s, r, c, CalcState, CalcStack), wi.Xls.OptionsDates1904);
                        if (v1 is TFlxFormulaErrorValue)
                        {
                            if (CalcState.IgnoreErrors) continue;
                            return v1;
                        }

						if (CountAnything)
						{
							if (v1 is bool)
								if ((bool)v1) v1 = 1.0; else v1 = 0.0;
							if (v1 is string)
								v1 = 0.0;

						}

						if (v1 is double)
						{
							double m = ((double)v1 - FAvg);
							ResultValue += Math.Abs(m);
							ValueCount++;
						}
					}
				}
			}
            
			return new TAverageValue(ResultValue, ValueCount);
		}

		internal override object AggValues(object val1, object val2)
		{
			if (val1 is TFlxFormulaErrorValue) return val1;
			if (val2 is TFlxFormulaErrorValue) return val2;
			TAverageValue a1 = (val1 as TAverageValue);
			TAverageValue a2 = (val2 as TAverageValue);
			if (a1 == null)
			{
				if (val1 == null) return a2;
				return TFlxFormulaErrorValue.ErrValue;
			}
			if (a2 == null)
			{
				if (val2 == null) return a1;
				return TFlxFormulaErrorValue.ErrValue;
			}

			return a1.AddModDiff(a2, FAvg);
		}

		internal override object AggArray(object[,] arr)
		{
			TAverageValue Result = new TAverageValue(0, 0);
			for (int i = 0; i < arr.GetLength(0); i++)
				for (int k = 0; k < arr.GetLength(1); k++)
				{
					if (arr[i, k] is double)
					{
						double x = (double)arr[i, k] - FAvg;
						Result.Sum += Math.Abs(x);
						Result.ValueCount++;
					}
				}
			return Result;
		}

		internal override void AggAverages(TAverageValue val1, double val2)
		{
			val1.Sum += Math.Abs(val2 - FAvg);
			val1.ValueCount++;
		}

	}

	internal class TAndAggregate : TBaseAggregate
	{
		internal static readonly TAndAggregate Instance = new TAndAggregate();

		private TAndAggregate()
		{
		}

		internal override object Agg(TWorkbookInfo wi, int Sheet1, int Sheet2, int Row1, int Col1, int Row2, int Col2, TCalcState CalcState, TCalcStack CalcStack)
		{
			bool OneArg = false;
			OrderRange(ref Sheet1, ref Sheet2, ref Row1, ref Col1, ref Row2, ref Col2);
			if (Sheet1 == Sheet2 && Row1 == Row2 && Col1 == Col2)
			{
				return ConvertToAllowedObject(wi.Xls.GetCellValueAndRecalc(Sheet1, Row1, Col1, CalcState, CalcStack), wi.Xls.OptionsDates1904);
			}

			for (int s = Sheet1; s <= Sheet2; s++)
			{
				int MaxRow = wi.Xls.GetRowCount(s);
				for (int r = Row1; r <= Row2; r++)
				{
					if (r > MaxRow) break;
					if (CalcState.IgnoreHidden && wi.Xls.GetRowHidden(s, r)) continue;

					for (int cIndex = wi.Xls.ColToIndex(s, r, Col2); cIndex > 0; cIndex--)
					{
						int c = wi.Xls.ColFromIndex(s, r, cIndex);
						if (c > Col2 || c == 0) continue;
						if (c < Col1) break;

						object v1 = ConvertToAllowedObject(wi.Xls.GetCellValueAndRecalc(s, r, c, CalcState, CalcStack), wi.Xls.OptionsDates1904);
                        if (v1 is TFlxFormulaErrorValue)
                        {
                            if (CalcState.IgnoreErrors) continue;
                            return v1;
                        }
                        if (CalcState.Aborted) return TFlxFormulaErrorValue.ErrNA;
                        bool res;
						if (TBaseParsedToken.ExtToBool(v1, out res))
						{
							if (!res) return false;
							OneArg = true;
						}
					}
				}
			}

			if (OneArg) return true;
			return TFlxFormulaErrorValue.ErrValue;
		}

		internal override object AggValues(object val1, object val2)
		{
			if (val1 is TFlxFormulaErrorValue) return val1;
			if (val2 is TFlxFormulaErrorValue) return val2;

			bool r1 = false;
			if (val1 == null) return val2;
			if (!TBaseParsedToken.ExtToBool(val1, out r1)) return TFlxFormulaErrorValue.ErrValue;
			bool r2 = false;
			if (val2 == null) return val1;
			if (!TBaseParsedToken.ExtToBool(val2, out r2)) return TFlxFormulaErrorValue.ErrValue;

			return r1 & r2;
		}

		internal override object AggArray(object[,] arr)
		{
			bool OneArg = false;
			for (int i = 0; i < arr.GetLength(0); i++)
				for (int k = 0; k < arr.GetLength(1); k++)
				{
					if (arr[i, k] is bool)
					{
						if (!(bool)arr[i, k]) return false;
						OneArg = true;
					}
				}
			if (OneArg) return true;
			return TFlxFormulaErrorValue.ErrValue;
		}

	}

	internal class TOrAggregate : TBaseAggregate
	{
		internal static readonly TOrAggregate Instance = new TOrAggregate();

		private TOrAggregate()
		{
		}

		internal override object Agg(TWorkbookInfo wi, int Sheet1, int Sheet2, int Row1, int Col1, int Row2, int Col2, TCalcState CalcState, TCalcStack CalcStack)
		{
			bool OneArg = false;
			OrderRange(ref Sheet1, ref Sheet2, ref Row1, ref Col1, ref Row2, ref Col2);
			if (Sheet1 == Sheet2 && Row1 == Row2 && Col1 == Col2)
			{
				return ConvertToAllowedObject(wi.Xls.GetCellValueAndRecalc(Sheet1, Row1, Col1, CalcState, CalcStack), wi.Xls.OptionsDates1904);
			}

			for (int s = Sheet1; s <= Sheet2; s++)
			{
				int MaxRow = wi.Xls.GetRowCount(s);
				for (int r = Row1; r <= Row2; r++)
				{
					if (r > MaxRow) break;
					if (CalcState.IgnoreHidden && wi.Xls.GetRowHidden(s, r)) continue;

					for (int cIndex = wi.Xls.ColToIndex(s, r, Col2); cIndex > 0; cIndex--)
					{
						int c = wi.Xls.ColFromIndex(s, r, cIndex);
						if (c > Col2 || c == 0) continue;
						if (c < Col1) break;
						
						object v1 = ConvertToAllowedObject(wi.Xls.GetCellValueAndRecalc(s, r, c, CalcState, CalcStack), wi.Xls.OptionsDates1904);
                        if (v1 is TFlxFormulaErrorValue)
                        {
                            if (CalcState.IgnoreErrors) continue;
                            return v1;
                        }
                        if (CalcState.Aborted) return TFlxFormulaErrorValue.ErrNA;
                        bool res;
						if (TBaseParsedToken.ExtToBool(v1, out res))
						{
							if (res) return true;
							OneArg = true;
						}
					}
				}
			}

			if (OneArg) return false;
			return TFlxFormulaErrorValue.ErrValue;
		}

		internal override object AggValues(object val1, object val2)
		{
			if (val1 is TFlxFormulaErrorValue) return val1;
			if (val2 is TFlxFormulaErrorValue) return val2;

			bool r1 = false;
			if (val1 == null) return val2;
			if (!TBaseParsedToken.ExtToBool(val1, out r1)) return TFlxFormulaErrorValue.ErrValue;
			bool r2 = false;
			if (val2 == null) return val1;
			if (!TBaseParsedToken.ExtToBool(val2, out r2)) return TFlxFormulaErrorValue.ErrValue;

			return r1 | r2;
		}

		internal override object AggArray(object[,] arr)
		{
			bool OneArg = false;
			for (int i = 0; i < arr.GetLength(0); i++)
				for (int k = 0; k < arr.GetLength(1); k++)
				{
					if (arr[i, k] is bool)
					{
						if ((bool)arr[i, k]) return true;
						OneArg = true;
					}
				}
			if (OneArg) return false;
			return TFlxFormulaErrorValue.ErrValue;
		}

	}

    internal class TPercentRankValue
    {
        internal long Count;
        internal double LowerValue;
        internal bool HasLower;
        internal long LowerCount;
        internal double UpperValue;
        internal bool HasUpper;
        // We should really have "UpperCount" here, with all numbers >= LowerValue and less than UpperValue. 
        //This is actually the count of LowerValue appearances. But Excel just assumes LoweVAlue appears just once, so this is LowerCount + 1


        internal object CalcPercentRank(double NumberToCompare, double Significance)
        {
            //Algorithm says it is interpolation, but it is not. In fact, way different from interpolation, and wrong (2.0001 can have a completely different result from 2.0)
            if (!HasLower || !HasUpper) return TFlxFormulaErrorValue.ErrNA;
            if (LowerCount == 0 && Count == 1) return 1;
            if (Count < 2) return TFlxFormulaErrorValue.ErrNA;
            
            double UpperResult = (double)LowerCount / (Count - 1);

            double Result = UpperResult;
            if (HasUpper && UpperValue > LowerValue)
            {
                double LowerResult = (double)(LowerCount -1) / (Count - 1);
                Result -= (UpperValue - NumberToCompare) / (UpperValue - LowerValue) * (UpperResult - LowerResult);
            }

            double factor = Math.Pow(10, Math.Floor(Significance)); //significance is > 0, so floor is the same as truncate (not in .net 1.1)
            if (factor == 0 || double.IsInfinity(factor)) return TFlxFormulaErrorValue.ErrDiv0;
            return Math.Floor(Result * factor) / factor; //always positive.
        }

    }

    internal class TPercentRankAggregate : TBaseAggregate
    {
        double NumberToCompare;
        internal TPercentRankAggregate(double aNumberToCompare)
        {
            NumberToCompare = aNumberToCompare;
        }

        internal override object Agg(TWorkbookInfo wi, int Sheet1, int Sheet2, int Row1, int Col1, int Row2, int Col2, TCalcState CalcState, TCalcStack CalcStack)
        {
            OrderRange(ref Sheet1, ref Sheet2, ref Row1, ref Col1, ref Row2, ref Col2);
            TPercentRankValue pv = new TPercentRankValue();

            for (int s = Sheet1; s <= Sheet2; s++)
            {
                int MaxRow = wi.Xls.GetRowCount(s);
                for (int r = Row1; r <= Row2; r++)
                {
                    if (r > MaxRow) break;
                    if (CalcState.IgnoreHidden && wi.Xls.GetRowHidden(s, r)) continue;

                    for (int cIndex = wi.Xls.ColToIndex(s, r, Col2); cIndex > 0; cIndex--)
                    {
                        int c = wi.Xls.ColFromIndex(s, r, cIndex);
                        if (c > Col2 || c == 0) continue;
                        if (c < Col1) break;

                        object v1 = ConvertToAllowedObject(wi.Xls.GetCellValueAndRecalc(s, r, c, CalcState, CalcStack), wi.Xls.OptionsDates1904);
                        if (v1 is TFlxFormulaErrorValue)
                        {
                            if (CalcState.IgnoreErrors) continue;
                            return v1;
                        }

                        if (v1 is double)
                        {
                            double d = (double)v1;
                            CalcPercentRank(pv, d);
                        }
                    }
                }
            }

            return pv;
        }

        private void CalcPercentRank(TPercentRankValue pv, double d)
        {
            if (d < NumberToCompare)
            {
                if (!pv.HasLower || pv.LowerValue < d) pv.LowerValue = d;
                pv.LowerCount++;
                pv.HasLower = true;
            }
            else
                if (d == NumberToCompare)
                {
                    pv.LowerValue = d;
                    pv.UpperValue = d;
                    pv.HasLower = true;
                    pv.HasUpper = true;
                }
                else //d > numbertocompare
                {
                    if (!pv.HasUpper || pv.UpperValue > d)
                    {
                        pv.UpperValue = d;
                        pv.HasUpper = true;
                    }
                }

            pv.Count++;
        }

        internal override object AggValues(object val1, object val2)
        {
            if (val1 is TFlxFormulaErrorValue) return val1;
            if (val2 is TFlxFormulaErrorValue) return val2;

            TPercentRankValue p1 = val1 as TPercentRankValue; if (p1 == null) return TFlxFormulaErrorValue.ErrNA;
            TPercentRankValue p2 = val2 as TPercentRankValue; if (p2 == null) return TFlxFormulaErrorValue.ErrNA;

            p1.Count += p2.Count;
            p1.LowerCount += p2.LowerCount;
            p1.LowerValue = Math.Max(p1.LowerValue, p2.LowerValue);
            if (p1.HasUpper)
            {
                if (p2.HasUpper)
                {
                    p1.UpperValue = Math.Min(p1.UpperValue, p2.UpperValue);
                }
                else { }
            }
            else
            {
                p1.HasUpper = p2.HasUpper;
                p1.UpperValue = p2.UpperValue;
            }

            return p1;
        }

        internal override object AggArray(object[,] arr)
        {
            TPercentRankValue pv = new TPercentRankValue();
            for (int i = 0; i < arr.GetLength(0); i++)
                for (int k = 0; k < arr.GetLength(1); k++)
                {
                    if (arr[i, k] is double)
                    {
                        double v = (double)arr[i, k];
                        CalcPercentRank(pv, v);
                    }
                }
            return pv;
        }
    }

    internal class TMedianAggregate : TBaseAggregate
    {
        internal TMedianAggregate()
        {
        }

        internal override object Agg(TWorkbookInfo wi, int Sheet1, int Sheet2, int Row1, int Col1, int Row2, int Col2, TCalcState CalcState, TCalcStack CalcStack)
        {
            TDoubleList FList = new TDoubleList();

            OrderRange(ref Sheet1, ref Sheet2, ref Row1, ref Col1, ref Row2, ref Col2);

            for (int s = Sheet1; s <= Sheet2; s++)
            {
                int MaxRow = wi.Xls.GetRowCount(s);
                for (int r = Row1; r <= Row2; r++)
                {
                    if (r > MaxRow) break;
                    if (CalcState.IgnoreHidden && wi.Xls.GetRowHidden(s, r)) continue;

                    for (int cIndex = wi.Xls.ColToIndex(s, r, Col2); cIndex > 0; cIndex--)
                    {
                        int c = wi.Xls.ColFromIndex(s, r, cIndex);
                        if (c > Col2 || c == 0) continue;
                        if (c < Col1) break;

                        object v1 = ConvertToAllowedObject(wi.Xls.GetCellValueAndRecalc(s, r, c, CalcState, CalcStack), wi.Xls.OptionsDates1904);
                        if (v1 is TFlxFormulaErrorValue)
                        {
                            if (CalcState.IgnoreErrors) continue;
                            return v1;
                        }

                        if (v1 is double)
                        {
                            double d = (double)v1;
                            FList.Add(d);
                        }
                    }
                }
            }

            return FList;
        }

        internal override object AggValues(object val1, object val2)
        {
            if (val1 is TFlxFormulaErrorValue) return val1;
            if (val2 is TFlxFormulaErrorValue) return val2;

            TDoubleList p1 = val1 as TDoubleList; if (p1 == null) return TFlxFormulaErrorValue.ErrNA;
            TDoubleList p2 = val2 as TDoubleList; if (p2 == null) return TFlxFormulaErrorValue.ErrNA;

            p1.AddRange(p2);
            return p1;
        }

        internal override object AggArray(object[,] arr)
        {
            TDoubleList FList = new TDoubleList();
            for (int i = 0; i < arr.GetLength(0); i++)
                for (int k = 0; k < arr.GetLength(1); k++)
                {
                    if (arr[i, k] is double)
                    {
                        FList.Add((double)arr[i, k]);
                    }
                }
            return FList;
        }
    }

    internal struct TDoubleInt : IComparable, IComparable<TDoubleInt>
    {
        public int Position;
        public double Value;

        public TDoubleInt(int aPosition, double aValue)
        {
            Position = aPosition;
            Value = aValue;
        }

        #region IComparable Members

        public int CompareTo(object obj)
        {
            if (!(obj is TDoubleInt)) return -1;
            return CompareTo((TDoubleInt)obj);
        }

        #endregion

        #region IComparable<TDoubleInt> Members

        public int CompareTo(TDoubleInt other)
        {
            return Value.CompareTo(other.Value);
        }

        #endregion
    }

    internal class TFrequencyAggregate : TBaseAggregate
    {
        private List<TDoubleInt> BinArray;

        private TFrequencyAggregate(List<TDoubleInt> aBinArray)
        {
            BinArray = aBinArray;
        }

        public static TFrequencyAggregate Create(object[,] aBinArray)
        {
            List<TDoubleInt> BinArray = CalcBinArray(aBinArray);
            if (BinArray == null) return null;

            return new TFrequencyAggregate(BinArray);
        }

        private static List<TDoubleInt> CalcBinArray(object[,] aBinArray)
        {
            int l0 = aBinArray.GetLength(0);
            int l1 = aBinArray.GetLength(1);

            List<TDoubleInt> BinArray = new List<TDoubleInt>(l0 * l1);
            for (int i = 0; i < l0; i++)
            {
                for (int k = 0; k < l1; k++)
                {
                    if (aBinArray[i, k] is TFlxFormulaErrorValue) return null;
                    if (aBinArray[i, k] is double)
                    {
                        BinArray.Add(new TDoubleInt(BinArray.Count, (double)aBinArray[i, k]));
                    }
                }
            }

            if (BinArray.Count == 0) BinArray.Add(new TDoubleInt(0, 0));
            BinArray.Sort();

            return BinArray;
        }

        internal override object Agg(TWorkbookInfo wi, int Sheet1, int Sheet2, int Row1, int Col1, int Row2, int Col2, TCalcState CalcState, TCalcStack CalcStack)
        {
            double[] dResult = new double[BinArray.Count + 1];
            OrderRange(ref Sheet1, ref Sheet2, ref Row1, ref Col1, ref Row2, ref Col2);

            for (int s = Sheet1; s <= Sheet2; s++)
            {
                int MaxRow = wi.Xls.GetRowCount(s);
                for (int r = Row1; r <= Row2; r++)
                {
                    if (r > MaxRow) break;
                    if (CalcState.IgnoreHidden && wi.Xls.GetRowHidden(s, r)) continue;

                    for (int cIndex = wi.Xls.ColToIndex(s, r, Col2); cIndex > 0; cIndex--)
                    {
                        int c = wi.Xls.ColFromIndex(s, r, cIndex);
                        if (c > Col2 || c == 0) continue;
                        if (c < Col1) break;

                        object o = wi.Xls.GetCellValueAndRecalc(s, r, c, CalcState, CalcStack);
                        if (o is TFlxFormulaErrorValue)
                        {
                            if (CalcState.IgnoreErrors) continue;
                            return o;
                        }

                        if (o is double)
                        {
                            CalcFrequency(dResult, (double)o);
                        }
                    }
                }
            }

            return CreateObjectArray(dResult);
        }

        private void CalcFrequency(double[] dResult, double d)
        {
            int index = BinArray.BinarySearch(new TDoubleInt(0, d));
            if (index < 0) index = ~index;
            if (index == BinArray.Count)
            {
                dResult[index]++;
            }
            else
            {
                dResult[BinArray[index].Position]++;
            }
        }

        private static object[,] CreateObjectArray(double[] dResult)
        {
            object[,] Result = new object[dResult.Length, 1];
            for (int i = 0; i < dResult.Length; i++)
            {
                Result[i, 0] = dResult[i];
            }
            return Result;
        }

        internal override object AggValues(object val1, object val2)
        {
            return TFlxFormulaErrorValue.ErrNA;
        }

        internal override object AggArray(object[,] arr)
        {
            double[] dResult = new double[BinArray.Count + 1];
            for (int i = 0; i < arr.GetLength(0); i++)
                for (int k = 0; k < arr.GetLength(1); k++)
                {
                    if (arr[i,k] is TFlxFormulaErrorValue) return arr[i, k];
                    if (arr[i, k] is double)
                    {
                        CalcFrequency(dResult, (double)arr[i, k]);
                    }
                }

            return CreateObjectArray(dResult);
        }
    }

	#endregion

	#region Token classes
	#region TParsedTokenListBuilder
	/// <summary>
	/// Class used to create a TParsedTokenList. It uses an ArrayList internally, so it is not well suited to be stored. 
	/// (An arraylist can have multiple not used bytes in it)
	/// </summary>
	internal class TParsedTokenListBuilder
	{
        private List<TBaseParsedToken> FList;

		internal TParsedTokenListBuilder()
		{
            FList = new List<TBaseParsedToken>();
		}

		internal TBaseParsedToken this[int index] { get { return (TBaseParsedToken)FList[index]; } set { FList[index] = value; } }

		internal int Count
		{
			get
			{
				return FList.Count;
			}
		}

		#region List
		internal void Add(TBaseParsedToken Data)
		{
			FList.Add(Data);
		}

        internal void RemoveAt(int Postition)
        {
            FList.RemoveAt(Postition);
        }

		internal void AddFrom(TParsedTokenList Source)
		{
			foreach (TBaseParsedToken bp in Source)
			{
				FList.Add(bp);
			}
		}

		internal void Clear()
		{
			FList.Clear();
		}

		internal TParsedTokenList ToParsedTokenList()
		{
            return new TParsedTokenList(FList.ToArray());
		}
		#endregion

		#region AddParsed
		public static TBaseParsedToken GetParsedOp(TOperator op)
		{
			switch (op)
			{

				case TOperator.Neg: return TNegToken.Instance;
				case TOperator.UPlus: return TUPlusToken.Instance;
				case TOperator.Percent: return TPercentToken.Instance;
				case TOperator.Power: return TPowerToken.Instance;
				case TOperator.Mul: return TMulToken.Instance;
				case TOperator.Div: return TDivToken.Instance;
				case TOperator.Add: return TAddToken.Instance;
				case TOperator.Sub: return TSubToken.Instance;
				case TOperator.Concat: return TConcatToken.Instance;
				case TOperator.GE: return TGEToken.Instance;
				case TOperator.LE: return TLEToken.Instance;
				case TOperator.NE: return TNEToken.Instance;
				case TOperator.EQ: return TEQToken.Instance;
				case TOperator.LT: return TLTToken.Instance;
				case TOperator.GT: return TGTToken.Instance;
				default: return new TUnsupportedToken(2, (ptg)((byte)op));
			}
		}

		public static TBaseFunctionToken GetParsedFormula(ptg FmlaPtg, TCellFunctionData Func, byte ArgCount)
		{
			TBaseFunctionToken NewFunction;
			if (FunctionCache.TryGetValue(Func.Index, FmlaPtg, ArgCount, out NewFunction)) return NewFunction;

			NewFunction = GetNewParsedFuntion(FmlaPtg, Func, ArgCount);

			return NewFunction;
		}

		private static TBaseFunctionToken GetNewParsedFuntion(ptg FmlaPtg, TCellFunctionData Func, byte ArgCount)
		{
			//When adding a new function here, search for AddFunctionHere and add it on all those places

			switch (Func.Index)
			{
				case 0: return new TCountToken(ArgCount, FmlaPtg, Func);
				case 169: return new TCountAToken(ArgCount, FmlaPtg, Func);
				case 347: return new TCountBlankToken(ArgCount, FmlaPtg, Func);
				case 1: return new TIFToken(ArgCount, FmlaPtg, Func);

				case 31: return new TMidToken(FmlaPtg, Func);
				case 32: return new TLengthToken(FmlaPtg, Func);
				case 115: return new TLeftToken(ArgCount, FmlaPtg, Func);
				case 116: return new TRightToken(ArgCount, FmlaPtg, Func);
				case 112: return new TLowerToken(FmlaPtg, Func);
				case 113: return new TUpperToken(FmlaPtg, Func);
				case 118: return new TTrimToken(FmlaPtg, Func);
				case 111: return new TCharToken(FmlaPtg, Func);
				case 214: return new TAscToken(FmlaPtg, Func);
				case 119: return new TReplaceToken(FmlaPtg, Func);
				case 124: return new TFindToken(ArgCount, FmlaPtg, Func);
				case 336: return new TConcatenateToken(ArgCount, FmlaPtg, Func);
				case 117: return new TExactToken(FmlaPtg, Func);
				case 114: return new TProperToken(FmlaPtg, Func);
				case 30: return new TReptToken(FmlaPtg, Func);
				case 82: return new TSearchToken(ArgCount, FmlaPtg, Func);
				case 162: return new TCleanToken(FmlaPtg, Func);
				case 120: return new TSubstituteToken(ArgCount, FmlaPtg, Func);
				case 48: return new TTextToken(FmlaPtg, Func);
				case 33: return new TValueToken(FmlaPtg, Func);
				case 121: return new TCodeToken(FmlaPtg, Func);
				case 130: return new TTToken(FmlaPtg, Func);
				case 131: return new TNToken(FmlaPtg, Func);

				case 13: return new TDollarToken(ArgCount, FmlaPtg, Func);
				case 14: return new TFixedToken(ArgCount, FmlaPtg, Func);

				case 359: return new THyperlinkToken(ArgCount, FmlaPtg, Func);

				case 34: return new TTrueToken(FmlaPtg, Func);
				case 35: return new TFalseToken(FmlaPtg, Func);

				case 36: return new TAndToken(ArgCount, FmlaPtg, Func);
				case 37: return new TOrToken(ArgCount, FmlaPtg, Func);
				case 38: return new TNotToken(FmlaPtg, Func);

				case 4: return new TSumToken(ArgCount, FmlaPtg, Func);
				case 5: return new TAverageToken(ArgCount, FmlaPtg, Func);
				case 361: return new TAverageAToken(ArgCount, FmlaPtg, Func);
				case 6: return new TMinToken(ArgCount, FmlaPtg, Func);
				case 363: return new TMinAToken(ArgCount, FmlaPtg, Func);
				case 7: return new TMaxToken(ArgCount, FmlaPtg, Func);
				case 362: return new TMaxAToken(ArgCount, FmlaPtg, Func);
				case 183: return new TProductToken(ArgCount, FmlaPtg, Func);
				case 321: return new TSumSqToken(ArgCount, FmlaPtg, Func);

				case 8: return new TRowToken(ArgCount, FmlaPtg, Func);
				case 9: return new TColToken(ArgCount, FmlaPtg, Func);
				case 75: return new TAreasToken(ArgCount, FmlaPtg, Func);
				case 76: return new TRowsToken(FmlaPtg, Func);
				case 77: return new TColumnsToken(FmlaPtg, Func);
				case 83: return new TTransposeToken(FmlaPtg, Func);


				case 39: return new TModToken(FmlaPtg, Func);
				case 27: return new TRoundToken(FmlaPtg, Func);
				case 24: return new TAbsToken(FmlaPtg, Func);
				case 288: return new TCeilingToken(FmlaPtg, Func);
				case 212: return new TRoundUpToken(FmlaPtg, Func);  //you can search for ROUNDUP
				case 213: return new TRoundDownToken(FmlaPtg, Func);
				case 279: return new TEvenToken(FmlaPtg, Func);
				case 298: return new TOddToken(FmlaPtg, Func);

				case 285: return new TFloorToken(FmlaPtg, Func);
				case 21: return new TExpToken(FmlaPtg, Func);
				case 25: return new TIntToken(FmlaPtg, Func);
				case 22: return new TLnToken(FmlaPtg, Func);
				case 109: return new TLogToken(ArgCount, FmlaPtg, Func);
				case 23: return new TLog10Token(FmlaPtg, Func);
				case 337: return new TPowerFuncToken(FmlaPtg, Func);
				case 63: return new TRandToken(FmlaPtg, Func);
				case 26: return new TSignToken(FmlaPtg, Func);
				case 20: return new TSqrtToken(FmlaPtg, Func);
				case 197: return new TTruncToken(ArgCount, FmlaPtg, Func);

				case 184: return new TFactToken(FmlaPtg, Func);
				case 276: return new TCombinToken(FmlaPtg, Func);
				case 299: return new TPermutToken(FmlaPtg, Func);

				case 344: return new TSubTotalToken(ArgCount, FmlaPtg, Func);

				case 148: return new TIndirectToken(ArgCount, FmlaPtg, Func);

				case 342: return new TRadiansToken(FmlaPtg, Func);
				case 343: return new TDegreesToken(FmlaPtg, Func);

				case 15: return new TSinToken(FmlaPtg, Func);
				case 16: return new TCosToken(FmlaPtg, Func);
				case 17: return new TTanToken(FmlaPtg, Func);
				case 18: return new TAtanToken(FmlaPtg, Func);
				case 19: return new TPiToken(FmlaPtg, Func);
				case 97: return new TAtan2Token(FmlaPtg, Func);
				case 98: return new TAsinToken(FmlaPtg, Func);
				case 99: return new TAcosToken(FmlaPtg, Func);

				case 229: return new TSinhToken(FmlaPtg, Func);
				case 230: return new TCoshToken(FmlaPtg, Func);
				case 231: return new TTanhToken(FmlaPtg, Func);
				case 232: return new TAsinhToken(FmlaPtg, Func);
				case 233: return new TAcoshToken(FmlaPtg, Func);
				case 234: return new TAtanhToken(FmlaPtg, Func);

				case 345: return new TSumIfToken(ArgCount, FmlaPtg, Func);
				case 346: return new TCountIfToken(FmlaPtg, Func);

				case 65: return new TDateToken(FmlaPtg, Func);
				case 140: return new TDateValueToken(FmlaPtg, Func);
				case 67: return new TDayToken(FmlaPtg, Func);
				case 68: return new TMonthToken(FmlaPtg, Func);
				case 69: return new TYearToken(FmlaPtg, Func);
				case 220: return new TDays360Token(ArgCount, FmlaPtg, Func);

				case 66: return new TTimeToken(FmlaPtg, Func);
				case 141: return new TTimeValueToken(FmlaPtg, Func);
				case 71: return new THourToken(FmlaPtg, Func);
				case 72: return new TMinuteToken(FmlaPtg, Func);
				case 73: return new TSecondToken(FmlaPtg, Func);
				case 70: return new TWeekDayToken(ArgCount, FmlaPtg, Func);

				case 74: return new TNowToken(FmlaPtg, Func);
				case 221: return new TTodayToken(FmlaPtg, Func);

				case 261: return new TErrorTypeToken(FmlaPtg, Func);
				case 2: return new TIsNAToken(FmlaPtg, Func);
				case 3: return new TIsErrorToken(FmlaPtg, Func);
				case 126: return new TIsErrToken(FmlaPtg, Func);
				case 127: return new TIsTextToken(FmlaPtg, Func);
				case 128: return new TIsNumberToken(FmlaPtg, Func);
				case 129: return new TIsBlankToken(FmlaPtg, Func);
				case 190: return new TIsNonTextToken(FmlaPtg, Func);
				case 198: return new TIsLogicalToken(FmlaPtg, Func);
				case 105: return new TIsRefToken(FmlaPtg, Func);

				case 86: return new TTypeToken(FmlaPtg, Func);
				case 10: return new TNaToken(FmlaPtg, Func);

				case 78: return new TOffsetToken(ArgCount, FmlaPtg, Func);
				case 100: return new TChooseToken(ArgCount, FmlaPtg, Func);
				case 101: return new THLookupToken(ArgCount, FmlaPtg, Func);
				case 102: return new TVLookupToken(ArgCount, FmlaPtg, Func);
				case 28: return new TLookupToken(ArgCount, FmlaPtg, Func);
				case 29: return new TIndexToken(ArgCount, FmlaPtg, Func);
				case 64: return new TMatchToken(ArgCount, FmlaPtg, Func);
				case 219: return new TAddressToken(ArgCount, FmlaPtg, Func);

				case 228: return new TSumProductToken(ArgCount, FmlaPtg, Func);

				case 46: return new TVarToken(ArgCount, FmlaPtg, Func, 1, false);
				case 194: return new TVarToken(ArgCount, FmlaPtg, Func, 0, false);
				case 367: return new TVarToken(ArgCount, FmlaPtg, Func, 1, true);
				case 365: return new TVarToken(ArgCount, FmlaPtg, Func, 0, true);
				case 12: return new TStDevToken(ArgCount, FmlaPtg, Func);
				case 193: return new TStDevPToken(ArgCount, FmlaPtg, Func);
				case 366: return new TStDevAToken(ArgCount, FmlaPtg, Func);
				case 364: return new TStDevPAToken(ArgCount, FmlaPtg, Func);
				case 307: return new TCorrelToken(FmlaPtg, Func);
				case 308: return new TCoVarToken(FmlaPtg, Func);
				case 322: return new TKurtToken(ArgCount, FmlaPtg, Func);
				case 323: return new TSkewToken(ArgCount, FmlaPtg, Func);

				case 125: return new TCellToken(ArgCount, FmlaPtg, Func);

				case 325: return new TLargeToken(FmlaPtg, Func);
				case 326: return new TSmallToken(FmlaPtg, Func);

				case 40: return new TDCountToken(FmlaPtg, Func);
				case 41: return new TDSumToken(FmlaPtg, Func);
				case 42: return new TDAverageToken(FmlaPtg, Func);
				case 43: return new TDMinToken(FmlaPtg, Func);
				case 44: return new TDMaxToken(FmlaPtg, Func);
				case 189: return new TDProductToken(FmlaPtg, Func);
				case 199: return new TDCountAToken(FmlaPtg, Func);
				case 235: return new TDGetToken(FmlaPtg, Func);
				case 47: return new TDVarToken(FmlaPtg, Func, 1); //dvar
				case 196: return new TDVarToken(FmlaPtg, Func, 0); //dvarp
				case 45: return new TDStDevToken(FmlaPtg, Func, 1); //dstdev
				case 195: return new TDStDevToken(FmlaPtg, Func, 0); //dstdevP

				case 354: return new TRomanToken(ArgCount, FmlaPtg, Func);

				case 247: return new TDBToken(ArgCount, FmlaPtg, Func);
				case 142: return new TSlnToken(FmlaPtg, Func);
				case 143: return new TSydToken(FmlaPtg, Func);
				case 144: return new TDDBToken(ArgCount, FmlaPtg, Func);

				case 11: return new TNPVToken(ArgCount, FmlaPtg, Func);
				case 56: return new TPVToken(ArgCount, FmlaPtg, Func);
				case 57: return new TFVToken(ArgCount, FmlaPtg, Func);
				case 58: return new TNPerToken(ArgCount, FmlaPtg, Func);
				case 59: return new TPMTToken(ArgCount, FmlaPtg, Func);
				case 60: return new TRateToken(ArgCount, FmlaPtg, Func);
				case 61: return new TMIRRToken(ArgCount, FmlaPtg, Func);
				case 62: return new TIRRToken(ArgCount, FmlaPtg, Func);
				case 167: return new TIPMTToken(ArgCount, FmlaPtg, Func);
				case 168: return new TPPMTToken(ArgCount, FmlaPtg, Func);

				case 271: return new TGammaLnToken(FmlaPtg, Func);
				case 273: return new TBinomDistToken(FmlaPtg, Func);
				case 274: return new TChiDistToken(FmlaPtg, Func);
				case 275: return new TChiInvToken(FmlaPtg, Func);
				case 277: return new TConfidenceToken(FmlaPtg, Func);
				case 280: return new TExponDistToken(FmlaPtg, Func);
				case 286: return new TGammaDistToken(FmlaPtg, Func);
				case 287: return new TGammaInvToken(FmlaPtg, Func);
				case 289: return new THypGeomDistToken(4, FmlaPtg, Func);
				case 290: return new TLogNormDistToken(3, FmlaPtg, Func);
				case 291: return new TLogInvToken(FmlaPtg, Func);
				case 292: return new TNegBinomDistToken(3, FmlaPtg, Func);
				case 293: return new TNormDistToken(FmlaPtg, Func);
				case 294: return new TNormsDistToken(FmlaPtg, Func);
				case 295: return new TNormInvToken(FmlaPtg, Func);
				case 296: return new TNormsInvToken(FmlaPtg, Func);
				case 297: return new TStandardizeToken(FmlaPtg, Func);
				case 300: return new TPoissonToken(FmlaPtg, Func);
				case 302: return new TWeibullToken(FmlaPtg, Func);
				case 306: return new TChiTestToken(FmlaPtg, Func);
				case 324: return new TZTestToken(ArgCount, FmlaPtg, Func);

				case 319: return new TGeoMeanToken(ArgCount, FmlaPtg, Func);
				case 320: return new THarMeanToken(ArgCount, FmlaPtg, Func);
				case 216: return new TRankToken(ArgCount, FmlaPtg, Func);

				case 303: return new TSumXmY2Token(FmlaPtg, Func);
				case 304: return new TSumX2mY2Token(FmlaPtg, Func);
				case 305: return new TSumX2pY2Token(FmlaPtg, Func);

				case 165: return new TMMultToken(FmlaPtg, Func);

				case 269: return new TAveDevToken(ArgCount, FmlaPtg, Func);
				case 318: return new TDevSqToken(ArgCount, FmlaPtg, Func);
				case 311: return new TInterceptToken(FmlaPtg, Func);
				case 312: return new TPearsonToken(FmlaPtg, Func);
				case 313: return new TRsqToken(FmlaPtg, Func);
				case 314: return new TSteyxToken(FmlaPtg, Func);
				case 315: return new TSlopeToken(FmlaPtg, Func);

				case 283: return new TFisherToken(FmlaPtg, Func);
				case 284: return new TFisherInvToken(FmlaPtg, Func);

                case 252: return new TFrequencyToken(FmlaPtg, Func);

				case 227: return new TMedianToken(ArgCount, FmlaPtg, Func);
				case 327: return new TPercentileToken(FmlaPtg, Func, true, false);
				case 328: return new TPercentileToken(FmlaPtg, Func, false, false);
                case 329: return new TPercentRankToken(ArgCount, FmlaPtg, Func);
				case 330: return new TModeToken(ArgCount, FmlaPtg, Func);

                //Excel 2007
                case 0x1E0: return new TIfErrorToken(FmlaPtg, Func);
                case 0x1E1: return new TCountIfsToken(ArgCount, FmlaPtg, Func);
                case 0x1E2: return new TSumIfsToken(ArgCount, FmlaPtg, Func);
                case 0x1E3: return new TAverageIfToken(ArgCount, FmlaPtg, Func);
                case 0x1E4: return new TAverageIfsToken(ArgCount, FmlaPtg, Func);

                //Our own
                case (int)TFutureFunctions.CeilingPrecise: return new TCeilingPreciseToken(ArgCount, FmlaPtg, Func);
                case (int)TFutureFunctions.IsoCeiling: return new TCeilingPreciseToken(ArgCount, FmlaPtg, Func); //seems to be the same.
                case (int)TFutureFunctions.FloorPrecise: return new TFloorPreciseToken(ArgCount, FmlaPtg, Func);
                case (int)TFutureFunctions.Aggregate: return new TAggregateToken(ArgCount, FmlaPtg, Func);
                case (int)TFutureFunctions.PercentileExc: return new TPercentileToken(FmlaPtg, Func, false, true);
                case (int)TFutureFunctions.QuartileExc: return new TPercentileToken(FmlaPtg, Func, true, true);
                
                case (int)TFutureFunctions.BinomDist: return new TBinomDistToken(FmlaPtg, Func);
                case (int)TFutureFunctions.ChisqDistRt: return new TChiDistToken(FmlaPtg, Func);
                case (int)TFutureFunctions.ChisqInvRt: return new TChiInvToken(FmlaPtg, Func);
                case (int)TFutureFunctions.ChisqTest: return new TChiTestToken(FmlaPtg, Func);
                case (int)TFutureFunctions.ConfidenceNorm: return new TConfidenceToken(FmlaPtg, Func);
                
                case (int)TFutureFunctions.CovarianceP: return new TCoVarToken(FmlaPtg, Func);
                case (int)TFutureFunctions.ExponDist: return new TExponDistToken(FmlaPtg, Func);
                case (int)TFutureFunctions.GammaDist: return new TGammaDistToken(FmlaPtg, Func);
                case (int)TFutureFunctions.GammaInv: return new TGammaInvToken(FmlaPtg, Func);
                //case (int)TFutureFunctions.HypGeomDist: return new THypGeomDistToken(ArgCount, FmlaPtg, Func); //diff params

                //case (int)TFutureFunctions.LogNormDist: return new TLogNormDistToken(ArgCount, FmlaPtg, Func); //diff params
                case (int)TFutureFunctions.LogNormInv: return new TLogInvToken(FmlaPtg, Func);
                case (int)TFutureFunctions.ModeSngl: return new TModeToken(ArgCount, FmlaPtg, Func);
                //case (int)TFutureFunctions.NegBinom: return new TNegBinomDistToken(ArgCount, FmlaPtg, Func); //diff params
                case (int)TFutureFunctions.NormDist: return new TNormDistToken(FmlaPtg, Func);

                case (int)TFutureFunctions.NormInv: return new TNormInvToken(FmlaPtg, Func);
                //case (int)TFutureFunctions.NormSDist: return new TNormsDistToken(FmlaPtg, Func); //diff params
                case (int)TFutureFunctions.NormSInv: return new TNormsInvToken(FmlaPtg, Func);
                case (int)TFutureFunctions.PercentileInc: return new TPercentileToken(FmlaPtg, Func, false, false);
                case (int)TFutureFunctions.QuartileInc: return new TPercentileToken(FmlaPtg, Func, true, false);
                
                case (int)TFutureFunctions.PercentRankInc: return new TPercentRankToken(ArgCount, FmlaPtg, Func);
                case (int)TFutureFunctions.PoissonDist: return new TPoissonToken(FmlaPtg, Func);
                case (int)TFutureFunctions.RankEq: return new TRankToken(ArgCount, FmlaPtg, Func);
                case (int)TFutureFunctions.StDevP: return new TStDevPToken(ArgCount, FmlaPtg, Func);
                case (int)TFutureFunctions.StDevS: return new TStDevToken(ArgCount, FmlaPtg, Func);
                
                case (int)TFutureFunctions.VarP: return new TVarToken(ArgCount, FmlaPtg, Func, 0, false);
                case (int)TFutureFunctions.VarS: return new TVarToken(ArgCount, FmlaPtg, Func, 1, false);
                case (int)TFutureFunctions.WeibullDist: return new TWeibullToken(FmlaPtg, Func);
                case (int)TFutureFunctions.ZTest: return new TZTestToken(ArgCount, FmlaPtg, Func);

				case 255: return new TUserDefinedToken(ArgCount, FmlaPtg, Func);

				default:
					return new TUnsupportedFunction(ArgCount, FmlaPtg, Func);
			}

		}
		#endregion

    }
	#endregion

	#region List and base
	/// <summary>
	/// This could descend from Stack, but an expression can be evaluated multiple times, and
	/// with the current stack implementation, each time we evaluate it we have to call pop(), so the stack is destroyed.
	/// </summary>
	internal class TParsedTokenList: IEnumerable
	{
		private TBaseParsedToken[] FList;
		internal int TextLenght;  //Only set when converting from text.
		private int Position;

		internal TParsedTokenList(TBaseParsedToken[] Tokens)
		{
			FList = Tokens;
		}

		internal void ResetPositionToLast()
		{
			Position = FList == null ? 0 : FList.Length;
		}

		internal void ResetPositionToStart()
		{
			Position = -1;
		}

		internal int SavePosition()
		{
			return Position;
		}

		internal void RestorePosition(int SavedPosition)
		{
			Position = SavedPosition;
		}

		internal bool Bof()
		{
			return FList == null || Position <= 0;
		}

		internal bool Eof()
		{
			return FList == null || Position >= FList.Length - 1;
		}

		internal TBaseParsedToken LightPop()
		{
			Position--;
			return FList[Position];
		}

		internal TBaseParsedToken ForwardPop()
		{
			Position++;
			return FList[Position];
		}

		internal int Count
		{
			get
			{
				if (FList == null) return 0;
				return FList.Length;
			}
		}

		internal bool IsEmpty
		{
			get
			{
				return FList == null || FList.Length == 0;
			}
		}

        internal void RemoveToken()
        {
            if (Position < 0 || Position >= Count) FlxMessages.ThrowException(FlxErr.ErrInternal);
            TBaseParsedToken[] NewList = new TBaseParsedToken[FList.Length - 1];

			for (int i = FList.Length - 1; i > Position; i--)
			{
				NewList[i - 1] = FList[i];
				FixJumps(FList[i], -1); //always decrease the abs pos.
			}

            for (int i = Position - 1; i >= 0; i--)
            {
                NewList[i] = FList[i];
                FixJumps(NewList[i], Position);
            }

            FList = NewList;
        }

        public static void FixJumps(TBaseParsedToken Token, int RemovePosition)
        {
            TAttrGotoToken tk = Token as TAttrGotoToken; //includes AttrOptIf
            if (tk != null)
            {
                if (tk.PositionOfNextPtg > RemovePosition) tk.PositionOfNextPtg--;
                return;
            }

            TAttrOptChooseToken ochoose = Token as TAttrOptChooseToken;
            if (ochoose != null)
            {
                for (int z = 0; z < ochoose.PositionOfNextPtg.Length; z++)
                {
                    if (ochoose.PositionOfNextPtg[z] > RemovePosition) ochoose.PositionOfNextPtg[z]--;
                }
                return;
            }

            TSimpleMemToken mem = Token as TSimpleMemToken;
            if (mem != null)
            {
                if (mem.PositionOfNextPtg > RemovePosition) mem.PositionOfNextPtg--;
            }
        }

        internal void MoveBack()
        {
            Position--;
        }

        internal void MoveTo(int posi)
        {
            Position = posi;
        }

		internal object EvaluateAll(TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack)
		{
			ResetPositionToLast();
			return EvaluateToken(wi, CalcState, CalcStack);
		}

		internal object EvaluateAll(TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			ResetPositionToLast();
			return EvaluateToken(wi, f, CalcState, CalcStack);
		}

		internal object EvaluateAllRef(TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack)
		{
			ResetPositionToLast();
			return EvaluateTokenRef(wi, CalcState, CalcStack);
		}

		internal object EvaluateToken(TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack)
		{
			if (wi.IsArrayFormula || wi.SumProductCount > 0)
			{
				return EvaluateToken(wi, TArrayAggregate.Instance, CalcState, CalcStack);
			}
			else
			{
				return EvaluateToken(wi, TErr2Aggregate.Instance, CalcState, CalcStack);
			}
		}

		internal object EvaluateToken(TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			if (f == null) return EvaluateToken(wi, CalcState, CalcStack);
			try
			{
				if (Position <= 0) return TFlxFormulaErrorValue.ErrNA;
				return LightPop().Evaluate(this, wi, f, CalcState, CalcStack);
			}
			catch (FlexCelException)
			{
				return TFlxFormulaErrorValue.ErrNA;
			}
			catch (FormatException)
			{
				return TFlxFormulaErrorValue.ErrValue;
			}
			catch (OverflowException)
			{
				return TFlxFormulaErrorValue.ErrNum;
			}
			catch (ArithmeticException)
			{
				return TFlxFormulaErrorValue.ErrValue;
			}
            catch (ArgumentOutOfRangeException)
            {
                return TFlxFormulaErrorValue.ErrNum;
            }
            catch (ArgumentException)
            {
                return TFlxFormulaErrorValue.ErrNum;
            }
        }

		internal object EvaluateTokenRef(TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack)
		{
			if (Position <= 0) return TFlxFormulaErrorValue.ErrNA;
			return LightPop().EvaluateRef(this, wi, CalcState, CalcStack);
		}

		//Flushes a parameter without evaluating it.
		internal void Flush()
		{
			TBaseParsedToken bp = LightPop();
			bp.Flush(this);
		}

		internal TParsedTokenList Clone()
		{
            if (FList == null) return new TParsedTokenList(FList);
			TBaseParsedToken[] NewTokens = new TBaseParsedToken[FList.Length];
			for (int i = 0; i < NewTokens.Length; i++)
			{
				NewTokens[i] = FList[i].Clone();
			}
			return new TParsedTokenList(NewTokens);
		}

		internal bool SameTokens(TParsedTokenList aParsedTokenList)
		{
			if (aParsedTokenList.Count != Count) return false;
			for (int i = Count - 1; i >= 0; i--)
			{
				if (FList[i] != aParsedTokenList.FList[i] //same instances are very common, so we can check this
					&& !FList[i].Same(aParsedTokenList.FList[i]))
				{
					return false;
				}
			}
			return true;
		}

		internal void UnShare(int Row, int Col, bool FromBiff8)
		{
			FList[Position] = ((TBaseRefToken)FList[Position]).UnShare(Row, Col, FromBiff8);
		}

		internal TBaseParsedToken GetToken(int index)
		{
			return FList[index];
		}

        #region IEnumerable Members

        IEnumerator IEnumerable.GetEnumerator()
        {
            return FList.GetEnumerator();
        }

        #endregion
    }

	/// <summary>
	/// Argument types allowed on functions
	/// </summary>
	internal enum TArgType
	{
		Int,
		UInt,
		Double,
		String,
		Boolean,
		Object
	}


	/// <summary>
	/// This class will be the base for all parsed tokens. As many tokens can be created on a singleton pattern
	/// (only one instance on app), there can be no mutable class variables.
	/// </summary>
	internal abstract class TBaseParsedToken
	{
		[ThreadStatic]
		internal static bool Dates1904; //STATIC*  It could be different on different threads. Remember, do not initialize threadstatic members
		protected int FArgCount;
		internal TBaseParsedToken(int aArgCount)
		{
			FArgCount = aArgCount;
		}

		internal int ArgumentCount { get { return FArgCount; } }

		internal virtual void Flush(TParsedTokenList FTokenList)
		{
			for (int i = 0; i < FArgCount; i++) FTokenList.Flush();
		}

		internal abstract object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack);
		internal virtual object EvaluateRef(TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack)
		{
			//FlxMessages.ThrowException(FlxErr.ErrFormulaInvalid, Convert.ToString(FTokenList.FormulaText));
			return TFlxFormulaErrorValue.ErrNA;
		}

		internal abstract ptg GetId { get; }

		internal ptg GetBaseId { get { return CalcBaseToken(GetId); } }

		internal static ptg CalcBaseToken(ptg RealToken)
		{
			if (((byte)RealToken & 0x40) == 0x40) return (ptg)(((byte)RealToken | 0x20) & 0x3F);
			else return (ptg)((byte)RealToken & 0x3F);
		}


		#region Convert
		internal static object UnPack(object v)
		{
			object[,] av = v as object[,];
			if (av != null && av.GetLength(0) == 1 && av.GetLength(1) == 1) return av[0, 0];
			return v;
		}

		internal static object UnPack(object[,] av)
		{
			if (av != null && av.GetLength(0) == 1 && av.GetLength(1) == 1) return av[0, 0];
			return av;
		}

		internal static bool ExtToBool(object Cond, out bool ResultValue)
		{
			Cond = UnPack(Cond);
			if (Cond is double)
			{
				if ((double)Cond == 0) Cond = false; else Cond = true;
			}

			string StrCond = Cond as string;
			if (StrCond != null)
			{
				if (String.Equals(StrCond, TFormulaMessages.TokenString(TFormulaToken.fmTrue), StringComparison.CurrentCultureIgnoreCase))
					Cond = true;
				else
					if (String.Equals(StrCond, TFormulaMessages.TokenString(TFormulaToken.fmFalse), StringComparison.CurrentCultureIgnoreCase))
					Cond = false;
			}

			//Cond = Convert.ToBoolean(Cond, CultureInfo.CurrentCulture);
			if (Cond is TMissingArg) Cond = false;

			if (!(Cond is Boolean))
			{
				ResultValue = false;
				return false;
			}

			ResultValue = (bool)Cond;
			return true;
		}

		internal static bool ExtToDouble(object o, out double ResultValue)
		{
			o = UnPack(o);
			ResultValue = 0;
			if (o is bool)
			{
				if ((bool)o)
				{
					ResultValue = 1;
					return true;
				}
				else
				{
					ResultValue = 0;
					return true;
				}
			}

			if (o is DateTime)
			{
				ResultValue = FlxDateTime.ToOADate((DateTime)o, Dates1904);
				return true;
			}

			return GetDouble(o, out ResultValue);
		}


		internal static bool GetUInt(object v, out int ResultValue, out TFlxFormulaErrorValue Err)
		{
			v = UnPack(v);
			ResultValue = 0;
			Err = TFlxFormulaErrorValue.ErrValue;
			try
			{
				string sv = v as string;
				if (sv != null && sv.Length == 0) return false; //Avoid unnecessary exceptions.

				double d;
				if (!FlxConvert.TryToDouble(v, out d)) return false;
				if (d > Int32.MaxValue) { Err = TFlxFormulaErrorValue.ErrNum; return false; }
				if (d < 0) { Err = TFlxFormulaErrorValue.ErrValue; return false; }
				ResultValue = (int)d;
			}
			catch (InvalidCastException)
			{
				return false;
			}
			if (ResultValue < 0) return false;
			return true;
		}

		internal static bool GetDouble(object v, out double ResultValue)
		{
			v = UnPack(v);
			ResultValue = 0;
			if (v is TFlxFormulaErrorValue) return false;
			try
			{
				string s = v as String;
				if (s != null)
				{
					if (s.Length == 0) return false;
                    if (TCompactFramework.ConvertToNumber(s, CultureInfo.CurrentCulture, out ResultValue)) return true;  //Try to avoid unnecessary exceptions.
					DateTime DateResult;
					if (TCompactFramework.ConvertDateToNumber(s, out DateResult))
					{
						return FlxDateTime.TryToOADate(DateResult, Dates1904, out ResultValue);
					}
					return false;
				}

				return FlxConvert.TryToDouble(v, out ResultValue);
			}
			catch (InvalidCastException)
			{
				return false;
			}
			catch (FormatException)
			{
				return false;
			}

		}

		internal static bool GetInt(object v, out int ResultValue, out TFlxFormulaErrorValue Err)
		{
			v = UnPack(v);
			ResultValue = 0;
			Err = TFlxFormulaErrorValue.ErrValue;
			try
			{
				string sv = v as string;
				if (sv != null && sv.Length == 0) return false; //Avoid unnecessary exceptions.

				double d;
				if (!FlxConvert.TryToDouble(v, out d)) return false;
				if (d > Int32.MaxValue || d < Int32.MinValue) { Err = TFlxFormulaErrorValue.ErrNum; return false; }
				ResultValue = (int)d;
			}
			catch (InvalidCastException)
			{
				return false;
			}
			return true;
		}


		internal static object GetRangeOrArray(TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack, out TAddress adr1, out TAddress adr2, out object[,] values, TFlxFormulaErrorValue DefaultError, bool AllowArrays)
		{
			TAddressList adr = null;
			adr1 = null; adr2 = null;
			object Result = GetRangeOrArray(FTokenList, wi, CalcState, CalcStack, out adr, out values, DefaultError, AllowArrays);
			if (Result != null) return Result;
			if (values != null) return null;
			if (adr.Count != 1) return DefaultError;
			adr1 = adr[0][0];
			adr2 = adr[0][1];
			return null;
		}

		internal static object GetRangeOrArray(TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack, out TAddress adr1, out TAddress adr2, out object[,] values, bool AllowArrays)
		{
			return GetRangeOrArray(FTokenList, wi, CalcState, CalcStack, out adr1, out adr2, out values, TFlxFormulaErrorValue.ErrRef, AllowArrays);
		}

		internal static object GetRange(object v1, out TAddress adr1, out TAddress adr2, TFlxFormulaErrorValue DefaultError)
		{
			TAddressList adr = null; adr1 = null; adr2 = null;
			object Result = GetRange(v1, out adr, DefaultError);
			if (Result != null) return Result;
			if (adr.Count != 1) return DefaultError;
			adr1 = adr[0][0];
			adr2 = adr[0][1];
			return null;
		}

		internal static object GetRangeOrArray(TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack, out TAddressList adr, out object[,] values, bool AllowArrays)
		{
			return GetRangeOrArray(FTokenList, wi, CalcState, CalcStack, out adr, out values, TFlxFormulaErrorValue.ErrRef, AllowArrays);
		}

		internal static object GetRangeOrArray(TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack, out TAddressList adr, out object[,] values, TFlxFormulaErrorValue DefaultError, bool AllowArrays)
		{
			adr = null; values = null;
			//Here we can have a reference or a value. As we don't know this before, we must try both.
			int OriginalPos = FTokenList.SavePosition();
			object v0 = FTokenList.EvaluateTokenRef(wi, CalcState, CalcStack);
			if (v0 is TFlxFormulaErrorValue && AllowArrays)
			{
				FTokenList.RestorePosition(OriginalPos);
				v0 = FTokenList.EvaluateToken(wi, TArrayAggregate.Instance, CalcState, CalcStack);
				if (v0 is TFlxFormulaErrorValue) return v0;

				values = v0 as object[,];
				if (values == null)
				{
					double dd;
					if (ExtToDouble(v0, out dd))
					{
						values = new object[1, 1];
						values[0, 0] = dd;
					}
				}
				if (values != null) return null;
			}
			if (v0 is TFlxFormulaErrorValue) return v0;

			return GetRange(v0, out adr, DefaultError);
		}

		internal static object GetRange(object v1, out TAddressList adr, TFlxFormulaErrorValue DefaultError)
		{
			adr = null;
			TAddressList alist = v1 as TAddressList;
			if (alist != null)
			{
				adr = alist;
				return null;
			}

			object v2 = v1;
			TAddress[] range = (v1 as TAddress[]);
			if (range != null)
			{
				v1 = range[0];
				v2 = range[1];
			}

			if (v1 is TFlxFormulaErrorValue) return v1;
			if (v2 is TFlxFormulaErrorValue) return v2;

			TAddress[] a = new TAddress[] { (v1 as TAddress), v2 as TAddress };
			if (a[0] == null) return DefaultError;
			if (a[1] == null) return DefaultError;
			adr = new TAddressList(a);
			return null;
		}

		#endregion

		#region GetArguments
		protected static object GetUIntArgument(TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack, ref int l, out object[,] ArrL)
		{
			ArrL = null;
			object v1 = FTokenList.EvaluateToken(wi, CalcState, CalcStack);
			if (v1 is TFlxFormulaErrorValue) return v1;
			if (!IsArrayArgument(v1, out ArrL))
			{
				if (!(v1 is TMissingArg)) //Keep the default
				{
					TFlxFormulaErrorValue Err;
					if (!GetUInt(v1, out l, out Err)) return Err;
				}
			}

			return null;
		}

		protected static object GetIntArgument(TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack, ref int l, out object[,] ArrL)
		{
			ArrL = null;
			object v1 = FTokenList.EvaluateToken(wi, CalcState, CalcStack);
			if (v1 is TFlxFormulaErrorValue) return v1;
			if (!IsArrayArgument(v1, out ArrL))
			{
				if (!(v1 is TMissingArg)) //Keep the default
				{
					TFlxFormulaErrorValue Err;
					if (!GetInt(v1, out l, out Err)) return Err;
				}
			}

			return null;
		}

		protected static object GetBoolArgument(TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack, ref bool b, out object[,] ArrL)
		{
			ArrL = null;
			object v1 = FTokenList.EvaluateToken(wi, CalcState, CalcStack);
			if (v1 is TFlxFormulaErrorValue) return v1;
			if (!IsArrayArgument(v1, out ArrL))
			{
				if (!(v1 is TMissingArg)) //Keep the default
				{
					if (!ExtToBool(v1, out b)) return TFlxFormulaErrorValue.ErrValue;
				}
			}

			return null;
		}

		protected static object GetDoubleArgument(TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack, ref double d, out object[,] ArrL)
		{
			ArrL = null;
			object v1 = FTokenList.EvaluateToken(wi, CalcState, CalcStack);
			if (v1 is TFlxFormulaErrorValue) return v1;
			if (!IsArrayArgument(v1, out ArrL))
			{
				if (!(v1 is TMissingArg)) //Keep the default
				{
					if (!GetDouble(v1, out d)) return TFlxFormulaErrorValue.ErrValue;
				}
			}

			return null;
		}

		protected static object GetStringArgument(TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack, ref string s, out object[,] ArrL)
		{
			ArrL = null;
			object v1 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v1 is TFlxFormulaErrorValue) return v1;
			if (!IsArrayArgument(v1, out ArrL))
			{
				if (!(v1 is TMissingArg)) //Keep the default
				{
					s = FlxConvert.ToString(v1);
				}
			}

			return null;
		}

		protected static object GetObjectArgument(TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack, ref object s, out object[,] ArrL)
		{
			ArrL = null;
			object v1 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v1 is TFlxFormulaErrorValue) return v1;
			if (!IsArrayArgument(v1, out ArrL))
			{
				if (!(v1 is TMissingArg)) //Keep the default
				{
					s = v1;
				}
			}

			return null;
		}

		protected static object GetArgList(TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack, long ValueCount, ref object[] Values, ref object[][,] ArrValues, TArgType[] ArgType, bool IgnoreErrors)
		{
			for (long i = ValueCount - 1; i >= 0; i--)
			{
				switch (ArgType[i])
				{
					case TArgType.UInt:
						int z = (int)Values[i];
						object Ret1 = GetUIntArgument(FTokenList, wi, CalcState, CalcStack, ref z, out ArrValues[i]);
						Values[i] = z;
						if (Ret1 != null)
							if (IgnoreErrors) Values[i] = Ret1; else return Ret1;
						break;
					case TArgType.Int:
						int y = (int)Values[i];
						object Ret2 = GetIntArgument(FTokenList, wi, CalcState, CalcStack, ref y, out ArrValues[i]);
						Values[i] = y;
						if (Ret2 != null)
							if (IgnoreErrors) Values[i] = Ret2; else return Ret2;
						break;
					case TArgType.Double:
						double d = (double)Values[i];
						object Ret3 = GetDoubleArgument(FTokenList, wi, CalcState, CalcStack, ref d, out ArrValues[i]);
						Values[i] = d;
						if (Ret3 != null)
							if (IgnoreErrors) Values[i] = Ret3; else return Ret3;
						break;
					case TArgType.String:
						string s = (string)Values[i];
						object Ret4 = GetStringArgument(FTokenList, wi, CalcState, CalcStack, ref s, out ArrValues[i]);
						Values[i] = s;
						if (Ret4 != null)
							if (IgnoreErrors) Values[i] = Ret4; else return Ret4;
						break;
					case TArgType.Boolean:
						bool b = (bool)Values[i];
						object Ret5 = GetBoolArgument(FTokenList, wi, CalcState, CalcStack, ref b, out ArrValues[i]);
						Values[i] = b;
						if (Ret5 != null)
							if (IgnoreErrors) Values[i] = Ret5; else return Ret5;
						break;
					case TArgType.Object:
						object o = Values[i];
						object Ret6 = GetObjectArgument(FTokenList, wi, CalcState, CalcStack, ref o, out ArrValues[i]);
						Values[i] = o;
						if (Ret6 != null)
							if (IgnoreErrors) Values[i] = Ret6; else return Ret6;
						break;
				}
			}
			return null;
		}


		protected static bool GetObjectItem(ref object[,] ArrResult, ref object[,] ArrL, int i, int j, object Default, out object Res, TArgType ArgType)
		{
			Res = TFlxFormulaErrorValue.ErrNA;
			object v1 = GetItem(ArrL, i, j); if (v1 is TFlxFormulaErrorValue) { ArrResult[i, j] = v1; return false; }
			if (v1 is TMissingArg)
			{
				Res = Default;
				return true;
			}

			TFlxFormulaErrorValue Err;

			switch (ArgType)
			{
				case TArgType.UInt:
					int l;
					if (!GetUInt(v1, out l, out Err)) { ArrResult[i, j] = Err; return false; }
					Res = l;
					return true;
				case TArgType.Int:
					int z;
					if (!GetInt(v1, out z, out Err)) { ArrResult[i, j] = Err; return false; }
					Res = z;
					return true;
				case TArgType.Double:
					double d;
					if (!ExtToDouble(v1, out d)) { ArrResult[i, j] = TFlxFormulaErrorValue.ErrValue; return false; }
					Res = d;
					return true;
				case TArgType.String:
					Res = FlxConvert.ToString(v1);
					return true;
				case TArgType.Boolean:
					bool b;
					if (!ExtToBool(v1, out b)) { ArrResult[i, j] = TFlxFormulaErrorValue.ErrValue; return false; }
					Res = b;
					return true;
				case TArgType.Object:
					Res = v1;
					return true;
			}
			return false;
		}

		protected static bool GetUIntItem(ref object[,] ArrResult, ref object[,] ArrL, int i, int j, out int l)
		{
			l = 0;
			object v1 = GetItem(ArrL, i, j); if (v1 is TFlxFormulaErrorValue) { ArrResult[i, j] = v1; return false; }
			TFlxFormulaErrorValue Err;
			if (!GetUInt(v1, out l, out Err)) { ArrResult[i, j] = Err; return false; }
			return true;
		}

		protected static bool GetIntItem(ref object[,] ArrResult, ref object[,] ArrL, int i, int j, out int l)
		{
			l = 0;
			object v1 = GetItem(ArrL, i, j); if (v1 is TFlxFormulaErrorValue) { ArrResult[i, j] = v1; return false; }
			TFlxFormulaErrorValue Err;
			if (!GetInt(v1, out l, out Err)) { ArrResult[i, j] = Err; return false; }
			return true;
		}

		protected static bool GetBoolItem(ref object[,] ArrResult, ref object[,] ArrL, int i, int j, out bool b)
		{
			b = false;
			object v1 = GetItem(ArrL, i, j); if (v1 is TFlxFormulaErrorValue) { ArrResult[i, j] = v1; return false; }
			if (!ExtToBool(v1, out b)) { ArrResult[i, j] = TFlxFormulaErrorValue.ErrValue; return false; }
			return true;
		}

		protected static bool GetDoubleItem(ref object[,] ArrResult, ref object[,] ArrL, int i, int j, out double d)
		{
			d = 0;
			object v1 = GetItem(ArrL, i, j); if (v1 is TFlxFormulaErrorValue) { ArrResult[i, j] = v1; return false; }
			if (!ExtToDouble(v1, out d)) { ArrResult[i, j] = TFlxFormulaErrorValue.ErrValue; return false; }
			return true;
		}

		protected static bool GetStringItem(ref object[,] ArrResult, ref object[,] ArrL, int i, int j, out string s)
		{
			s = null;
			object v1 = GetItem(ArrL, i, j); if (v1 is TFlxFormulaErrorValue) { ArrResult[i, j] = v1; return false; }
			s = FlxConvert.ToString(v1);
			return true;
		}

		#endregion

		#region GetAllArguments
		protected object DoArguments(TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack, int ArgCount, ref object[] Values, TArgType[] ArgType, bool IgnoreErrors)
		{
			object[] Defaults = (object[])Values.Clone();
			object[][,] ArrValues = new object[Values.Length][,];
			object ret2 = GetArgList(FTokenList, wi, CalcState, CalcStack, ArgCount, ref Values, ref ArrValues, ArgType, IgnoreErrors);
			if (ret2 != null) return ret2;

			bool HasArray = false;
			for (int i = 0; i < ArrValues.Length; i++)
			{
				if (ArrValues[i] != null)
				{
					HasArray = true;
					break;
				}
			}

			if (HasArray)
			{
				int RowCount = 1; int ColCount = 1;
				for (int i = 0; i < ArrValues.Length - 1; i++)
				{
					SetupArray(ref ArrValues[i], ref ArrValues[i + 1], Values[i], Values[i + 1], ref RowCount, ref ColCount);
				}

				object[,] ArrResult = new object[RowCount, ColCount];
				for (int i = 0; i < ArrResult.GetLength(0); i++)
					for (int j = 0; j < ArrResult.GetLength(1); j++)
					{
						bool AllOk = true;
						for (int z = 0; z < ArrValues.Length; z++)
						{
							object obj;
							if (!GetObjectItem(ref ArrResult, ref ArrValues[z], i, j, Defaults[z], out obj, ArgType[z]))
							{
								AllOk = false;
								break;
							}
							Values[z] = obj;
						}
						if (!AllOk) continue;
						ArrResult[i, j] = DoOneArg(wi, Values);
					}
				return UnPack(ArrResult);
			}

			return DoOneArg(wi, Values);
		}

        protected virtual object DoOneArg(TWorkbookInfo wi, object[] Values)
		{
			return TFlxFormulaErrorValue.ErrNA;
		}


		#endregion

		#region Compare
		internal static object CompareValues(object v1, object v2)
		{
			v1 = ConvertToAllowedObject(v1);
			v2 = ConvertToAllowedObject(v2);
			if (v1 is TFlxFormulaErrorValue) return v1;
			if (v2 is TFlxFormulaErrorValue) return v2;
			if (v1 == null)
			{
				if (v2 == null) return 0;
				if (v2 is double) v1 = (double)0;
				if (v2 is string) v1 = String.Empty;
				if (v2 is bool) v1 = false;
			}
			if (v2 == null)
			{
				if (v1 is double) v2 = (double)0;
				if (v1 is string) v2 = String.Empty;
				if (v1 is bool) v2 = false;
			}

			//double<string<bool.
			if (v1 is Double)
			{
				if (v2 is Double)
					return ((double)v1).CompareTo(v2);
				return -1;
			}

			if (v1 is string)
			{
				if (v2 is string)
					return String.Compare((string)v1, (string)v2, StringComparison.CurrentCultureIgnoreCase);
				if (v2 is double) return 1; else return -1;
			}
			if (v1 is bool)
			{
				if (v2 is bool)
					return ((bool)v1).CompareTo(v2);
				if (v2 is double || v2 is string) return 1;  //We must check here, because if it weren't a bool or a string, we need to return an error.
			}
			return TFlxFormulaErrorValue.ErrNA;

		}

		internal static object CompareWithWildcards(object o1, object SearchPattern, bool UseWildcards)
		{
			if (!UseWildcards) return CompareValues(o1, SearchPattern);
			string s1 = o1 as string;
			if (s1 == null) return CompareValues(o1, SearchPattern);
			string Pattern = SearchPattern as string;
			if (Pattern == null) return CompareValues(o1, SearchPattern);

			s1 = s1.ToUpper(CultureInfo.CurrentCulture); Pattern = Pattern.ToUpper(CultureInfo.CurrentCulture);

			if (TWildcardMatch.Matches(Pattern, s1))
				return 0;
			return -1;

		}

		protected static bool SameType(object a, object b)
		{
			if (a == null) return false;
			if (b == null) b = (double)0;
			return a.GetType() == b.GetType();
		}


		#endregion

		#region General utility
		protected static object ConvertToAllowedObject(object o)
		{
			return TExcelTypes.ConvertToAllowedObject(o, Dates1904);
		}

		protected static object ConvertToAllowedParam(object o)
		{
			if (o is TMissingArg) return null;
			TXls3DRange range = o as TXls3DRange;
			if (range != null) return range;

            return ConvertToAllowedObjectWithArrays(o);

		}

        private static object ConvertToAllowedObjectWithArrays(object o)
        {
            object[,] ArrResult = o as object[,];
            if (ArrResult != null)
            {
                for (int i = 0; i < ArrResult.GetLength(0); i++)
                {
                    for (int j = 0; j < ArrResult.GetLength(1); j++)
                    {
                        ArrResult[i, j] = ConvertToAllowedParam(ArrResult[i, j]);
                    }
                }
                return ArrResult;
            }


            return ConvertToAllowedObject(o);
        }

		internal static decimal CalcFactor(double d)
		{
			int i = d > 15? 15: d < -15? -15: (int)d;
			if (i == 0) return 0;
			if (i > 0)
			{
				if (i > 15) i = 15;
				decimal Result = 10;
				for (int k = 1; k < i; k++)
					Result = Result * 10;
				return Result;

			}

			if (i < -15) i = -15;
			decimal Result2 = (decimal)1.0 / (decimal)10.0;
			for (int k = 1; k < -i; k++)
				Result2 = Result2 / 10;
			return Result2;
		}

        internal static double UpRound(double d, double decimals)
		{
			decimal factor = CalcFactor(decimals);
			//In .NET 1.0 Round always rounds with banker's round. As Excel uses arithmethic rounding, we can't use:
			//return Math.Round(d2, i);

			if (factor == 0) return Math.Floor(Math.Abs(d) + 0.5) * Math.Sign(d);
			return ((double)(Decimal.Floor(((Decimal)Math.Abs(d) * factor + 0.5M)) / factor)) * Math.Sign(d);
		}

		internal static void OrderRange(ref int Sheet1, ref int Sheet2, ref int Row1, ref int Col1, ref int Row2, ref int Col2)
		{
			if (Sheet2 < Sheet1)
			{
				int s = Sheet2;
				Sheet2 = Sheet1;
				Sheet1 = s;
			}
			if (Row2 < Row1)
			{
				int r = Row2;
				Row2 = Row1;
				Row1 = r;
			}
			if (Col2 < Col1)
			{
				int c = Col2;
				Col2 = Col1;
				Col1 = c;
			}
		}
		#endregion

		#region Array formulas
		protected static void SetupArray(ref object[,] Arr1, ref object[,] Arr2, object o1, object o2, ref int MaxRow, ref int MaxCol)
		{
			FillArray(ref Arr1, o1);
			FillArray(ref Arr2, o2);
			GetMaxArrayDim(Arr1, ref MaxRow, ref MaxCol);
			GetMaxArrayDim(Arr2, ref MaxRow, ref MaxCol);
		}

		protected static void SetupArray(ref object[,] ArrL, ref object[,] ArrS, ref object[,] ArrSp, object l, object s, object sp, out object[,] ArrResult)
		{
			FillArray(ref ArrL, l);
			FillArray(ref ArrS, s);
			FillArray(ref ArrSp, sp);
			int MaxCol = 1;
			int MaxRow = 1;
			GetMaxArrayDim(ArrL, ref MaxRow, ref MaxCol);
			GetMaxArrayDim(ArrS, ref MaxRow, ref MaxCol);
			GetMaxArrayDim(ArrSp, ref MaxRow, ref MaxCol);

			ArrResult = new object[MaxRow, MaxCol];
		}

		internal static bool IsArrayArgument(object Value, out object[,] ArrResult, out object[,] ArrValue)
		{
			ArrValue = null;
			ArrResult = null;
			ArrValue = Value as object[,];
			if (ArrValue == null) return false;
			ArrResult = new object[ArrValue.GetLength(0), ArrValue.GetLength(1)];
			return true;
		}

		internal static bool IsArrayArgument(object Value, out object[,] ArrValue)
		{
			ArrValue = null;
			ArrValue = Value as object[,];
			if (ArrValue == null) return false;
			return true;
		}

		internal static void FillArray(ref object[,] a, object value)
		{
			if (a == null)
			{
				a = new object[1, 1];
				a[0, 0] = value;
			}
		}

		internal static bool SameDimensions(object[,] Arr1, object[,] Arr2)
		{
			return Arr1.GetLength(0) == Arr2.GetLength(0) && Arr1.GetLength(1) == Arr2.GetLength(1);
		}

		internal static bool IsUniDim(object[,] Arr)
		{
			return Arr.GetLength(0) == 1 && Arr.GetLength(1) == 1;
		}

		internal static bool CompatibleDimensions(object[,] Arr1, object[,] Arr2)
		{
			if (Arr1 == null || Arr2 == null) return true;
			if (Arr1.GetLength(0) == Arr2.GetLength(0) && Arr1.GetLength(1) == Arr2.GetLength(1)) return true;
			if (IsUniDim(Arr1) || IsUniDim(Arr2)) return true;
			return false;
		}

		internal static object[,] GetArrayObj(object[,] a1, object[,] a2)
		{
			if (a1 == null || (a1.GetLength(0) <= 1 && a1.GetLength(1) <= 1))
			{
				if (a2 == null) return new object[1, 1];
				else
					return new object[a2.GetLength(0), a2.GetLength(1)];
			}
			return new object[a1.GetLength(0), a1.GetLength(1)];
		}

		internal static void GetMaxArrayDim(object[,] a1, ref int MaxRow, ref int MaxCol)
		{
			if (MaxRow < 1) MaxRow = 1;
			if (MaxCol < 1) MaxCol = 1;
			if (a1 == null) return;
			if (a1.GetLength(0) > MaxRow) MaxRow = a1.GetLength(0);
			if (a1.GetLength(1) > MaxCol) MaxCol = a1.GetLength(1);
		}

		internal static object GetItem(object[,] a, int i, int j)
		{
			if (a == null) return null;
			if (a.GetLength(0) == 1) i = 0;
			if (a.GetLength(1) == 1) j = 0;
			if (i >= a.GetLength(0) || j >= a.GetLength(1)) return TFlxFormulaErrorValue.ErrNA;
			return a[i, j];
		}

		#endregion

		internal virtual TBaseParsedToken Clone()
		{
			return (TBaseParsedToken)MemberwiseClone();
		}

		internal virtual TBaseParsedToken SetId(ptg aId)
		{
			return this;
		}

		internal virtual bool Same(TBaseParsedToken aBaseParsedToken)
		{
			if (aBaseParsedToken == null) return false;
			return aBaseParsedToken.GetType() == this.GetType();
		}

		internal virtual int ExternSheet{get{return -1;} set{} }
	}

	/// <summary>
	/// This class includes support for static instances.
	/// </summary>
	internal abstract class TStaticToken: TBaseParsedToken
	{
		protected TStaticToken(int aArgCount)
			: base(aArgCount)
		{
		}

		internal override TBaseParsedToken Clone()
		{
			return this;
		}
	}
	#endregion

	#region Data Tokens
	internal abstract class TDataToken : TBaseParsedToken
	{
		protected object Data;
		internal TDataToken(object aData) : base(0) {Data = aData; }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			return Data;
		}

		internal override object EvaluateRef(TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack)
		{
			if (Data is TFlxFormulaErrorValue) return Data;
			return base.EvaluateRef(FTokenList, wi, CalcState, CalcStack);
		}

		internal override bool Same(TBaseParsedToken aBaseParsedToken)
		{
			if (!base.Same(aBaseParsedToken)) return false;
			TDataToken dt = aBaseParsedToken as TDataToken;
			if (dt == null) return false;
			if (dt.Data == null) return Data == null;
			return dt.Data.Equals(Data);
		}
	}

	internal sealed class TIntDataToken : TDataToken
	{
		internal TIntDataToken(int aData) : base((double)aData) { }
		internal override ptg GetId
		{
			get { return ptg.Int; }
		}

		internal double GetData()
		{
			return (double)Data;
		}

        #region Clone
        internal override TBaseParsedToken Clone()
        {
            return new TIntDataToken((int)((double)Data));
        }
        #endregion


	}

	internal sealed class TBoolDataToken : TDataToken
	{
		internal TBoolDataToken(bool aData) : base(aData) { }
		internal override ptg GetId
		{
			get { return ptg.Bool; }
		}

		internal bool GetData()
		{
			return (bool)Data;
		}

        #region Clone
        internal override TBaseParsedToken Clone()
        {
            return new TBoolDataToken((bool)Data);
        }
        #endregion
	}

	internal sealed class TNumDataToken : TDataToken
	{
		internal TNumDataToken(double aData) : base(aData) { }
		internal override ptg GetId
		{
			get { return ptg.Num; }
		}

		internal double GetData()
		{
			return (double)Data;
		}

        #region Clone
        internal override TBaseParsedToken Clone()
        {
            return new TNumDataToken((double)Data);
        }
        #endregion

	}

	internal sealed class TErrDataToken : TDataToken
	{
		internal TErrDataToken(TFlxFormulaErrorValue aData) : base(aData) { }
		internal override ptg GetId
		{
			get { return ptg.Err; }
		}

		internal TFlxFormulaErrorValue GetData()
		{
			return (TFlxFormulaErrorValue)Data;
		}

        #region Clone
        internal override TBaseParsedToken Clone()
        {
            return new TErrDataToken((TFlxFormulaErrorValue)Data);
        }
        #endregion
	}

	internal sealed class TStrDataToken : TDataToken
	{
        internal TStrDataToken(string aData, string FormulaText, bool CheckForErrors)
            : base(aData)
        {
            if (CheckForErrors && aData.Length > FlxConsts.Max_FormulaStringConstant)
            {
                FlxMessages.ThrowException(FlxErr.ErrStringConstantInFormulaTooLong, aData, FormulaText);
            }
        }
		
        internal override ptg GetId
		{
			get { return ptg.Str; }
		}

		internal string GetData()
		{
			return FlxConvert.ToString(Data);
		}

        #region Clone
        internal override TBaseParsedToken Clone()
        {
            return new TStrDataToken((string)Data, String.Empty, false);
        }
        #endregion
	}

	internal class TMissingArgDataToken : TBaseParsedToken
	{
		internal static readonly TMissingArgDataToken Instance = new TMissingArgDataToken();

		private TMissingArgDataToken() : base(0) { }
		internal override ptg GetId
		{
			get { return ptg.MissArg; }
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			return TMissingArg.Instance;
		}
	}

	internal class TArrayDataToken : TBaseParsedToken
	{
		private ptg FId;
		private object[,] Data;
		internal TArrayDataToken(ptg aId, object[,] aData) : base(0) { FId = aId; Data = aData; }

		internal override ptg GetId
		{
			get { return FId; }
		}

		internal override TBaseParsedToken SetId(ptg aId)
		{
			if (FId == aId) return this;
			return new TArrayDataToken(aId, Data); //no need to clone data.
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			if (Data == null || Data.Length == 0) return TFlxFormulaErrorValue.ErrValue;
			return f.AggArray((object[,])Data.Clone()); //AggArray might change the value of array, we don't want that
		}

		internal override object EvaluateRef(TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack)
		{
			return base.EvaluateRef(FTokenList, wi, CalcState, CalcStack);
		}

		internal object[,] GetData { get { return Data; } }

		internal override bool Same(TBaseParsedToken aBaseParsedToken)
		{
			if (!base.Same(aBaseParsedToken)) return false;
            
			TArrayDataToken v2 = aBaseParsedToken as TArrayDataToken;
			if (v2 == null) return false;
			if (v2.Data.GetLength(0) != Data.GetLength(0) || v2.Data.GetLength(1) != Data.GetLength(1)) return false;

			for (int r = Data.GetLength(0) - 1; r >= 0; r--)
			{
				for (int c = Data.GetLength(1) - 1; c >= 0; c--)
				{
					if (Data[r, c] != v2.Data[r, c]) return false;
				}
			}

			return true;
		}

		internal override TBaseParsedToken Clone()
		{
			return this;
		}


	}
	#endregion

	#region Arithmetic Tokens

	internal abstract class TBaseArithOpToken : TStaticToken
	{
		protected TBaseArithOpToken(int aArgCount) : base(aArgCount) { }
		protected abstract object Op(double a1, double a2);

		/// <summary>
		/// a1 and are guaranteed not to be null.
		/// </summary>
		/// <param name="a1"></param>
		/// <param name="a2"></param>
		/// <returns></returns>
		private object ProcessArray(object[,] a1, object[,] a2)
		{
			//optimized most common case. would work also commenting this block.
			if (SameDimensions(a1, a2))
			{
				for (int i = 0; i < a1.GetLength(0); i++)
					for (int k = 0; k < a1.GetLength(1); k++)
					{
						double d1; if (!ExtToDouble(a1[i, k], out d1)) return TFlxFormulaErrorValue.ErrValue;
						double d2; if (!ExtToDouble(a2[i, k], out d2)) return TFlxFormulaErrorValue.ErrValue; ;
						a1[i, k] = Op(d1, d2);
					}
				return a1;
			}

			if (a1.GetLength(0) != a2.GetLength(0))
				if (a1.GetLength(0) != 1 && a2.GetLength(0) != 1)
					return TFlxFormulaErrorValue.ErrNA;
			if (a1.GetLength(1) != a2.GetLength(1))
				if (a1.GetLength(1) != 1 && a2.GetLength(1) != 1)
					return TFlxFormulaErrorValue.ErrNA;

			int lx = Math.Max(a1.GetLength(0), a2.GetLength(0));
			int ly = Math.Max(a1.GetLength(1), a2.GetLength(1));

			object[,] Result = new object[lx, ly];


			for (int i = 0; i < lx; i++)
			{
				int i1 = 0; if (a1.GetLength(0) > 1) i1 = i;
				int i2 = 0; if (a2.GetLength(0) > 1) i2 = i;

				for (int k = 0; k < ly; k++)
				{
					int k1 = 0; if (a1.GetLength(1) > 1) k1 = k;
					int k2 = 0; if (a2.GetLength(1) > 1) k2 = k;

					if (a1[i1, k1] is TFlxFormulaErrorValue) return a1[i1, k1];
					if (a2[i2, k2] is TFlxFormulaErrorValue) return a2[i2, k2];
					double d1; if (!ExtToDouble(a1[i1, k1], out d1)) return TFlxFormulaErrorValue.ErrValue;
					double d2; if (!ExtToDouble(a2[i2, k2], out d2)) return TFlxFormulaErrorValue.ErrValue;
					Result[i, k] = Op(d1, d2);
				}
			}
			return Result;
		}


		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			object v1 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v1 is TFlxFormulaErrorValue) return v1;
			object v2 = FArgCount < 2 ?
			v1 :
				FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v2 is TFlxFormulaErrorValue) return v2;
			object[,] a1 = v1 as object[,];
			object[,] a2 = v2 as object[,];

			if (a1 != null)
			{
				if (a2 == null)
				{
					double d2; if (!ExtToDouble(v2, out d2)) return TFlxFormulaErrorValue.ErrValue;
					a2 = new object[1, 1];
					a2[0, 0] = d2;
				}
				return ProcessArray(a1, a2);
			}

			if (a2 != null)
			{
				double d1; if (!ExtToDouble(v1, out d1)) return TFlxFormulaErrorValue.ErrValue;
				a1 = new object[1, 1];
				a1[0, 0] = d1;
				return ProcessArray(a1, a2);
			}

			double dd1; if (!ExtToDouble(v1, out dd1)) return TFlxFormulaErrorValue.ErrValue;
			double dd2; if (!ExtToDouble(v2, out dd2)) return TFlxFormulaErrorValue.ErrValue;
			return Op(dd1, dd2);
		}
	}

	internal sealed class TPercentToken : TBaseArithOpToken
	{
		private TPercentToken() : base(1) { }
		internal static readonly TPercentToken Instance = new TPercentToken();

		internal override ptg GetId { get { return ptg.Percent; } }

		protected override object Op(double a1, double a2)
		{
			return a1 / 100;
		}
	}

	internal sealed class TPowerToken : TBaseArithOpToken
	{
		private TPowerToken() : base(2) { }
		internal static readonly TPowerToken Instance = new TPowerToken();

		internal override ptg GetId { get { return ptg.Power; } }

		protected override object Op(double a1, double a2)
		{
			return Math.Pow(a2, a1);
		}
	}

	internal sealed class TNegToken : TBaseArithOpToken
	{
		private TNegToken() : base(1) { }
		internal static readonly TNegToken Instance = new TNegToken();

		internal override ptg GetId { get { return ptg.Uminus; } }

		protected override object Op(double a1, double a2)
		{
			return -a1;
		}
	}

	internal sealed class TUPlusToken : TStaticToken
	{
		private TUPlusToken() : base(1) { }
		internal static readonly TUPlusToken Instance = new TUPlusToken();

		internal override ptg GetId { get { return ptg.Uplus; } }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			return FTokenList.EvaluateToken(wi, f, CalcState, CalcStack);
		}
		internal override object EvaluateRef(TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack)
		{
			return FTokenList.EvaluateTokenRef(wi, CalcState, CalcStack);
		}
	}

	internal sealed class TAddToken : TBaseArithOpToken
	{
		private TAddToken() : base(2) { }
		internal static readonly TAddToken Instance = new TAddToken();

		internal override ptg GetId { get { return ptg.Add; } }

		protected override object Op(double a1, double a2)
		{
			return a1 + a2;
		}
	}

	internal sealed class TSubToken : TBaseArithOpToken
	{
		private TSubToken() : base(2) { }
		internal static readonly TSubToken Instance = new TSubToken();

		internal override ptg GetId { get { return ptg.Sub; } }

		protected override object Op(double a1, double a2)
		{
			return -a1 + a2;
		}
	}

	internal sealed class TMulToken : TBaseArithOpToken
	{
		private TMulToken() : base(2) { }
		internal static readonly TMulToken Instance = new TMulToken();

		internal override ptg GetId { get { return ptg.Mul; } }

		protected override object Op(double a1, double a2)
		{
			return a1 * a2;
		}
	}

	internal sealed class TDivToken : TBaseArithOpToken
	{
		private TDivToken() : base(2) { }
		internal static readonly TDivToken Instance = new TDivToken();

		internal override ptg GetId { get { return ptg.Div; } }

		protected override object Op(double a1, double a2)
		{
			if (a1 == 0) return TFlxFormulaErrorValue.ErrDiv0;
			return a2 / a1;

		}
	}

	#endregion

	#region String Tokens

	internal sealed class TConcatToken : TStaticToken
	{
		private TConcatToken() : base(2) { }
		internal static readonly TConcatToken Instance = new TConcatToken();

		internal override ptg GetId { get { return ptg.Concat; } }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			object[] Values = { String.Empty, String.Empty };
			return DoArguments(FTokenList, wi, CalcState, CalcStack, 2, ref Values, new TArgType[] { TArgType.String, TArgType.String }, false);
		}
        protected override object DoOneArg(TWorkbookInfo wi, object[] Values)
		{
			return FlxConvert.ToString(Values[0]) + FlxConvert.ToString(Values[1]);
		}

	}

	#endregion

	#region Equality Tokens

	internal abstract class TEqualityToken : TStaticToken
	{
		protected TEqualityToken(int aArgCount) : base(aArgCount) { }

		protected object CompareArrValues(object v1, object v2)
		{
			object[,] a1 = v1 as object[,];
			object[,] a2 = v2 as object[,];

			if (a1 == null && a2 == null)
			{
				object res = CompareValues(v1, v2);
				if (res is TFlxFormulaErrorValue) return res;
				return Op((int)res);
			}

			if (a1 == null)
			{
				a1 = new object[1, 1];
				a1[0, 0] = v1;
			}
			if (a2 == null)
			{
				a2 = new object[1, 1];
				a2[0, 0] = v2;
			}

			if (a1.GetLength(0) != a2.GetLength(0))
				if (a1.GetLength(0) != 1 && a2.GetLength(0) != 1)
					return TFlxFormulaErrorValue.ErrNA;
			if (a1.GetLength(1) != a2.GetLength(1))
				if (a1.GetLength(1) != 1 && a2.GetLength(1) != 1)
					return TFlxFormulaErrorValue.ErrNA;

			int lx = Math.Max(a1.GetLength(0), a2.GetLength(0));
			int ly = Math.Max(a1.GetLength(1), a2.GetLength(1));

			object[,] Result = new object[lx, ly];


			for (int i = 0; i < lx; i++)
			{
				int i1 = 0; if (a1.GetLength(0) > 1) i1 = i;
				int i2 = 0; if (a2.GetLength(0) > 1) i2 = i;

				for (int k = 0; k < ly; k++)
				{
					int k1 = 0; if (a1.GetLength(1) > 1) k1 = k;
					int k2 = 0; if (a2.GetLength(1) > 1) k2 = k;

					object res = CompareValues(a1[i1, k1], a2[i2, k2]);
					if (res is TFlxFormulaErrorValue)
						Result[i, k] = res;
					else
						Result[i, k] = Op((int)res);
				}
			}
			return Result;
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			TBaseAggregate f2 = (f.PropagateOnEquality) ? f : TErr2Aggregate.Instance;  //Equality must go down in subproduct.
			if (wi.IsArrayFormula) f2 = TArrayAggregate.Instance;

			object obj1 = FTokenList.EvaluateToken(wi, f2, CalcState, CalcStack); if (obj1 is TFlxFormulaErrorValue) return obj1;
			object obj2 = FTokenList.EvaluateToken(wi, f2, CalcState, CalcStack); if (obj2 is TFlxFormulaErrorValue) return obj2;
			object res = CompareArrValues(obj1, obj2);
			return res;
		}


		protected abstract bool Op(int res);
	}

	internal sealed class TEQToken : TEqualityToken
	{
		private TEQToken() : base(2) { }
		internal static readonly TEQToken Instance = new TEQToken();

		internal override ptg GetId { get { return ptg.EQ; } }

		protected override bool Op(int res)
		{
			return res == 0;
		}
	}

	internal sealed class TNEToken : TEqualityToken
	{
		private TNEToken() : base(2) { }
		internal static readonly TNEToken Instance = new TNEToken();

		internal override ptg GetId { get { return ptg.NE; } }

		protected override bool Op(int res)
		{
			return res != 0;
		}

	}

	internal sealed class TGEToken : TEqualityToken
	{
		private TGEToken() : base(2) { }
		internal static readonly TGEToken Instance = new TGEToken();

		internal override ptg GetId { get { return ptg.GE; } }

		protected override bool Op(int res)
		{
			return res <= 0;
		}
	}

	internal sealed class TLEToken : TEqualityToken
	{
		private TLEToken() : base(2) { }
		internal static readonly TLEToken Instance = new TLEToken();

		internal override ptg GetId { get { return ptg.LE; } }
		protected override bool Op(int res)
		{
			return res >= 0;
		}
	}

	internal sealed class TGTToken : TEqualityToken
	{
		private TGTToken() : base(2) { }
		internal static readonly TGTToken Instance = new TGTToken();

		internal override ptg GetId { get { return ptg.GT; } }
		protected override bool Op(int res)
		{
			return res < 0;
		}

	}

	internal sealed class TLTToken : TEqualityToken
	{
		private TLTToken() : base(2) { }
		internal static readonly TLTToken Instance = new TLTToken();

		internal override ptg GetId { get { return ptg.LT; } }
		protected override bool Op(int res)
		{
			return res > 0;
		}

	}

	#endregion

	#region Reference Tokens
	internal abstract class TBaseRefToken : TBaseParsedToken
	{
		#region Variables
		protected ptg FId;
		#endregion

		#region Constructor
		internal TBaseRefToken(int aArgCount, ptg aId)
			: base(aArgCount)
		{
			FId = aId;
		}
		#endregion

		internal static int WrapRow(int a, bool FromBiff8)
		{
			unchecked
			{
                int MaxRow = FromBiff8 ? FlxConsts.Max_Rows97_2003 : FlxConsts.Max_Rows;
                while (a < 0) a += MaxRow + 1;
				while (a > MaxRow) a -= MaxRow + 1;
				return a;
			}
		}

        internal static int WrapRowSigned(int a, bool IsAbsolute, bool FromBiff8)
        {
            if (IsAbsolute) return a;
            unchecked
            {
                int MaxRow = FromBiff8 ? FlxConsts.Max_Rows97_2003 : FlxConsts.Max_Rows;
                int middle = (MaxRow + 1) / 2;
                int low = -middle; int high = middle - 1;
                while (a < low) a += MaxRow + 1;
                while (a > high) a -= MaxRow + 1;
                return a;
            }
        }

		internal static int WrapColumn(int a, bool FromBiff8)
		{
			unchecked
			{
                int MaxCol = FromBiff8 ? FlxConsts.Max_Columns97_2003 : FlxConsts.Max_Columns;
				while (a < 0) a += MaxCol + 1;
				while (a > MaxCol) a -= MaxCol + 1;
				return a;
			}
		}

        internal static int WrapColumnSigned(int a, bool IsAbsolute, bool FromBiff8)
        {
            if (IsAbsolute) return a;
            unchecked
            {
                int MaxCol = FromBiff8 ? FlxConsts.Max_Columns97_2003 : FlxConsts.Max_Columns;
                int middle = (MaxCol + 1) / 2;
                int low = -middle; int high = middle - 1;
                while (a < low) a += MaxCol + 1;
                while (a > high) a -= MaxCol + 1;
                return a;
            }
        }


        internal static int WrapSigned(int a, int Max)
        {
            unchecked
            {
                int middle = (Max + 1) / 2;
                int low = -middle; int high = middle - 1;

                while (a < low) a += Max + 1;
                while (a > high) a -= Max + 1;
                return a;
            }
        }

		internal static bool CheckRef(TWorkbookInfo wi, int SheetIndexBase1, int r, int c)
		{
			if (SheetIndexBase1 < 1 || SheetIndexBase1 > wi.Xls.SheetCount)
				return false;

			if (r < 0 || r > FlxConsts.Max_Rows || c < 0 || c > FlxConsts.Max_Columns)
				return false;
			return true;
		}

		internal override ptg GetId { get { return FId; } }

		internal override TBaseParsedToken SetId(ptg aId)
		{
			FId = aId;
			return this;
		}

		internal bool IsErr()
		{
			ptg BaseId = CalcBaseToken(FId);
			return (BaseId == ptg.RefErr) || (BaseId == ptg.AreaErr) || (BaseId == ptg.Ref3dErr) || (BaseId == ptg.Area3dErr);
		}

		//This method must be called every time, since the externsheet could change.
		internal object ParseExternSheet(ExcelFile Xls, out string ExternalBook, out int Sheet1, out int Sheet2)
		{
			ExternalBook = null;
			bool ExternalSheets; string ExternBookName;
			Xls.GetSheetsFromExternSheet(ExternSheet, out Sheet1, out Sheet2, out ExternalSheets, out ExternBookName);
			Sheet1++;
			Sheet2++;

			if (ExternalSheets)
			{
				ExcelFile ExternalXls = Xls.GetSupportingFile(ExternBookName);
				if (ExternalXls != null)
				{
					ExternalBook = ExternBookName;
				}
				else
				{
					Xls.AddUnsupported(TUnsupportedFormulaErrorType.ExternalReference, ExternBookName);

					return TFlxFormulaErrorValue.ErrNA;
				}
			}
			return null;
		}

		internal override bool Same(TBaseParsedToken aBaseParsedToken)
		{
			if (!base.Same(aBaseParsedToken)) return false;
			TBaseRefToken tk = aBaseParsedToken as TBaseRefToken;
			if (tk == null) return false;
			if (tk.FId != FId) return false;
			return true;
		}

		#region Relative references
		internal virtual int GetRow(int Row, bool RefAbsolute, bool IgnoreRowOfs, int CellRow, int RowOfs)
		{
			return IgnoreRowOfs? Row: Row + RowOfs;
		}

		internal virtual int GetCol(int Col, bool RefAbsolute, bool IgnoreColOfs, int CellCol, int ColOfs)
		{
			return IgnoreColOfs? Col: Col + ColOfs;
		}

		internal int GetRow(int Row, bool RefAbsolute, bool IgnoreRowOfs, TWorkbookInfo wi)
		{
			return GetRow(Row, RefAbsolute, IgnoreRowOfs, wi.Row, wi.RowOfs);
		}

		internal int GetCol(int Col, bool RefAbsolute, bool IgnoreColOfs, TWorkbookInfo wi)
		{
			return GetCol(Col, RefAbsolute, IgnoreColOfs, wi.Col, wi.ColOfs);
		}
        
		internal virtual bool CanHaveRelativeOffsets
		{
			get
			{
				return false;
			}
		}
		#endregion

		#region Token Manipulator

        internal virtual TBaseRefToken UnShare(int aRow, int aCol, bool FromBiff8)
		{
			FlxMessages.ThrowException(FlxErr.ErrInternal); //should never be here.
			return null;
		}

		internal abstract void CreateInvalidRef();

		/// <summary>
		/// Increments a row/col, up to a Maximum value. If Max is reached, then Max will be returned.
		/// If a value less than 0 is reached, or the row/col is in the deleted range, in invalid ref is created. 
		/// </summary>
		/// <param name="RowCol">Variable to update. </param>
		/// <param name="InsPos">  Row/Col where we insert. To know the deleted range</param>
		/// <param name="Offset"> Amount to increment. Might be negative</param>
		/// <param name="Max"> Maximum value</param>
		/// <param name="CheckInside">When true and the cell is inside the deleted range, an invalidref will be created.</param>
		internal void IncRowCol(ref int RowCol, int InsPos, int Offset, int Max, bool CheckInside)
		{
			long w = RowCol;
			//Handle deletes...
			if (CheckInside && (Offset < 0) && (InsPos >= 0) && (w >= InsPos) && (w < InsPos - Offset))
			{
				CreateInvalidRef();
				return;
			}

			w += Offset;

			if ((w < 0) || (w > Max))
			{
				CreateInvalidRef();
				return;
			}

			RowCol = (int)w;
		}

		#endregion
	}
	internal class TRefToken : TBaseRefToken
	{
		#region Variables
		internal int Row;
		internal int Col;
		internal bool RowAbs;
		internal bool ColAbs;
		#endregion

		#region Constructor
		internal TRefToken(ptg aId, int aRow, int aCol, bool aRowAbs, bool aColAbs)
			: base(0, aId)
		{
			Row = aRow;
			Col = aCol;
			RowAbs = aRowAbs;
			ColAbs = aColAbs;
		}
		#endregion

		#region Evaluate
		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			if (IsErr()) return TFlxFormulaErrorValue.ErrRef;

			int ro = GetRow(Row, RowAbs, RowAbs, wi);
			int co = GetCol(Col, ColAbs, ColAbs, wi);
			if (!CheckRef(wi, wi.SheetIndexBase1, ro, co))
				return TFlxFormulaErrorValue.ErrRef;

			return f.Agg(wi, wi.SheetIndexBase1, wi.SheetIndexBase1, ro + 1, co + 1, ro + 1, co + 1, CalcState, CalcStack);
		}

		internal override object EvaluateRef(TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack)
		{
			if (IsErr()) return TFlxFormulaErrorValue.ErrRef;

			int ro = GetRow(Row, RowAbs, RowAbs, wi);
			int co = GetCol(Col, ColAbs, ColAbs, wi);
			if (!CheckRef(wi, wi.SheetIndexBase1, ro, co))
				return TFlxFormulaErrorValue.ErrRef;

			return new TAddress(wi, wi.SheetIndexBase1, ro + 1, co + 1);

		}

		internal override bool Same(TBaseParsedToken aBaseParsedToken)
		{
			if (!base.Same(aBaseParsedToken)) return false;
			TRefToken tk = aBaseParsedToken as TRefToken;
			if (tk == null) return false;
			if (tk.Row != Row || tk.Col != Col || tk.RowAbs != RowAbs || tk.ColAbs != ColAbs) return false;
			return true;
		}
		#endregion

		#region Rows and Cols

		internal int GetRow1(int CellRow) { return GetRow(Row, RowAbs, RowAbs, CellRow, 0); }
		internal int GetCol1(int CellCol) { return GetCol(Col, ColAbs, ColAbs, CellCol, 0); }

		internal override void CreateInvalidRef()
		{
			if (IsErr()) return;
			FId += ptg.RefErr - ptg.Ref;
		}

		internal void IncRow(int InsPos, int Offset, int Max, bool CheckInside)
		{
			IncRowCol(ref Row, InsPos, Offset, Max, CheckInside);
		}

		internal void IncCol(int InsPos, int Offset, int Max, bool CheckInside)
		{
			IncRowCol(ref Col, InsPos, Offset, Max, CheckInside);
		}
		#endregion

        #region Clone
        internal override TBaseParsedToken Clone()
        {
            return new TRefToken(FId, Row, Col, RowAbs, ColAbs);
        }
        #endregion

	}

	internal sealed class TRefNToken : TRefToken
	{
		#region Constructor
		internal TRefNToken(ptg aId, int aRow, int aCol, bool aRowAbs, bool aColAbs, bool FromBiff8)
			: base(aId, WrapRowSigned(aRow, aRowAbs, FromBiff8), WrapColumnSigned(aCol, aColAbs, FromBiff8), aRowAbs, aColAbs)
		{ }
		#endregion

		#region Rows and Cols
		internal override int GetRow(int Row, bool RefAbsolute, bool IgnoreRowOfs, int CellRow, int RowOfs)
		{
			//Relative references wrap.
			if (RefAbsolute) return base.GetRow(Row, RefAbsolute, IgnoreRowOfs, CellRow, RowOfs);
			return WrapRow(Row + CellRow + RowOfs, false);
		}

		internal override int GetCol(int Col, bool RefAbsolute, bool IgnoreColOfs, int CellCol, int ColOfs)
		{
			if (RefAbsolute) return base.GetCol(Col, RefAbsolute, IgnoreColOfs, CellCol, ColOfs);
			return WrapColumn(Col + CellCol + ColOfs, false);
		}

		internal override bool CanHaveRelativeOffsets
		{
			get
			{
				return true;
			}
		}

		internal override void CreateInvalidRef()
		{
			if (IsErr()) return;
			FId += ptg.RefErr - ptg.RefN; //there is no RefNErr. We will change it to RefErr
		}

		#endregion

		#region Token Manipulator

		internal override TBaseRefToken UnShare(int aRow, int aCol, bool FromBiff8)
		{
			ptg NewId = (ptg)((byte)GetId + (byte)ptg.Ref - (byte)ptg.RefN);
			int NewRow = RowAbs? Row: WrapRow(Row + aRow, FromBiff8);
			int NewCol = ColAbs? Col: WrapColumn(Col + aCol, FromBiff8);
			return new TRefToken(NewId, NewRow, NewCol, RowAbs, ColAbs);
		}

		#endregion

        #region Clone
        internal override TBaseParsedToken Clone()
        {
            return new TRefNToken(FId, Row, Col, RowAbs, ColAbs, false);
        }
        #endregion
    }

	internal class TAreaToken : TBaseRefToken
	{
		#region Variables
		internal int Row1;
		internal int Col1;
		internal int Row2;
		internal int Col2;
		internal bool RowAbs1;
		internal bool ColAbs1;
		internal bool RowAbs2;
		internal bool ColAbs2;
		#endregion

		#region Constructor
		internal TAreaToken(ptg aId, int aRow1, int aCol1, bool aRowAbs1, bool aColAbs1,
			int aRow2, int aCol2, bool aRowAbs2, bool aColAbs2)
			: base(0, aId)
		{
			Row1 = aRow1;
			Col1 = aCol1;
			RowAbs1 = aRowAbs1;
			ColAbs1 = aColAbs1;

			Row2 = aRow2;
			Col2 = aCol2;
			RowAbs2 = aRowAbs2;
			ColAbs2 = aColAbs2;
		}
		#endregion

		#region Evaluate
		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			if (IsErr()) return TFlxFormulaErrorValue.ErrRef;

			int ro1 = GetRow(Row1, RowAbs1, RowAbs1, wi);
			int co1 = GetCol(Col1, ColAbs1, ColAbs1, wi);
			int ro2 = GetRow(Row2, RowAbs2, RowAbs2, wi);
			int co2 = GetCol(Col2, ColAbs2, ColAbs2, wi);

			if (!CheckRef(wi, wi.SheetIndexBase1, ro1, co1))
				return TFlxFormulaErrorValue.ErrRef;
			if (!CheckRef(wi, wi.SheetIndexBase1, ro2, co2))
				return TFlxFormulaErrorValue.ErrRef;

			return f.Agg(wi, wi.SheetIndexBase1, wi.SheetIndexBase1, ro1 + 1, co1 + 1, ro2 + 1, co2 + 1, CalcState, CalcStack);
		}

		internal override object EvaluateRef(TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack)
		{
			if (IsErr()) return TFlxFormulaErrorValue.ErrRef;

			int ro1 = GetRow(Row1, RowAbs1, RowAbs1, wi);
			int co1 = GetCol(Col1, ColAbs1, ColAbs1, wi);
			int ro2 = GetRow(Row2, RowAbs2, RowAbs2, wi);
			int co2 = GetCol(Col2, ColAbs2, ColAbs2, wi);

			if (!CheckRef(wi, wi.SheetIndexBase1, ro1, co1))
				return TFlxFormulaErrorValue.ErrRef;
			if (!CheckRef(wi, wi.SheetIndexBase1, ro2, co2))
				return TFlxFormulaErrorValue.ErrRef;

			TAddress[] Result = new TAddress[2];
			Result[0] = new TAddress(wi, wi.SheetIndexBase1, ro1 + 1, co1 + 1);
			Result[1] = new TAddress(wi, wi.SheetIndexBase1, ro2 + 1, co2 + 1);
			return Result;
		}

		internal override bool Same(TBaseParsedToken aBaseParsedToken)
		{
			if (!base.Same(aBaseParsedToken)) return false;
			TAreaToken tk = aBaseParsedToken as TAreaToken;
			if (tk == null) return false;
			if (tk.Row1 != Row1 || tk.Col1 != Col1 || tk.RowAbs1 != RowAbs1 || tk.ColAbs1 != ColAbs1) return false;
			if (tk.Row2 != Row2 || tk.Col2 != Col2 || tk.RowAbs2 != RowAbs2 || tk.ColAbs2 != ColAbs2) return false;
			return true;
		}

		#endregion

		#region Rows and Cols
		internal int GetRow1(int CellRow) { return GetRow(Row1, RowAbs1, RowAbs1, CellRow, 0); }
		internal int GetCol1(int CellCol) { return GetCol(Col1, ColAbs1, ColAbs1, CellCol, 0); }
		internal int GetRow2(int CellRow) { return GetRow(Row2, RowAbs2, RowAbs2, CellRow, 0); }
		internal int GetCol2(int CellCol) { return GetCol(Col2, ColAbs2, ColAbs2, CellCol, 0); }

		internal override void CreateInvalidRef()
		{
			if (IsErr()) return;
			FId += ptg.AreaErr - ptg.Area;
		}

		internal void IncRow1(int InsPos, int Offset, int Max, bool CheckInside)
		{
			IncRowCol(ref Row1, InsPos, Offset, Max, CheckInside);
		}

		internal void IncCol1(int InsPos, int Offset, int Max, bool CheckInside)
		{
			IncRowCol(ref Col1, InsPos, Offset, Max, CheckInside);
		}

		internal void IncRow2(int InsPos, int Offset, int Max, bool CheckInside)
		{
			IncRowCol(ref Row2, InsPos, Offset, Max, CheckInside);
		}

		internal void IncCol2(int InsPos, int Offset, int Max, bool CheckInside)
		{
			IncRowCol(ref Col2, InsPos, Offset, Max, CheckInside);
		}
        
		internal void DeleteRowsArea(int CellRow, int aRowCount, TXlsCellRange CellRange)
		{
			if (CanHaveRelativeOffsets && (!RowAbs1 || !RowAbs2)) return; //need to check it is a name too for this condition to apply
			if (GetRow1(CellRow) >= CellRange.Top)
				if (GetRow2(CellRow) < CellRange.Top + CellRange.RowCount * -aRowCount) //range is all inside the deleted range
				{
					CreateInvalidRef();
				}
				else
				{
					//Do NOT DELETE full columns sum(a:a). should always remain this. (Except when deleting the full range on another sheet);
					//A funny one: Sum(a65536:a65536), when deleting a row, will convert to sum(a65535:a65536). Only second row will wrap.
					IncRow1(CellRange.Top, Math.Max(CellRange.Top - GetRow1(CellRow), aRowCount * CellRange.RowCount), FlxConsts.Max_Rows, false); //Use max here, as we are using negative numbers.
					if (GetRow2(CellRow) < FlxConsts.Max_Rows)
						IncRow2(CellRange.Top, aRowCount * CellRange.RowCount, FlxConsts.Max_Rows, false);
				}
			else
			{
				if (GetRow2(CellRow) >= FlxConsts.Max_Rows) return; //Do NOT DELETE full columns sum(a:a) should always remain this. (Except when deleting the full range on another sheet);

				if (GetRow2(CellRow) >= CellRange.Top)
					IncRow2(CellRange.Top, Math.Max(CellRange.Top - GetRow2(Row2) - 1, aRowCount * CellRange.RowCount), FlxConsts.Max_Rows, false);
			}
		}

		internal void DeleteColsArea(int CellCol, int aColCount, TXlsCellRange CellRange)
		{
			if (GetCol1(CellCol) >= CellRange.Left)
				if (GetCol2(CellCol) < CellRange.Left + CellRange.ColCount * -aColCount) //range is all inside the deleted range
				{
					CreateInvalidRef();
				}
				else
				{
					IncCol1(CellRange.Left, Math.Max(CellRange.Left - GetCol1(CellCol), aColCount * CellRange.ColCount), FlxConsts.Max_Columns, false); //Use max here, as we are using negative numbers.
					if (GetCol2(CellCol) < FlxConsts.Max_Columns)
						IncCol2(CellRange.Left, aColCount * CellRange.ColCount, FlxConsts.Max_Columns, false);
				}
			else
			{
				if (GetCol2(CellCol) >= FlxConsts.Max_Columns) return; //Do NOT DELETE full rows sum(1:1) should always remain this. (Except when deleting the full range on another sheet);

				if (GetCol2(CellCol) >= CellRange.Left)
					IncCol2(CellRange.Left, Math.Max(CellRange.Left - GetCol2(CellCol) - 1, aColCount * CellRange.ColCount), FlxConsts.Max_Columns, false);
			}
		}

		#endregion

        #region Clone
        internal override TBaseParsedToken Clone()
        {
            return new TAreaToken(FId, Row1, Col1, RowAbs1, ColAbs1, Row2, Col2, RowAbs2, ColAbs2);
        }
        #endregion

	}

	internal sealed class TAreaNToken : TAreaToken
	{
		#region Constructor
		internal TAreaNToken(ptg aId, int aRow1, int aCol1, bool aRowAbs1, bool aColAbs1, int aRow2, int aCol2, bool aRowAbs2, bool aColAbs2, bool FromBiff8)
            : base(aId, WrapRowSigned(aRow1, aRowAbs1, FromBiff8), WrapColumnSigned(aCol1, aColAbs1, FromBiff8), aRowAbs1, aColAbs1,
            WrapRowSigned(aRow2, aRowAbs2, FromBiff8), WrapColumnSigned(aCol2, aColAbs2, FromBiff8), aRowAbs2, aColAbs2)
		{ }

		#endregion

		#region Rows and Cols
		internal override int GetRow(int Row, bool RefAbsolute, bool IgnoreRowOfs, int CellRow, int RowOfs)
		{
			//Relative references wrap.
			if (RefAbsolute) return base.GetRow(Row, RefAbsolute, IgnoreRowOfs, CellRow, RowOfs);
			return WrapRow(Row + CellRow + RowOfs, false);
		}

		internal override int GetCol(int Col, bool RefAbsolute, bool IgnoreColOfs, int CellCol, int ColOfs)
		{
			if (RefAbsolute) return base.GetCol(Col, RefAbsolute, IgnoreColOfs, CellCol, ColOfs);
			return WrapColumn(Col + CellCol + ColOfs, false);
		}

		internal override bool CanHaveRelativeOffsets
		{
			get
			{
				return true;
			}
		}

		internal override void CreateInvalidRef()
		{
			if (IsErr()) return;
			FId += ptg.AreaErr - ptg.AreaN; //there is no AreaNErr. We will change it to RefErr
		}

		#endregion

		#region Token Manipulator

        internal override TBaseRefToken UnShare(int aRow, int aCol, bool FromBiff8)
		{
			ptg NewId = (ptg)((byte)GetId + (byte)ptg.Area - (byte)ptg.AreaN);
			int NewRow1 = RowAbs1 ? Row1 : WrapRow(Row1 + aRow, FromBiff8);
			int NewCol1 = ColAbs1 ? Col1 : WrapColumn(Col1 + aCol, FromBiff8);
			int NewRow2 = RowAbs2 ? Row2 : WrapRow(Row2 + aRow, FromBiff8);
			int NewCol2 = ColAbs2 ? Col2 : WrapColumn(Col2 + aCol, FromBiff8);
			return new TAreaToken(NewId, NewRow1, NewCol1, RowAbs1, ColAbs1, NewRow2, NewCol2, RowAbs2, ColAbs2);
		}
		#endregion

        #region Clone
        internal override TBaseParsedToken Clone()
        {
            return new TAreaNToken(FId, Row1, Col1, RowAbs1, ColAbs1, Row2, Col2, RowAbs2, ColAbs2, false);
        }
        #endregion

	}

	internal class TRef3dToken : TRefToken
	{
		#region Variables
		internal int FExternSheet;
		#endregion

		#region Constructor
		internal TRef3dToken(ptg aId, int aExternSheet, int aRow, int aCol, bool aRowAbs, bool aColAbs)
			: base(aId, aRow, aCol, aRowAbs, aColAbs)
		{
			FExternSheet = aExternSheet;
		}
		#endregion

		#region Evaluate
		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			if (IsErr()) return TFlxFormulaErrorValue.ErrRef;

			string ExternalBook; int Sheet1; int Sheet2;
			object res = ParseExternSheet(wi.Xls, out ExternalBook, out Sheet1, out Sheet2);
			if (res != null) return res; 

			int ro = GetRow(Row, RowAbs, RowAbs || wi.SheetIndexBase1 != Sheet1 || Sheet1 != Sheet2 || ExternalBook != null, wi);
			int co = GetCol(Col, ColAbs, ColAbs || wi.SheetIndexBase1 != Sheet1 || Sheet1 != Sheet2 || ExternalBook != null, wi);

			TWorkbookInfo wi2 = ExternalBook == null? wi: wi.ShallowClone();
			if (ExternalBook != null)
			{
				wi2.Xls = wi.Xls.GetSupportingFile(ExternalBook);
				if (wi2.Xls == null) return TFlxFormulaErrorValue.ErrRef;
			}

			if (!CheckRef(wi2, Sheet1, ro, co))
				return TFlxFormulaErrorValue.ErrRef;
			if (!CheckRef(wi2, Sheet2, ro, co))
				return TFlxFormulaErrorValue.ErrRef;

			return f.Agg(wi2, Sheet1, Sheet2, ro + 1, co + 1, ro + 1, co + 1, CalcState, CalcStack);
		}

		internal override object EvaluateRef(TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack)
		{
			if (IsErr()) return TFlxFormulaErrorValue.ErrRef;

			string ExternalBook; int Sheet1; int Sheet2;
			object res = ParseExternSheet(wi.Xls, out ExternalBook, out Sheet1, out Sheet2);
			if (res != null) return res;

			int ro = GetRow(Row, RowAbs, RowAbs || wi.SheetIndexBase1 != Sheet1 || Sheet1 != Sheet2 || ExternalBook != null, wi);
			int co = GetCol(Col, ColAbs, ColAbs || wi.SheetIndexBase1 != Sheet1 || Sheet1 != Sheet2 || ExternalBook != null, wi);

			TWorkbookInfo wi2 = ExternalBook == null? wi: wi.ShallowClone();
			if (ExternalBook != null)
			{
				wi2.Xls = wi.Xls.GetSupportingFile(ExternalBook);
				if (wi2.Xls == null) return TFlxFormulaErrorValue.ErrRef;
			}

			if (!CheckRef(wi2, Sheet1, ro, co))
				return TFlxFormulaErrorValue.ErrRef;
			if (!CheckRef(wi2, Sheet2, ro, co))
				return TFlxFormulaErrorValue.ErrRef;

			if (Sheet1 == Sheet2)
				return new TAddress(wi2, ExternalBook, Sheet1, ro +  1, co + 1);

			TAddress[] Result = new TAddress[2];
			Result[0] = new TAddress(wi2, ExternalBook, Sheet1, ro + 1, co + 1);
			Result[1] = new TAddress(wi2, ExternalBook, Sheet2, ro + 1, co + 1);
			return Result;
		}

		internal override bool Same(TBaseParsedToken aBaseParsedToken)
		{
			if (!base.Same(aBaseParsedToken)) return false;
			TRef3dToken tk = aBaseParsedToken as TRef3dToken;
			if (tk == null) return false;
			if (tk.FExternSheet != FExternSheet) return false;
			return true;
		}

		#endregion

		#region Rows and Cols
		internal override int ExternSheet
		{
			get
			{
				return FExternSheet;
			}
			set
			{
				FExternSheet = value;
			}
		}

		internal override void CreateInvalidRef()
		{
			if (IsErr()) return;
			FId += ptg.Ref3dErr - ptg.Ref3d;
		}
		#endregion

        #region Clone
        internal override TBaseParsedToken Clone()
        {
            return new TRef3dToken(FId, FExternSheet, Row, Col, RowAbs, ColAbs);
        }
        #endregion

	}

	/// <summary>
	/// This class doesn't really exist in Excel, but named ranges use this behavior, even if they use TRef3dTokens.
	/// </summary>
	internal sealed class TRef3dNToken : TRef3dToken
	{
		#region Constructor
		internal TRef3dNToken(ptg aId, int aExternSheet, int aRow, int aCol, bool aRowAbs, bool aColAbs, bool FromBiff8)
			: base(aId, aExternSheet, WrapRowSigned(aRow, aRowAbs, FromBiff8), WrapColumnSigned(aCol, aColAbs, FromBiff8), aRowAbs, aColAbs)
		{ }
		#endregion

		#region Rows and Cols
		internal override int GetRow(int Row, bool RefAbsolute, bool IgnoreRowOfs, int CellRow, int RowOfs)
		{
			//Relative references wrap.
			if (RefAbsolute) return base.GetRow(Row, RefAbsolute, IgnoreRowOfs, CellRow, RowOfs);
			return WrapRow(Row + CellRow + RowOfs, false);
		}

		internal override int GetCol(int Col, bool RefAbsolute, bool IgnoreColOfs, int CellCol, int ColOfs)
		{
			if (RefAbsolute) return base.GetCol(Col, RefAbsolute, IgnoreColOfs, CellCol, ColOfs);
			return WrapColumn(Col + CellCol + ColOfs, false);
		}

		internal override bool CanHaveRelativeOffsets
		{
			get
			{
				return true;
			}
		}
		#endregion
        
        #region Clone
        internal override TBaseParsedToken Clone()
        {
            return new TRef3dNToken(FId, FExternSheet, Row, Col, RowAbs, ColAbs, false);
        }
        #endregion

	}

	internal class TArea3dToken : TAreaToken
	{
		#region Variables
		internal int FExternSheet;
		#endregion

		#region Constructor
		internal TArea3dToken(ptg aId, int aExternSheet, int aRow1, int aCol1, bool aRowAbs1, bool aColAbs1,
			int aRow2, int aCol2, bool aRowAbs2, bool aColAbs2)
			: base(aId, aRow1, aCol1, aRowAbs1, aColAbs1, aRow2, aCol2, aRowAbs2, aColAbs2)
		{
			FExternSheet = aExternSheet;
		}
		#endregion

		#region Evaluate
		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			if (IsErr()) return TFlxFormulaErrorValue.ErrRef;

			string ExternalBook; int Sheet1; int Sheet2;
			object res = ParseExternSheet(wi.Xls, out ExternalBook, out Sheet1, out Sheet2);
			if (res != null) return res;

			int ro1 = GetRow(Row1, RowAbs1, RowAbs1 || wi.SheetIndexBase1 != Sheet1 || Sheet1 != Sheet2 || ExternalBook != null, wi);
			int co1 = GetCol(Col1, ColAbs1, ColAbs1 || wi.SheetIndexBase1 != Sheet1 || Sheet1 != Sheet2 || ExternalBook != null, wi);
			int ro2 = GetRow(Row2, RowAbs2, RowAbs2 || wi.SheetIndexBase1 != Sheet1 || Sheet1 != Sheet2 || ExternalBook != null, wi);
			int co2 = GetCol(Col2, ColAbs2, ColAbs2 || wi.SheetIndexBase1 != Sheet1 || Sheet1 != Sheet2 || ExternalBook != null, wi);

			TWorkbookInfo wi2 = ExternalBook == null? wi: wi.ShallowClone();
			if (ExternalBook != null)
			{
				wi2.Xls = wi.Xls.GetSupportingFile(ExternalBook);
				if (wi2.Xls == null) return TFlxFormulaErrorValue.ErrRef;
			}

			if (!CheckRef(wi2, Sheet1, ro1, co1))
				return TFlxFormulaErrorValue.ErrRef;
			if (!CheckRef(wi2, Sheet2, ro2, co2))
				return TFlxFormulaErrorValue.ErrRef;

			return f.Agg(wi2, Sheet1, Sheet2, ro1 + 1, co1 + 1, ro2 + 1, co2 + 1, CalcState, CalcStack);
		}

		internal override object EvaluateRef(TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack)
		{
			if (IsErr()) return TFlxFormulaErrorValue.ErrRef;

			string ExternalBook; int Sheet1; int Sheet2;
			object res = ParseExternSheet(wi.Xls, out ExternalBook, out Sheet1, out Sheet2);
			if (res != null) return res;

			int ro1 = GetRow(Row1, RowAbs1, RowAbs1 || wi.SheetIndexBase1 != Sheet1 || Sheet1 != Sheet2 || ExternalBook != null, wi);
			int co1 = GetCol(Col1, ColAbs1, ColAbs1 || wi.SheetIndexBase1 != Sheet1 || Sheet1 != Sheet2 || ExternalBook != null, wi);
			int ro2 = GetRow(Row2, RowAbs2, RowAbs2 || wi.SheetIndexBase1 != Sheet1 || Sheet1 != Sheet2 || ExternalBook != null, wi);
			int co2 = GetCol(Col2, ColAbs2, ColAbs2 || wi.SheetIndexBase1 != Sheet1 || Sheet1 != Sheet2 || ExternalBook != null, wi);

			TWorkbookInfo wi2 = ExternalBook == null? wi: wi.ShallowClone();
			if (ExternalBook != null)
			{
				wi2.Xls = wi.Xls.GetSupportingFile(ExternalBook);
				if (wi2.Xls == null) return TFlxFormulaErrorValue.ErrRef;
			}

			if (!CheckRef(wi2, Sheet1, ro1, co2))
				return TFlxFormulaErrorValue.ErrRef;
			if (!CheckRef(wi2, Sheet2, ro2, co2))
				return TFlxFormulaErrorValue.ErrRef;

			TAddress[] Result = new TAddress[2];
			Result[0] = new TAddress(wi2, ExternalBook, Sheet1, ro1 + 1, co1 + 1);
			Result[1] = new TAddress(wi2, ExternalBook, Sheet2, ro2 + 1, co2 + 1);
			return Result;
		}

		internal override bool Same(TBaseParsedToken aBaseParsedToken)
		{
			if (!base.Same(aBaseParsedToken)) return false;
			TArea3dToken tk = aBaseParsedToken as TArea3dToken;
			if (tk == null) return false;
			if (tk.FExternSheet != FExternSheet) return false;
			return true;
		}

		#endregion

		#region Rows and Cols

		internal override int ExternSheet
		{
			get
			{
				return FExternSheet;
			}
			set
			{
				FExternSheet = value;
			}
		}
    
		internal override void CreateInvalidRef()
		{
			if (IsErr()) return;
			FId += ptg.Area3dErr - ptg.Area3d; //there is no RefNErr. We will change it to RefErr
		}
		#endregion

        #region Clone
        internal override TBaseParsedToken Clone()
        {
            return new TArea3dToken(FId, ExternSheet, Row1, Col1, RowAbs1, ColAbs1, Row2, Col2, RowAbs2, ColAbs2);
        }
        #endregion

	}

	/// <summary>
	/// This class doesn't really exist in Excel, but named ranges use this behavior, even if they use TRef3dTokens.
	/// </summary>
	internal sealed class TArea3dNToken : TArea3dToken
	{
		#region Constructor
		internal TArea3dNToken(ptg aId, int aExternSheet, int aRow1, int aCol1, bool aRowAbs1, bool aColAbs1, int aRow2, int aCol2, bool aRowAbs2, bool aColAbs2, bool FromBiff8)
            : base(aId, aExternSheet, WrapRowSigned(aRow1, aRowAbs1, FromBiff8), WrapColumnSigned(aCol1, aColAbs1, FromBiff8), aRowAbs1, aColAbs1,
            WrapRowSigned(aRow2, aRowAbs2, FromBiff8), WrapColumnSigned(aCol2, aColAbs2, FromBiff8), aRowAbs2, aColAbs2)
		{ }
		#endregion

		#region Rows and Cols
		internal override int GetRow(int Row, bool RefAbsolute, bool IgnoreRowOfs, int CellRow, int RowOfs)
		{
			//Relative references wrap.
			if (RefAbsolute) return base.GetRow(Row, RefAbsolute, IgnoreRowOfs, CellRow, RowOfs);
			return WrapRow(Row + CellRow + RowOfs, false);
		}

		internal override int GetCol(int Col, bool RefAbsolute, bool IgnoreColOfs, int CellCol, int ColOfs)
		{
			if (RefAbsolute) return base.GetCol(Col, RefAbsolute, IgnoreColOfs, CellCol, ColOfs);
			return WrapColumn(Col + CellCol + ColOfs, false);
		}

		internal override bool CanHaveRelativeOffsets
		{
			get
			{
				return true;
			}
		}
		#endregion

        #region Clone
        internal override TBaseParsedToken Clone()
        {
            return new TArea3dNToken(FId, ExternSheet, Row1, Col1, RowAbs1, ColAbs1, Row2, Col2, RowAbs2, ColAbs2, false);
        }
        #endregion

	}

	#endregion

	#region Reference Operators
	internal abstract class TBaseRefOpToken : TStaticToken
	{
		internal TBaseRefOpToken()
			: base(2)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			object v1 = EvaluateRef(FTokenList, wi, CalcState, CalcStack); if (v1 is TFlxFormulaErrorValue) return v1;

			TAddress[] adr = v1 as TAddress[];
			if (adr != null)
			{
				return f.Agg(adr[0].wi, adr[0].Sheet, adr[0].Sheet, adr[0].Row, adr[0].Col, adr[1].Row, adr[1].Col, CalcState, CalcStack);
			}

			TAddressList adrl = v1 as TAddressList;
			if (adrl != null)
			{
				object Result = null;

				for (int i = 0; i < adrl.Count; i++)
				{
					if (Result is TFlxFormulaErrorValue) return Result;
					adr = adrl[i];
					if (adr[0].wi.Xls != adr[1].wi.Xls) return TFlxFormulaErrorValue.ErrRef;
					Result = f.AggValues(Result, f.Agg(adr[0].wi, adr[0].Sheet, adr[1].Sheet, adr[0].Row, adr[0].Col, adr[1].Row, adr[1].Col, CalcState, CalcStack));
				}
				return Result;
			}


			return TFlxFormulaErrorValue.ErrRef;
		}

		protected abstract object ManageRange(TAddressList adr1, TAddressList adr2);

		internal override object EvaluateRef(TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack)
		{
			TAddressList adr1 = null;
			object[,] values = null;
			object ret = GetRangeOrArray(FTokenList, wi, CalcState, CalcStack, out adr1, out values, false);
			if (ret != null) return ret;

			TAddressList adr2 = null;
			ret = GetRangeOrArray(FTokenList, wi, CalcState, CalcStack, out adr2, out values, false);
			if (ret != null) return ret;

			return ManageRange(adr1, adr2);
		}
	}

	internal class TISectToken : TBaseRefOpToken
	{
		private TISectToken() : base() { }
		internal static readonly TISectToken Instance = new TISectToken();

		internal override ptg GetId { get { return ptg.Isect; } }

		protected override object ManageRange(TAddressList adr1, TAddressList adr2)
		{
            List<TAddress[]> ResultValue = new List<TAddress[]>();
			for (int i = 0; i < adr1.Count; i++)
				for (int k = 0; k < adr2.Count; k++)
				{
					TAddress[] o = IntersectRanges(adr1[i], adr2[k]);
					if (o != null)
						ResultValue.Add(o);
				}
			if (ResultValue.Count == 0) return TFlxFormulaErrorValue.ErrNull;
			if (ResultValue.Count == 1) return (TAddress[])ResultValue[0];
			return new TAddressList(ResultValue);
		}

		protected static TAddress[] IntersectRanges(TAddress[] c1, TAddress[] c2)
		{
			if (c1[0].Sheet != c1[1].Sheet) return null;
			if (c2[0].Sheet != c1[1].Sheet) return null;
			if (c1[0].Sheet != c2[0].Sheet) return null;

			if (c1[0].BookName != c1[1].BookName) return null;
			if (c2[0].BookName != c1[1].BookName) return null;
			if (c1[0].BookName != c2[0].BookName) return null;

			TAddress[] Result =
				new TAddress[]
			{
				new TAddress (
				c1[0].wi,
				c1[0].BookName,
				c1[0].Sheet, 
				Math.Max(
				Math.Min(c1[0].Row, c1[1].Row),
				Math.Min(c2[0].Row, c2[1].Row)),
				Math.Max(
				Math.Min(c1[0].Col, c1[1].Col),
				Math.Min(c2[0].Col, c2[1].Col))),
								
				new TAddress (
				c1[0].wi,
				c1[0].BookName,
				c1[0].Sheet, 
				Math.Min(
				Math.Max(c1[0].Row, c1[1].Row),
				Math.Max(c2[0].Row, c2[1].Row)),
				Math.Min(
				Math.Max(c1[0].Col, c1[1].Col),
				Math.Max(c2[0].Col, c2[1].Col)))
			};

			if (Result[0].Col > Result[1].Col || Result[0].Row > Result[1].Row) return null;
			return Result;
		}
	}

	internal class TRangeToken : TBaseRefOpToken
	{
		private TRangeToken() : base() { }
		internal static readonly TRangeToken Instance = new TRangeToken();

		internal override ptg GetId { get { return ptg.Range; } }

		protected override object ManageRange(TAddressList adr1, TAddressList adr2)
		{
			TAddress[] Result = adr1[0];
			for (int i = 1; i < adr1.Count; i++)
				Result = ExpandRanges(Result, adr1[i]);
			for (int i = 0; i < adr2.Count; i++)
				Result = ExpandRanges(Result, adr2[i]);

			return Result;
		}

		protected static TAddress[] ExpandRanges(TAddress[] c1, TAddress[] c2)
		{
			if (c1[0].Sheet != c1[1].Sheet) return null;
			if (c2[0].Sheet != c1[1].Sheet) return null;
			if (c1[0].Sheet != c2[0].Sheet) return null;

			if (c1[0].BookName != c1[1].BookName) return null;
			if (c2[0].BookName != c1[1].BookName) return null;
			if (c1[0].BookName != c2[0].BookName) return null;

			TAddress[] Result =
				new TAddress[]
			{
				new TAddress (
				c1[0].wi,
				c1[0].BookName,
				c1[0].Sheet, 
				Math.Min(
				Math.Min(c1[0].Row, c1[1].Row),
				Math.Min(c2[0].Row, c2[1].Row)),
				Math.Min(
				Math.Min(c1[0].Col, c1[1].Col),
				Math.Min(c2[0].Col, c2[1].Col))),
								
				new TAddress (
				c1[0].wi,
				c1[0].BookName,
				c1[0].Sheet, 
				Math.Max(
				Math.Max(c1[0].Row, c1[1].Row),
				Math.Max(c2[0].Row, c2[1].Row)),
				Math.Max(
				Math.Max(c1[0].Col, c1[1].Col),
				Math.Max(c2[0].Col, c2[1].Col)))
			};

			if (Result[0].Col > Result[1].Col || Result[0].Row > Result[1].Row) return null;
			return Result;
		}
	}

	internal class TUnionToken : TStaticToken
	{
		private TUnionToken() : base(2) { }
		internal static readonly TUnionToken Instance = new TUnionToken();
        
		internal override ptg GetId { get { return ptg.Union; } }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			object v1 = FTokenList.EvaluateToken(wi, f, CalcState, CalcStack); if (v1 is TFlxFormulaErrorValue) return v1;
			object v2 = FTokenList.EvaluateToken(wi, f, CalcState, CalcStack); if (v2 is TFlxFormulaErrorValue) return v2;

			return f.AggValues(v1, v2);
		}

		internal override object EvaluateRef(TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack)
		{
			TAddressList adr1 = null; object[,] values = null;
			object ret = GetRangeOrArray(FTokenList, wi, CalcState, CalcStack, out adr1, out values, false);
			if (ret != null) return ret;

			TAddressList adr2 = null;
			ret = GetRangeOrArray(FTokenList, wi, CalcState, CalcStack, out adr2, out values, false);
			if (ret != null) return ret;

			adr1.Add(adr2);
			return adr1;
		}

	}
	#endregion

	#region Unsupported Token
	internal sealed class TUnsupportedToken : TBaseParsedToken
	{
		private ptg FId;

		internal TUnsupportedToken(int ArgCount, ptg aId)
			: base(ArgCount)
		{
			FId = aId;
		}

		internal override ptg GetId
		{
			get { return FId; }
		}

		internal override TBaseParsedToken SetId(ptg aId)
		{
			FId = aId;
			return this;
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			wi.AddUnsupported(TUnsupportedFormulaErrorType.FormulaTooComplex, String.Empty);

			return TFlxFormulaErrorValue.ErrNA;
		}

		internal override bool Same(TBaseParsedToken aBaseParsedToken)
		{
			if (!base.Same(aBaseParsedToken)) return false;
			TUnsupportedToken tk = this as TUnsupportedToken;
			return (tk.FArgCount == FArgCount && tk.FId == FId);
		}
	}
	#endregion

	#region Function Tokens
    
	#region Base
	internal abstract class TBaseFunctionToken : TBaseParsedToken
	{
		private ptg FId;
		protected TCellFunctionData FuncData;

		protected TBaseFunctionToken(int ArgCount, ptg aId, TCellFunctionData aFuncData) : base(ArgCount)
		{
			FId = aId;
			FuncData = aFuncData;
			FunctionCache.Add(aFuncData.Index, aId, ArgCount, this);
		}

		internal override ptg GetId
		{
			get { return FId; }
		}

		internal override TBaseParsedToken SetId(ptg aId)
		{
			if (aId == FId) return this;

			TBaseFunctionToken NewFunction;
			if (FunctionCache.TryGetValue(FuncData.Index, aId, FArgCount, out NewFunction)) return NewFunction;

			NewFunction = (TBaseFunctionToken) MemberwiseClone();
			NewFunction.FId = aId;
			FunctionCache.Add(FuncData.Index, aId, FArgCount, NewFunction); //need to add, because memberwiseclone doesn't go through the constructor.

			return NewFunction;
		}

		/// <summary>
		/// By default all functions can be cloned freely. No need to copy the token, since the class is inmutable.
		/// </summary>
		/// <returns></returns>
		internal override TBaseParsedToken Clone()
		{
			return this;
		}


		internal TCellFunctionData GetFunctionData()
		{
			return FuncData;
		}

		internal override bool Same(TBaseParsedToken aBaseParsedToken)
		{
			TBaseFunctionToken v2 = aBaseParsedToken as TBaseFunctionToken;
			if (v2 == null) return false;
			if (v2.FId != FId || v2.FArgCount != FArgCount || v2.FuncData != FuncData) return false;
			return true;
		}
	}

	#endregion

	#region Generic tokens for common functions
	internal abstract class TOneDoubleArgToken : TBaseFunctionToken
	{
		protected TOneDoubleArgToken(ptg aId, TCellFunctionData aFuncData) : base(1, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			double d = 0;
			object[,] ArrF;
			object Ret = GetDoubleArgument(FTokenList, wi, CalcState, CalcStack, ref d, out ArrF); if (Ret != null) return Ret;

			if (ArrF != null)
			{
				Object[,] ArrResult = new object[ArrF.GetLength(0), ArrF.GetLength(1)];
				for (int i = 0; i < ArrResult.GetLength(0); i++)
					for (int j = 0; j < ArrResult.GetLength(1); j++)
					{
						if (!GetDoubleItem(ref ArrResult, ref ArrF, i, j, out d)) continue;
						try
						{
							ArrResult[i, j] = Calc(d);
						}
						catch (FlexCelException)
						{
							ArrResult[i, j] = TFlxFormulaErrorValue.ErrNA;
						}
						catch (FormatException)
						{
							ArrResult[i, j] = TFlxFormulaErrorValue.ErrValue;
						}
						catch (OverflowException)
						{
							ArrResult[i, j] = TFlxFormulaErrorValue.ErrNum;
						}
						catch (ArithmeticException)
						{
							ArrResult[i, j] = TFlxFormulaErrorValue.ErrValue;
						}
						catch (ArgumentOutOfRangeException)
						{
							ArrResult[i, j] = TFlxFormulaErrorValue.ErrNum;
						}
					}

				return UnPack(ArrResult);
			}


			return Calc(d);
		}

		protected abstract object Calc(double x);
	}

	internal abstract class TNDoubleArgToken : TBaseFunctionToken
	{
		protected TNDoubleArgToken(int aArgCount, ptg aId, TCellFunctionData aFuncData) : base(aArgCount, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			double[] d = new double[FArgCount];
			object[][,] Arr = new object[FArgCount][,];
			bool HasArr = false;
			for (int i = 0; i < FArgCount; i++)
			{
				object Ret = GetDoubleArgument(FTokenList, wi, CalcState, CalcStack, ref d[i], out Arr[i]); if (Ret != null) return Ret;
				if (Arr[i] != null) HasArr = true;
			}

			if (HasArr)
			{
				int RowCount = 1; int ColCount = 1;
				for (int i = 0; i < FArgCount - 1; i++)
				{
					SetupArray(ref Arr[i], ref Arr[i + 1], d[i], d[i + 1], ref RowCount, ref ColCount);
				}
				object[,] ArrResult = new object[RowCount, ColCount];

				for (int i = 0; i < ArrResult.GetLength(0); i++)
					for (int j = 0; j < ArrResult.GetLength(1); j++)
					{
						bool Ok = true;
						for (int n = 0; n < FArgCount; n++)
						{
							if (!GetDoubleItem(ref ArrResult, ref Arr[n], i, j, out d[n])) { Ok = false; break; }
						}
						if (!Ok) continue;

						try
						{
							ArrResult[i, j] = Calc(d);
						}
						catch (FlexCelException)
						{
							ArrResult[i, j] = TFlxFormulaErrorValue.ErrNA;
						}
						catch (FormatException)
						{
							ArrResult[i, j] = TFlxFormulaErrorValue.ErrValue;
						}
						catch (OverflowException)
						{
							ArrResult[i, j] = TFlxFormulaErrorValue.ErrNum;
						}
						catch (ArithmeticException)
						{
							ArrResult[i, j] = TFlxFormulaErrorValue.ErrValue;
						}
						catch (ArgumentOutOfRangeException)
						{
							ArrResult[i, j] = TFlxFormulaErrorValue.ErrNum;
						}
					}

				return UnPack(ArrResult);
			}


			return Calc(d);
		}

		protected abstract object Calc(double[] x);
	}

	internal abstract class TOneStringArgToken : TBaseFunctionToken
	{
		protected TOneStringArgToken(ptg aId, TCellFunctionData aFuncData) : base(1, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			string s = string.Empty;
			object[,] ArrF;
			object Ret = GetStringArgument(FTokenList, wi, CalcState, CalcStack, ref s, out ArrF); if (Ret != null) return Ret;

			if (ArrF != null)
			{
				Object[,] ArrResult = new object[ArrF.GetLength(0), ArrF.GetLength(1)];
				for (int i = 0; i < ArrResult.GetLength(0); i++)
					for (int j = 0; j < ArrResult.GetLength(1); j++)
					{
						if (!GetStringItem(ref ArrResult, ref ArrF, i, j, out s)) continue;
						try
						{
							ArrResult[i, j] = Calc(s);
						}
						catch (FlexCelException)
						{
							ArrResult[i, j] = TFlxFormulaErrorValue.ErrNA;
						}
						catch (FormatException)
						{
							ArrResult[i, j] = TFlxFormulaErrorValue.ErrValue;
						}
						catch (OverflowException)
						{
							ArrResult[i, j] = TFlxFormulaErrorValue.ErrNum;
						}
						catch (ArithmeticException)
						{
							ArrResult[i, j] = TFlxFormulaErrorValue.ErrValue;
						}
						catch (ArgumentOutOfRangeException)
						{
							ArrResult[i, j] = TFlxFormulaErrorValue.ErrNum;
						}
					}
				return UnPack(ArrResult);
			}

			return Calc(s);
		}

		protected abstract object Calc(string s);
	}

	#endregion

	#region Unsupported function
	internal sealed class TUnsupportedFunction : TBaseFunctionToken
	{
		internal TUnsupportedFunction(int aArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(aArgCount, aId, aFuncData)
		{
		}

        internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
        {
            string FuncName = FuncData.Name != null ? FuncData.Name : "Unknown Function";
            wi.AddUnsupported(TUnsupportedFormulaErrorType.MissingFunction, FuncName);

            return TFlxFormulaErrorValue.ErrNA;
        }
	}
	#endregion

	#region Aggregated functions
	internal abstract class TRangeParsedToken : TBaseFunctionToken
	{
		protected int FStartArg;
		protected bool CountAnything;
		protected bool IgnoreMissingArg;
        
		internal TRangeParsedToken(int ArgCount, ptg aId, TCellFunctionData aFuncData, int StartArg, bool aCountAnything, bool aIgnoreMissingArg)
			: base(ArgCount, aId, aFuncData)
		{
			FStartArg = StartArg;
			CountAnything = aCountAnything;
			IgnoreMissingArg = aIgnoreMissingArg;
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			bool First = true; //Can hold an extra argument for example for the min function.
			double Result = 0;
			for (int i = FStartArg; i < FArgCount; i++)
			{
				object v1 = FTokenList.EvaluateToken(wi, f, CalcState, CalcStack);
				object[,] arr1 = v1 as object[,];
				if (arr1 != null)
					v1 = f.AggArray(arr1);

				/* 
				 if (v1 is string)
					 return TFlxFormulaErrorValue.ErrValue;

				 v1 = ConvertToAllowedObject(v1);

				 if (v1 is TFlxFormulaErrorValue) return v1;
				 if (v1 is bool) // && CountAnything) Here bools are always converted to int, even on not "A" functions. Max(true, false) = 1
					 if ((bool)v1) v1=1.0; else v1=0.0;

				 if (v1 is double)
				 {
					 EvalOne ((double)v1, ref Result);
				 } 
				 */

				if (v1 is TFlxFormulaErrorValue) return v1;

				if (v1 != null && (!IgnoreMissingArg || !(v1 is TMissingArg)))
				{
					double dResult;
					if (ExtToDouble(v1, out dResult))
					{
						EvalOne(dResult, ref Result, First, f);
						First = false;
					}
					else return TFlxFormulaErrorValue.ErrValue;
				}

			}
			return Result;
		}

		internal object EvaluateAvg(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack, bool ReturnCount)
		{
			object Result = EvaluateAvg2(FTokenList, wi, f, CalcState, CalcStack, ReturnCount, false);
			if (Result is TFlxFormulaErrorValue) return Result;

			TAverageValue Avg = Result as TAverageValue;
			if (Avg == null) return TFlxFormulaErrorValue.ErrNA;

			if (ReturnCount)
				return (double)Avg.ValueCount;

			if (Avg.ValueCount == 0) return TFlxFormulaErrorValue.ErrDiv0;
			else return Avg.Sum / (double)Avg.ValueCount;

		}

		internal object EvaluateAvg2(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack, bool IgnoreErrs, bool Multiply)
		{
            return EvaluateAvg2St(FArgCount, FStartArg, IgnoreMissingArg, CountAnything, FTokenList, wi, f, CalcState, CalcStack, IgnoreErrs, Multiply);
		}

        internal static object EvaluateAvg2St(int FArgCount, int FStartArg, bool IgnoreMissingArg, bool CountAnything, TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack, bool IgnoreErrs, bool Multiply)
        {
            long ValueCount = 0;

            double ResultValue = 0;
            for (int i = FStartArg; i < FArgCount; i++)
            {
                object v1 = FTokenList.EvaluateToken(wi, f, CalcState, CalcStack);
                if (!IgnoreErrs && v1 is TFlxFormulaErrorValue) return v1;

                object[,] arr1 = v1 as object[,];
                if (arr1 != null)
                    v1 = f.AggArray(arr1);

                TAverageValue av = v1 as TAverageValue;
                if (av != null)
                {
                    if (Multiply && ValueCount > 0)
                    {
                        if (av.ValueCount > 0) ResultValue *= av.Sum;
                    }
                    else
                    {
                        ResultValue += av.Sum;
                    }
                    ValueCount += av.ValueCount;
                }
                else
                {
                    if (v1 is TMissingArg && !IgnoreMissingArg) v1 = 0.0;
                    v1 = ConvertToAllowedObject(v1);

                    if (CountAnything)
                    {
                        if (v1 is bool)
                            if ((bool)v1) v1 = 1.0; else v1 = 0.0;
                        else if (v1 != null && !(v1 is double)) //this includes errors.
                            v1 = 0.0;
                    }

                    string s1 = v1 as string;
                    if (s1 != null && s1.Length > 0)
                    {
                        double dResult = 0;
                        if (TCompactFramework.ConvertToNumber(s1, CultureInfo.CurrentCulture, out dResult)) v1 = dResult;  //Try to avoid unnecessary exceptions.
                        else
                        {
                            DateTime DateResult;
                            if (TCompactFramework.ConvertDateToNumber(s1, out DateResult))
                            {
                                v1 = FlxDateTime.ToOADate(DateResult, Dates1904);
                            }
                        }
                    }
                    if (v1 is double)
                    {
                        if (Multiply && (double)v1 <= 0) return TFlxFormulaErrorValue.ErrNum; //GeoMean only uses positives.

                        if (Multiply && ValueCount > 0)
                        {
                            ResultValue *= (double)v1;
                        }
                        else
                        {
                            ResultValue += (double)v1;
                        }
                        ValueCount++;
                    }

                }
            }

            return new TAverageValue(ResultValue, ValueCount);
        }

		internal virtual void EvalOne(double d, ref double ResultValue, bool First, TBaseAggregate f)
		{
		}
	}

	internal class TAverageValue
	{
		internal double Sum;
		internal long ValueCount; //1 million rows * 16000 cols can overflow 32 bits.

		internal bool HasErr;
		internal TFlxFormulaErrorValue Err;

		internal TAverageValue(double aSum, long aValueCount)
		{
			Sum = aSum;
			ValueCount = aValueCount;
			HasErr = false;
		}

        internal TAverageValue Add(TAverageValue v)
        {
            return new TAverageValue(Sum + v.Sum, ValueCount + v.ValueCount);
        }

        internal TAverageValue Mult(TAverageValue v)
        {
            if (ValueCount == 0) return v;
            if (v.ValueCount == 0) return this;
            return new TAverageValue(Sum * v.Sum, ValueCount + v.ValueCount);
        }

        internal TAverageValue AddSquaredDiff(TAverageValue v, double Avg)
        {
            return new TAverageValue(Sum + v.Sum, ValueCount + v.ValueCount); //this is already aggregated
        }

        internal TAverageValue AddNSquaredDiff(TAverageValue v, double Avg, int N)
        {
            return new TAverageValue(Sum + v.Sum, ValueCount + v.ValueCount); //this is already aggregated
        }

        internal TAverageValue AddModDiff(TAverageValue v, double Avg)
        {
            return new TAverageValue(Sum + v.Sum, ValueCount + v.ValueCount); //this is already aggregated
        }

	}

	internal sealed class TAverageToken : TRangeParsedToken
	{
		internal TAverageToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData, 0, false, false)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			return EvaluateAvg(FTokenList, wi, TAverageAggregate.Instance0, CalcState, CalcStack, false);
		}

	}
	internal sealed class TAverageAToken : TRangeParsedToken
	{
		internal TAverageAToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData, 0, true, false)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			return EvaluateAvg(FTokenList, wi, TAverageAggregate.InstanceA, CalcState, CalcStack, false);
		}

	}

	internal sealed class TGeoMeanToken : TRangeParsedToken
	{
		internal TGeoMeanToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData, 0, false, false)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			object Result = EvaluateAvg2(FTokenList, wi, TGeoMeanAggregate.Instance, CalcState, CalcStack, false, true);
			if (Result is TFlxFormulaErrorValue) return Result;

			TAverageValue Avg = Result as TAverageValue;
			if (Avg == null) return TFlxFormulaErrorValue.ErrNA;

			if (Avg.ValueCount == 0) return TFlxFormulaErrorValue.ErrNum;
			else return Math.Pow(Avg.Sum, 1 / (double)Avg.ValueCount);
		}

	}

	internal sealed class THarMeanToken : TRangeParsedToken
	{
		internal THarMeanToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData, 0, false, false)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			object Result = EvaluateAvg3(FTokenList, wi, THarMeanAggregate.Instance, CalcState, CalcStack);
			if (Result is TFlxFormulaErrorValue) return Result;

			TAverageValue Avg = Result as TAverageValue;
			if (Avg == null) return TFlxFormulaErrorValue.ErrNA;

			if (Avg.ValueCount == 0) return TFlxFormulaErrorValue.ErrNA;
			double Hi = Avg.Sum / (double)Avg.ValueCount;
			if (Hi == 0) return TFlxFormulaErrorValue.ErrNum;
			return 1 / Hi;
		}

		internal object EvaluateAvg3(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			long ValueCount = 0;

			double ResultValue = 0;
			for (int i = FStartArg; i < FArgCount; i++)
			{
				object v1 = FTokenList.EvaluateToken(wi, f, CalcState, CalcStack);
				if (v1 is TFlxFormulaErrorValue) return v1;

				object[,] arr1 = v1 as object[,];
				if (arr1 != null)
					v1 = f.AggArray(arr1);

				TAverageValue av = v1 as TAverageValue;
				if (av != null)
				{
					ResultValue += av.Sum;
					ValueCount += av.ValueCount;
				}
				else
				{
					v1 = ConvertToAllowedObject(v1);

					if (CountAnything)
					{
						if (v1 is bool)
							if ((bool)v1) v1 = 1.0; else v1 = 0.0;
						else if (v1 != null && !(v1 is double)) //this includes errors.
							v1 = 0.0;
					}

					if (v1 is string)
					{
						double d;
						if (ExtToDouble(v1, out d)) v1 = d;
					}

					if (v1 is double)
					{
						if ((double)v1 <= 0) return TFlxFormulaErrorValue.ErrNum; // only uses positives.
						ResultValue += 1 / (double)v1;
						ValueCount++;
					}

				}
			}

			return new TAverageValue(ResultValue, ValueCount);
		}
	}

	internal class TSumToken : TRangeParsedToken
	{
		internal TSumToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData, 0, false, true)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			return base.Evaluate(FTokenList, wi, TSumAggregate.Instance, CalcState, CalcStack);
		}

		internal override void EvalOne(double d, ref double ResultValue, bool First, TBaseAggregate f)
		{
			ResultValue += d;
		}
	}

	internal class TSumSqToken : TRangeParsedToken
	{
		internal TSumSqToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData, 0, false, true)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			return base.Evaluate(FTokenList, wi, TSumSqAggregate.Instance, CalcState, CalcStack);
		}

		internal override void EvalOne(double d, ref double ResultValue, bool First, TBaseAggregate f)
		{
			ResultValue += d * d;
		}
	}

	internal class TProductToken : TRangeParsedToken
	{
		internal TProductToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData, 0, false, true)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			return base.Evaluate(FTokenList, wi, TProductAggregate.Instance, CalcState, CalcStack);
		}

		internal override void EvalOne(double d, ref double ResultValue, bool First, TBaseAggregate f)
		{
			if (!First) ResultValue *= d; else ResultValue = d;
		}
	}

	internal class TCountToken : TRangeParsedToken
	{
		internal TCountToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData, 0, false, false)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			return EvaluateAvg(FTokenList, wi, TCountAggregate.Instance, CalcState, CalcStack, true);
		}
	}

	internal class TCountAToken : TRangeParsedToken
	{
		internal TCountAToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData, 0, true, false)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			return EvaluateAvg(FTokenList, wi, TCountAAggregate.InstanceAll, CalcState, CalcStack, true);
		}
	}

	internal class TCountBlankToken : TRangeParsedToken
	{
		internal TCountBlankToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData, 0, true, false)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			return EvaluateAvg(FTokenList, wi, TCountAAggregate.InstanceBlank, CalcState, CalcStack, true);
		}
	}

	internal sealed class TMinToken : TRangeParsedToken
	{
		internal TMinToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData, 0, false, false)
		{
		}

		internal override void EvalOne(double d, ref double ResultValue, bool First, TBaseAggregate f)
		{
			if (First || (d < ResultValue)) ResultValue = d;
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			return base.Evaluate(FTokenList, wi, TMinAggregate.Instance0, CalcState, CalcStack);
		}
	}

	internal sealed class TMinAToken : TRangeParsedToken
	{
		internal TMinAToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData, 0, true, false)
		{
		}

		internal override void EvalOne(double d, ref double ResultValue, bool First, TBaseAggregate f)
		{
			if (First || (d < ResultValue)) ResultValue = d;
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			return base.Evaluate(FTokenList, wi, TMinAggregate.InstanceA, CalcState, CalcStack);
		}
	}

	internal sealed class TMaxToken : TRangeParsedToken
	{
		internal TMaxToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData, 0, false, false)
		{
		}

		internal override void EvalOne(double d, ref double ResultValue, bool First, TBaseAggregate f)
		{
			if (First || (d > ResultValue)) ResultValue = d;
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			return base.Evaluate(FTokenList, wi, TMaxAggregate.Instance0, CalcState, CalcStack);
		}
	}

	internal sealed class TMaxAToken : TRangeParsedToken
	{
		internal TMaxAToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData, 0, true, false)
		{
		}

		internal override void EvalOne(double d, ref double ResultValue, bool First, TBaseAggregate f)
		{
			if (First || (d > ResultValue)) ResultValue = d;
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			return base.Evaluate(FTokenList, wi, TMaxAggregate.InstanceA, CalcState, CalcStack);
		}
	}
	#endregion

	#region SumIf and CountIf
	internal enum TCriteriaType
	{
		EQ,
		NE,
		GT,
		GE,
		LT,
		LE
	}

    internal abstract class TBaseRangeIfToken : TBaseFunctionToken
    {
        internal TBaseRangeIfToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
            : base(ArgCount, aId, aFuncData)
        {
        }

        protected static object GetSumRange(TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, ref TCalcStack CalcStack, ref TAddressList SumRange)
        {
            object[,] values;
            object ret = GetRangeOrArray(FTokenList, wi, CalcState, CalcStack, out SumRange, out values, false);
            if (ret != null) return ret;
            if (SumRange.Has3dRef() || SumRange.Count != 1) return TFlxFormulaErrorValue.ErrValue;
            return null;
        }

        protected static TValueCriteria[,] GetCriteria(TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack)
        {
            TValueCriteria[,] CriteriaEval = null;
            object Criteria = FTokenList.EvaluateToken(wi, CalcState, CalcStack); //Not here. if (Criteria is TFlxFormulaErrorValue) return Criteria;

            object[,] ArrCriteria;
            if (IsArrayArgument(Criteria, out ArrCriteria))
            {
                CriteriaEval = new TValueCriteria[ArrCriteria.GetLength(0), ArrCriteria.GetLength(1)];
                for (int i = 0; i < ArrCriteria.GetLength(0); i++)
                {
                    for (int j = 0; j < ArrCriteria.GetLength(1); j++)
                    {
                        object CriteriaElem = ConvertToAllowedObject(ArrCriteria[i, j]);
                        CriteriaEval[i, j] = TValueCriteria.Create(CriteriaElem, -1, true);  //Column is not important here, since we are not going to evaluate.
                    }
                }

                return CriteriaEval;
            }

            Criteria = ConvertToAllowedObject(Criteria);
            CriteriaEval = new TValueCriteria[1, 1];
            CriteriaEval[0, 0] = TValueCriteria.Create(Criteria, -1, true);  //Column is not important here, since we are not going to evaluate.

            return CriteriaEval;
        }

        protected bool MeetsAnyCriteria(TCriteriaAndAddress[] CriteriaEval, object v1, int rr, int cc)
        {
            foreach (TCriteriaAndAddress CriteriaAddr in CriteriaEval)
            {
                if (CriteriaAddr.Criteria.GetLength(0) == 1 && CriteriaAddr.Criteria.GetLength(1) == 1)
                {
                    if (!CriteriaAddr.Criteria[0, 0].MeetsCriteria(v1)) return false;
                }
                else
                {
                    if (rr >= CriteriaAddr.Criteria.GetLength(0)) return false;
                    if (cc >= CriteriaAddr.Criteria.GetLength(1)) return false;
                    if (!CriteriaAddr.Criteria[rr, cc].MeetsCriteria(v1)) return false;
                }
            }

            return true;
        }

        protected abstract void Process(object v, ref TAverageValue ResultValue);
        protected abstract object GetFinalValue(TAverageValue Result);

    }

	internal abstract class TRangeIfToken : TBaseRangeIfToken
	{
		internal TRangeIfToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			TAddressList SumRange = null;

			if (FArgCount > 2)
			{
                object res = GetSumRange(FTokenList, wi, CalcState, ref CalcStack, ref SumRange);
                if (res != null) return res;
			}

			object Criteria = FTokenList.EvaluateToken(wi, CalcState, CalcStack); //Not here. if (Criteria is TFlxFormulaErrorValue) return Criteria;

			TAddressList Range = null;
            object[,] values;
            object ret2 = GetRangeOrArray(FTokenList, wi, CalcState, CalcStack, out Range, out values, false);
			if (ret2 != null) return ret2;
			if (Range.Has3dRef() || Range.Count != 1) return TFlxFormulaErrorValue.ErrValue;

			if (SumRange == null) SumRange = Range;

			object[,] ArrResult;
			object[,] ArrCriteria;
			if (IsArrayArgument(Criteria, out ArrResult, out ArrCriteria))
			{
                for (int i = 0; i < ArrResult.GetLength(0); i++)
                {
                    for (int j = 0; j < ArrResult.GetLength(1); j++)
                    {
                        Criteria = ConvertToAllowedObject(ArrCriteria[i, j]);
                        ArrResult[i, j] = CalcValues(Range[0][0].wi.Xls, Range[0][0].Sheet, Range[0][0].Row, Range[0][0].Col, Range[0][1].Row, Range[0][1].Col,
                            SumRange[0][0].wi.Xls, SumRange[0][0].Sheet, SumRange[0][0].Row, SumRange[0][0].Col, Criteria, CalcState, CalcStack);
                    }
                }
				return UnPack(ArrResult);
			}

			Criteria = ConvertToAllowedObject(Criteria);
			return CalcValues(Range[0][0].wi.Xls, Range[0][0].Sheet, Range[0][0].Row, Range[0][0].Col, Range[0][1].Row, Range[0][1].Col,
				SumRange[0][0].wi.Xls, SumRange[0][0].Sheet, SumRange[0][0].Row, SumRange[0][0].Col, Criteria, CalcState, CalcStack);
		}

		internal object CalcValues(ExcelFile Xls1, int RSheet, int RRow1, int RCol1, int RRow2, int RCol2, ExcelFile SXls, int SSheet, int SRow1, int SCol1, object Criteria, TCalcState CalcState, TCalcStack CalcStack)
		{
			int x = 0;
			OrderRange(ref x, ref x, ref RRow1, ref RCol1, ref RRow2, ref RCol2);
            TAverageValue ResultValue = new TAverageValue(0, 0);

			TValueCriteria CriteriaEval = TValueCriteria.Create(Criteria, -1, true);  //Column is not important here, since we are not going to evaluate.


			bool SumRangeDifferent = Xls1 != SXls || RSheet != SSheet || RRow1 != SRow1 || RCol1 != SCol1;

			long ValuesProcessed = 0;
			int MaxRow = Xls1.GetRowCount(RSheet);

            ProcessRange(Xls1, RSheet, RRow1, RCol1, RRow2, RCol2, SXls, SSheet, SRow1, SCol1, ref CalcState, CalcStack, ref ResultValue, SumRangeDifferent, ref ValuesProcessed, MaxRow, CriteriaEval);

            ProcessFinal(ValuesProcessed, ((long)RRow2 - RRow1 + 1) * (RCol2 - RCol1 + 1), ref ResultValue, CriteriaEval);
			return GetFinalValue(ResultValue);
		}

        protected void ProcessRange(ExcelFile Xls1, int RSheet, int RRow1, int RCol1, int RRow2, int RCol2,
     ExcelFile SXls, int SSheet, int SRow1, int SCol1, ref TCalcState CalcState, TCalcStack CalcStack, ref TAverageValue Result,
     bool SumRangeDifferent, ref long ValuesProcessed, int MaxRow, TValueCriteria CriteriaEval)
        {
            bool NeedsAllCols = SumRangeDifferent && CriteriaEval.MeetsCriteria(null);

            for (int r = RRow1; r <= RRow2; r++)
            {
                if (r > MaxRow) break;

                for (int cIndex = Xls1.ColToIndex(RSheet, r, RCol2); cIndex > 0; cIndex--)
                {
                    int c = Xls1.ColFromIndex(RSheet, r, cIndex);
                    if (c > RCol2 || c == 0) continue;
                    if (c < RCol1) break;

                    ValuesProcessed++;
                    object v1 = ConvertToAllowedObject(Xls1.GetCellValueAndRecalc(RSheet, r, c, CalcState, CalcStack));
                    if (CalcState.Aborted) return;
                    object v2 = v1;
                    //if (SumRangeDifferent)  Do not convert here!. It can raise a circular error even when because of the criteria we will not use the value.
                    //    v2=ConvertToAllowedObject(Xls.GetCellValueAndRecalc(SSheet, SRow1+r-RRow1, SCol1+c-RCol1));

                    if (CriteriaEval.MeetsCriteria(v1))
                    {
                        if (SumRangeDifferent) v2 = ConvertToAllowedObject(SXls.GetCellValueAndRecalc(SSheet, SRow1 + r - RRow1, SCol1 + c - RCol1, CalcState, CalcStack));
                        if (CalcState.Aborted) return;
                        Process(v2, ref Result);
                    }

                }

                if (NeedsAllCols)
                {
                    int SCol2Real = SCol1 + RCol2 - RCol1;
                    for (int cIndex = SXls.ColToIndex(SSheet, SRow1 + r - RRow1, SCol2Real); cIndex > 0; cIndex--)
                    {
                        int c = SXls.ColFromIndex(SSheet, SRow1 + r - RRow1, cIndex);
                        if (c > SCol2Real || c == 0) continue;
                        if (c < SCol1) break;

                        int rIndex = Xls1.ColToIndex(RSheet, r, RCol1 + c - SCol1);
                        if (Xls1.ColFromIndex(RSheet, r, rIndex) == RCol1 + c - SCol1) continue; //it has already been processed in the main loop

                        ValuesProcessed++;
                        //Here criteria is always met. 
                        object v2 = ConvertToAllowedObject(SXls.GetCellValueAndRecalc(SSheet, SRow1 + r - RRow1, c, CalcState, CalcStack));
                        if (CalcState.Aborted) return;
                        Process(v2, ref Result);
                    }
                }
            }
        }
        protected abstract void ProcessFinal(long ValuesProcessed, long TotalValues, ref TAverageValue ResultValue, TValueCriteria CriteriaEval);


	}

    internal struct TCriteriaAndAddress
    {
        public TValueCriteria[,] Criteria;
        public TAddress[] Range;
    }

    internal abstract class TRangeIfsToken : TBaseRangeIfToken
    {
        private int FixedArgs;
        internal TRangeIfsToken(int ArgCount, ptg aId, TCellFunctionData aFuncData, int aFixedArgs)
            : base(ArgCount, aId, aFuncData)
        {
            FixedArgs = aFixedArgs;
        }

        internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
        {
            TAddress[] SumRange = null;
            if ((FArgCount - FixedArgs) % 2 != 0) return TFlxFormulaErrorValue.ErrNA; //shouldn't happen.
            
            TCriteriaAndAddress[] Criterias = new TCriteriaAndAddress[(FArgCount - FixedArgs) / 2];
            if (Criterias.Length < 1) return TFlxFormulaErrorValue.ErrNA;

            int MaxRows = 1;
            int MaxCols = 1;

            for (int i = 0; i < Criterias.Length; i++)
            {
                Criterias[i].Criteria = GetCriteria(FTokenList, wi, CalcState, CalcStack);
                if (Criterias[i].Criteria.GetLength(0) > MaxRows) MaxRows = Criterias[i].Criteria.GetLength(0);
                if (Criterias[i].Criteria.GetLength(1) > MaxCols) MaxCols = Criterias[i].Criteria.GetLength(1);

                object[,] values; TAddressList Range;
                object ret2 = GetRangeOrArray(FTokenList, wi, CalcState, CalcStack, out Range, out values, false);
                if (ret2 != null) return ret2;
                if (Range.Has3dRef() || Range.Count != 1) return TFlxFormulaErrorValue.ErrValue;
                Criterias[i].Range = Range[0];

                if (Math.Abs(Criterias[i].Range[1].Row - Criterias[i].Range[0].Row) !=
                    Math.Abs(Criterias[0].Range[1].Row - Criterias[0].Range[0].Row)) return TFlxFormulaErrorValue.ErrValue;
                if (Math.Abs(Criterias[i].Range[1].Col - Criterias[i].Range[0].Col) !=
                    Math.Abs(Criterias[0].Range[1].Col - Criterias[0].Range[0].Col)) return TFlxFormulaErrorValue.ErrValue;

            }

            if (FixedArgs > 0)
            {
                TAddressList SumRangeList = null;
                object res = GetSumRange(FTokenList, wi, CalcState, ref CalcStack, ref SumRangeList);
                if (res != null) return res;
                if (SumRangeList.Has3dRef() || SumRangeList.Count != 1) return TFlxFormulaErrorValue.ErrValue;

                SumRange = SumRangeList[0];
                if (Math.Abs(SumRange[1].Row - SumRange[0].Row) !=
                    Math.Abs(Criterias[0].Range[1].Row - Criterias[0].Range[0].Row)) return TFlxFormulaErrorValue.ErrValue;
                if (Math.Abs(SumRange[1].Col - SumRange[0].Col) !=
                    Math.Abs(Criterias[0].Range[1].Col - Criterias[0].Range[0].Col)) return TFlxFormulaErrorValue.ErrValue;
            }
            else
            {
                SumRange = Criterias[0].Range; //this will only happen in countifs, where SumRange doesn't matter. We don't want it to be null.
            }

            if (MaxRows > 1 || MaxCols > 1)
            {
                object[,] ArrResult = new object[MaxRows, MaxCols];
                for (int i = 0; i < ArrResult.GetLength(0); i++)
                {
                    for (int j = 0; j < ArrResult.GetLength(1); j++)
                    {
                        ArrResult[i, j] = CalcValues(SumRange, Criterias, i, j, CalcState, CalcStack);
                    }
                }
                return UnPack(ArrResult);
            }

            return CalcValues(SumRange, Criterias, 0, 0, CalcState, CalcStack);
        }


        internal object CalcValues(TAddress[] SumRange, TCriteriaAndAddress[] Criterias, int rr, int cc, TCalcState CalcState, TCalcStack CalcStack)
        {
            TAverageValue ResultValue = new TAverageValue(0, 0);

            long ValuesProcessed = 0;
            ProcessRange(CalcMaxRow(Criterias), SumRange, Criterias, rr, cc, CalcState, CalcStack, ref ResultValue, ref ValuesProcessed);

            ProcessFinal(ValuesProcessed, 
                ((long)Criterias[0].Range[1].Row - Criterias[0].Range[0].Row + 1) * (Criterias[0].Range[1].Col - Criterias[0].Range[0].Col + 1), 
                ref ResultValue, 
                rr, cc,
                Criterias);
            return GetFinalValue(ResultValue);
        }

        private int CalcMaxRow(TCriteriaAndAddress[] Criterias)
        {
            int Result = 0;
            foreach (TCriteriaAndAddress Crit in Criterias)
            {
                int SheetRowCount = Crit.Range[0].wi.Xls.GetRowCount(Crit.Range[0].Sheet);
                int SheetMax = SheetRowCount - Crit.Range[0].Row;
                if (Result < SheetMax) Result = SheetMax;
            }

            return Result;
        }

        protected void ProcessRange(int MaxRow, TAddress[] SumRange, TCriteriaAndAddress[] Criterias, int rr, int cc, TCalcState CalcState, 
            TCalcStack CalcStack, ref TAverageValue ResultValue, ref long ValuesProcessed)
        {
            int RowCount = Criterias[0].Range[1].Row - Criterias[0].Range[0].Row + 1;
            RowCount = Math.Min(RowCount, MaxRow);
            int ColCount = Criterias[0].Range[1].Col - Criterias[0].Range[0].Col + 1; 
            for (int r = 0; r < RowCount; r++)
            {
                for (int c = 0; c < ColCount; c++)
                {
                    bool RowHasMoreCols;
                    ProcessCell(SumRange[0], r, c, rr, cc, CalcState, CalcStack, Criterias, ref ResultValue, out RowHasMoreCols);
                    ValuesProcessed++;
                    if (!RowHasMoreCols) break;
                    if (CalcState.Aborted) return;
                }
            }
        }

        private void ProcessCell(TAddress SumStart, int r, int c, int rr, int cc, TCalcState CalcState, TCalcStack CalcStack, TCriteriaAndAddress[] Criterias, 
            ref TAverageValue Result, out bool RowHasMoreCols)
        {
            RowHasMoreCols = RowStillGoing(SumStart, r, c);
            for (int i = 0; i < Criterias.Length; i++)
            {
                TAddress FirstCell = Criterias[i].Range[0];
                object v1 = ConvertToAllowedObject(FirstCell.wi.Xls.GetCellValueAndRecalc(FirstCell.Sheet, FirstCell.Row + r, FirstCell.Col + c, CalcState, CalcStack));
                if (!RowHasMoreCols && RowStillGoing(FirstCell, r, c)) RowHasMoreCols = true;
                if (CalcState.Aborted) return;
                
                //if (SumRangeDifferent)  Do not convert here!. It can raise a circular error even when because of the criteria we will not use the value.
                //    v2=ConvertToAllowedObject(Xls.GetCellValueAndRecalc(SSheet, SRow1+r-RRow1, SCol1+c-RCol1));

                TValueCriteria[,] Crit = Criterias[i].Criteria;
                if (Crit.GetLength(0) == 1 && Crit.GetLength(1) == 1)
                {
                    if (!Crit[0, 0].MeetsCriteria(v1)) return;
                }
                else
                {
                    if (Crit.GetLength(0) <= rr || Crit.GetLength(1) <= cc || !Crit[rr, cc].MeetsCriteria(v1)) return;
                }
            }

            object v2 = ConvertToAllowedObject(SumStart.wi.Xls.GetCellValueAndRecalc(SumStart.Sheet, SumStart.Row + r, SumStart.Col + c, CalcState, CalcStack));
            if (CalcState.Aborted) return;
            Process(v2, ref Result);
        }

        private bool RowStillGoing(TAddress addr, int r, int c)
        {
            int CCount = addr.wi.Xls.ColCountInRow(addr.Row + r);
            if (CCount == 0) return false;
            if (addr.Col + c > addr.wi.Xls.ColFromIndex(addr.Row + r, CCount - 1)) return false;
            return true;
        }

        protected abstract void ProcessFinal(long ValuesProcessed, long TotalValues, ref TAverageValue ResultValue, int rr, int cc, TCriteriaAndAddress[] CriteriaEval);

    }

	internal class TSumIfToken : TRangeIfToken
	{
		internal TSumIfToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData)
		{
		}

		protected override void Process(object v, ref TAverageValue ResultValue)
		{
			if (v is double) ResultValue.Sum += (double)v;
		}

        protected override void ProcessFinal(long ValuesProcessed, long TotalValues, ref TAverageValue ResultValue, TValueCriteria CriteriaEval)
		{
			//Nothing to do here. Values not processed don't add to the sum.
		}

        protected override object GetFinalValue(TAverageValue ResultValue)
        {
            return ResultValue.Sum;
        }
	}

    internal class TAverageIfToken : TRangeIfToken
    {
        internal TAverageIfToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
            : base(ArgCount, aId, aFuncData)
        {
        }

        protected override void Process(object v, ref TAverageValue ResultValue)
        {
            if (v is double)
            {
                ResultValue.Sum += (double)v;
                ResultValue.ValueCount++;
            }
        }

        protected override void ProcessFinal(long ValuesProcessed, long TotalValues, ref TAverageValue ResultValue, TValueCriteria CriteriaEval)
        {
            //nothing here.
        }

        protected override object GetFinalValue(TAverageValue ResultValue)
        {
            if (ResultValue.ValueCount <= 0) return TFlxFormulaErrorValue.ErrDiv0;
            return ResultValue.Sum / ResultValue.ValueCount;
        }
    }

	internal class TCountIfToken : TRangeIfToken
	{
		internal TCountIfToken(ptg aId, TCellFunctionData aFuncData)
			: base(2, aId, aFuncData)
		{
		}

		protected override void Process(object v, ref TAverageValue ResultValue)
		{
			ResultValue.ValueCount++;
		}

        protected override void ProcessFinal(long ValuesProcessed, long TotalValues, ref TAverageValue ResultValue, TValueCriteria CriteriaEval)
        {
            if (CriteriaEval.MeetsCriteria(null)) ResultValue.ValueCount += TotalValues - ValuesProcessed;
		}

        protected override object GetFinalValue(TAverageValue ResultValue)
        {
            return (double)ResultValue.ValueCount;
        }

	}

    internal class TSumIfsToken : TRangeIfsToken
	{
		internal TSumIfsToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData, 1)
		{
		}

		protected override void Process(object v, ref TAverageValue ResultValue)
		{
			if (v is double) ResultValue.Sum += (double)v;
		}

        protected override void ProcessFinal(long ValuesProcessed, long TotalValues, ref TAverageValue ResultValue, int rr, int cc, TCriteriaAndAddress[] CriteriaEval)
		{
			//Nothing to do here. Values not processed don't add to the sum.
		}

        protected override object GetFinalValue(TAverageValue ResultValue)
        {
            return ResultValue.Sum;
        }
	}

    internal class TAverageIfsToken : TRangeIfsToken
    {
        internal TAverageIfsToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
            : base(ArgCount, aId, aFuncData, 1)
        {
        }

        protected override void Process(object v, ref TAverageValue ResultValue)
        {
            if (v is double)
            {
                ResultValue.Sum += (double)v;
                ResultValue.ValueCount++;
            }
        }

        protected override void ProcessFinal(long ValuesProcessed, long TotalValues, ref TAverageValue ResultValue, int rr, int cc, TCriteriaAndAddress[] CriteriaEval)
        {
            //nothing here.
        }

        protected override object GetFinalValue(TAverageValue ResultValue)
        {
            if (ResultValue.ValueCount <= 0) return TFlxFormulaErrorValue.ErrDiv0;
            return ResultValue.Sum / ResultValue.ValueCount;
        }
    }

	internal class TCountIfsToken : TRangeIfsToken
	{
		internal TCountIfsToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData, 0)
		{
		}

		protected override void Process(object v, ref TAverageValue ResultValue)
		{
			ResultValue.ValueCount++;
		}

        protected override void ProcessFinal(long ValuesProcessed, long TotalValues, ref TAverageValue ResultValue, int rr, int cc, TCriteriaAndAddress[] CriteriaEval)
        {
            if (MeetsAnyCriteria(CriteriaEval, null, rr, cc)) ResultValue.ValueCount += TotalValues - ValuesProcessed;
		}

        protected override object GetFinalValue(TAverageValue ResultValue)
        {
            return (double)ResultValue.ValueCount;
        }
	}

	#endregion

	#region Logical
	internal sealed class TIFToken : TBaseFunctionToken
	{
		internal TIFToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			//We need to get the positions for true and false without evaluating them, because an 
			//exception on the false condition would trash the results.
			int FalsePos = FTokenList.SavePosition();
			if (FArgCount > 2)
				FTokenList.Flush();
			int TruePos = FTokenList.SavePosition();
			FTokenList.Flush();

			object Cond = FTokenList.EvaluateToken(wi, CalcState, CalcStack); //No aggregated here.
			if (Cond is TFlxFormulaErrorValue) return Cond;

			if (Cond == null) Cond = false;

			object[,] Conds;
			if (IsArrayArgument(Cond, out Conds))
			{
				object TrueEval = EvalOne(true, TruePos, FalsePos, FTokenList, wi, null, CalcState, CalcStack);  //Even if there is no true item on condition, this counts up to the size of the final array.
				object FalseEval = EvalOne(false, TruePos, FalsePos, FTokenList, wi, null, CalcState, CalcStack);
				object[,] ArrTrueEval = TrueEval as object[,];
				object[,] ArrFalseEval = FalseEval as object[,];
				object[,] ArrResult;

				SetupArray(ref Conds, ref ArrTrueEval, ref ArrFalseEval, Cond, TrueEval, FalseEval, out ArrResult);

				for (int i = 0; i < ArrResult.GetLength(0); i++)
				{
					for (int j = 0; j < ArrResult.GetLength(1); j++)
					{
						ArrResult[i, j] = GetCondition(GetItem(Conds, i, j));
						if (ArrResult[i, j] is Boolean)
						{
							if ((bool)ArrResult[i, j])
								ArrResult[i, j] = GetItem(ArrTrueEval, i, j);
							else
								ArrResult[i, j] = GetItem(ArrFalseEval, i, j);
						}
					}
				}

				return UnPack(ArrResult);
			}

			object OCond = GetCondition(Cond);
			if (!(OCond is bool)) return OCond;
			return EvalOne((bool)OCond, TruePos, FalsePos, FTokenList, wi, f, CalcState, CalcStack);
		}

		private static object GetCondition(object Cond)
		{
			if (Cond is TFlxFormulaErrorValue) return Cond;
			bool BoolCond;
			if (!ExtToBool(Cond, out BoolCond)) return TFlxFormulaErrorValue.ErrValue;
			return BoolCond;
		}

		private object EvalOne(bool BoolCond, int TruePos, int FalsePos, TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			int CondPos = FTokenList.SavePosition();
			if (BoolCond)
			{
				FTokenList.RestorePosition(TruePos);
				object IfTrue = FTokenList.EvaluateToken(wi, f, CalcState, CalcStack);
				FTokenList.RestorePosition(CondPos);
				return IfTrue;
			}
			else
			{
				if (FArgCount < 3) return false;

				FTokenList.RestorePosition(FalsePos);
				object IfFalse = FTokenList.EvaluateToken(wi, f, CalcState, CalcStack); //For a formula like =Min(if(a1=1,a1:a2,b1:b2));
				FTokenList.RestorePosition(CondPos);
				return IfFalse;
			}

		}
	}

	internal sealed class TAndToken : TBaseFunctionToken
	{
		internal TAndToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			bool Result = true;
			for (int i = 0; i < FArgCount; i++)
			{
				object Cond = FTokenList.EvaluateToken(wi, TAndAggregate.Instance, CalcState, CalcStack); if (Cond is TFlxFormulaErrorValue) return Cond;

				object[,] Conds = Cond as Object[,];
				if (Conds != null)
				{
					for (int j = 0; j < Conds.GetLength(0); j++)
						for (int k = 0; k < Conds.GetLength(1); k++)
						{
							if (Conds[j, k] is TFlxFormulaErrorValue) return Conds[j, k];
							bool BC;
							if (!ExtToBool(Conds[j, k], out BC)) return TFlxFormulaErrorValue.ErrValue;
							if (!BC) Result = false;
						}
				}
				else
				{

					bool BoolCond;
					if (!ExtToBool(Cond, out BoolCond)) return TFlxFormulaErrorValue.ErrValue;
					Result = Result & BoolCond;
				}
			}

			return Result;
		}
	}

	internal sealed class TOrToken : TBaseFunctionToken
	{
		internal TOrToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			bool Result = false;
			for (int i = 0; i < FArgCount; i++)
			{
				object Cond = FTokenList.EvaluateToken(wi, TOrAggregate.Instance, CalcState, CalcStack); if (Cond is TFlxFormulaErrorValue) return Cond;

				object[,] Conds = Cond as Object[,];
				if (Conds != null)
				{
					for (int j = 0; j < Conds.GetLength(0); j++)
						for (int k = 0; k < Conds.GetLength(1); k++)
						{
							if (Conds[j, k] is TFlxFormulaErrorValue) return Conds[j, k];
							bool BC;
							if (!ExtToBool(Conds[j, k], out BC)) return TFlxFormulaErrorValue.ErrValue;
							if (BC) Result = true;
						}
				}
				else
				{
					bool BoolCond;
					if (!ExtToBool(Cond, out BoolCond)) return TFlxFormulaErrorValue.ErrValue;
					Result = Result || BoolCond;
				}
			}

			return Result;
		}
	}

	internal sealed class TNotToken : TBaseFunctionToken
	{
		internal TNotToken(ptg aId, TCellFunctionData aFuncData) : base(1, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			object[,] ArrCond;
			object[,] ArrResult;

			object Cond = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (Cond is TFlxFormulaErrorValue) return Cond;
			if (IsArrayArgument(Cond, out ArrResult, out ArrCond))
			{
				for (int i = 0; i < ArrResult.GetLength(0); i++)
					for (int j = 0; j < ArrResult.GetLength(1); j++)
					{
						bool BoolCond2;
						if (!GetBoolItem(ref ArrResult, ref ArrCond, i, j, out BoolCond2)) continue;
						ArrResult[i, j] = !BoolCond2;

					}
				return UnPack(ArrResult);
			}
			bool BoolCond;
			if (!ExtToBool(Cond, out BoolCond)) return TFlxFormulaErrorValue.ErrValue;
			return !BoolCond;
		}
	}

	internal sealed class TTrueToken : TBaseFunctionToken
	{
		internal TTrueToken(ptg aId, TCellFunctionData aFuncData) : base(0, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			return true;
		}
	}

	internal sealed class TFalseToken : TBaseFunctionToken
	{
		internal TFalseToken(ptg aId, TCellFunctionData aFuncData) : base(0, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			return false;
		}
	}

    internal sealed class TIfErrorToken : TBaseFunctionToken
    {
		internal TIfErrorToken(ptg aId, TCellFunctionData aFuncData) : base(2, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
            object Err = FTokenList.EvaluateToken(wi, CalcState, CalcStack);
            object Val = FTokenList.EvaluateToken(wi, CalcState, CalcStack);

            object[,] ValArr = Val as object[,];
            object[,] ValErr = Err as object[,];
            if (ValArr != null)
            {
                for (int r = 0; r < ValArr.GetLength(0); r++)
                {
                    for (int c = 0; c < ValArr.GetLength(1); c++)
                    {
                        if (ValArr[r, c] is TFlxFormulaErrorValue)
                        {
                            if (ValErr != null)
                                ValArr[r, c] = GetItem(ValErr, r, c);
                            else
                                ValArr[r, c] = Err;
                        }
                    }
                }
            }

            if (Val is TFlxFormulaErrorValue) return Err;
            return Val;
		}

    }
	#endregion

	#region Indirect and references

	internal class TIndirectToken : TBaseFunctionToken
	{
		internal TIndirectToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			object Ref = EvaluateRef(FTokenList, wi, CalcState, CalcStack);
			TAddress CellRef1 = null, CellRef2 = null;
			object ret = GetRange(Ref, out CellRef1, out CellRef2, TFlxFormulaErrorValue.ErrRef);
			if (ret != null) return ret;

			return f.Agg(CellRef1.wi, CellRef1.Sheet, CellRef2.Sheet, CellRef1.Row, CellRef1.Col, CellRef2.Row, CellRef2.Col, CalcState, CalcStack);
		}

		private static bool SplitSheet(ref ExcelFile xls, string SheetPlusName, ref int LocalIndex, out string OnlyName, out bool LocalName)
		{
			OnlyName = null;

			int w = SheetPlusName.LastIndexOf(TBaseFormulaParser.fts(TFormulaToken.fmExternalRef));
			if (w > 0)
			{
				LocalName = true;
				string SheetName = SheetPlusName.Substring(0, w);
				LocalIndex = xls.GetSheetIndex(SheetName, false);
				if (LocalIndex < 1) 
				{
					//Try to see if it is not an external name. something like "book1.xls!name"
					ExcelFile LocalXls = xls.GetSupportingFile(SheetName);
					if (LocalXls != null)
					{
						xls = LocalXls;
						LocalIndex = 0;
						LocalName = false;
						OnlyName = SheetPlusName.Substring(w + 1);
						return true;
					}

					return false;
				}
				OnlyName = SheetPlusName.Substring(w + 1);
				return true;
			}

			LocalName = false;
			OnlyName = SheetPlusName;
			return true;

		}

		private static TXlsNamedRange FindNamedRange(TWorkbookInfo wi, out ExcelFile LocalXls, string s)
		{
			LocalXls = wi.Xls;

			string FileName; string Sheets;
			TCellAddress.SplitFileName(s, out FileName, out Sheets);
			if (FileName != null)
			{
				LocalXls = wi.Xls.GetSupportingFile(FileName);
			}
            if (LocalXls == null) return null;
			
			int LocalIndex = wi.SheetIndexBase1;
			string OnlyName; bool LocalName;
			if (!SplitSheet(ref LocalXls, Sheets, ref LocalIndex, out OnlyName, out LocalName)) return null;

			TXlsNamedRange Nr = LocalXls.GetNamedRange(OnlyName, -1, LocalIndex); //try a local range.
			if (Nr == null && !LocalName)
				Nr = LocalXls.GetNamedRange(Sheets, -1, 0);  //try global ranges.
			
			return Nr;
		}

		internal override object EvaluateRef(TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack)
		{
            bool IsA1 = true;

            if (FArgCount > 1)
            {
                object a1 = FTokenList.EvaluateToken(wi, CalcState, CalcStack);
                if (a1 is TFlxFormulaErrorValue) return a1;

                if (!ExtToBool(a1, out IsA1)) return TFlxFormulaErrorValue.ErrValue;
            }

			object v1 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v1 is TFlxFormulaErrorValue) return v1;
			v1 = ConvertToAllowedObject(v1); //May be an array. Indirect does not work with arrays, so we take the first element.
			try
			{
				string s = FlxConvert.ToString(v1);
				int Sheet1; int Sheet2; int Row1; int Col1; int Row2; int Col2;

				ExcelFile LocalXls;

				if (!TCellAddress.TryParseAddress(wi.Xls, s, wi.SheetIndexBase1, out LocalXls, 
                    out Sheet1, out Sheet2, out Row1, out Col1, out Row2, out Col2,
                    IsA1 ? TReferenceStyle.A1 : TReferenceStyle.R1C1, wi.Row + 1 + wi.RowOfs, wi.Col + 1 + wi.ColOfs))
				{
					//The named range must be tried only after we know it is not a sheet name.
					TXlsNamedRange Nr = FindNamedRange(wi, out LocalXls, s);  
                
					if (Nr != null)
					{
						Sheet1 = Nr.SheetIndex;
						Sheet2 = Nr.SheetIndex;
						Row1 = Nr.Top;
						Col1 = Nr.Left;
						Row2 = Nr.Bottom;
						Col2 = Nr.Right;
					}
					else return TFlxFormulaErrorValue.ErrRef;
				}
				
				TWorkbookInfo wi2 = wi;
				if (wi.Xls != LocalXls)
				{
					wi2 = wi.ShallowClone();
					wi2.Xls = LocalXls;
				}

				TAddress Ca1 = new TAddress(wi2, Sheet1, Row1, Col1);
				if (Sheet1 == Sheet2 && Row1 == Row2 && Col1 == Col2) return Ca1;

				TAddress Ca2 = new TAddress(wi2, Sheet2, Row2, Col2);
				TAddress[] Result = new TAddress[2];
				Result[0] = Ca1;
				Result[1] = Ca2;
				return Result;


			}
			catch (Exception) //Whatever.
			{
			}

			return TFlxFormulaErrorValue.ErrRef;
		}
	}

	internal sealed class TRowToken : TBaseFunctionToken
	{
		internal TRowToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			if (FArgCount == 0)
			{
				int FColCount = 1; //ColCount is always 1.
				int FRowCount = wi.RowCount;
				if (FRowCount <= 0 || (FColCount == 1 && FRowCount == 1))
					return (double)(wi.Row + 1 + wi.RowOfs);
				object[,] ArrResult = new object[FRowCount, FColCount];
				for (int i = 0; i < FRowCount; i++)
					for (int j = 0; j < FColCount; j++)
					{
						ArrResult[i, j] = (double)(wi.Row + 1 + i + wi.RowOfs);
					}
				return ArrResult;
			}

			TAddress CellRef1 = null, CellRef2 = null;
			object[,] values = null;
			object ret = GetRangeOrArray(FTokenList, wi, CalcState, CalcStack, out CellRef1, out CellRef2, out values, false);
			if (ret != null) return ret;

			if (wi.IsArrayFormula)
			{
				object[,] ArrResult = new object[CellRef2.Row - CellRef1.Row + 1, 1];
				for (int r = CellRef1.Row; r <= CellRef2.Row; r++)
				{
					ArrResult[r - CellRef1.Row, 0] = (double)r;
				}
				return UnPack(ArrResult);
			}

			return (double)(CellRef1.Row);  //Row a1:a2 returns 1.
		}
	}

	internal sealed class TColToken : TBaseFunctionToken
	{
		internal TColToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			if (FArgCount == 0)
			{
				int FRowCount = 1; //Always 1.
				int FColCount = wi.ColCount;

				if (FColCount <= 0 || (FColCount == 1 && FRowCount == 1))
					return (double)(wi.Col + 1 + wi.ColOfs);
				object[,] ArrResult = new object[FRowCount, FColCount];
				for (int i = 0; i < FRowCount; i++)
					for (int j = 0; j < FColCount; j++)
					{
						ArrResult[i, j] = (double)(wi.Col + 1 + j + wi.ColOfs);
					}
				return ArrResult;
			}

			TAddress CellRef1 = null, CellRef2 = null; object[,] values = null;
			object ret = GetRangeOrArray(FTokenList, wi, CalcState, CalcStack, out CellRef1, out CellRef2, out values, false);
			if (ret != null) return ret;

			if (wi.IsArrayFormula)
			{
				object[,] ArrResult = new object[CellRef2.Col - CellRef1.Col + 1, 1];
				for (int c = CellRef1.Col; c <= CellRef2.Col; c++)
				{
					ArrResult[c - CellRef1.Col, 0] = (double)c;
				}
				return UnPack(ArrResult);
			}

			return (double)(CellRef1.Col);  //column a1:b2 returns 1.
		}
	}

	internal sealed class TAreasToken : TBaseFunctionToken
	{
		internal TAreasToken(int ArgCount, ptg aId, TCellFunctionData aFuncData) : base(ArgCount, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			TAddressList CellRefs;
			object[,] values = null;
			object ret = GetRangeOrArray(FTokenList, wi, CalcState, CalcStack, out CellRefs, out values, false);
			if (ret != null) return ret;

			if (CellRefs == null || CellRefs.Count == 0) return TFlxFormulaErrorValue.ErrNull;
			return CellRefs.Count;
		}
	}

	internal sealed class TRowsToken : TBaseFunctionToken
	{
		internal TRowsToken(ptg aId, TCellFunctionData aFuncData) : base(1, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			TAddress CellRef1 = null, CellRef2 = null;
			object[,] values = null;
			object ret = GetRangeOrArray(FTokenList, wi, CalcState, CalcStack, out CellRef1, out CellRef2, out values, true);
			if (ret != null) return ret;

			if (CellRef1 != null)
			{
				if (CellRef1.wi.Xls != CellRef2.wi.Xls || CellRef1.Sheet != CellRef2.Sheet) return TFlxFormulaErrorValue.ErrRef;
				return Math.Abs(CellRef2.Row - CellRef1.Row) + 1;
			}

			if (values != null)
			{
				return values.GetLength(0);
			}

			return TFlxFormulaErrorValue.ErrRef;
		}
	}

	internal sealed class TColumnsToken : TBaseFunctionToken
	{
		internal TColumnsToken(ptg aId, TCellFunctionData aFuncData) : base(1, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			TAddress CellRef1 = null, CellRef2 = null;
			object[,] values = null;
			object ret = GetRangeOrArray(FTokenList, wi, CalcState, CalcStack, out CellRef1, out CellRef2, out values, true);
			if (ret != null) return ret;

			if (CellRef1 != null)
			{
				if (CellRef1.wi.Xls != CellRef2.wi.Xls || CellRef1.Sheet != CellRef2.Sheet) return TFlxFormulaErrorValue.ErrRef;
				return Math.Abs(CellRef2.Col - CellRef1.Col) + 1;
			}

			if (values != null)
			{
				return values.GetLength(1);
			}

			return TFlxFormulaErrorValue.ErrRef;
		}
	}

	internal sealed class TTransposeToken : TBaseFunctionToken
	{
		internal TTransposeToken(ptg aId, TCellFunctionData aFuncData) : base(1, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			TAddress CellRef1 = null, CellRef2 = null;
			object[,] values = null;
			object ret = GetRangeOrArray(FTokenList, wi, CalcState, CalcStack, out CellRef1, out CellRef2, out values, TFlxFormulaErrorValue.ErrValue, true);
			if (ret != null) return ret;

			if (CellRef1 != null)
			{
				if (CellRef1.wi.Xls != CellRef2.wi.Xls || CellRef1.Sheet != CellRef2.Sheet) return TFlxFormulaErrorValue.ErrRef;
				if (!wi.IsArrayFormula) return TFlxFormulaErrorValue.ErrValue;
				object[,] ArrResult = new object[(CellRef2.Col - CellRef1.Col) + 1, (CellRef2.Row - CellRef1.Row) + 1];
				for (int r = CellRef1.Row; r <= CellRef2.Row; r++)
				{
					for (int c = CellRef1.Col; c <= CellRef2.Col; c++)
					{
						ArrResult[c - CellRef1.Col, r - CellRef1.Row] = CellRef1.wi.Xls.GetCellValueAndRecalc(CellRef1.Sheet, r, c, CalcState, CalcStack);
                        if (CalcState.Aborted) return TFlxFormulaErrorValue.ErrNA;
                    }
				}
				return UnPack(ArrResult);
			}

			if (values != null)
			{
				object[,] ArrResult = new object[values.GetLength(1), values.GetLength(0)];
				for (int r = 0; r < values.GetLength(0); r++)
				{
					for (int c = 0; c < values.GetLength(1); c++)
					{
						ArrResult[c, r] = values[r, c];
					}
				}
				return UnPack(ArrResult);
			}

			return TFlxFormulaErrorValue.ErrRef;
		}
	}

	#endregion

	#region Math functions

	internal sealed class TModToken : TNDoubleArgToken
	{
		internal TModToken(ptg aId, TCellFunctionData aFuncData) : base(2, aId, aFuncData) { }

		protected override object Calc(double[] x)
		{
			return (double)(x[1] % x[0]);
		}
	}

	internal sealed class TRoundToken : TNDoubleArgToken
	{
		internal TRoundToken(ptg aId, TCellFunctionData aFuncData) : base(2, aId, aFuncData) { }

		protected override object Calc(double[] x)
		{
            return UpRound(x[1], x[0]);
		}
	}

	internal sealed class TAbsToken : TOneDoubleArgToken
	{
		internal TAbsToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }

		protected override object Calc(double x)
		{
			return Math.Abs(x);
		}
	}

	internal sealed class TCeilingToken : TNDoubleArgToken
	{
		internal TCeilingToken(ptg aId, TCellFunctionData aFuncData) : base(2, aId, aFuncData) { }
        
		protected override object Calc(double[] x)
		{
			double signif = x[0]; double v = x[1];
			if (signif * v < 0) //Math.Sign is wrong here, because 0 has no sign;
				return TFlxFormulaErrorValue.ErrNum;

			if (signif == 0) return 0;
			signif = Math.Abs(signif);
			/*decimal uniform=(decimal)Math.Abs(d2)/(decimal)d1;
			decimal Temp = decimal.Floor(uniform);
			decimal z=1;
			if (Temp== uniform) z=0;
			return ((double)(Decimal.Floor(Temp+z)*(decimal)d1))*Math.Sign(d2);
			*/
			return Math.Floor(Math.Abs(v) / -signif) * -signif * Math.Sign(v);
		}
	}

    internal sealed class TCeilingPreciseToken : TNDoubleArgToken
    {
        internal TCeilingPreciseToken(int aArgCount, ptg aId, TCellFunctionData aFuncData) : base(aArgCount, aId, aFuncData) { }

        protected override object Calc(double[] x)
        {
            double v;
            double signif = 1;

            if (x.Length > 1)
            {
                v = x[1];
                signif = x[0];
            }
            else v = x[0];

            if (signif == 0) return 0;
            if (v == 0) return 0;
            signif = Math.Abs(signif);

            return Math.Floor(v / -signif) * -signif;
        }
    }

	internal sealed class TFloorToken : TNDoubleArgToken
	{
		internal TFloorToken(ptg aId, TCellFunctionData aFuncData) : base(2, aId, aFuncData) { }
        
		protected override object Calc(double[] x)
		{
			double v = x[0]; double signif = x[1];
			if (v * signif < 0) //Math.Sign is wrong here, because 0 has no sign;
				return TFlxFormulaErrorValue.ErrNum;

			if (v == 0) return 0;
			v = Math.Abs(v);
			return Math.Floor(Math.Abs(signif) / v) * v * Math.Sign(signif);
		}
	}

    internal sealed class TFloorPreciseToken : TNDoubleArgToken
    {
        internal TFloorPreciseToken(int aArgCount, ptg aId, TCellFunctionData aFuncData) : base(aArgCount, aId, aFuncData) { }

        protected override object Calc(double[] x)
        {
            double v;
            double signif = 1;

            if (x.Length > 1)
            {
                v = x[1];
                signif = x[0];
            }
            else v = x[0];

            if (signif == 0) return 0;
            if (v == 0) return 0;
            signif = Math.Abs(signif);

            return Math.Floor(v / signif) * signif;
        }
    }

	internal abstract class TBaseRoundToken : TNDoubleArgToken
	{
		internal TBaseRoundToken(ptg aId, TCellFunctionData aFuncData) : base(2, aId, aFuncData) { }

		protected override object Calc(double[] x)
		{
			int d1;
			TFlxFormulaErrorValue Err;
			if (!GetInt(x[0], out d1, out Err)) return Err;
			double d2 = x[1];

			double dx2 = d2;
			double Result = DoRound(dx2);
			double mult = 1;

			for (int i = 0; i < d1; i++)
			{
				if (Result == dx2) return d2;
				dx2 *= 10;
				mult /= 10;
				Result = DoRound(dx2);
			}

			for (int i = 0; i > d1; i--)
			{
				if (Result == dx2) return d2;
				dx2 /= 10;
				mult *= 10;
				Result = DoRound(dx2);
			}

			return Result * mult;
		}

		protected abstract double DoRound(double value);
	}

	internal sealed class TRoundUpToken : TBaseRoundToken
	{
		internal TRoundUpToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override double DoRound(double value)
		{
			if (value > 0)
				return Math.Ceiling(value);
			return Math.Floor(value);
		}
	}

	internal sealed class TRoundDownToken : TBaseRoundToken
	{
		internal TRoundDownToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override double DoRound(double value)
		{
			if (value < 0)
				return Math.Ceiling(value);
			return Math.Floor(value);
		}
	}

	internal abstract class TBaseOddEvenToken : TOneDoubleArgToken
	{
		internal TBaseOddEvenToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }

		protected override object Calc(double x)
		{
			double Result = 0;
			if (x > 0)
				Result = Math.Ceiling(x);
			else
				Result = Math.Floor(x);

			if (AddOne(Result))
				if (Result < 0) Result--;
				else
					Result++;

			return Result;
		}

		protected abstract bool AddOne(double value);
	}

	internal sealed class TEvenToken : TBaseOddEvenToken
	{
		internal TEvenToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override bool AddOne(double value)
		{
			return Convert.ToDecimal(value) % 2 != 0;
		}
	}

	internal sealed class TOddToken : TBaseOddEvenToken
	{
		internal TOddToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override bool AddOne(double value)
		{
			return Convert.ToDecimal(value) % 2 == 0;
		}
	}

	internal sealed class TExpToken : TOneDoubleArgToken
	{
		internal TExpToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(double x)
		{
			return Math.Exp(x);
		}
	}

	internal sealed class TIntToken : TOneDoubleArgToken
	{
		internal TIntToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(double x)
		{
			return Math.Floor(x);
		}

	}

	internal sealed class TLnToken : TOneDoubleArgToken
	{
		internal TLnToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(double x)
		{
			if (x < 0) return TFlxFormulaErrorValue.ErrNum;
			return Math.Log(x);
		}
	}

	internal class TLogToken : TBaseFunctionToken
	{
		internal TLogToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			double d1 = 10;
			if (FArgCount > 1)
			{
				object v1 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v1 is TFlxFormulaErrorValue) return v1;
				if (!ExtToDouble(v1, out d1)) return TFlxFormulaErrorValue.ErrValue;
			}

			object v2 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v2 is TFlxFormulaErrorValue) return v2;
			double d2; if (!ExtToDouble(v2, out d2)) return TFlxFormulaErrorValue.ErrValue;
			if (d2 < 0) return TFlxFormulaErrorValue.ErrNum;
			return Math.Log(d2) / Math.Log(d1);
		}
	}

	internal sealed class TLog10Token : TOneDoubleArgToken
	{
		internal TLog10Token(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }

		protected override object Calc(double x)
		{
			if (x < 0) return TFlxFormulaErrorValue.ErrNum;
			return Math.Log10(x);
		}
	}

	internal sealed class TPowerFuncToken : TNDoubleArgToken
	{
		internal TPowerFuncToken(ptg aId, TCellFunctionData aFuncData) : base(2, aId, aFuncData) { }

		protected override object Calc(double[] x)
		{
			return Math.Pow(x[1], x[0]);
		}
	}


	internal sealed class TPiToken : TBaseFunctionToken
	{
		internal TPiToken(ptg aId, TCellFunctionData aFuncData) : base(0, aId, aFuncData) { }
        
		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			return Math.PI;
		}
	}

	internal sealed class TRandToken : TBaseFunctionToken
	{
		internal TRandToken(ptg aId, TCellFunctionData aFuncData) : base(0, aId, aFuncData) { }
		internal readonly Random rnd = new Random();  //initialized only once.


		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			return rnd.NextDouble();
		}
	}

	internal sealed class TSignToken : TOneDoubleArgToken
	{
		internal TSignToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(double x)
		{
			return (double)Math.Sign(x);
		}
	}

	internal sealed class TSqrtToken : TOneDoubleArgToken
	{
		internal TSqrtToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(double x)
		{
			if (x < 0) return TFlxFormulaErrorValue.ErrNum;
			return Math.Sqrt(x);
		}
	}

	internal sealed class TTruncToken : TBaseFunctionToken
	{
		internal TTruncToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			double d1 = 0;
			if (FArgCount > 1)
			{
				object v1 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v1 is TFlxFormulaErrorValue) return v1;
				if (!ExtToDouble(v1, out d1)) return TFlxFormulaErrorValue.ErrValue;
			}

			object v2 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v2 is TFlxFormulaErrorValue) return v2;
			double d2; if (!ExtToDouble(v2, out d2)) return TFlxFormulaErrorValue.ErrValue;

			decimal factor = CalcFactor(d1);

			if (factor == 0) return Math.Floor(Math.Abs(d2)) * Math.Sign(d2);
			return ((double)(Decimal.Floor(((Decimal)Math.Abs(d2) * factor)) / factor)) * Math.Sign(d2);
		}
	}

	internal sealed class TFactToken : TOneDoubleArgToken
	{
		internal TFactToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(double x)
		{
			if (x < 0 || x > MaxFact) return TFlxFormulaErrorValue.ErrNum;
			int Result = (int)x;
			return Factorial(Result);
		}

		internal const int MaxFact = 170;

		internal static double Factorial(int n)
		{
			if (n < 2) return 1;

			double p = 1;
			double r = 1;
			long NN = 1;

			int h = 0, shift = 0, high = 1;
			//int log2n = (int)Math.Floor(Math.Log(n, 2));  This does not run on CF
			int log2n = (int)Math.Floor(Math.Log(n) / Math.Log(2));

			while (h != n)
			{
				shift += h;
				h = n >> log2n--;
				int len = high;
				high = (h & 1) == 1 ? h : h - 1;
				len = (high - len) / 2;

				if (len > 0)
				{
					p *= Product(len, ref NN);
					r *= p;
				}
			}
			return r * Math.Pow(2, shift);
		}

		private static double Product(int n, ref long NN)
		{
			int m = n / 2;
			if (m == 0) return NN += 2;
			if (n == 2) return (NN += 2) * (NN += 2);
			return Product(n - m, ref NN) * Product(m, ref NN);
		}
	}

	internal abstract class TBasePermutCombinToken : TNDoubleArgToken
	{
		protected TBasePermutCombinToken(ptg aId, TCellFunctionData aFuncData) : base(2, aId, aFuncData) { }

		protected override object Calc(double[] x)
		{
			int NumberChosen;
			TFlxFormulaErrorValue Err;
			if (!GetInt(x[0], out NumberChosen, out Err)) return Err;

			int Number;
			if (!GetInt(x[1], out Number, out Err)) return Err;

			if (NumberChosen < 0 || Number < NumberChosen) return TFlxFormulaErrorValue.ErrNum;

			return DoCombin(Number, NumberChosen);
		}

		protected abstract object DoCombin(long n, long k);

	}

	internal sealed class TCombinToken : TBasePermutCombinToken
	{
		internal TCombinToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object DoCombin(long n, long k)
		{
			return Combin(n, k);
		}

		internal static double Combin(long n, long k)
		{
			if (n == k || k == 0) return 1;

			long delta, iMax;

			if (k < n - k)
			{
				delta = n - k;
				iMax = k;
			}
			else
			{
				delta = k;
				iMax = n - k;
			}

			double ans = delta + 1;

			for (long i = 2; i <= iMax; ++i)
			{
				checked { ans = (ans * (delta + i)) / i; }
				if (double.IsNaN(ans) || double.IsInfinity(ans)) return ans;
			}
			return ans;
		}
	}

	internal sealed class TPermutToken : TBasePermutCombinToken
	{
		internal TPermutToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object DoCombin(long n, long k)
		{
			double c = TCombinToken.Combin(n, k);
			return (double)c * TFactToken.Factorial((int)k);
		}
	}

	internal sealed class TRomanToken : TBaseFunctionToken
	{
		internal TRomanToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			int Form = 0;
			if (FArgCount > 1)
			{
				object v1 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v1 is TFlxFormulaErrorValue) return v1;
				if (v1 is bool) //here we cannot treat bools as numbers.
				{
					if (!(bool)v1) Form = 4;
				}
				else
				{
					TFlxFormulaErrorValue Err;
					if (!GetUInt(v1, out Form, out Err)) return Err;
					if (Form > 4) return TFlxFormulaErrorValue.ErrValue;
				}
			}

			object v2 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v2 is TFlxFormulaErrorValue) return v2;
			int Number;
			TFlxFormulaErrorValue Err2;
			if (!GetUInt(v2, out Number, out Err2)) return Err2;
			if (Number < 0 || Number > 3999) return TFlxFormulaErrorValue.ErrValue;

			int[] Numerals;
			string[] NumToChar;

			switch (Form)
			{
				case 1:
					Numerals = new int[] { 1, 4, 5, 9, 10, 40, 45, 50, 90, 95, 100, 400, 450, 500, 900, 950, 1000 };
					NumToChar = new string[] { "I", "IV", "V", "IX", "X", "XL", "VL", "L", "XC", "VC", "C", "CD", "LD", "D", "CM", "LM", "M" };
					break;
				case 2:
					Numerals = new int[] { 1, 4, 5, 9, 10, 40, 45, 49, 50, 90, 95, 99, 100, 400, 450, 490, 500, 900, 950, 990, 1000 };
					NumToChar = new string[] { "I", "IV", "V", "IX", "X", "XL", "VL", "IL", "L", "XC", "VC", "IC", "C", "CD", "LD", "XD", "D", "CM", "LM", "XM", "M" };
					break;
				case 3:
					Numerals = new int[] { 1, 4, 5, 9, 10, 40, 45, 49, 50, 90, 95, 99, 100, 400, 450, 490, 495, 500, 900, 950, 990, 995, 1000 };
					NumToChar = new string[] { "I", "IV", "V", "IX", "X", "XL", "VL", "IL", "L", "XC", "VC", "IC", "C", "CD", "LD", "XD", "VD", "D", "CM", "LM", "XM", "VM", "M" };
					break;
				case 4:
					Numerals = new int[] { 1, 4, 5, 9, 10, 40, 45, 49, 50, 90, 95, 99, 100, 400, 450, 490, 495, 499, 500, 900, 950, 990, 995, 999, 1000 };
					NumToChar = new string[] { "I", "IV", "V", "IX", "X", "XL", "VL", "IL", "L", "XC", "VC", "IC", "C", "CD", "LD", "XD", "VD", "ID", "D", "CM", "LM", "XM", "VM", "IM", "M" };
					break;
				default:
					Numerals = new int[] { 1, 4, 5, 9, 10, 40, 50, 90, 100, 400, 500, 900, 1000 };
					NumToChar = new string[] { "I", "IV", "V", "IX", "X", "XL", "L", "XC", "C", "CD", "D", "CM", "M" };
					break;
			}


			StringBuilder Result = new StringBuilder();
			while (Number > 0)
			{
				int n = Numerals.Length - 1;
				int r = 0;

				int k = r;

				while (n >= r)
				{
					k = (r + n) / 2;
					if (Number > Numerals[k]) r = k + 1; else if (Number < Numerals[k]) n = k - 1; else break;
				}

				if (Number < Numerals[k]) k--; //Position not found.
				Debug.Assert(Number >= Numerals[k]);

				Number -= Numerals[k];
				Result.Append(NumToChar[k]);
			}
			Debug.Assert(Number == 0);

			return Result.ToString();
		}
	}

	internal sealed class TMMultToken : TBaseFunctionToken
	{
		internal TMMultToken(ptg aId, TCellFunctionData aFuncData) : base(2, aId, aFuncData) { }
        
		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			TAddress CellRefB1 = null, CellRefB2 = null;
			object[,] valuesB = null;
			object ret = GetRangeOrArray(FTokenList, wi, CalcState, CalcStack, out CellRefB1, out CellRefB2, out valuesB, TFlxFormulaErrorValue.ErrValue, true);
			if (ret != null) return ret;

			int RowsB = 0; int ColsB = 0;
			int RB = -1; int CB = -1; int SB = -1; ExcelFile XB = null;
			if (CellRefB1 != null)
			{
				if (CellRefB1.wi.Xls != CellRefB2.wi.Xls || CellRefB1.Sheet != CellRefB2.Sheet) return TFlxFormulaErrorValue.ErrRef;
				RowsB = Math.Abs(CellRefB1.Row - CellRefB2.Row) + 1;
				ColsB = Math.Abs(CellRefB1.Col - CellRefB2.Col) + 1;
				RB = Math.Min(CellRefB1.Row, CellRefB2.Row);
				CB = Math.Min(CellRefB1.Col, CellRefB2.Col);
				SB = CellRefB1.Sheet;
				XB = CellRefB1.wi.Xls;
			}
			if (valuesB != null)
			{
				RowsB = valuesB.GetLength(0);
				ColsB = valuesB.GetLength(1);
			}

			TAddress CellRefA1 = null, CellRefA2 = null;
			object[,] valuesA = null;
			object ret2 = GetRangeOrArray(FTokenList, wi, CalcState, CalcStack, out CellRefA1, out CellRefA2, out valuesA, TFlxFormulaErrorValue.ErrValue, true);
			if (ret2 != null) return ret2;

			int ColsA = 0; int RowsA = 0;
			int RA = -1; int CA = -1; int SA = -1; ExcelFile XA = null;
			if (CellRefA1 != null)
			{
				if (CellRefA1.wi.Xls != CellRefA2.wi.Xls || CellRefA1.Sheet != CellRefA2.Sheet) return TFlxFormulaErrorValue.ErrRef;
				ColsA = Math.Abs(CellRefA1.Col - CellRefA2.Col) + 1;
				RowsA = Math.Abs(CellRefA1.Row - CellRefA2.Row) + 1;
				RA = Math.Min(CellRefA1.Row, CellRefA2.Row);
				CA = Math.Min(CellRefA1.Col, CellRefA2.Col);
				SA = CellRefA1.Sheet;
				XA = CellRefA1.wi.Xls;
			}
			if (valuesA != null)
			{
				RowsA = valuesA.GetLength(0);
				ColsA = valuesA.GetLength(1);
			}

			if (ColsA != RowsB || RowsA <= 0 || ColsB <= 0) return TFlxFormulaErrorValue.ErrValue;

			object[,] ArrResult = new object[RowsA, ColsB];
			for (int r = 0; r < RowsA; r++)
			{
				for (int c = 0; c < ColsB; c++)
				{
					object Arc = MultRow(r, c, ColsA, XA, SA, RA, CA, valuesA, XB, SB, RB, CB, valuesB, CalcState, CalcStack);
					if (!(Arc is Double)) return Arc;
					ArrResult[r, c] = (double)Arc;
				}
			}
			return UnPack(ArrResult);
		}

		private static object MultRow(int r, int c, int RC, ExcelFile XA, int SheetA, int RA, int CA, object[,] valuesA, ExcelFile XB, int SheetB, int RB, int CB, object[,] valuesB, TCalcState CalcState, TCalcStack CalcStack)
		{
			double Result = 0;
			for (int i = 0; i < RC; i++)
			{
				object a = GetElement(XA, SheetA, RA, CA, valuesA, r, i, CalcState, CalcStack);
				if (a is TFlxFormulaErrorValue) return a;
				if (!(a is Double)) return TFlxFormulaErrorValue.ErrValue;
				object b = GetElement(XB, SheetB, RB, CB, valuesB, i, c, CalcState, CalcStack);
				if (b is TFlxFormulaErrorValue) return b;
				if (!(b is Double)) return TFlxFormulaErrorValue.ErrValue;

				Result += (double)a * (double)b;
			}
			return Result;
		}

		private static object GetElement(ExcelFile Xls, int Sheet, int Row, int Col, object[,] Values, int r, int c, TCalcState CalcState, TCalcStack CalcStack)
		{
			if (Values == null)
			{
				return Xls.GetCellValueAndRecalc(Sheet, Row + r, Col + c, CalcState, CalcStack);
			}
			else
			{
				return Values[r, c];
			}
		}

	}

	#endregion

	#region Trigonometric functions
	internal sealed class TRadiansToken : TOneDoubleArgToken
	{
		internal TRadiansToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(double x)
		{
			return x * Math.PI / 180;
		}
	}

	internal sealed class TDegreesToken : TOneDoubleArgToken
	{
		internal TDegreesToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(double x)
		{
			return x / Math.PI * 180;
		}
	}

	#region Sin/Cos
	internal sealed class TSinToken : TOneDoubleArgToken
	{
		internal TSinToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(double x)
		{
			return Math.Sin(x);
		}
	}

	internal sealed class TCosToken : TOneDoubleArgToken
	{
		internal TCosToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(double x)
		{
			return Math.Cos(x);
		}
	}

	internal sealed class TTanToken : TOneDoubleArgToken
	{
		internal TTanToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(double x)
		{
			return Math.Tan(x);
		}
	}

	internal sealed class TAsinToken : TOneDoubleArgToken
	{
		internal TAsinToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(double x)
		{
			if (Math.Abs(x) > 1) return TFlxFormulaErrorValue.ErrNum;
			return Math.Asin(x);
		}
	}

	internal sealed class TAcosToken : TOneDoubleArgToken
	{
		internal TAcosToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(double x)
		{
			if (Math.Abs(x) > 1) return TFlxFormulaErrorValue.ErrNum;
			return Math.Acos(x);
		}
	}

	internal sealed class TAtanToken : TOneDoubleArgToken
	{
		internal TAtanToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(double x)
		{
			return Math.Atan(x);
		}
	}

	internal sealed class TAtan2Token : TNDoubleArgToken
	{
		internal TAtan2Token(ptg aId, TCellFunctionData aFuncData) : base(2, aId, aFuncData) { }
        
		protected override object Calc(double[] x)
		{
			if (x[0] == 0 && x[1] == 0) return TFlxFormulaErrorValue.ErrDiv0;
			return Math.Atan2(x[0], x[1]);
		}
	}

	#endregion

	#region Sin/Cos Hyperbolic
	internal sealed class TSinhToken : TOneDoubleArgToken
	{
		internal TSinhToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(double x)
		{
			return Math.Sinh(x);
		}
	}

	internal sealed class TCoshToken : TOneDoubleArgToken
	{
		internal TCoshToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(double x)
		{
			return Math.Cosh(x);
		}
	}

	internal sealed class TTanhToken : TOneDoubleArgToken
	{
		internal TTanhToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(double x)
		{
			return Math.Tanh(x);
		}
	}

	internal sealed class TAsinhToken : TOneDoubleArgToken
	{
		internal TAsinhToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(double x)
		{
			return Math.Log(x + Math.Sqrt(1.0 + x * x));
		}
	}

	internal sealed class TAcoshToken : TOneDoubleArgToken
	{
		internal TAcoshToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(double x)
		{
			return Math.Log(x + Math.Sqrt(x + 1.0) * Math.Sqrt(x - 1.0));
		}
	}

	internal sealed class TAtanhToken : TOneDoubleArgToken
	{
		internal TAtanhToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(double x)
		{
			return (Math.Log(1.0 + x) - Math.Log(1.0 - x)) / 2.0;
		}
	}
	#endregion
	#endregion

	#region String functions
	internal sealed class TMidToken : TBaseFunctionToken
	{
		public TMidToken(ptg aId, TCellFunctionData aFuncData) : base(3, aId, aFuncData) { }
        
		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			object[,] ArrL = null;
			object[,] ArrSp = null;
			object[,] ArrS = null;

			int l = 0;
			object Ret = GetUIntArgument(FTokenList, wi, CalcState, CalcStack, ref l, out ArrL); if (Ret != null) return Ret;

			int sp = 0;
			Ret = GetUIntArgument(FTokenList, wi, CalcState, CalcStack, ref sp, out ArrSp); if (Ret != null) return Ret;

			string s = String.Empty;
			Ret = GetStringArgument(FTokenList, wi, CalcState, CalcStack, ref s, out ArrS); if (Ret != null) return Ret;

			if (ArrL != null || ArrS != null || ArrSp != null)
			{
				Object[,] ArrResult;
				SetupArray(ref ArrL, ref ArrS, ref ArrSp, l, s, sp, out ArrResult);
				for (int i = 0; i < ArrResult.GetLength(0); i++)
					for (int j = 0; j < ArrResult.GetLength(1); j++)
					{
						if (!GetUIntItem(ref ArrResult, ref ArrL, i, j, out l)) continue;
						if (!GetUIntItem(ref ArrResult, ref ArrSp, i, j, out sp)) continue;
						if (!GetStringItem(ref ArrResult, ref ArrS, i, j, out s)) continue;
						ArrResult[i, j] = CalcMid(s, sp, l);
					}
				return UnPack(ArrResult);
			}

			return CalcMid(s, sp, l);
		}
		private static object CalcMid(string s, int sp, int l)
		{
			if (sp < 1) return TFlxFormulaErrorValue.ErrValue;

			if (sp - 1 < 0 || sp - 1 >= s.Length) return string.Empty;
			if (sp - 1 + l > s.Length) l = s.Length - (sp - 1);
			return s.Substring(sp - 1, l);
		}
	}

	internal abstract class TLeftRightToken : TBaseFunctionToken
	{
		protected TLeftRightToken(int ArgCount, ptg aId, TCellFunctionData aFuncData) : base(ArgCount, aId, aFuncData) { }


		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			object[] Values ={ String.Empty, 1 };
			return DoArguments(FTokenList, wi, CalcState, CalcStack, FArgCount, ref Values, new TArgType[] { TArgType.String, TArgType.UInt }, false);
		}

        protected override object DoOneArg(TWorkbookInfo wi, object[] Values)
		{
			return Calc((string)Values[0], (int)Values[1]);
		}

		internal abstract object Calc(string s, int l);
	}

	internal sealed class TLeftToken : TLeftRightToken
	{
		internal TLeftToken(int ArgCount, ptg aId, TCellFunctionData aFuncData) : base(ArgCount, aId, aFuncData) { }

		internal override object Calc(string s, int l)
		{
			if (s.Length > l) return s.Substring(0, l); else return s;
		}
	}

	internal sealed class TRightToken : TLeftRightToken
	{
		internal TRightToken(int ArgCount, ptg aId, TCellFunctionData aFuncData) : base(ArgCount, aId, aFuncData) { }

		internal override object Calc(string s, int l)
		{
			if (s.Length > l) return s.Substring(s.Length - l, l); else return s;
		}
	}

	internal sealed class TLengthToken : TOneStringArgToken
	{
		internal TLengthToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(string s)
		{
			return (double)s.Length;
		}
	}

	internal sealed class TLowerToken : TOneStringArgToken
	{
		internal TLowerToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(string s)
		{
			return s.ToLower(CultureInfo.CurrentCulture);
		}
	}

	internal sealed class TUpperToken : TOneStringArgToken
	{
		internal TUpperToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(string s)
		{
			return s.ToUpper(CultureInfo.CurrentCulture);
		}
	}

	internal sealed class TProperToken : TOneStringArgToken
	{
		internal TProperToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(string s)
		{
			//ToTitleCase is in fact nicer than Excel's Proper (for example it converts "dont's" into "Dont's" and not "Dont'S"
			//but we need to mimic Excel exactly.
			//return CultureInfo.CurrentCulture.TextInfo.ToTitleCase(Convert.ToString(v3, CultureInfo.CurrentCulture));
			char[] sb = s.ToCharArray();
			for (int i = 0; i < sb.Length; i++)
			{
				if (i > 0 && Char.IsLetter(sb[i - 1]))
					sb[i] = Char.ToLower(sb[i], CultureInfo.CurrentCulture);
				else
                    sb[i] = Char.ToUpper(sb[i], CultureInfo.CurrentCulture);
			}
			return new string(sb);
		}
	}

	internal sealed class TTrimToken : TOneStringArgToken
	{
		internal TTrimToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(string s)
		{
			return s.Trim();
		}
	}

	internal sealed class TCharToken : TOneDoubleArgToken
	{
		internal TCharToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(double x)
		{
			TFlxFormulaErrorValue Err;
			int l; if (!GetUInt(x, out l, out Err)) return Err;
			if (l<0 || l > 255) return TFlxFormulaErrorValue.ErrValue;

			return CharUtils.GetUniFromWin1252((byte)l).ToString(CultureInfo.InvariantCulture);
		}
	}

	internal sealed class TAscToken : TOneStringArgToken
	{
		internal TAscToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(string s)
		{
			return s;
		}

	}

	internal sealed class TCodeToken : TOneStringArgToken
	{
		internal TCodeToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(string s)
		{
			if (s == null || s.Length <= 0) return TFlxFormulaErrorValue.ErrValue;
			return Encoding.GetEncoding(0).GetBytes(s)[0];
		}
	}

	internal sealed class TReplaceToken : TBaseFunctionToken
	{
		internal TReplaceToken(ptg aId, TCellFunctionData aFuncData) : base(4, aId, aFuncData) { }
        
		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			object[] Values ={ String.Empty, 0, 0, String.Empty };
			return DoArguments(FTokenList, wi, CalcState, CalcStack, FArgCount, ref Values, new TArgType[] { TArgType.String, TArgType.UInt, TArgType.UInt, TArgType.String }, false);
		}
        protected override object DoOneArg(TWorkbookInfo wi, object[] Values)
		{
			string s = (string)Values[0];
			int rpPos = (int)Values[1];
			int rpCount = (int)Values[2];
			string rp = (string)Values[3];

			if (rpPos <= 0) return TFlxFormulaErrorValue.ErrValue;
			string s1 = ((rpPos - 1) < s.Length) ? s.Substring(0, rpPos - 1) : s;
			string s2 = ((rpPos - 1 + rpCount) < s.Length) ? s.Substring(rpPos - 1 + rpCount) : string.Empty;
			return s1 + rp + s2;
		}

	}

	internal sealed class TSubstituteToken : TBaseFunctionToken
	{
		internal TSubstituteToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			int instNo = 0;
			if (FArgCount > 3)
			{
				object v2 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v2 is TFlxFormulaErrorValue) return v2;
				TFlxFormulaErrorValue Err;
				if (!GetUInt(v2, out instNo, out Err)) return Err;
				if (instNo <= 0) return TFlxFormulaErrorValue.ErrValue;
			}
			object v4 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v4 is TFlxFormulaErrorValue) return v4;
			string newtext = FlxConvert.ToString(v4);

			v4 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v4 is TFlxFormulaErrorValue) return v4;
			string oldtext = FlxConvert.ToString(v4);

			v4 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v4 is TFlxFormulaErrorValue) return v4;
			string text = FlxConvert.ToString(v4);

			if (FArgCount <= 3)
				return text.Replace(oldtext, newtext);

			int start = -1;
			for (int i = 0; i < instNo; i++)
			{
				start = text.IndexOf(oldtext, start + 1);
				if (start < 0) return text;
			}
			if (start >= 0)
				return text.Substring(0, start) + newtext + text.Substring(start + oldtext.Length);
			return text;
		}
	}

	internal sealed class TFindToken : TBaseFunctionToken
	{
		internal TFindToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			int l;
			if (FArgCount > 2)
			{
				object v1 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v1 is TFlxFormulaErrorValue) return v1;
				TFlxFormulaErrorValue Err;
				if (!GetUInt(v1, out l, out Err)) return Err;
			}
			else l = 1;

			object v3 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v3 is TFlxFormulaErrorValue) return v3;
			string WithinText = FlxConvert.ToString(v3);
			object v4 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v4 is TFlxFormulaErrorValue) return v4;
			string FindText = FlxConvert.ToString(v4);

			if (l - 1 < 0 || l - 1 > WithinText.Length) return TFlxFormulaErrorValue.ErrValue;
			if (FindText.Length == 0) return l;
			int i = WithinText.IndexOf(FindText, l - 1);
			if (i >= 0) return i + 1; else return TFlxFormulaErrorValue.ErrValue;
		}
	}

	internal sealed class TConcatenateToken : TBaseFunctionToken
	{
		internal TConcatenateToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			object[] Values = new string[FArgCount];
			TArgType[] ArgTypes = new TArgType[FArgCount];
			for (int i = 0; i < ArgTypes.Length; i++)
			{
				ArgTypes[i] = TArgType.String;
				Values[i] = String.Empty;
			}
			return DoArguments(FTokenList, wi, CalcState, CalcStack, FArgCount, ref Values, ArgTypes, false);
		}

        protected override object DoOneArg(TWorkbookInfo wi, object[] Values)
		{
			StringBuilder Result = new StringBuilder();
			for (int i = 0; i < Values.Length; i++)
			{
				Result.Append(FlxConvert.ToString(Values[i]));
			}
			return Result.ToString();
		}
	}

	internal sealed class TExactToken : TBaseFunctionToken
	{
		internal TExactToken(ptg aId, TCellFunctionData aFuncData) : base(2, aId, aFuncData) { }
        
		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			object obj1 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (obj1 is TFlxFormulaErrorValue) return obj1;
			object obj2 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (obj2 is TFlxFormulaErrorValue) return obj2;
			string s1 = FlxConvert.ToString(obj1); if (obj1 is bool) { if ((bool)obj1) s1 = TFormulaMessages.TokenString(TFormulaToken.fmTrue); else TFormulaMessages.TokenString(TFormulaToken.fmFalse); }
			string s2 = FlxConvert.ToString(obj2); if (obj2 is bool) { if ((bool)obj2) s2 = TFormulaMessages.TokenString(TFormulaToken.fmTrue); else TFormulaMessages.TokenString(TFormulaToken.fmFalse); }
			return s1 == s2;
		}
	}

	internal sealed class TReptToken : TBaseFunctionToken
	{
		internal TReptToken(ptg aId, TCellFunctionData aFuncData) : base(2, aId, aFuncData) { }
        
		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			object v1 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v1 is TFlxFormulaErrorValue) return v1;
			TFlxFormulaErrorValue Err;
			int l; if (!GetUInt(v1, out l, out Err)) return Err;
			if (l < 0) return TFlxFormulaErrorValue.ErrValue;

			object obj2 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (obj2 is TFlxFormulaErrorValue) return obj2;
			string s2 = FlxConvert.ToString(obj2);
			StringBuilder sb = new StringBuilder(s2.Length * l + 1);
			for (int i = 0; i < l; i++)
				sb.Append(s2);
			return sb.ToString();
		}
	}

	internal sealed class TFixedToken : TBaseFunctionToken
	{
		internal TFixedToken(int ArgCount, ptg aId, TCellFunctionData aFuncData) : base(ArgCount, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			bool NoCommas = false;
			if (FArgCount > 2)
			{
				object v3 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v3 is TFlxFormulaErrorValue) return v3;
				if (!ExtToBool(v3, out NoCommas)) return TFlxFormulaErrorValue.ErrValue;
			}

			int Decimals = 2;
			if (FArgCount > 1)
			{
				object v2 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v2 is TFlxFormulaErrorValue) return v2;
				TFlxFormulaErrorValue Err;
				if (!GetInt(v2, out Decimals, out Err)) return Err;
			}

			if (Decimals > 127) return TFlxFormulaErrorValue.ErrValue;


			object v1 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v1 is TFlxFormulaErrorValue) return v1;
			double n; if (!GetDouble(v1, out n)) return TFlxFormulaErrorValue.ErrValue;

			int Decs = Decimals;
			if (Decs < 0) Decs = 0;
			if (Decs > 15) Decs = 15;

            n = UpRound(n, Decimals);
			string Sp = NoCommas ? "F" : "N";
			string Result = n.ToString(Sp + Decs.ToString(CultureInfo.InvariantCulture), CultureInfo.CurrentCulture);
			if (Decimals > Decs) Result += new string('0', Decimals - Decs); //up to 127 zeroes.
			return Result;
		}
	}

	internal sealed class TDollarToken : TBaseFunctionToken
	{
		internal TDollarToken(int ArgCount, ptg aId, TCellFunctionData aFuncData) : base(ArgCount, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			int Decimals = 2;
			if (FArgCount > 1)
			{
				object v2 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v2 is TFlxFormulaErrorValue) return v2;
				TFlxFormulaErrorValue Err;
				if (!GetInt(v2, out Decimals, out Err)) return Err;
			}

			if (Decimals > 127) return TFlxFormulaErrorValue.ErrValue;


			object v1 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v1 is TFlxFormulaErrorValue) return v1;
			double n; if (!GetDouble(v1, out n)) return TFlxFormulaErrorValue.ErrValue;

			int Decs = Decimals;
			if (Decs < 0) Decs = 0;
			if (Decs > 15) Decs = 15;

            n = UpRound(n, Decimals);

			CultureInfo ExcelProvider = (CultureInfo)CultureInfo.CurrentCulture.Clone();

            if (FlxConsts.ExcelVersion == TExcelVersion.v97_2003)
            {
                ExcelProvider.NumberFormat.CurrencyNegativePattern = 12; //Excel always seems to use this.
            }

			string Result = n.ToString("C" + Decs.ToString(CultureInfo.InvariantCulture), ExcelProvider);
			if (Decimals > Decs) Result += new string('0', Decimals - Decs); //up to 127 zeroes.
			return Result;
		}
	}

	internal sealed class TCleanToken : TOneStringArgToken
	{
		internal TCleanToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(string s)
		{
			StringBuilder sb = new StringBuilder(s);
			for (int i = s.Length - 1; i >= 0; i--)
				if (sb[i] < ' ') sb.Remove(i, 1);
			return sb.ToString();
		}
	}

	internal sealed class TSearchToken : TBaseFunctionToken
	{
		internal TSearchToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			object[,] ArrL = null;
			object[,] ArrWt = null;
			object[,] ArrFt = null;

			int l = 0;
			if (FArgCount > 2)
			{
				object Res2 = GetUIntArgument(FTokenList, wi, CalcState, CalcStack, ref l, out ArrL); if (Res2 != null) return Res2;
			}
			else l = 1;

			string WithinText = string.Empty;
			object Res = GetStringArgument(FTokenList, wi, CalcState, CalcStack, ref WithinText, out ArrWt); if (Res != null) return Res;
			if (WithinText != null) WithinText = WithinText.ToUpper(CultureInfo.CurrentCulture);

			string FindText = string.Empty;
			Res = GetStringArgument(FTokenList, wi, CalcState, CalcStack, ref FindText, out ArrFt); if (Res != null) return Res;
			if (FindText != null) FindText = FindText.ToUpper(CultureInfo.CurrentCulture);

			if (ArrL != null || ArrWt != null || ArrFt != null)
			{
				object[,] ArrResult;
				SetupArray(ref ArrL, ref ArrWt, ref ArrFt, l, WithinText, FindText, out ArrResult);

				for (int i = 0; i < ArrResult.GetLength(0); i++)
					for (int j = 0; j < ArrResult.GetLength(1); j++)
					{
						if (!GetUIntItem(ref ArrResult, ref ArrL, i, j, out l)) continue;
						if (!GetStringItem(ref ArrResult, ref ArrWt, i, j, out WithinText)) continue;
						WithinText = WithinText.ToUpper(CultureInfo.CurrentCulture);
						if (!GetStringItem(ref ArrResult, ref ArrFt, i, j, out FindText)) continue;
						FindText = FindText.ToUpper(CultureInfo.CurrentCulture);
						ArrResult[i, j] = DoSearch(FindText, WithinText, l);
					}
				return UnPack(ArrResult);
			}

			return DoSearch(FindText, WithinText, l);
		}

		private static object DoSearch(string FindText, string WithinText, int l)
		{
			if (l - 1 < 0 || l - 1 > WithinText.Length) return TFlxFormulaErrorValue.ErrValue;
			if (FindText.Length == 0) return (double)l;
			int i = TWildcardMatch.IndexOf(FindText, WithinText, l - 1);
			if (i >= 0) return (double)(i + 1); else return TFlxFormulaErrorValue.ErrValue;
		}
	}

	internal sealed class TTextToken : TBaseFunctionToken
	{
		internal TTextToken(ptg aId, TCellFunctionData aFuncData) : base(2, aId, aFuncData) { }
        
		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			object v1 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v1 is TFlxFormulaErrorValue) return v1;
			string format = FlxConvert.ToString(v1);

			object obj2 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (obj2 is TFlxFormulaErrorValue) return obj2;

			double d = 0;
			if (GetDouble(obj2, out d)) obj2 = d;
			Color mycolor = Colors.Black;
			return TFlxNumberFormat.FormatValue(obj2, format, ref mycolor, null).ToString();
		}
	}

	internal sealed class TTToken : TBaseFunctionToken
	{
		internal TTToken(ptg aId, TCellFunctionData aFuncData) : base(1, aId, aFuncData) { }
        
		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			object v1 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v1 is TFlxFormulaErrorValue) return v1;
			string Result = ConvertToAllowedObject(v1) as string;
			if (Result == null) return String.Empty;
			return Result;
		}
	}

	internal sealed class TNToken : TBaseFunctionToken
	{
		internal TNToken(ptg aId, TCellFunctionData aFuncData) : base(1, aId, aFuncData) { }
        
		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			object v1 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v1 is TFlxFormulaErrorValue) return v1;
			v1 = ConvertToAllowedObject(v1);

			if (v1 is bool)
			{
				if ((bool)v1) return 1;
				else return 0;
			}
			if (v1 is DateTime)
				return FlxDateTime.ToOADate((DateTime)v1, Dates1904);

			if (v1 is double) return v1;
			return 0;
		}
	}

	internal sealed class TValueToken : TBaseFunctionToken
	{
		internal TValueToken(ptg aId, TCellFunctionData aFuncData) : base(1, aId, aFuncData) { }
        
		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			object v1 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v1 is TFlxFormulaErrorValue) return v1;

			object[,] ArrResult;
			object[,] ArrValue;
			if (IsArrayArgument(v1, out ArrResult, out ArrValue))
			{
				for (int i = 0; i < ArrResult.GetLength(0); i++)
					for (int j = 0; j < ArrResult.GetLength(1); j++)
						ArrResult[i, j] = CalcValue(ArrValue[i, j]);

				return UnPack(ArrResult);
			}

			return CalcValue(v1);
		}

		private static object CalcValue(object v1)
		{
			string s1 = (v1 as String);
			if (s1 != null)
			{
				if (s1.Length == 0) return TFlxFormulaErrorValue.ErrValue;
				double Result = 0;
                if (TCompactFramework.ConvertToNumber(s1, CultureInfo.CurrentCulture, out Result)) return Result;  //Try to avoid unnecessary exceptions.
				DateTime DateResult;
				if (TCompactFramework.ConvertDateToNumber(s1, out DateResult))
				{
					return FlxDateTime.ToOADate(DateResult, Dates1904);
				}
				return TFlxFormulaErrorValue.ErrValue;
			}

			if (v1 is bool) return TFlxFormulaErrorValue.ErrValue;  //Here bool is not equal to 0/1
			if (v1 is DateTime) return FlxDateTime.ToOADate((DateTime)v1, Dates1904);
			double d = 0;
			if ((GetDouble(v1, out d))) return d;

			return TFlxFormulaErrorValue.ErrValue;
		}
	}

	#endregion

	#region Date and time
	internal sealed class TDateToken : TNDoubleArgToken
	{
		internal TDateToken(ptg aId, TCellFunctionData aFuncData) : base(3, aId, aFuncData) { }
        
		protected override object Calc(double[] x)
		{
			TFlxFormulaErrorValue Err;
			int d1; if (!GetInt(x[0], out d1, out Err)) return TFlxFormulaErrorValue.ErrNum; //Different from standard, -1 should return errnum
			int d2; if (!GetInt(x[1], out d2, out Err)) return TFlxFormulaErrorValue.ErrNum;
			int d3; if (!GetUInt(x[2], out d3, out Err)) return TFlxFormulaErrorValue.ErrNum;

			if (d3 < 1900) d3 += 1900;
			//We have a nice issue here!!!! Excel has january 29, 1900 but c# does not. Values from january1, 1900 to january 29 1900 will be different from excel.
			if (d3 < DateTime.MinValue.Year || d3 > DateTime.MaxValue.Year) return TFlxFormulaErrorValue.ErrNum;  //avoid exceptions.
			DateTime d = new DateTime(d3, 1, 1).AddMonths(d2 - 1).AddDays(d1 - 1);
			return FlxDateTime.ToOADate(d, Dates1904);
		}
	}

	internal sealed class TDateValueToken : TOneStringArgToken
	{
		internal TDateValueToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }

        protected override object Calc(string s)
        {
            DateTime d;
            if (!TCompactFramework.ConvertDateToNumber(s, out d)) return TFlxFormulaErrorValue.ErrValue;
            return Math.Floor(FlxDateTime.ToOADate(d, Dates1904));
        }
	}

	internal sealed class TDayToken : TOneDoubleArgToken
	{
		internal TDayToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(double x)
		{
			DateTime d = FlxDateTime.FromOADate(x, Dates1904);
			return (double)d.Day;
		}
	}

	internal sealed class TMonthToken : TOneDoubleArgToken
	{
		internal TMonthToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(double x)
		{
			DateTime d = FlxDateTime.FromOADate(x, Dates1904);
			return (double)d.Month;
		}
	}

	internal sealed class TYearToken : TOneDoubleArgToken
	{
		internal TYearToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(double x)
		{
			DateTime d = FlxDateTime.FromOADate(x, Dates1904);
			return (double)d.Year;
		}
	}

	internal sealed class TTimeToken : TNDoubleArgToken
	{
		internal TTimeToken(ptg aId, TCellFunctionData aFuncData) : base(3, aId, aFuncData) { }
        
		protected override object Calc(double[] x)
		{
			TFlxFormulaErrorValue Err;
			int d1; if (!GetUInt(x[0], out d1, out Err)) return TFlxFormulaErrorValue.ErrNum;
			int d2; if (!GetUInt(x[1], out d2, out Err)) return TFlxFormulaErrorValue.ErrNum;
			int d3; if (!GetUInt(x[2], out d3, out Err)) return TFlxFormulaErrorValue.ErrNum;

			DateTime d = new DateTime(1, 1, 1, (d3 + d2 / 60 + d1 / 3600) % 24, (d2 + d1 / 60) % 60, d1 % 60);
			double xx = FlxDateTime.ToOADate(d, Dates1904);
			return xx - Math.Floor(xx);
		}
	}

	internal sealed class TTimeValueToken : TOneStringArgToken
	{
		internal TTimeValueToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }

        protected override object Calc(string s)
        {
            DateTime d;
            if (!TCompactFramework.ConvertDateToNumber(s, out d)) return TFlxFormulaErrorValue.ErrValue;
            double x = FlxDateTime.ToOADate(d, Dates1904);
            return x - Math.Floor(x);
        }
	}

	internal sealed class THourToken : TOneDoubleArgToken
	{
		internal THourToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(double x)
		{
			DateTime d = FlxDateTime.FromOADate(x, Dates1904);
			return (double)d.Hour;
		}
	}

	internal sealed class TMinuteToken : TOneDoubleArgToken
	{
		internal TMinuteToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(double x)
		{
			DateTime d = FlxDateTime.FromOADate(x, Dates1904);
			return (double)d.Minute;
		}
	}

	internal sealed class TSecondToken : TOneDoubleArgToken
	{
		internal TSecondToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(double x)
		{
			DateTime d = FlxDateTime.FromOADate(x, Dates1904);
			return (double)d.Second;
		}
	}

	internal sealed class TNowToken : TBaseFunctionToken
	{
		internal TNowToken(ptg aId, TCellFunctionData aFuncData) : base(0, aId, aFuncData) {}
        
		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			DateTime d = DateTime.Now;
			return FlxDateTime.ToOADate(d, Dates1904);
		}
	}

	internal sealed class TTodayToken : TBaseFunctionToken
	{
		internal TTodayToken(ptg aId, TCellFunctionData aFuncData) : base(0, aId, aFuncData) {}
        
		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			DateTime d = DateTime.Now;
			return Math.Floor(FlxDateTime.ToOADate(d, Dates1904));
		}
	}
    
	internal sealed class TWeekDayToken : TNDoubleArgToken
	{
		internal TWeekDayToken(int ArgCount, ptg aId, TCellFunctionData aFuncData) : base(ArgCount, aId, aFuncData) { }


        protected override object Calc(double[] x)
        {
            int ResultType = 1;
            double d = x[0];
            if (x.Length > 1)
            {
                ResultType = (int)Math.Round(x[0]);
                if (ResultType < 1 || ResultType > 3) return TFlxFormulaErrorValue.ErrNum;

                d = x[1];
            }

            return CalcWeekDay(ResultType, d);
        }

        private static object CalcWeekDay(int ResultType, double d1)
        {
            DateTime d = FlxDateTime.FromOADate(d1, Dates1904);
            switch (ResultType)
            {
                case 1: return (double)d.DayOfWeek + 1;
                case 2: double Result = (double)d.DayOfWeek;
                    if (Result == 0) return 7; else return Result;

                case 3: double Result2 = (double)d.DayOfWeek;
                    if (Result2 == 0) return 6; else return Result2 - 1;


                default: return TFlxFormulaErrorValue.ErrNum;
            }
        }
	}

	internal sealed class TDays360Token : TBaseFunctionToken
	{
		internal TDays360Token(int ArgCount, ptg aId, TCellFunctionData aFuncData) : base(ArgCount, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			bool Eu360 = false;
			if (FArgCount > 2)
			{
				object v1 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v1 is TFlxFormulaErrorValue) return v1;
				if (!ExtToBool(v1, out Eu360)) return TFlxFormulaErrorValue.ErrValue;
			}

			object o2 = GetOneDate(FTokenList, wi, CalcState, CalcStack);
			if (o2 is TFlxFormulaErrorValue) return o2;
			DateTime d2 = (DateTime)o2;
			object o1 = GetOneDate(FTokenList, wi, CalcState, CalcStack);
			if (o1 is TFlxFormulaErrorValue) return o1;
			DateTime d1 = (DateTime)o1;

			int Months = (d2.Month - d1.Month) + (12 * (d2.Year - d1.Year));

			int EndDay = d2.Day;
			int StartDay = d1.Day;

			if (StartDay > 30)
				StartDay = 30;

			if (!Eu360 && d1.Month == 2)
				if (d1.Day == 29 || (d1.Day == 28 && !DateTime.IsLeapYear(d1.Year)))
				{
					StartDay = 30;
				}

            if (EndDay > 30)
            {
                if (!Eu360)
                {
                    if (StartDay >= 30)
                    {
                        //Months++;
                        //EndDay = 1;
                        EndDay = 30;
                    }
                }
                else EndDay = 30;
            }

			return Months * 30 + EndDay - StartDay;

		}

		private static object GetOneDate(TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack)
		{
			object v2 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v2 is TFlxFormulaErrorValue) return v2;
			string sv2 = v2 as string;
			if (sv2 != null)
			{
				DateTime d;
				if (!TCompactFramework.ConvertDateToNumber(sv2, out d)) return TFlxFormulaErrorValue.ErrValue;
				return d;
			}
			else
			{
				double d2; if (!ExtToDouble(v2, out d2)) return TFlxFormulaErrorValue.ErrValue;
				return FlxDateTime.FromOADate(d2, Dates1904);
			}
		}
	}


	#endregion

	#region Information
	internal abstract class TLogicalToken : TBaseFunctionToken
	{
		protected TLogicalToken(ptg aId, TCellFunctionData aFuncData) : base(1, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			object v1 = FTokenList.EvaluateToken(wi, CalcState, CalcStack);
			object[,] ArrResult;
			object[,] ArrValue;
			if (IsArrayArgument(v1, out ArrResult, out ArrValue))
			{
				for (int i = 0; i < ArrResult.GetLength(0); i++)
					for (int j = 0; j < ArrResult.GetLength(1); j++)
						ArrResult[i, j] = Calc(ArrValue[i, j]);

				return UnPack(ArrResult);
			}
			return Calc(v1);
		}

		protected abstract object Calc(object v1);

	}

	internal sealed class TErrorTypeToken : TLogicalToken
	{
		internal TErrorTypeToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(object v1)
		{
			if (v1 is TFlxFormulaErrorValue)
				switch ((TFlxFormulaErrorValue)v1)
				{
					case TFlxFormulaErrorValue.ErrNull: return (double)1;
					case TFlxFormulaErrorValue.ErrDiv0: return (double)2;
					case TFlxFormulaErrorValue.ErrValue: return (double)3;
					case TFlxFormulaErrorValue.ErrRef: return (double)4;
					case TFlxFormulaErrorValue.ErrName: return (double)5;
					case TFlxFormulaErrorValue.ErrNum: return (double)6;
					case TFlxFormulaErrorValue.ErrNA: return (double)7;
				}
			return TFlxFormulaErrorValue.ErrNA;
		}
	}

	internal sealed class TIsBlankToken : TLogicalToken
	{
		internal TIsBlankToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(object v1)
		{
			return v1 == null;
		}
	}

	internal sealed class TIsErrToken : TLogicalToken
	{
		internal TIsErrToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(object v1)
		{
			return (v1 is TFlxFormulaErrorValue) && ((TFlxFormulaErrorValue)v1 != TFlxFormulaErrorValue.ErrNA);
		}
	}

	internal sealed class TIsErrorToken : TLogicalToken
	{
		internal TIsErrorToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(object v1)
		{
			return v1 is TFlxFormulaErrorValue;
		}
	}

	internal sealed class TIsLogicalToken : TLogicalToken
	{
		internal TIsLogicalToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(object v1)
		{
			return (v1 is bool);
		}
	}

	internal sealed class TIsNAToken : TLogicalToken
	{
		internal TIsNAToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(object v1)
		{
			return (v1 is TFlxFormulaErrorValue) && ((TFlxFormulaErrorValue)v1 == TFlxFormulaErrorValue.ErrNA);
		}
	}

	internal sealed class TIsNonTextToken : TLogicalToken
	{
		internal TIsNonTextToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(object v1)
		{
			v1 = ConvertToAllowedObject(v1);
			return !(v1 is string);
		}
	}

	internal sealed class TIsNumberToken : TLogicalToken
	{
		internal TIsNumberToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(object v1)
		{
			v1 = ConvertToAllowedObject(v1);
			return (v1 is double);
		}
	}

	internal sealed class TIsRefToken : TBaseFunctionToken
	{
		internal TIsRefToken(ptg aId, TCellFunctionData aFuncData) : base(1, aId, aFuncData) { }
        
		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			object v1 = FTokenList.EvaluateTokenRef(wi, CalcState, CalcStack);
			//bug in excel? when there is more than one sheet on a ref formula, result is false.
			TAddress[] AdrList = (v1 as TAddress[]);
			if (AdrList != null && (AdrList[0].wi.Xls != AdrList[1].wi.Xls || AdrList[0].Sheet != AdrList[1].Sheet)) return false;

			return (v1 is TAddress || AdrList != null);
		}
	}

	internal sealed class TIsTextToken : TLogicalToken
	{
		internal TIsTextToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }
        
		protected override object Calc(object v1)
		{
			v1 = ConvertToAllowedObject(v1);
			return (v1 is string);
		}
	}

	internal sealed class TTypeToken : TBaseFunctionToken
	{
		internal TTypeToken(ptg aId, TCellFunctionData aFuncData) : base(1, aId, aFuncData) { }
        
		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			object v1;
			try
			{
				object z = FTokenList.EvaluateToken(wi, CalcState, CalcStack);
				if (z is Array) return 64;
				v1 = ConvertToAllowedObject(z);
			}
			catch (Exception)
			{
				return 16;
			}
			if (v1 is double || v1 == null) return 1;
			if (v1 is string) return 2;
			if (v1 is bool) return 4;
			return 16;
		}
	}

	internal sealed class TNaToken : TBaseFunctionToken
	{
		internal TNaToken(ptg aId, TCellFunctionData aFuncData) : base(0, aId, aFuncData) { }
        
		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			return TFlxFormulaErrorValue.ErrNA;
		}
	}

	internal sealed class TCellToken : TBaseFunctionToken
	{

		internal TCellToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			if (FArgCount < 2) return TFlxFormulaErrorValue.ErrNA; //this means for the last cell changed, and we will not implement this.

			TAddress CellRef1 = null, CellRef2 = null;
			object[,] values = null;
			object ret = GetRangeOrArray(FTokenList, wi, CalcState, CalcStack, out CellRef1, out CellRef2, out values, false);
			if (ret != null) return ret;

			if (CellRef1.wi.Xls != CellRef2.wi.Xls || CellRef1.Sheet != CellRef2.Sheet) return TFlxFormulaErrorValue.ErrValue;

			object v2 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v2 is TFlxFormulaErrorValue) return v2;
			string FuncType = FlxConvert.ToString(v2);
			if (FuncType == null) return TFlxFormulaErrorValue.ErrValue;
			FuncType = FuncType.ToLower(CultureInfo.CurrentCulture);

			switch (FuncType)
			{
				case "row":
					return (double)(CellRef1.Row);  //Row a1:a2 returns 1.
				case "col":
					return (double)(CellRef1.Col);
				case "address":

					string SheetName = String.Empty;
					if (wi.Xls != CellRef1.wi.Xls && CellRef1.BookName != null) 
						SheetName += TFormulaMessages.TokenChar(TFormulaToken.fmWorkbookOpen) + CellRef1.BookName + TFormulaMessages.TokenChar(TFormulaToken.fmWorkbookClose);
					if (wi.Xls != CellRef1.wi.Xls || wi.SheetIndexBase1 != CellRef1.Sheet) SheetName += CellRef1.wi.Xls.GetSheetName(CellRef1.Sheet);

					TCellAddress ca = new TCellAddress(SheetName, CellRef1.Row, CellRef1.Col, true, true);
                    if (wi.Xls.OptionsR1C1) return ca.CellRefR1C1(CellRef1.Row, CellRef1.Col); else return ca.CellRef;

				case "contents":
					return CellRef1.wi.Xls.GetCellValueAndRecalc(CellRef1.Sheet, CellRef1.Row, CellRef1.Col, CalcState, CalcStack);

				case "type":
					object v = CellRef1.wi.Xls.GetCellValueAndRecalc(CellRef1.Sheet, CellRef1.Row, CellRef1.Col, CalcState, CalcStack);
					if (v == null) return "b";
					if (v is string || v is TRichString) return "l";
					return "v";

				case "protect":
					int XF = CellRef1.wi.Xls.GetCellFormat(CellRef1.Sheet, CellRef1.Row, CellRef1.Col);
					TFlxFormat fmt = CellRef1.wi.Xls.GetFormat(XF);
					if (fmt.Locked) return 1; else return 0;

				case "prefix":
					int XF2 = -1;
                    object value = CellRef1.wi.Xls.GetCellValue(CellRef1.Sheet, CellRef1.Row, CellRef1.Col, ref XF2); //Here we don't want the recalculated value, but what really is in the cell. If it is a formula, we want a formula object, not its result.

					TFlxFormat fmt2 = CellRef1.wi.Xls.GetFormat(XF2);
					if (value is string || value is TRichString) //formulas with strings do not count.
					{
						switch (fmt2.HAlignment)
						{
							case THFlxAlignment.center_across_selection:
							case THFlxAlignment.center: return "^";			
							case THFlxAlignment.right: return "\"";
							case THFlxAlignment.fill: return "\\";
							default: return "'";
						}
					}
					return String.Empty;

				case "width":
					double w = Math.Round(CellRef1.wi.Xls.GetColWidth(CellRef1.Sheet, CellRef1.Col, true) / 256F - 0.75);
					if (w < 0) return 0;
					return w;
			}

			return TFlxFormulaErrorValue.ErrValue;

		}

	}

	#endregion

	#region Lookup
	internal class TChooseToken : TBaseFunctionToken
	{
		internal TChooseToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			object[] Values = new object[FArgCount];
			TArgType[] ArgTypes = new TArgType[FArgCount];
			for (int i = 1; i < ArgTypes.Length; i++)
			{
				ArgTypes[i] = TArgType.Object;
				Values[i] = null;
			}
			ArgTypes[0] = TArgType.UInt;
			Values[0] = 0;

			return DoArguments(FTokenList, wi, CalcState, CalcStack, FArgCount, ref Values, ArgTypes, true);
		}

        protected override object DoOneArg(TWorkbookInfo wi, object[] Values)
		{
			int index = (int)Values[0];
			if (index <= 0 || index > FArgCount - 1) return TFlxFormulaErrorValue.ErrValue;

			return Values[index];
		}
	}

	internal class TOffsetToken : TBaseFunctionToken
	{
		internal TOffsetToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData)
		{
		}

		private object CalcRange(TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack, out TWorkbookInfo wi1, out TWorkbookInfo wi2, ref int Sheet1, ref int Sheet2, ref int r1, ref int c1, ref int r2, ref int c2)
		{
			TFlxFormulaErrorValue Err;
			wi1 = null;
			wi2 = null;
			int Width = -1;
			if (FArgCount > 4)
			{
				object v1 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v1 is TFlxFormulaErrorValue) return v1;
				if (!GetUInt(v1, out Width, out Err)) return Err;
				if (Width <= 0) return TFlxFormulaErrorValue.ErrValue;
			}
			int Height = -1;
			if (FArgCount > 3)
			{
				object v2 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v2 is TFlxFormulaErrorValue) return v2;
				if (!GetUInt(v2, out Height, out Err)) return Err;
				if (Height <= 0) return TFlxFormulaErrorValue.ErrValue;
			}

			object v3 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v3 is TFlxFormulaErrorValue) return v3;
			int co;
			object[,] x3 = v3 as object[,];
			if (x3 != null && x3.GetLength(0) > 0 && x3.GetLength(1) > 0)
			{
				v3 = x3[0, 0]; //this is the way offset behaves. not excel standard, but...
			}

			if (!GetInt(v3, out co, out Err)) return Err;

			object v4 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v4 is TFlxFormulaErrorValue) return v4;
			int ro;
			object[,] x4 = v4 as object[,];
			if (x4 != null && x4.GetLength(0) > 0 && x4.GetLength(1) > 0)
			{
				v4 = x4[0, 0]; //this is the way offset behaves. not excel standard, but...
			}

			if (!GetInt(v4, out ro, out Err)) return Err;

			TAddress adr1 = null, adr2 = null; object[,] values = null;
			object ret = GetRangeOrArray(FTokenList, wi, CalcState, CalcStack, out adr1, out adr2, out values, false);
			if (ret != null) return ret;

			r1 = adr1.Row + ro; if (r1 < 1 || r1 > FlxConsts.Max_Rows + 1) return TFlxFormulaErrorValue.ErrRef;
			c1 = adr1.Col + co; if (c1 < 1 || c1 > FlxConsts.Max_Columns + 1) return TFlxFormulaErrorValue.ErrRef;

			r2 = adr2.Row + ro;
			if (Height > 0) r2 = r1 + Height - 1;
			if (r2 < 1 || r2 > FlxConsts.Max_Rows + 1) return TFlxFormulaErrorValue.ErrRef;

			c2 = adr2.Col + co;
			if (Width > 0) c2 = c1 + Width - 1;
			if (c2 < 1 || c2 > FlxConsts.Max_Columns + 1) return TFlxFormulaErrorValue.ErrRef;

			wi1 = adr1.wi;
			wi2 = adr2.wi;
			Sheet1 = adr1.Sheet;
			Sheet2 = adr2.Sheet;
			return null;
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			TWorkbookInfo wi1, wi2;
			int s1 = 0, s2 = 0, r1 = 0, r2 = 0, c1 = 0, c2 = 0;
			object err = CalcRange(FTokenList, wi, CalcState, CalcStack, out wi1, out wi2, ref s1, ref s2, ref r1, ref c1, ref r2, ref c2);
			if (err != null) return err;
			if (wi1.Xls != wi2.Xls) return TFlxFormulaErrorValue.ErrValue;

			if (wi.IsArrayFormula)
			{
				if (s1 != s2) return TFlxFormulaErrorValue.ErrValue;
				OrderRange(ref s1, ref s2, ref r1, ref c1, ref r2, ref c2);
				object[,] ArrResult = new object[r2 - r1 + 1, c2 - c1 + 1];
				for (int i = 0; i < ArrResult.GetLength(0); i++)
					for (int j = 0; j < ArrResult.GetLength(1); j++)
					{
						ArrResult[i, j] = wi1.Xls.GetCellValueAndRecalc(s1, r1 + i, c1 + j, CalcState, CalcStack);
                        if (CalcState.Aborted) return TFlxFormulaErrorValue.ErrNA;
                    }
				return UnPack(ArrResult);
			}
			return f.Agg(wi1, s1, s2, r1, c1, r2, c2, CalcState, CalcStack);
		}

		internal override object EvaluateRef(TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack)
		{
			TWorkbookInfo wi1, wi2;
			int s1 = 0, s2 = 0, r1 = 0, r2 = 0, c1 = 0, c2 = 0;
			object err = CalcRange(FTokenList, wi, CalcState, CalcStack, out wi1, out wi2, ref s1, ref s2, ref r1, ref c1, ref r2, ref c2);
			if (err != null) return err;

			if (wi1.Xls == wi2.Xls && s1 == s2 && r1 == r2 && c1 == c2) return new TAddress(wi1, s1, r1, c1);
			TAddress[] Result = new TAddress[2];

			Result[0] = new TAddress(wi1, s1, r1, c1);
			Result[1] = new TAddress(wi2, s2, r2, c2);
			return Result;
		}
	}

	internal abstract class TBaseLookupToken : TBaseFunctionToken
	{
		internal TBaseLookupToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			bool RangeLookup = true;
			if (FArgCount > 3)
			{
				object v1 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v1 is TFlxFormulaErrorValue) return v1;
				if (!ExtToBool(v1, out RangeLookup)) return TFlxFormulaErrorValue.ErrValue;
			}

			object v2 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v2 is TFlxFormulaErrorValue) return v2;
			TFlxFormulaErrorValue Err;
			int index; if (!GetInt(v2, out index, out Err) || index < 1) return Err;

			TAddress adr1 = null, adr2 = null; object[,] valArray = null;
			object ret = GetRangeOrArray(FTokenList, wi, CalcState, CalcStack, out adr1, out adr2, out valArray, TFlxFormulaErrorValue.ErrValue, true);
			if (ret != null) return ret;

			if (adr1 != null && (adr1.wi.Xls != adr2.wi.Xls || adr1.Sheet != adr2.Sheet)) return TFlxFormulaErrorValue.ErrRef;

			object v4 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v4 is TFlxFormulaErrorValue) return v4;

			if (valArray != null)
			{
				if (RangeLookup)
					return LookupBinary(valArray, v4, index);
				else
				{
					if (v4 == null || v4 is TMissingArg) return TFlxFormulaErrorValue.ErrNA;
					return LookupLinear(valArray, v4, index);
				}
			}
			else
			{
				if (RangeLookup)
					return LookupBinary(adr1.wi.Xls, adr1.Sheet, adr1.Row, adr1.Col, adr2.Row, adr2.Col, v4, index, CalcState, CalcStack);
				else
				{
					if (v4 == null || v4 is TMissingArg) return TFlxFormulaErrorValue.ErrNA;
					return LookupLinear(adr1.wi.Xls, adr1.Sheet, adr1.Row, adr1.Col, adr2.Row, adr2.Col, v4, index, CalcState, CalcStack);
				}
			}
		}

		protected abstract object LookupLinear(ExcelFile Xls, int Sheet, int Row1, int Col1, int Row2, int Col2, object SearchPattern, int index, TCalcState CalcState, TCalcStack CalcStack);
		protected abstract object LookupLinear(object[,] valArray, object SearchPattern, int index);

		protected abstract object LookupBinary(ExcelFile Xls, int Sheet, int Row1, int Col1, int Row2, int Col2, object SearchPattern, int index, TCalcState CalcState, TCalcStack CalcStack);
		protected abstract object LookupBinary(object[,] valArray, object SearchPattern, int index);
	}

	internal class THLookupToken : TBaseLookupToken
	{
		internal THLookupToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData)
		{
		}

		private static object CellValue(ExcelFile Xls, int Sheet, int Row1, int Row2, int index, ref TCalcState CalcState, TCalcStack CalcStack, int c)
		{
			if (Row1 + index - 1 > Row2) return TFlxFormulaErrorValue.ErrRef; //We must check this in place
			return Xls.GetCellValueAndRecalc(Sheet, Row1 + index - 1, c, CalcState, CalcStack);
		}

		private static object CellValue(object[,] valArray, int index, int c)
		{
			if (index - 1 >= valArray.GetLength(0) || index - 1 < 0) return TFlxFormulaErrorValue.ErrRef; //Must be checked here, to behave like Excel
			return valArray[index - 1, c];
		}

		protected override object LookupLinear(ExcelFile Xls, int Sheet, int Row1, int Col1, int Row2, int Col2, object SearchPattern, int index, TCalcState CalcState, TCalcStack CalcStack)
		{
			int x = 0, y = 0;
			OrderRange(ref x, ref y, ref Row1, ref Col1, ref Row2, ref Col2);

			int c2Index = Xls.ColToIndex(Sheet, Row1, Col2);
			for (int cIndex = Xls.ColToIndex(Sheet, Row1, Col1); cIndex <= c2Index; cIndex++)
			{
				int c = Xls.ColFromIndex(Sheet, Row1, cIndex);
				if (c < Col1 || c == 0) continue;
				if (c > Col2) break;
		
				object o = CompareValues(Xls.GetCellValueAndRecalc(Sheet, Row1, c, CalcState, CalcStack), SearchPattern);
				if (o is int)
				{
					int z = (int)o;
					if (z == 0)
					{
						return CellValue(Xls, Sheet, Row1, Row2, index, ref CalcState, CalcStack, c);
					}
				}

			}
			return TFlxFormulaErrorValue.ErrNA;
		}

		protected override object LookupLinear(object[,] valArray, object SearchPattern, int index)
		{
			for (int c = 0; c < valArray.GetLength(1); c++)
			{
				object o = CompareValues(valArray[0, c], SearchPattern);
				if (o is int)
				{
					int z = (int)o;
					if (z == 0)
					{
						return CellValue(valArray, index, c);
					}
				}

			}
			return TFlxFormulaErrorValue.ErrNA;
		}

		protected override object LookupBinary(ExcelFile Xls, int Sheet, int Row1, int Col1, int Row2, int Col2, object SearchPattern, int index, TCalcState CalcState, TCalcStack CalcStack)
		{
			int x = 0, y = 0;
			OrderRange(ref x, ref y, ref Row1, ref Col1, ref Row2, ref Col2);
			bool Found;
			return LookupBinary(Xls, Sheet, Row1, Col1, Row2, Col2, SearchPattern, index, CalcState, CalcStack, out Found);
		}
        
		private object LookupBinary(ExcelFile Xls, int Sheet, int Row1, int Col1, int Row2, int Col2, object SearchPattern, int index, TCalcState CalcState, TCalcStack CalcStack, out bool Found)
		{

			int N = Col2 - Col1 + 1;
			int low = 0;
			int high = N;

			while (low < high)
			{
				int mid = (low + high) / 2;

				object Cell = Xls.GetCellValueAndRecalc(Sheet, Row1, Col1 + mid, CalcState, CalcStack);
				object o = Cell==null? null: CompareValues(Cell, SearchPattern);

				if (!(o is int))  //z is  tflxformulaerror or null.
				{
					object o2 = LookupBinary(Xls, Sheet, Row1, Col1 + mid + 1, Row2, Col2, SearchPattern, index, CalcState, CalcStack, out Found);
					if (Found) return o2;
					return LookupBinary(Xls, Sheet, Row1, Col1, Row2, Col1 + mid - 1, SearchPattern, index, CalcState, CalcStack, out Found);
				}

				int z = (int)o;
				if (z < 0)
					low = mid + 1;
				else
				{
					if (z == 0) // We check here, since getting the value is a slow operation.
					{
						Found = true;
                        return CellValue(Xls, Sheet, Row1, Row2, index, ref CalcState, CalcStack, Col1 + mid);
					}
					high = mid;
				}
			}


			low--;
			Found = low >= 0;  //Found means that the value is in the interval, not an exact match. We need it because checking against TFlxFormulaErrorValue.ErrNA is not enough to know if there was a match or not.
			if (Found && SameType(Xls.GetCellValueAndRecalc(Sheet, Row1, Col1 + low, CalcState, CalcStack), SearchPattern))
			{
                return CellValue(Xls, Sheet, Row1, Row2, index, ref CalcState, CalcStack, Col1 + low);
			}
			else
				return TFlxFormulaErrorValue.ErrNA; // not found     

		}

		protected override object LookupBinary(object[,] valArray, object SearchPattern, int index)
		{
			bool Found;
			return LookupBinary(valArray, 0, valArray.GetLength(1), SearchPattern, index, out Found);
		}

		private object LookupBinary(object[,] valArray, int low, int high, object SearchPattern, int index, out bool Found)
		{
			while (low < high)
			{
				int mid = (low + high) / 2;

				object o = CompareValues(valArray[0, mid], SearchPattern);
				if (!(o is int))  //z is  tflxformulaerror or null.
				{
					object o2 = LookupBinary(valArray, mid + 1, high, SearchPattern, index, out Found);
					if (Found) return o2;
					return LookupBinary(valArray, low, mid - 1, SearchPattern, index, out Found);
				}

				int z = (int)o;
				if (z < 0)
					low = mid + 1;
				else
				{
					if (z == 0)
					{
						Found = true;
						return CellValue(valArray, index, mid);
					}
					high = mid;
				}
			}

			low--;
			Found = low >= 0;
			if (Found && SameType(valArray[0,low], SearchPattern))
			{
				return CellValue(valArray, index, low);
			}
			else
				return TFlxFormulaErrorValue.ErrNA; // not found     
		}
	}

	internal class TVLookupToken : TBaseLookupToken
	{
		internal TVLookupToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData)
		{
		}

		private static object CellValue(ExcelFile Xls, int Sheet, int Col1, int Col2, int index, ref TCalcState CalcState, TCalcStack CalcStack, int r)
		{
			if (Col1 + index - 1 > Col2) return TFlxFormulaErrorValue.ErrRef; //We must check this in place
			return Xls.GetCellValueAndRecalc(Sheet, r, Col1 + index - 1, CalcState, CalcStack);
		}

		private static object CellValue(object[,] valArray, int index, int r)
		{
			if (index - 1 >= valArray.GetLength(1) || index - 1 < 0) return TFlxFormulaErrorValue.ErrRef; //Must be checked here, to behave like Excel
			return valArray[r, index - 1];
		}


		protected override object LookupLinear(ExcelFile Xls, int Sheet, int Row1, int Col1, int Row2, int Col2, object SearchPattern, int index, TCalcState CalcState, TCalcStack CalcStack)
		{
			int x = 0, y = 0;
			OrderRange(ref x, ref y, ref Row1, ref Col1, ref Row2, ref Col2);

			int R2 = Math.Min(Row2, Xls.GetRowCount(Sheet));
			for (int r = Row1; r <= R2; r++)
			{
				object o = CompareValues(Xls.GetCellValueAndRecalc(Sheet, r, Col1, CalcState, CalcStack), SearchPattern);
				if (o is int)
				{
					int z = (int)o;
					if (z == 0)
					{
                        return CellValue(Xls, Sheet, Col1, Col2, index, ref CalcState, CalcStack, r);
					}
				}

			}
			return TFlxFormulaErrorValue.ErrNA;
		}

		protected override object LookupLinear(object[,] valArray, object SearchPattern, int index)
		{
			for (int r = 0; r < valArray.GetLength(0); r++)
			{
				object o = CompareValues(valArray[r, 0], SearchPattern);
				if (o is int)
				{
					int z = (int)o;
					if (z == 0)
					{
						return CellValue(valArray, index, r);
					}
				}

			}
			return TFlxFormulaErrorValue.ErrNA;
		}

		protected override object LookupBinary(ExcelFile Xls, int Sheet, int Row1, int Col1, int Row2, int Col2, object SearchPattern, int index, TCalcState CalcState, TCalcStack CalcStack)
		{
			int x = 0, y = 0;
			OrderRange(ref x, ref y, ref Row1, ref Col1, ref Row2, ref Col2);
			bool Found;
			return LookupBinary(Xls, Sheet, Row1, Col1, Row2, Col2, SearchPattern, index, CalcState, CalcStack, out Found);
		}
        
		private object LookupBinary(ExcelFile Xls, int Sheet, int Row1, int Col1, int Row2, int Col2, object SearchPattern, int index, TCalcState CalcState, TCalcStack CalcStack, out bool Found)
		{
			int N = Row2 - Row1 + 1;
			int low = 0;
			int high = N;

			while (low < high)
			{
				int mid = (low + high) / 2;

				object Cell = Xls.GetCellValueAndRecalc(Sheet, Row1 + mid, Col1, CalcState, CalcStack);
				object o = Cell==null? null: CompareValues(Cell, SearchPattern);

				if (!(o is int))  //z is  tflxformulaerror or null.
				{
					object o2 = LookupBinary(Xls, Sheet, Row1 + mid + 1, Col1, Row2, Col2, SearchPattern, index, CalcState, CalcStack, out Found);
					if (Found) return o2;
					return LookupBinary(Xls, Sheet, Row1, Col1, Row1 + mid - 1, Col2, SearchPattern, index, CalcState, CalcStack, out Found);
				}
                
				int z = (int)o;

				if (z < 0)
					low = mid + 1;
				else
				{
					if (z == 0) // We check here, since getting the value is a slow operation.
					{
						Found = true;
                        return CellValue(Xls, Sheet, Col1, Col2, index, ref CalcState, CalcStack, Row1 + mid);
					}
					high = mid;
				}
			}


			low--;
			Found = low >= 0;
			if (Found && SameType(Xls.GetCellValueAndRecalc(Sheet, Row1 + low, Col1, CalcState, CalcStack), SearchPattern))
			{
                return CellValue(Xls, Sheet, Col1, Col2, index, ref CalcState, CalcStack, Row1 + low);
			}
			else
				return TFlxFormulaErrorValue.ErrNA; // not found     

		}


		protected override object LookupBinary(object[,] valArray, object SearchPattern, int index)
		{
			bool Found;
			return LookupBinary(valArray, 0, valArray.GetLength(0), SearchPattern, index, out Found);
		}

		private object LookupBinary(object[,] valArray, int low, int high, object SearchPattern, int index, out bool Found)
		{
			while (low < high)
			{
				int mid = (low + high) / 2;

				object o = CompareValues(valArray[mid, 0], SearchPattern);
				if (!(o is int))  //z is  tflxformulaerror or null.
				{
					object o2 = LookupBinary(valArray, mid + 1, high, SearchPattern, index, out Found);
					if (Found) return o2;
					return LookupBinary(valArray, low, mid - 1, SearchPattern, index, out Found);
				}

				int z = (int)o;
				if (z < 0)
					low = mid + 1;
				else
				{
					if (z == 0)
					{
						Found = true;
						return CellValue(valArray, index, mid);
					}
					high = mid;
				}
			}

			low--;
			Found = low >= 0;
			if (Found && SameType(valArray[low, 0], SearchPattern))
			{
				return CellValue(valArray, index, low);
			}
			else
				return TFlxFormulaErrorValue.ErrNA; // not found     
		}

	}

	internal class TIndexToken : TBaseFunctionToken
	{
		internal TIndexToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
            return EvaluateValOrRef(FTokenList, wi, f, CalcState, ref CalcStack, false);
		}

        internal override object EvaluateRef(TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack)
        {
            return EvaluateValOrRef(FTokenList, wi, TErr2Aggregate.Instance, CalcState, ref CalcStack, true);
        }

        private object EvaluateValOrRef(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, ref TCalcStack CalcStack, bool IsRef)
        {
            int Area = 1;
            if (FArgCount > 3)
            {
                object v3 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v3 is TFlxFormulaErrorValue) return v3;
                TFlxFormulaErrorValue Err;
                if (!GetUInt(v3, out Area, out Err)) return Err;
            }

            int Col = -1;
            if (FArgCount > 2)
            {
                object v2 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v2 is TFlxFormulaErrorValue) return v2;
                TFlxFormulaErrorValue Err;
                if (!GetUInt(v2, out Col, out Err)) return Err;
            }

            int Row = -1;
            object v1 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v1 is TFlxFormulaErrorValue) return v1;
            if (!(v1 is TMissingArg))
            {
                TFlxFormulaErrorValue Err;
                if (!GetUInt(v1, out Row, out Err)) return Err;
            }

            TAddressList adr = null; object[,] arr = null;
            object ret = GetRangeOrArray(FTokenList, wi, CalcState, CalcStack, out adr, out arr, TFlxFormulaErrorValue.ErrRef, true);
            if (ret is TFlxFormulaErrorValue) return ret;

            if (arr != null)
            {
                if (Area != 1 || IsRef)
                    return TFlxFormulaErrorValue.ErrRef;
                if (Col <= 0)
                {
                    if (Row <= 0) Col = 1;
                    else
                    {
                        return SliceH(arr, Row);
                    }
                }

                if (Row <= 0)
                {
                    return SliceV(arr, Col);
                }

                if (Row > arr.GetLength(0)) return TFlxFormulaErrorValue.ErrRef;
                if (Col > arr.GetLength(1)) return TFlxFormulaErrorValue.ErrRef;

                return arr[Row - 1, Col - 1];
            }

            if (Area < 1 || Area > adr.Count)
                return TFlxFormulaErrorValue.ErrRef;

            Area = adr.Count - Area;
            TAddress adr1 = adr[Area][0];
            TAddress adr2 = adr[Area][1];

            if (adr1.wi.Xls != adr2.wi.Xls || adr1.Sheet != adr2.Sheet) return TFlxFormulaErrorValue.ErrRef;

            if (Col < 0) //ommited
            {
                if (adr1.Col != adr2.Col)
                {
                    if (adr1.Row != adr2.Row || Row < 0) return TFlxFormulaErrorValue.ErrRef;
                    //Col and Row are switched.
                    Col = Row;
                    Row = 1;
                }
                else
                {
                    Col = 1;
                }
            }

            if (Row < 0)
            {
                //if (adr1.Row != adr2.Row)
                //    return TFlxFormulaErrorValue.ErrRef;
                Row = 0;
            }

            if (Row > adr2.Row - adr1.Row + 1) return TFlxFormulaErrorValue.ErrRef;
            if (Col > adr2.Col - adr1.Col + 1) return TFlxFormulaErrorValue.ErrRef;

            if (Row > 0 && Col > 0)
            {
                int r = adr1.Row + Row - 1;
                int c = adr1.Col + Col - 1;

                if (IsRef) return new TAddress(adr1.wi, adr1.wi.SheetIndexBase1, r, c);

                return adr1.wi.Xls.GetCellValueAndRecalc(adr1.Sheet, r, c, CalcState, CalcStack);
            }
            if (Row > 0)
            {
                int r = adr1.Row + Row - 1;
                adr1.Row = r; adr2.Row = r;
            }

            if (Col > 0)
            {
                int c = adr1.Col + Col - 1;
                adr1.Col = c; adr2.Col = c;
            }

            if (IsRef) return new TAddress[] { adr1, adr2 };
            return f.Agg(adr1.wi, adr1.Sheet, adr2.Sheet, adr1.Row, adr1.Col, adr2.Row, adr2.Col, CalcState, CalcStack);
        }

        private object SliceH(object[,] arr, int Row)
        {
            if (Row > arr.GetLength(0)) return TFlxFormulaErrorValue.ErrRef;
            object[,] Result = new object[1, arr.GetLength(1)];
            for (int c = 0; c < Result.Length; c++)
            {
                Result[0, c] = arr[Row - 1, c];
            }

            return Result;
        }

        private object SliceV(object[,] arr, int Col)
        {
            if (Col > arr.GetLength(1)) return TFlxFormulaErrorValue.ErrRef;
            object[,] Result = new object[arr.GetLength(0), 1];
            for (int r = 0; r < Result.Length; r++)
            {
                Result[r, 0] = arr[r, Col - 1];
            }

            return Result;
        }

	}

	internal class TMatchToken : TBaseFunctionToken
	{
		internal TMatchToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			int MatchType = 1;
			object[,] ArrMatchType = null;
			object[,] Arrv4 = null;
			bool MatchTypeIsArray = false;

			if (FArgCount > 2)
			{
				object v1 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v1 is TFlxFormulaErrorValue) return v1;
				MatchTypeIsArray = IsArrayArgument(v1, out ArrMatchType);
				if (!MatchTypeIsArray)
				{
					TFlxFormulaErrorValue Err;
					if (!GetInt(v1, out MatchType, out Err)) return Err;
				}
			}

			TAddress adr1 = null, adr2 = null; object[,] valArray = null;
			object ret = GetRangeOrArray(FTokenList, wi, CalcState, CalcStack, out adr1, out adr2, out valArray, TFlxFormulaErrorValue.ErrNA, true);
			if (ret != null) return ret;

			if (adr1 != null && (adr1.wi.Xls != adr2.wi.Xls || adr1.Sheet != adr2.Sheet)) return TFlxFormulaErrorValue.ErrRef;

			object v4 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v4 is TFlxFormulaErrorValue) return v4;

			if (IsArrayArgument(v4, out Arrv4) || ArrMatchType != null) //Array formula
			{
				FillArray(ref ArrMatchType, MatchType);
				FillArray(ref Arrv4, v4);
				if (!CompatibleDimensions(Arrv4, ArrMatchType)) return TFlxFormulaErrorValue.ErrNA;
				object[,] ArrResult = GetArrayObj(Arrv4, ArrMatchType);

				if (valArray != null)
					for (int i = 0; i < ArrResult.GetLength(0); i++)
						for (int j = 0; j < ArrResult.GetLength(1); j++)
						{
							v4 = GetItem(Arrv4, i, j); if (v4 is TFlxFormulaErrorValue) { ArrResult[i, j] = v4; continue; }
							object v1 = GetItem(ArrMatchType, i, j); if (v1 is TFlxFormulaErrorValue) { ArrResult[i, j] = v1; continue; }
							TFlxFormulaErrorValue Err;
							if (!GetInt(v1, out MatchType, out Err)) { ArrResult[i, j] = Err; continue; }
							ArrResult[i, j] = ConvertToAllowedObject(Match(valArray, v4, MatchType));
						}
				else
					for (int i = 0; i < ArrResult.GetLength(0); i++)
						for (int j = 0; j < ArrResult.GetLength(1); j++)
						{
							v4 = GetItem(Arrv4, i, j); if (v4 is TFlxFormulaErrorValue) { ArrResult[i, j] = v4; continue; }
							object v1 = GetItem(ArrMatchType, i, j); if (v1 is TFlxFormulaErrorValue) { ArrResult[i, j] = v1; continue; }
							TFlxFormulaErrorValue Err;
							if (!GetInt(v1, out MatchType, out Err)) { ArrResult[i, j] = Err; continue; }
							ArrResult[i, j] = ConvertToAllowedObject(Match(adr1.wi.Xls, adr1.Sheet, adr1.Row, adr1.Col, adr2.Row, adr2.Col, v4, MatchType, CalcState, CalcStack));
						}

				return UnPack(ArrResult);
			}

			if (valArray != null)
				return Match(valArray, v4, MatchType);
			else
				return Match(adr1.wi.Xls, adr1.Sheet, adr1.Row, adr1.Col, adr2.Row, adr2.Col, v4, MatchType, CalcState, CalcStack);
		}

		internal static object Match(ExcelFile Xls, int Sheet, int Row1, int Col1, int Row2, int Col2, object SearchPattern, int MatchType, TCalcState CalcState, TCalcStack CalcStack)
		{
			int x = 0, y = 0;
			OrderRange(ref x, ref y, ref Row1, ref Col1, ref Row2, ref Col2);

			if (Row1 != Row2 && Col1 != Col2) return TFlxFormulaErrorValue.ErrNA;
			int a = Row1 != Row2 ? Row1 : Col1;
			int b = Row1 != Row2 ? Row2 : Col2;
			if (a - 1 > b) return TFlxFormulaErrorValue.ErrRef;

			int r1 = 0;
			for (int r = a; r <= b; r++)
			{
				object v = Row2 != Row1 ?
					Xls.GetCellValueAndRecalc(Sheet, r, Col1, CalcState, CalcStack) :
					Xls.GetCellValueAndRecalc(Sheet, Row1, r, CalcState, CalcStack);
				if (v != null)
				{
					object o = CompareWithWildcards(v, SearchPattern, MatchType == 0);
					if (o is int)
					{
						int z = (int)o;
						if (z == 0 && MatchType == 0)
							return r - a + 1;

						if ((MatchType > 0 && z > 0) || (MatchType < 0 && z < 0))
						{
							if (r1 > 0) return r1;
							else return TFlxFormulaErrorValue.ErrNA;
						}

						r1 = r - a + 1;
					}
				}

			}

			if (r1 > 0 && MatchType != 0) 
			{
				object v0 = Row2 != Row1 ?
					Xls.GetCellValueAndRecalc(Sheet, Row1 + r1 - 1, Col1, CalcState, CalcStack) :
					Xls.GetCellValueAndRecalc(Sheet, Row1, Col1 + r1 - 1, CalcState, CalcStack);

				bool ValuesAreSameType = SameType(v0, SearchPattern);

				if (ValuesAreSameType) return r1;
			}
			return TFlxFormulaErrorValue.ErrNA;
		}

		internal static object Match(object[,] valArray, object SearchPattern, int MatchType)
		{
			if (valArray.GetLength(0) > 1 && valArray.GetLength(1) > 1) return TFlxFormulaErrorValue.ErrNA;
			int a = valArray.GetLength(0) > 1 ? 0 : 1;

			for (int r = 0; r < valArray.GetLength(a); r++)
			{
				object o = a == 0 ?
					CompareWithWildcards(valArray[r, 0], SearchPattern, MatchType == 0) :
					CompareWithWildcards(valArray[0, r], SearchPattern, MatchType == 0);

				if (o is int)
				{
					int z = (int)o;
					if (z == 0 && MatchType == 0)
						return r + 1;

					if ((MatchType > 0 && z > 0) || (MatchType < 0 && z < 0))
					{
						if (r > 0) return r;
						else return TFlxFormulaErrorValue.ErrNA;
					}
				}

			}
			if ((MatchType != 0) && valArray.GetLength(a) > 1) return valArray.GetLength(a);
			return TFlxFormulaErrorValue.ErrNA;
		}

	}

	internal class TLookupToken : TBaseFunctionToken
	{
		internal TLookupToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			TAddress adr21 = null, adr22 = null; object[,] valArray2 = null;
			if (FArgCount > 2)
			{
				object ret2 = GetRangeOrArray(FTokenList, wi, CalcState, CalcStack, out adr21, out adr22, out valArray2, TFlxFormulaErrorValue.ErrNA, true);
				if (ret2 != null) return ret2;
				if (adr21 != null && (adr21.wi.Xls != adr22.wi.Xls || adr21.Sheet != adr22.Sheet)) return TFlxFormulaErrorValue.ErrValue;
			}

			TAddress adr11 = null, adr12 = null; object[,] valArray1 = null;
			object ret1 = GetRangeOrArray(FTokenList, wi, CalcState, CalcStack, out adr11, out adr12, out valArray1, TFlxFormulaErrorValue.ErrNA, true);
			if (ret1 != null) return ret1;

			if (adr11 != null && (adr11.wi.Xls != adr12.wi.Xls || adr11.Sheet != adr12.Sheet)) return TFlxFormulaErrorValue.ErrValue;

			if (FArgCount <= 2) //Missing the last argument, we need to get both from the second.
			{
				SplitArrays(ref valArray1, ref valArray2);
				SplitAddresses(ref adr11, ref adr12, ref adr21, ref adr22);
			}

			object v4 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v4 is TFlxFormulaErrorValue) return v4;

			object index = null;
			if (valArray1 != null)
				index = TMatchToken.Match(valArray1, v4, 1);
			else
				index = TMatchToken.Match(adr11.wi.Xls, adr11.Sheet, adr11.Row, adr11.Col, adr12.Row, adr12.Col, v4, 1, CalcState, CalcStack);

			if (index is TFlxFormulaErrorValue) return index;
			if (!(index is int)) return TFlxFormulaErrorValue.ErrValue;

			int idx = ((int)index) - 1;
			if (valArray2 != null)
			{
				bool SecondDim = false;
				int len = valArray2.GetLength(0);
				if (len <= 1)
				{
					len = valArray2.GetLength(1);
					SecondDim = true;
				}
				else
				{
					if (valArray2.GetLength(1) != 1) return TFlxFormulaErrorValue.ErrNA;
				}

				if (idx < 0 || idx >= len) return TFlxFormulaErrorValue.ErrNA;
				if (SecondDim)
					return valArray2[0, idx];
				else
					return valArray2[idx, 0];
			}

			int aRow = adr21.Row;
			int aCol = adr21.Col;
			if (adr21.Row == adr22.Row) aCol += idx;
			else
				if (adr21.Col == adr22.Col) aRow += idx;
			else
				return TFlxFormulaErrorValue.ErrNA;

			return adr21.wi.Xls.GetCellValueAndRecalc(adr21.Sheet, aRow, aCol, CalcState, CalcStack);
		}

		private static void SplitArrays(ref object[,] valArray1, ref object[,] valArray2)
		{
			if (valArray1 == null) return; //Not an array.
			int dim = 1;
			int LowDimLen = valArray1.GetLength(0);
			if (LowDimLen >= valArray1.GetLength(1))
			{
				dim = 0;
				LowDimLen = valArray1.GetLength(1);
			}

			object[,] Tmp = new object[1, valArray1.GetLength(dim)];
			valArray2 = new object[1, valArray1.GetLength(dim)];

			if (dim == 0)
			{
				for (int i = 0; i < valArray1.GetLength(dim); i++)
				{
					Tmp[0, i] = valArray1[i, 0];
					valArray2[0, i] = valArray1[i, LowDimLen - 1];
				}
			}
			else
			{
				for (int i = 0; i < valArray1.GetLength(dim); i++)
				{
					Tmp[0, i] = valArray1[0, i];
					valArray2[0, i] = valArray1[LowDimLen - 1, i];
				}
			}

			valArray1 = Tmp;
		}

		private static void SplitAddresses(ref TAddress adr11, ref TAddress adr12, ref TAddress adr21, ref TAddress adr22)
		{
			if (adr11 == null) return; //not an address
			if (Math.Abs(adr12.Row - adr11.Row) >= Math.Abs(adr12.Col - adr11.Col))
			{
				adr21 = new TAddress(adr11.wi, adr11.BookName, adr11.Sheet, adr11.Row, adr12.Col);
				adr22 = new TAddress(adr11.wi, adr11.BookName, adr11.Sheet, adr12.Row, adr12.Col);
				adr12.Col = adr11.Col;
			}
			else
			{
				adr21 = new TAddress(adr11.wi, adr11.BookName, adr11.Sheet, adr12.Row, adr11.Col);
				adr22 = new TAddress(adr11.wi, adr11.BookName, adr11.Sheet, adr12.Row, adr12.Col);
				adr12.Row = adr11.Row;
			}
		}
	}

	internal sealed class TRankToken : TBaseFunctionToken
	{
		internal TRankToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			int Order = 0;
			if (FArgCount > 2)
			{
				object v1 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v1 is TFlxFormulaErrorValue) return v1;
				TFlxFormulaErrorValue Err;
				if (!GetInt(v1, out Order, out Err)) return Err;
			}

			TAddressList adr = null; object[,] valArray = null;
			object ret = GetRangeOrArray(FTokenList, wi, CalcState, CalcStack, out adr, out valArray, TFlxFormulaErrorValue.ErrNA, false);
			if (ret != null) return ret;

			if (adr == null) return TFlxFormulaErrorValue.ErrRef;

			double Number;
			object v4 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v4 is TFlxFormulaErrorValue) return v4;
			if (!GetDouble(v4, out Number)) return TFlxFormulaErrorValue.ErrValue;

			bool Found = false;
			int Result = 1;

			for (int i = 0; i < adr.Count; i++)
			{
				TAddress[] adr1 = (TAddress[])adr[i];
				if (adr1.Length != 2) continue;
				if (adr1[0].wi.Xls != adr1[1].wi.Xls) return TFlxFormulaErrorValue.ErrRef;
				Rank(adr1[0].wi.Xls, adr1[0].Sheet, adr1[1].Sheet, adr1[0].Row, adr1[0].Col, adr1[1].Row, adr1[1].Col, Number, Order, CalcState, CalcStack, ref Found, ref Result);
			}
			if (Found) return Result;
			return TFlxFormulaErrorValue.ErrNA;
		}

		internal static void Rank(ExcelFile Xls, int Sheet1, int Sheet2, int Row1, int Col1, int Row2, int Col2, double Number, int Order, TCalcState CalcState, TCalcStack CalcStack, ref bool Found, ref int Result)
		{
			OrderRange(ref Sheet1, ref Sheet2, ref Row1, ref Col1, ref Row2, ref Col2);
			for (int Sheet = Sheet1; Sheet <= Sheet2; Sheet++)
			{
				int MaxRow = Xls.GetRowCount(Sheet);
				for (int r = Row1; r <= Row2; r++)
				{
					if (r > MaxRow) break;

					for (int cIndex = Xls.ColToIndex(Sheet, r, Col2); cIndex > 0; cIndex--)
					{
						int c = Xls.ColFromIndex(Sheet, r, cIndex);
						if (c > Col2 || c == 0) continue;
						if (c < Col1) break;

						object o = ConvertToAllowedObject(Xls.GetCellValueAndRecalc(Sheet, r, c, CalcState, CalcStack));
						if (o is double)
						{
							double z = (double)o;
							if (z == Number) Found = true;
							else
							{
								if (z < Number && Order != 0) Result++;
								if (z > Number && Order == 0) Result++;
							}

						}
					}
				}
			}
		}
	}

    internal sealed class TPercentRankToken : TBaseFunctionToken
    {
        internal TPercentRankToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
            : base(ArgCount, aId, aFuncData)
        {
        }

        internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
        {
            int Significance = 3;
            if (FArgCount > 2)
            {
                object v1 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v1 is TFlxFormulaErrorValue) return v1;
                TFlxFormulaErrorValue Err;
                if (!GetInt(v1, out Significance, out Err)) return Err;
                if (Significance < 1) return TFlxFormulaErrorValue.ErrNum;
            }

            double Number;
            object v4 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v4 is TFlxFormulaErrorValue) return v4;
            if (!GetDouble(v4, out Number)) return TFlxFormulaErrorValue.ErrValue;

            return EvaluatePercentRank(Number, Significance, FTokenList, wi, new TPercentRankAggregate(Number), CalcState, CalcStack);

        }

        private object EvaluatePercentRank(double NumberToCompare, double Significance, TParsedTokenList FTokenList, TWorkbookInfo wi, TPercentRankAggregate f, TCalcState CalcState, TCalcStack CalcStack)
        {
            double Result = 0;
            object v1 = FTokenList.EvaluateToken(wi, f, CalcState, CalcStack);
            object[,] arr1 = v1 as object[,];
            if (arr1 != null)
                v1 = f.AggArray(arr1);


            if (v1 is TFlxFormulaErrorValue) return v1;

            if (v1 != null)
            {
                TPercentRankValue pv = v1 as TPercentRankValue;
                if (pv == null) return TFlxFormulaErrorValue.ErrNA;
                return pv.CalcPercentRank(NumberToCompare, Significance);
            }


            return Result;
        }

    }

	internal sealed class TAddressToken : TBaseFunctionToken
	{
		internal TAddressToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			object[] Values ={ 0, 0, 1, true, String.Empty };
			return DoArguments(FTokenList, wi, CalcState, CalcStack, FArgCount, ref Values, new TArgType[] { TArgType.Int, TArgType.Int, TArgType.Int, TArgType.Boolean, TArgType.String }, false);
		}

		protected override object DoOneArg(TWorkbookInfo wi, object[] Values)
		{
			return CalcAddress(wi, (int)Values[0], (int)Values[1], (int)Values[2], (bool)Values[3], (string)Values[4]);
		}


		private static object CalcAddress(TWorkbookInfo wi, int Row, int Col, int AbsNum, bool NotRC, string SheetName)
		{
            if (FlxConsts.ExcelVersion == TExcelVersion.v97_2003)
            {
                if (AbsNum <= 0 || AbsNum > 8) return TFlxFormulaErrorValue.ErrValue;
            }
            else
            {
                if (AbsNum <= 0 || AbsNum > 4) return TFlxFormulaErrorValue.ErrValue;
            }
			AbsNum = (AbsNum - 1) % 4 + 1;

			if (Col <= 0 || Col > FlxConsts.Max_Columns + 1) return TFlxFormulaErrorValue.ErrValue;
			if (Row <= 0 || Row > FlxConsts.Max_Rows + 1) return TFlxFormulaErrorValue.ErrValue;

			TCellAddress adr = new TCellAddress(SheetName, Row, Col, AbsNum <= 2, (AbsNum & 1) != 0);
            if (NotRC)
            {
                return adr.CellRef;
            }
            else
            {
                //return adr.CellRefR1C1(wi.Row + 1 + wi.RowOfs, wi.Col + 1 + wi.ColOfs);
                return adr.CellRefR1C1(-1, -1); //it is relative to always the start of the sheet.
            }
		}
	}

	#endregion

	#region SubTotal // Aggregate
	internal class TBaseAggregateToken : TRangeParsedToken
	{
		internal TBaseAggregateToken(int ArgCount, ptg aId, TCellFunctionData aFuncData, int StartArg, bool aCountAnything, bool aIgnoreMissingArg)
			: base(ArgCount, aId, aFuncData, StartArg, aCountAnything, aIgnoreMissingArg)
		{
		}

        protected object CalcAggregate(TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, ref TCalcStack CalcStack, int AggFunction, TCalcState NewState)
        {
            object Result = null;
            switch (AggFunction)
            {
                case 1: //Average
                    Result = base.EvaluateAvg(FTokenList, wi, TAverageAggregate.Instance0, NewState, CalcStack, false);
                    FlushParams(FTokenList);
                    return Result;

                case 2: //Count
                    Result = base.EvaluateAvg(FTokenList, wi, TCountAggregate.Instance, NewState, CalcStack, true);
                    FlushParams(FTokenList);
                    return Result;

                case 3: //countA
                    Result = base.EvaluateAvg(FTokenList, wi, TCountAAggregate.InstanceAll, NewState, CalcStack, true);
                    FlushParams(FTokenList);
                    return Result;

                case 4: //Max
                    Result = base.Evaluate(FTokenList, wi, TMaxAggregate.Instance0, NewState, CalcStack);
                    FlushParams(FTokenList);
                    return Result;

                case 5: //Min
                    Result = base.Evaluate(FTokenList, wi, TMinAggregate.Instance0, NewState, CalcStack);
                    FlushParams(FTokenList);
                    return Result;

                case 6: //product
                    Result = base.Evaluate(FTokenList, wi, TProductAggregate.Instance, NewState, CalcStack);
                    FlushParams(FTokenList);
                    return Result;

                case 7: //stdev
                    Result = TBaseStDevToken.CalcStDev(FArgCount - FStartArg, false, 1, FTokenList, wi, NewState, CalcStack);
                    FlushParams(FTokenList);
                    return Result;

                case 8: //stdevp
                    Result = TBaseStDevToken.CalcStDev(FArgCount - FStartArg, false, 0, FTokenList, wi, NewState, CalcStack);
                    FlushParams(FTokenList);
                    return Result;
                
                case 9: //Sum
                    Result = base.Evaluate(FTokenList, wi, TSumAggregate.Instance, NewState, CalcStack);
                    FlushParams(FTokenList);
                    return Result;
                
                case 10: //var
                    Result = TVarToken.CalcVar(FArgCount, FStartArg, false, false, 1, FTokenList, wi, NewState, CalcStack);
                    FlushParams(FTokenList);
                    return Result;

                case 11: //varp
                    Result = TVarToken.CalcVar(FArgCount, FStartArg, false, false, 0, FTokenList, wi, NewState, CalcStack);
                    FlushParams(FTokenList);
                    return Result;

                case 12: //median
                    Result = TMedianToken.CalcMedian(FTokenList, wi, FArgCount - FStartArg, NewState, CalcStack);
                    FlushParams(FTokenList);
                    return Result;                    

                case 13: //mode
                    Result = TModeToken.CalcMode(FArgCount - FStartArg, FTokenList, wi, NewState, CalcStack);
                    FlushParams(FTokenList);
                    return Result;

                case 14: //large
                    if (FArgCount - FStartArg != 2) return TFlxFormulaErrorValue.ErrValue;
                    Result = TSmallLargeToken.Calc(true, FTokenList, wi, NewState, CalcStack);
                    FlushParams(FTokenList);
                    return Result;

                case 15: //small
                    if (FArgCount - FStartArg != 2) return TFlxFormulaErrorValue.ErrValue;
                    Result = TSmallLargeToken.Calc(false, FTokenList, wi, NewState, CalcStack);
                    FlushParams(FTokenList);
                    return Result;

                case 16: //Percentile
                    if (FArgCount - FStartArg != 2) return TFlxFormulaErrorValue.ErrValue;
                    Result = TPercentileToken.CalcPercentile(false, false, FTokenList, wi, NewState, CalcStack);
                    FlushParams(FTokenList);
                    return Result;

                case 17:  //quartile
                    if (FArgCount - FStartArg != 2) return TFlxFormulaErrorValue.ErrValue;
                    Result = TPercentileToken.CalcPercentile(true, false, FTokenList, wi, NewState, CalcStack);
                    FlushParams(FTokenList);
                    return Result;

                case 18: //Percentile.Exc
                    if (FArgCount - FStartArg != 2) return TFlxFormulaErrorValue.ErrValue;
                    Result = TPercentileToken.CalcPercentile(false, true, FTokenList, wi, NewState, CalcStack);
                    FlushParams(FTokenList);
                    return Result;

                case 19:  //quartile.exc
                    if (FArgCount - FStartArg != 2) return TFlxFormulaErrorValue.ErrValue;
                    Result = TPercentileToken.CalcPercentile(true, true, FTokenList, wi, NewState, CalcStack);
                    FlushParams(FTokenList);
                    return Result;


            }

            wi.AddUnsupported(TUnsupportedFormulaErrorType.FunctionalityNotImplemented, FuncData.Name);
            return TFlxFormulaErrorValue.ErrValue;
        }

        private void FlushParams(TParsedTokenList FTokenList)
        {
            for (int i = 0; i < FStartArg; i++)
            {
                FTokenList.Flush();
            }
        }

		internal override void EvalOne(double d, ref double ResultValue, bool First, TBaseAggregate f)
		{
			if (f == TMaxAggregate.Instance0)
			{
				if (First || (d > ResultValue)) ResultValue = d;
				return;
			}

			if (f == TMinAggregate.Instance0)
			{
				if (First || (d < ResultValue)) ResultValue = d;
				return;
			}

			if (f == TSumAggregate.Instance)
			{
				ResultValue += d;
				return;
			}
            if (f == TProductAggregate.Instance)
            {
                if (First) ResultValue = d;else  ResultValue *= d;
                return;
            }

		}
	}

    internal sealed class TSubTotalToken : TBaseAggregateToken
    {
        internal TSubTotalToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
            : base(ArgCount, aId, aFuncData, 1, false, false)
        {
        }

        internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
        {
            //We need to know the SubTotal function *before* we can evaluate the arguments.
            //But as this is RPN and on the reverse order, Subtotal is the last Parameter.
            //One thing that helps here is that all parameters except the last must be references.

            int SavePosition = FTokenList.SavePosition();
            for (int i = FStartArg; i < FArgCount; i++)
            {
                FTokenList.Flush();
            }

            object v1 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v1 is TFlxFormulaErrorValue) return v1;
            TFlxFormulaErrorValue Err;
            int AggFunction; if (!GetUInt(v1, out AggFunction, out Err)) return Err;

            FTokenList.RestorePosition(SavePosition);

            TCalcState NewState = CalcState.Clone();
            NewState.InSubTotal = true;
            NewState.IgnoreHidden = false;

            if (AggFunction > 100)
            {
                AggFunction -= 100;
                NewState.IgnoreHidden = true;
            }

            if (AggFunction > 11) return TFlxFormulaErrorValue.ErrValue;

            return CalcAggregate(FTokenList, wi, CalcState, ref CalcStack, AggFunction, NewState);
        }
    }

    internal sealed class TAggregateToken : TBaseAggregateToken
    {
        internal TAggregateToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
            : base(ArgCount, aId, aFuncData, 2, false, false)
        {
        }

        internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
        {
            //We need to know the SubTotal function and the agg type *before* we can evaluate the arguments.
            //But as this is RPN and on the reverse order, Subtotal is the last Parameter.
            //Note that here, different from Subtotal, parameters might not be refs.

            int SavePosition = FTokenList.SavePosition();
            for (int i = FStartArg; i < FArgCount; i++)
            {
                FTokenList.Flush();
            }

            object v1 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v1 is TFlxFormulaErrorValue) return v1;
            TFlxFormulaErrorValue Err;
            int AggMethod;
            if (!GetUInt(v1, out AggMethod, out Err)) return Err;
            if (AggMethod < 0 || AggMethod > 7) return TFlxFormulaErrorValue.ErrValue; 

            object v2 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v2 is TFlxFormulaErrorValue) return v2;
            int AggFunction; 
            if (!GetUInt(v2, out AggFunction, out Err)) return Err;

            FTokenList.RestorePosition(SavePosition);

            TCalcState NewState = CalcState.Clone();
            NewState.InSubTotal = AggMethod <= 3;
            NewState.IgnoreHidden = AggMethod == 1 || AggMethod == 3 || AggMethod == 5 || AggMethod == 7;
            NewState.IgnoreErrors = AggMethod == 2 || AggMethod == 3 || AggMethod == 6 || AggMethod == 7;


            if (AggFunction > 19) return TFlxFormulaErrorValue.ErrValue;
            return CalcAggregate(FTokenList, wi, CalcState, ref CalcStack, AggFunction, NewState);
        }
    }

	#endregion

	#region SumProduct
	internal abstract class TBaseSumProductToken : TBaseFunctionToken
	{
		bool AllowTransposed;

		internal TBaseSumProductToken(int ArgCount, ptg aId, TCellFunctionData aFuncData, bool aAllowTransposed)
			: base(ArgCount, aId, aFuncData)
		{
			AllowTransposed = aAllowTransposed;
		}

		protected abstract double DoY(double val);
		protected abstract double DoX(double ResultValue, double val);

		private object ProcessArray(ref double[,] ResultValue, ref double[,] Result2, ref bool[,] InvalidLines, object[,] arr, bool UseResult2)
		{
			if (ResultValue == null)
			{
				ResultValue = new double[arr.GetLength(0), arr.GetLength(1)];
				if (UseResult2) Result2 = new double[arr.GetLength(0), arr.GetLength(1)];
				InvalidLines = new bool[arr.GetLength(0), arr.GetLength(1)];
				for (int i = 0; i < ResultValue.GetLength(0); i++)
					for (int k = 0; k < ResultValue.GetLength(1); k++)
					{
						object o = ConvertToAllowedObject(arr[i, k]);
						if (o is TFlxFormulaErrorValue) return o;
						if (o is double)
						{
							if (UseResult2)
								Result2[i, k] = DoY((double)o);
							else
								ResultValue[i, k] = DoY((double)o);
						}
						else InvalidLines[i, k] = true;
					}
			}
			else
			{
				if (AllowTransposed)
				{
					if (ResultValue.GetLength(0) != arr.GetLength(0) || ResultValue.GetLength(1) != arr.GetLength(1))
					{
						if (ResultValue.GetLength(0) * ResultValue.GetLength(1) != arr.GetLength(0) * arr.GetLength(1))
						{
							return TFlxFormulaErrorValue.ErrNA; //This is different on x2y2 than in sumproduct.
						}
					}
				}
				else
				{
					if (ResultValue.GetLength(0) != arr.GetLength(0) || ResultValue.GetLength(1) != arr.GetLength(1))
						return TFlxFormulaErrorValue.ErrValue;
				}

				int mi = 0;
				int mk = 0;
				int MaxMk = arr.GetLength(1) - 1;
				for (int i = 0; i < ResultValue.GetLength(0); i++)
					for (int k = 0; k < ResultValue.GetLength(1); k++)
					{
						object a = arr[mi, mk];
						object o = ConvertToAllowedObject(a);
						if (o is TFlxFormulaErrorValue) return o;
						if (o is double)
							ResultValue[i, k] = DoX(ResultValue[i, k], (double)o);
						else
						{
							InvalidLines[i, k] = true;
							ResultValue[i, k] = 0;
						}
						if (mk < MaxMk) mk++; else { mk = 0; mi++; }
					}
			}
			return null;
		}

		private object ProcessRef(ExcelFile Xls, ref double[,] ResultValue, ref double[,] Result2, ref bool[,] InvalidLines, TAddressList adr, TCalcState CalcState, TCalcStack CalcStack, bool UseResult2)
		{
			if (adr.Count != 1) return TFlxFormulaErrorValue.ErrValue;
			if (adr[0].Length < 1 || adr[0].Length > 2) return TFlxFormulaErrorValue.ErrValue;
			TAddress a1 = adr[0][0];
			TAddress a2 = a1;
			if (adr[0].Length > 1)
				a2 = adr[0][1];

			if (a1 == null || a2 == null) return TFlxFormulaErrorValue.ErrNA;

			if (a1.wi.Xls != a2.wi.Xls || a1.Sheet != a2.Sheet)
				return TFlxFormulaErrorValue.ErrRef;

			int MinRow = Math.Min(a1.Row, a2.Row);
			int MinCol = Math.Min(a1.Col, a2.Col);
			int MaxCol = Math.Max(a1.Col, a2.Col);

			if (ResultValue == null)
			{
				ResultValue = new double[Math.Abs(a1.Row - a2.Row) + 1, Math.Abs(a1.Col - a2.Col) + 1];
				if (UseResult2) Result2 = new double[Math.Abs(a1.Row - a2.Row) + 1, Math.Abs(a1.Col - a2.Col) + 1];
				InvalidLines = new bool[Math.Abs(a1.Row - a2.Row) + 1, Math.Abs(a1.Col - a2.Col) + 1];
				for (int i = 0; i < ResultValue.GetLength(0); i++)
					for (int k = 0; k < ResultValue.GetLength(1); k++)
					{
						object o = ConvertToAllowedObject(a1.wi.Xls.GetCellValueAndRecalc(a1.Sheet, i + MinRow, k + MinCol, CalcState, CalcStack));
						if (o is TFlxFormulaErrorValue) return o;
						if (o is double)
						{
							if (UseResult2)
								Result2[i, k] = DoY((double)o); //here strings and booleans are 0.
							else
								ResultValue[i, k] = DoY((double)o); //here strings and booleans are 0.
						}
						else
							InvalidLines[i, k] = true;
					}
			}
			else
			{
				if (AllowTransposed)
				{
					if (ResultValue.GetLength(0) * ResultValue.GetLength(1) != (Math.Abs(a1.Row - a2.Row) + 1) * (Math.Abs(a1.Col - a2.Col) + 1))
					{
						return TFlxFormulaErrorValue.ErrNA; //This is different on x2y2 than in sumproduct.
					}
				}
				else
				{
					if (ResultValue.GetLength(0) != Math.Abs(a1.Row - a2.Row) + 1 || ResultValue.GetLength(1) != Math.Abs(a1.Col - a2.Col) + 1)
						return TFlxFormulaErrorValue.ErrValue;
				}

				int mi = MinRow;
				int mk = MinCol;
				for (int i = 0; i < ResultValue.GetLength(0); i++)
					for (int k = 0; k < ResultValue.GetLength(1); k++)
					{
						object a = a1.wi.Xls.GetCellValueAndRecalc(a1.Sheet, mi, mk, CalcState, CalcStack);
						object o = ConvertToAllowedObject(a);
						if (o is TFlxFormulaErrorValue) return o;
						if (o is double)
							ResultValue[i, k] = DoX(ResultValue[i, k], (double)o);
						else
						{
							InvalidLines[i, k] = true;
							ResultValue[i, k] = 0;
						}
						if (mk < MaxCol) mk++; else { mk = MinCol; mi++; }
					}
			}

			return null;
		}


		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			double[,] Result = null;
			bool[,] InvalidLines = null;
			wi.SumProductCount++;
			try
			{
				object res = CalcResult(FTokenList, wi, CalcState, CalcStack, ref InvalidLines, ref Result); if (res != null) return res;
				double dResult = 0;
				for (int i = 0; i < Result.GetLength(0); i++)
				{
					for (int k = 0; k < Result.GetLength(1); k++)
					{
						if (!InvalidLines[i, k])
						{
							dResult += Result[i, k];
						}
					}
				}
				return dResult;
			}
			finally
			{
				wi.SumProductCount--;
			}
		}

		protected object CalcResult(TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack, ref bool[,] InvalidLines, ref double[,] ResultValue)
		{
			double[,] Result2 = null;
			return CalcResult(FTokenList, wi, CalcState, CalcStack, ref InvalidLines, ref ResultValue, ref Result2, false);
		}
		protected object CalcResult(TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack, ref bool[,] InvalidLines, ref double[,] ResultValue, ref double[,] Result2, bool UseResult2)
		{
			for (int a = 0; a < FArgCount; a++)
			{
				TAddressList adr = null; object[,] arr = null;
				object ret = GetRangeOrArray(FTokenList, wi, CalcState, CalcStack, out adr, out arr, TFlxFormulaErrorValue.ErrRef, true);
				if (ret is TFlxFormulaErrorValue) return ret;

				if (arr != null)
				{
					object res = ProcessArray(ref ResultValue, ref Result2, ref InvalidLines, arr, UseResult2);
					if (res != null) return res;
				}
				else
				{
					object res = ProcessRef(wi.Xls, ref ResultValue, ref Result2, ref InvalidLines, adr, CalcState, CalcStack, UseResult2);
					if (res != null) return res;
				}
			}
			if (ResultValue == null) return TFlxFormulaErrorValue.ErrNA;
			return null;
		}
	}

	internal sealed class TSumProductToken : TBaseSumProductToken
	{
		internal TSumProductToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData, false)
		{
		}

		protected override double DoY(double val)
		{
			return val;
		}

		protected override double DoX(double ResultValue, double val)
		{
			return ResultValue * val;
		}
	}

	internal sealed class TSumX2mY2Token : TBaseSumProductToken
	{
		internal TSumX2mY2Token(ptg aId, TCellFunctionData aFuncData)
			: base(2, aId, aFuncData, true)
		{
		}

		protected override double DoY(double val)
		{
			return -val * val;
		}

		protected override double DoX(double ResultValue, double val)
		{
			return ResultValue + val * val;
		}
	}

	internal sealed class TSumX2pY2Token : TBaseSumProductToken
	{
		internal TSumX2pY2Token(ptg aId, TCellFunctionData aFuncData)
			: base(2, aId, aFuncData, true)
		{
		}

		protected override double DoY(double val)
		{
			return val * val;
		}

		protected override double DoX(double ResultValue, double val)
		{
			return ResultValue + val * val;
		}
	}

	internal sealed class TSumXmY2Token : TBaseSumProductToken
	{
		internal TSumXmY2Token(ptg aId, TCellFunctionData aFuncData)
			: base(2, aId, aFuncData, true)
		{
		}

		//We will implement (x - y)2 = x2 + y2 - 2xy
		protected override double DoY(double val)
		{
			return val;
		}

		protected override double DoX(double ResultValue, double val)
		{
			return ResultValue * ResultValue + val * val - 2 * ResultValue * val;
		}
	}
	#endregion

	#region Statistical

	internal abstract class TStatToken : TBaseFunctionToken
	{
		internal TStatToken(int ArgCount, ptg aId, TCellFunctionData aFuncData) : base(ArgCount, aId, aFuncData) { }

		internal static TAverageValue ProcessRange(int ArgCount, TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack, bool CountAnything)
		{
			TAverageValue Result = new TAverageValue(0, 0);

			for (int i = 0; i < ArgCount; i++)
			{
				object v1 = FTokenList.EvaluateToken(wi, f, CalcState, CalcStack);
				if (v1 is TFlxFormulaErrorValue)
				{
					Result.HasErr = true;
					Result.Err = (TFlxFormulaErrorValue)v1;
					return Result;
				}

				object[,] arr1 = v1 as object[,];
				if (arr1 != null)
					v1 = f.AggArray(arr1);

				TAverageValue av = v1 as TAverageValue;
				if (av != null)
				{
					Result.Sum += av.Sum; //already aggregated.
					Result.ValueCount += av.ValueCount;
				}
				else
				{
					v1 = ConvertToAllowedObject(v1);
					//if (CountAnything)  When in arguments, they are always converted.
				{
					if (v1 is bool)
						if ((bool)v1) v1 = 1.0; else v1 = 0.0;
				}
					if (v1 is double)
					{
						f.AggAverages(Result, (double)v1);
					}
					else
					{
						Result.HasErr = true;
						Result.Err = TFlxFormulaErrorValue.ErrValue;
						return Result;
					}
				}
			}

			return Result;
		}

		internal static object Multiply(object[,] XValueArray, object[,] YValueArray, double XAvg, double YAvg, int d00, int d10, int d01, int d11)
		{
			if (XValueArray.GetLength(d00) != YValueArray.GetLength(d10) || XValueArray.GetLength(d01) != YValueArray.GetLength(d11))
				return TFlxFormulaErrorValue.ErrNA;

			double Result = 0;

			int MaxCols = YValueArray.GetLength(1);
			int c1 = MaxCols - 1;
			int r1 = YValueArray.GetLength(0) - 1;
			for (int r = XValueArray.GetLength(d00) - 1; r >= 0; r--)
			{
				for (int c = XValueArray.GetLength(d01) - 1; c >= 0; c--)
				{
					if (XValueArray[r, c] is double && YValueArray[r1, c1] is double)
					{
						Result += ((double)XValueArray[r, c] - XAvg) * ((double)YValueArray[r1, c1] - YAvg);
					}

					c1--;
					if (c1 < 0)
					{
						c1 = MaxCols - 1;
						r1--;
					}
				}
			}

			return Result;
		}

		internal static object CalcAvg(object[,] XValueArray, object[,] YValueArray, out double Count, out double XAvg, out double YAvg, int d00, int d10, int d01, int d11)
		{
			Count = 0;
			XAvg = 0;
			YAvg = 0;

			if (XValueArray.GetLength(d00) != YValueArray.GetLength(d10) || XValueArray.GetLength(d01) != YValueArray.GetLength(d11))
				return TFlxFormulaErrorValue.ErrNA;

			int MaxCols = YValueArray.GetLength(1);
			int c1 = MaxCols - 1;
			int r1 = YValueArray.GetLength(0) - 1;
			for (int r = XValueArray.GetLength(d00) - 1; r >= 0; r--)
			{
				for (int c = XValueArray.GetLength(d01) - 1; c >= 0; c--)
				{
					if (XValueArray[r, c] is double && YValueArray[r1, c1] is double)
					{
						XAvg += (double)XValueArray[r, c];
						YAvg += (double)YValueArray[r1, c1];
						Count++;
					}

					c1--;
					if (c1 < 0)
					{
						c1 = MaxCols - 1;
						r1--;
					}
				}
			}

			if (Count <= 0) return TFlxFormulaErrorValue.ErrDiv0;
			XAvg /= Count;
			YAvg /= Count;
			return null;
		}

		internal static object CalcStDev(object[,] XValueArray, object[,] YValueArray, double XAvg, double YAvg, out double XStDev, out double YStDev, int d00, int d10, int d01, int d11)
		{
			XStDev = 0;
			YStDev = 0;

			if (XValueArray.GetLength(d00) != YValueArray.GetLength(d10) || XValueArray.GetLength(d01) != YValueArray.GetLength(d11))
				return TFlxFormulaErrorValue.ErrNA;

			int MaxCols = YValueArray.GetLength(1);
			int c1 = MaxCols - 1;
			int r1 = YValueArray.GetLength(0) - 1;
			for (int r = XValueArray.GetLength(d00) - 1; r >= 0; r--)
			{
				for (int c = XValueArray.GetLength(d01) - 1; c >= 0; c--)
				{
					if (XValueArray[r, c] is double && YValueArray[r1, c1] is double)
					{
						double x = (double)XValueArray[r, c] - XAvg;
						XStDev += x * x;
						double y = (double)YValueArray[r1, c1] - YAvg;
						YStDev += y * y;
					}

					c1--;
					if (c1 < 0)
					{
						c1 = MaxCols - 1;
						r1--;
					}
				}
			}
			return null;
		}


	}

	internal abstract class TBaseStDevToken : TStatToken
	{
		private int FNOfs;
		protected bool FCountAnything;

		internal TBaseStDevToken(int ArgCount, ptg aId, TCellFunctionData aFuncData, int NOfs, bool aCountAnything)
			: base(ArgCount, aId, aFuncData)
		{
			FNOfs = NOfs;
			FCountAnything = aCountAnything;
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
            return CalcStDev(FArgCount, FCountAnything, FNOfs, FTokenList, wi, CalcState, CalcStack);
		}

        internal static object CalcStDev(int ArgCount, bool CountAnything, int NOfs, TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack)
        {
            TAverageAggregate AvgInstance = CountAnything ? TAverageAggregate.InstanceA : TAverageAggregate.Instance0;
            int StartPosition = FTokenList.SavePosition();
            TAverageValue Average = ProcessRange(ArgCount, FTokenList, wi, AvgInstance, CalcState, CalcStack, CountAnything);
            if (Average.HasErr) return Average.Err;
            if (Average.ValueCount <= 0) return TFlxFormulaErrorValue.ErrDiv0;
            double Avg = Average.Sum / Average.ValueCount;

            FTokenList.RestorePosition(StartPosition);
            TAverageValue StDev = ProcessRange(ArgCount, FTokenList, wi, new TSquaredDiffAggregate(Avg, CountAnything), CalcState, CalcStack, CountAnything);
            if (StDev.HasErr) return StDev.Err;
            if (StDev.ValueCount <= 1) return TFlxFormulaErrorValue.ErrDiv0;
            return Math.Sqrt(StDev.Sum / (StDev.ValueCount - NOfs));
        }
	}

	internal sealed class TStDevToken : TBaseStDevToken
	{
		internal TStDevToken(int ArgCount, ptg aId, TCellFunctionData aFuncData) : base(ArgCount, aId, aFuncData, 1, false) { }
	}

	internal sealed class TStDevPToken : TBaseStDevToken
	{
		internal TStDevPToken(int ArgCount, ptg aId, TCellFunctionData aFuncData) : base(ArgCount, aId, aFuncData, 0, false) { }
	}

	internal sealed class TStDevAToken : TBaseStDevToken
	{
		internal TStDevAToken(int ArgCount, ptg aId, TCellFunctionData aFuncData) : base(ArgCount, aId, aFuncData, 1, true) { }
	}

	internal sealed class TStDevPAToken : TBaseStDevToken
	{
		internal TStDevPAToken(int ArgCount, ptg aId, TCellFunctionData aFuncData) : base(ArgCount, aId, aFuncData, 0, true) { }
	}

	internal sealed class TCorrelToken : TStatToken
	{
		internal TCorrelToken(ptg aId, TCellFunctionData aFuncData) : base(2, aId, aFuncData) { }
        
		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			object xv = FTokenList.EvaluateToken(wi, TArrayAggregate.Instance, CalcState, CalcStack);
			object[,] XValueArray = xv as object[,];
			object yv = FTokenList.EvaluateToken(wi, TArrayAggregate.Instance, CalcState, CalcStack);
			object[,] YValueArray = yv as object[,];

			if (XValueArray == null || YValueArray == null) return TFlxFormulaErrorValue.ErrDiv0;

			object Mul = null;
			double XAvg;
			double YAvg;
			double n;
			double XStDev;
			double YStDev;

			if (XValueArray.GetLength(0) != YValueArray.GetLength(0))
			{
				object Result = CalcAvg(XValueArray, YValueArray, out n, out XAvg, out YAvg, 0, 1, 1, 0);
				if (Result != null) return Result;

				Result = CalcStDev(XValueArray, YValueArray, XAvg, YAvg, out XStDev, out YStDev, 0, 1, 1, 0);
				if (Result != null) return Result;

				Mul = Multiply(XValueArray, YValueArray, XAvg, YAvg, 0, 1, 1, 0);
				if (Mul is TFlxFormulaErrorValue) return Mul;
				if (!(Mul is double)) return TFlxFormulaErrorValue.ErrNA;


			}
			else
			{
				object Result = CalcAvg(XValueArray, YValueArray, out n, out XAvg, out YAvg, 0, 0, 1, 1);
				if (Result != null) return Result;

				Result = CalcStDev(XValueArray, YValueArray, XAvg, YAvg, out XStDev, out YStDev, 0, 0, 1, 1);
				if (Result != null) return Result;

				Mul = Multiply(XValueArray, YValueArray, XAvg, YAvg, 0, 0, 1, 1);
				if (Mul is TFlxFormulaErrorValue) return Mul;
				if (!(Mul is double)) return TFlxFormulaErrorValue.ErrNA;
			}


			if (XStDev <= 0 || YStDev <= 0) return TFlxFormulaErrorValue.ErrDiv0;
			return (double)Mul / Math.Sqrt(XStDev * YStDev);
		}

	}

	internal sealed class TCoVarToken : TStatToken
	{
		internal TCoVarToken(ptg aId, TCellFunctionData aFuncData) : base(2, aId, aFuncData) { }
        
		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			object xv = FTokenList.EvaluateToken(wi, TArrayAggregate.Instance, CalcState, CalcStack);
			object[,] XValueArray = xv as object[,];
			object yv = FTokenList.EvaluateToken(wi, TArrayAggregate.Instance, CalcState, CalcStack);
			object[,] YValueArray = yv as object[,];

			if (xv is double)
			{
				XValueArray = new object[1, 1];
				XValueArray[0, 0] = xv;
			}
			if (yv is double)
			{
				YValueArray = new object[1, 1];
				YValueArray[0, 0] = yv;
			}

			if (XValueArray == null || YValueArray == null) return TFlxFormulaErrorValue.ErrValue;

			object Mul = null;
			double XAvg;
			double YAvg;
			double n;

			if (XValueArray.GetLength(0) != YValueArray.GetLength(0))
			{
				object Result = CalcAvg(XValueArray, YValueArray, out n, out XAvg, out YAvg, 0, 1, 1, 0);
				if (Result != null) return Result;

				Mul = Multiply(XValueArray, YValueArray, XAvg, YAvg, 0, 1, 1, 0);
				if (Mul is TFlxFormulaErrorValue) return Mul;
				if (!(Mul is double)) return TFlxFormulaErrorValue.ErrNA;
			}
			else
			{
				object Result = CalcAvg(XValueArray, YValueArray, out n, out XAvg, out YAvg, 0, 0, 1, 1);
				if (Result != null) return Result;

				Mul = Multiply(XValueArray, YValueArray, XAvg, YAvg, 0, 0, 1, 1);
				if (Mul is TFlxFormulaErrorValue) return Mul;
				if (!(Mul is double)) return TFlxFormulaErrorValue.ErrNA;
			}


			if (n <= 0) return TFlxFormulaErrorValue.ErrDiv0;
			return (double)Mul / n;
		}
	}

	internal sealed class TVarToken : TRangeParsedToken
	{
		int P;

		internal TVarToken(int ArgCount, ptg aId, TCellFunctionData aFuncData, int aP, bool aCountAnything)
			: base(ArgCount, aId, aFuncData, 0, aCountAnything, false)
		{
			P = aP;
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
            return CalcVar(FArgCount, FStartArg, IgnoreMissingArg, CountAnything, P, FTokenList, wi, CalcState, CalcStack);
		}

        internal static object CalcVar(int FArgCount, int FStartArg, bool IgnoreMissingArg, bool CountAnything, int P, TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack)
        {
            TAverageAggregate AvgInstance = CountAnything ? TAverageAggregate.InstanceA : TAverageAggregate.Instance0;
            int SavePosition = FTokenList.SavePosition();
            object Result = EvaluateAvg2St(FArgCount, FStartArg, IgnoreMissingArg, CountAnything, FTokenList, wi, AvgInstance, CalcState, CalcStack, false, false);
            if (Result is TFlxFormulaErrorValue) return Result;

            TAverageValue Avg = Result as TAverageValue;
            if (Avg == null) return TFlxFormulaErrorValue.ErrNA;

            if (Avg.ValueCount <= 0) return TFlxFormulaErrorValue.ErrDiv0;
            FTokenList.RestorePosition(SavePosition);

            double XAvg = Avg.Sum / Avg.ValueCount;
            TSquaredDiffAggregate SqDiffAvg = new TSquaredDiffAggregate(XAvg, CountAnything);

            double Total = 0;
            long n = 0;
            for (int i = FStartArg; i < FArgCount; i++)
            {
                object oSqDiff = FTokenList.EvaluateToken(wi, SqDiffAvg, CalcState, CalcStack);
                if (oSqDiff is TFlxFormulaErrorValue) return oSqDiff;

                TAverageValue SqDiff = oSqDiff as TAverageValue;
                if (SqDiff == null)
                {
                    double v;
                    if (GetDouble(oSqDiff, out v))
                    {
                        Total += (v - XAvg) * (v - XAvg);
                        n++;
                    }
                    else return TFlxFormulaErrorValue.ErrValue;
                }
                else
                {
                    if (SqDiff.HasErr) return SqDiff.Err;
                    Total += SqDiff.Sum;
                    n += SqDiff.ValueCount;
                }
            }
            if (n - P <= 0) return TFlxFormulaErrorValue.ErrDiv0;
            return Total / (n - P);
        }
	}

	#endregion

	#region Statistical II
	internal abstract class TSmallLargeToken : TBaseFunctionToken
	{
		bool DoMax;
		protected TSmallLargeToken(bool aDoMax, ptg aId, TCellFunctionData aFuncData) : base(2, aId, aFuncData)
		{
			DoMax = aDoMax;
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
            return Calc(DoMax, FTokenList, wi, CalcState, CalcStack);
		}

        internal static object Calc(bool DoMax, TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack)
        {
            object v2 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v2 is TFlxFormulaErrorValue) return v2;

            int k;
            object[,] ArrResult;
            object[,] ArrV2;
            TFlxFormulaErrorValue Err;
            if (!IsArrayArgument(v2, out ArrResult, out ArrV2))
            {
                if (!GetInt(v2, out k, out Err)) return Err;
                if (k <= 0) return TFlxFormulaErrorValue.ErrNum;
            }
            else
            {
                int SavePosition = FTokenList.SavePosition();
                for (int i = 0; i < ArrResult.GetLength(0); i++)
                    for (int j = 0; j < ArrResult.GetLength(1); j++)
                    {
                        if (!GetInt(ArrV2[i, j], out k, out Err)) return Err;
                        if (k <= 0) return TFlxFormulaErrorValue.ErrNum;

                        FTokenList.RestorePosition(SavePosition);
                        ArrResult[i, j] = CalcSmallLarge(DoMax, FTokenList, wi, k, CalcState, CalcStack);
                    }
                return UnPack(ArrResult);
            }

            return CalcSmallLarge(DoMax, FTokenList, wi, k, CalcState, CalcStack);
        }

		private static object CalcSmallLarge(bool DoMax, TParsedTokenList FTokenList, TWorkbookInfo wi, int k, TCalcState CalcState, TCalcStack CalcStack)
		{
			TMaxMinKAggregate MaxK = new TMaxMinKAggregate(k, DoMax);
			object v1 = FTokenList.EvaluateToken(wi, MaxK, CalcState, CalcStack);

			object[,] arr1 = v1 as object[,];
			if (arr1 != null)
				v1 = MaxK.AggArray(arr1);

			v1 = ConvertToAllowedObject(v1);

			if (!MaxK.Used) //We didn't agregate
			{
				if (k != 1) return TFlxFormulaErrorValue.ErrNum;
			}

			if (v1 is TFlxFormulaErrorValue) return v1;
			if (v1 is double)
			{
				return v1;
			}

			double d1; if (!ExtToDouble(v1, out d1)) return TFlxFormulaErrorValue.ErrValue;
			return d1;

		}


	}

	internal sealed class TLargeToken : TSmallLargeToken
	{
		internal TLargeToken(ptg aId, TCellFunctionData aFuncData)
			: base(true, aId, aFuncData)
		{
		}
	}

	internal sealed class TSmallToken : TSmallLargeToken
	{
		internal TSmallToken(ptg aId, TCellFunctionData aFuncData)
			: base(false, aId, aFuncData)
		{
		}
	}

	#endregion

	#region Statistical Distributions
	internal abstract class TStatDistToken : TBaseFunctionToken
	{
		internal TStatDistToken(int ArgCount, ptg aId, TCellFunctionData aFuncData) : base(ArgCount, aId, aFuncData) { }
		protected static double MachinePrecision()
		{
			return Math.Pow(2, -52);

		}

		/// <summary>
		/// This method is ported from dcdflib, based on CODY's algorithm.
		/// </summary>
		/// <param name="x"></param>
		/// <returns></returns>
		internal static double CalcCumulativeNormal(double x)
		{
			#region Constants
			const double thrsh = 0.66291e0;
			double[] a = 
								{
									2.2352520354606839287e00,1.6102823106855587881e02,1.0676894854603709582e03,
									1.8154981253343561249e04,6.5682337918207449113e-2
								};
			double[] b = 
								{
									4.7202581904688241870e01,9.7609855173777669322e02,1.0260932208618978205e04,
									4.5507789335026729956e04
								};
			double[] c = 
								{
									3.9894151208813466764e-1,8.8831497943883759412e00,9.3506656132177855979e01,
									5.9727027639480026226e02,2.4945375852903726711e03,6.8481904505362823326e03,
									1.1602651437647350124e04,9.8427148383839780218e03,1.0765576773720192317e-8
								};
			double[] d = 
								{
									2.2266688044328115691e01,2.3538790178262499861e02,1.5193775994075548050e03,
									6.4855582982667607550e03,1.8615571640885098091e04,3.4900952721145977266e04,
									3.8912003286093271411e04,1.9685429676859990727e04
								};

			double[] p = 
								{
									2.1589853405795699e-1,1.274011611602473639e-1,2.2235277870649807e-2,
									1.421619193227893466e-3,2.9112874951168792e-5,2.307344176494017303e-2
								};

			double[] q = 
								{
									1.28426009614491121e00,4.68238212480865118e-1,6.59881378689285515e-2,
									3.78239633202758244e-3,7.29751555083966205e-5
								};


			double sixten = 1.6; //yep
			double sqrpi = 3.9894228040143267794e-1;

			#endregion

			double Result = 0;
			double ccum = 0;
			double y = Math.Abs(x);
			if (y <= thrsh)
			{
				//
				//  Evaluate  anorm  for  |X| <= 0.66291
				//
				double xsq = 0;
				if (y > MachinePrecision()) xsq = x * x;
				double xnum = a[4] * xsq;
				double xden = xsq;
				for (int i = 0; i < 3; i++)
				{
					xnum = (xnum + a[i]) * xsq;
					xden = (xden + b[i]) * xsq;
				}
				double temp = x * (xnum + a[3]) / (xden + b[3]);
				ccum = 0.5 - temp;
				Result = 0.5 + temp;
			}

				//
				//  Evaluate  anorm  for 0.66291 <= |X| <= sqrt(32)
				//
			else if (y <= Math.Sqrt(32))
			{
				double xnum = c[8] * y;
				double xden = y;
				for (int i = 0; i < 7; i++)
				{
					xnum = (xnum + c[i]) * y;
					xden = (xden + d[i]) * y;
				}

				double tmp0 = (xnum + c[7]) / (xden + d[7]);
				double xsq = Math.Floor(y * sixten) / sixten;
				double del = (y - xsq) * (y + xsq);
				Result = Math.Exp(-(xsq * xsq * 0.5)) * Math.Exp(-(del * 0.5)) * tmp0;
				ccum = 1 - Result;
				if (x > 0)
				{
					double temp = Result;
					Result = ccum;
					ccum = temp;
				}
			}
				//
				//  Evaluate  anorm  for |X| > sqrt(32)
				//
			else
			{
				Result = 0;
				double xsq = 1 / (x * x);
				double xnum = p[5] * xsq;
				double xden = xsq;
				for (int i = 0; i < 4; i++)
				{
					xnum = (xnum + p[i]) * xsq;
					xden = (xden + q[i]) * xsq;
				}
				Result = xsq * (xnum + p[4]) / (xden + q[4]);
				Result = (sqrpi - Result) / y;
				xsq = Math.Floor(x * sixten) / sixten;
				double del = (x - xsq) * (x + xsq);
				Result = Math.Exp(-(xsq * xsq / 2)) * Math.Exp(-(del / 2)) * Result;
				ccum = 1 - Result;
				if (x > 0)
				{
					double temp = Result;
					Result = ccum;
					ccum = temp;
				}
			}
			if (Result < Double.Epsilon || double.IsNaN(Result)) Result = 0.0e0;

			return Result;
		}


		protected static double eval_pol(double[] A, double x)
		{
			double Result = 0;
			double z = 1;
			for (int i = 0; i < A.Length; i++)
			{
				Result += A[i] * z;
				z *= x;
			}
			return Result;
		}

		protected static double eval_polR(double[] A, double x)
		{
			double Result = 0;
			double z = 1;
			for (int i = A.Length - 1; i >= 0; i--)
			{
				Result += A[i] * z;
				z *= x;
			}
			return Result;
		}

		protected static double eval_polR1(double[] A, double x)
		{
			double Result = 0;
			double z = 1;
			for (int i = A.Length - 1; i >= 0; i--)
			{
				Result += A[i] * z;
				z *= x;
			}
			Result += z;

			return Result;
		}

		/// <summary>
		/// provides starting values for the inverse of the normal distribution. ported from dcdflib
		/// </summary>
		/// <param name="p"></param>
		/// <returns></returns>
		private static double stvaln(double p)
		{
			double[] xden = {
								0.993484626060e-1,0.588581570495e0,0.531103462366e0,0.103537752850e0,
								0.38560700634e-2
							};
			double[] xnum = {
								-0.322232431088e0,-1.000000000000e0,-0.342242088547e0,-0.204231210245e-1,
								-0.453642210148e-4
							};

			double sign = -1;
			double z = p;
			if (p <= 0.5e0)
			{
				sign = 1;
				z = 1 - p;
			}

			double y = Math.Sqrt(-(2.0e0 * Math.Log(z)));
			double Result = y + eval_pol(xnum, y) / eval_pol(xden, y);
			return sign * Result;
		}

		/// <summary>
		/// Inverse of Cumulative Normal distribution. ported from dcdflib
		/// </summary>
		/// <param name="P"></param>
		/// <returns></returns>
		internal static double CalcZ(double P)
		{
			const int maxit = 100; //As stated on Excel help.
			const double eps = 1.0e-13;
			const double r2pi = 0.3989422804014326e0;

			//
			//     FIND MINIMUM OF P AND Q
			//
			double pp = P > 0.5 ? 1 - P : P;
			//
			//     INITIALIZATION STEP
			//
			double strtx = stvaln(pp);
			double xcur = strtx;
			//
			//     NEWTON INTERATIONS
			//
			for (int i = 0; i < maxit; i++)
			{
				double cum = CalcCumulativeNormal(xcur);
				double dx = (cum - pp) / (r2pi * Math.Exp(-0.5 * (xcur) * (xcur)));
				xcur -= dx;
				if (Math.Abs(dx / xcur) < eps)
				{
					//NEWTON HAS SUCCEDED
					return P > 0.5 ? -xcur : xcur;
				}
			}

			//
			//     IF WE GET HERE, NEWTON HAS FAILED
			//
			return P > 0.5 ? -strtx : strtx;
		}
	}

	internal sealed class TNormDistToken : TStatDistToken
	{
		internal TNormDistToken(ptg aId, TCellFunctionData aFuncData) : base(4, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			bool Cumulative;
			object v5 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v5 is TFlxFormulaErrorValue) return v5;
			if (!ExtToBool(v5, out Cumulative)) return TFlxFormulaErrorValue.ErrValue;

			double Sigma;
			object v4 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v4 is TFlxFormulaErrorValue) return v4;
			if (!ExtToDouble(v4, out Sigma)) return TFlxFormulaErrorValue.ErrValue;
			if (Sigma <= 0) return TFlxFormulaErrorValue.ErrNum;

			double Mean;
			object v3 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v3 is TFlxFormulaErrorValue) return v3;
			if (!ExtToDouble(v3, out Mean)) return TFlxFormulaErrorValue.ErrValue;

			double x;
			object v2 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v2 is TFlxFormulaErrorValue) return v2;
			if (!ExtToDouble(v2, out x)) return TFlxFormulaErrorValue.ErrValue;

			double z = (x - Mean) / Sigma;
			if (Cumulative)
			{
				return CalcCumulativeNormal(z);
			}
			else
			{
				double ex = z * z / 2;
				return 1 / Math.Sqrt(2 * Math.PI) / Sigma * Math.Exp(-ex);
			}
		}

	}

	internal sealed class TNormInvToken : TStatDistToken
	{
		internal TNormInvToken(ptg aId, TCellFunctionData aFuncData) : base(3, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			double Sigma;
			object v4 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v4 is TFlxFormulaErrorValue) return v4;
			if (!ExtToDouble(v4, out Sigma)) return TFlxFormulaErrorValue.ErrValue;
			if (Sigma <= 0) return TFlxFormulaErrorValue.ErrNum;

			double Mean;
			object v3 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v3 is TFlxFormulaErrorValue) return v3;
			if (!ExtToDouble(v3, out Mean)) return TFlxFormulaErrorValue.ErrValue;

			double P;
			object v2 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v2 is TFlxFormulaErrorValue) return v2;
			if (!ExtToDouble(v2, out P)) return TFlxFormulaErrorValue.ErrValue;
			if (P <= 0 || P >= 1) return TFlxFormulaErrorValue.ErrNum;

			double z = CalcZ(P);
			//z = (x - Mean) / Sigma;
			return z * Sigma + Mean;

		}
	}

	internal sealed class TNormsDistToken : TOneDoubleArgToken
	{
		internal TNormsDistToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }

		protected override object Calc(double x)
		{
			return TStatDistToken.CalcCumulativeNormal(x);
		}
	}

	internal sealed class TNormsInvToken : TOneDoubleArgToken
	{
		internal TNormsInvToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }

		protected override object Calc(double x)
		{
			double P = x;
			if (P <= 0 || P >= 1) return TFlxFormulaErrorValue.ErrNum;

			double z = TStatDistToken.CalcZ(P);
			//z = (x - Mean) / Sigma;
			return z;
		}
	}

	internal sealed class TLogNormDistToken : TStatDistToken
	{
		internal TLogNormDistToken(int aArgCount, ptg aId, TCellFunctionData aFuncData) : base(aArgCount, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			double Sigma;
			object v4 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v4 is TFlxFormulaErrorValue) return v4;
			if (!ExtToDouble(v4, out Sigma)) return TFlxFormulaErrorValue.ErrValue;
			if (Sigma <= 0) return TFlxFormulaErrorValue.ErrNum;

			double Mean;
			object v3 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v3 is TFlxFormulaErrorValue) return v3;
			if (!ExtToDouble(v3, out Mean)) return TFlxFormulaErrorValue.ErrValue;

			double x;
			object v2 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v2 is TFlxFormulaErrorValue) return v2;
			if (!ExtToDouble(v2, out x)) return TFlxFormulaErrorValue.ErrValue;
			if (x <= 0) return TFlxFormulaErrorValue.ErrNum;

			double z = (Math.Log(x) - Mean) / Sigma;
			return CalcCumulativeNormal(z);
		}
	}

	internal sealed class TLogInvToken : TStatDistToken
	{
		internal TLogInvToken(ptg aId, TCellFunctionData aFuncData) : base(3, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			double Sigma;
			object v4 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v4 is TFlxFormulaErrorValue) return v4;
			if (!ExtToDouble(v4, out Sigma)) return TFlxFormulaErrorValue.ErrValue;
			if (Sigma <= 0) return TFlxFormulaErrorValue.ErrNum;

			double Mean;
			object v3 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v3 is TFlxFormulaErrorValue) return v3;
			if (!ExtToDouble(v3, out Mean)) return TFlxFormulaErrorValue.ErrValue;

			double P;
			object v2 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v2 is TFlxFormulaErrorValue) return v2;
			if (!ExtToDouble(v2, out P)) return TFlxFormulaErrorValue.ErrValue;
			if (P <= 0 || P >= 1) return TFlxFormulaErrorValue.ErrNum;

			double z = CalcZ(P);
			return Math.Exp(Mean + Sigma * z);
		}
	}

	internal sealed class TZTestToken : TStatDistToken
	{
		internal TZTestToken(int ArgCount, ptg aId, TCellFunctionData aFuncData) : base(ArgCount, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			double Sigma = -1;  //If not provided, we have to calculate it.
			if (FArgCount > 2)
			{
				object v4 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v4 is TFlxFormulaErrorValue) return v4;
				if (!ExtToDouble(v4, out Sigma)) return TFlxFormulaErrorValue.ErrValue;
				if (Sigma <= 0) return TFlxFormulaErrorValue.ErrNum;
			}

			double Mean;
			object v3 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v3 is TFlxFormulaErrorValue) return v3;
			if (!ExtToDouble(v3, out Mean)) return TFlxFormulaErrorValue.ErrValue;

			int DataPosition = FTokenList.SavePosition();
			object v2 = FTokenList.EvaluateToken(wi, TAverageAggregate.Instance0, CalcState, CalcStack); if (v2 is TFlxFormulaErrorValue) return v2;
			TAverageValue Avg = v2 as TAverageValue;
			if (Avg == null) return TFlxFormulaErrorValue.ErrValue;
			if (Avg.HasErr) return Avg.Err;

			if (Avg.ValueCount == 0) return TFlxFormulaErrorValue.ErrNA;
			double xBar = Avg.Sum / Avg.ValueCount;

			if (Sigma < 0) //it was not specified
			{
				FTokenList.RestorePosition(DataPosition);
				object v0 = FTokenList.EvaluateToken(wi, new TSquaredDiffAggregate(xBar, false), CalcState, CalcStack); if (v0 is TFlxFormulaErrorValue) return v0;
				TAverageValue StDev = v0 as TAverageValue;
				if (StDev == null) return TFlxFormulaErrorValue.ErrValue;
				if (StDev.HasErr) return StDev.Err;
				if (StDev.ValueCount <= 1) return TFlxFormulaErrorValue.ErrDiv0;
				double SV = StDev.Sum / (StDev.ValueCount - 1);
				if (SV <= 0) return TFlxFormulaErrorValue.ErrDiv0;
				Sigma = Math.Sqrt(SV);
				if (Sigma <= 0) return TFlxFormulaErrorValue.ErrDiv0;
			}


			double n = Avg.ValueCount;
			return 1 - CalcCumulativeNormal((xBar - Mean) / Sigma * Math.Sqrt(n));

		}

	}

	internal sealed class TExponDistToken : TStatDistToken
	{
		internal TExponDistToken(ptg aId, TCellFunctionData aFuncData) : base(3, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			bool Cumulative;
			object v5 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v5 is TFlxFormulaErrorValue) return v5;
			if (!ExtToBool(v5, out Cumulative)) return TFlxFormulaErrorValue.ErrValue;

			double Lambda;
			object v4 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v4 is TFlxFormulaErrorValue) return v4;
			if (!ExtToDouble(v4, out Lambda)) return TFlxFormulaErrorValue.ErrValue;
			if (Lambda <= 0) return TFlxFormulaErrorValue.ErrNum;

			double x;
			object v2 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v2 is TFlxFormulaErrorValue) return v2;
			if (!ExtToDouble(v2, out x)) return TFlxFormulaErrorValue.ErrValue;
			if (x < 0) return TFlxFormulaErrorValue.ErrNum;

			if (Cumulative)
			{
				return 1 - Math.Exp(-Lambda * x);
			}
			return Lambda * Math.Exp(-Lambda * x);
		}
	}

	internal sealed class TWeibullToken : TStatDistToken
	{
		internal TWeibullToken(ptg aId, TCellFunctionData aFuncData) : base(4, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			bool Cumulative;
			object v5 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v5 is TFlxFormulaErrorValue) return v5;
			if (!ExtToBool(v5, out Cumulative)) return TFlxFormulaErrorValue.ErrValue;

			double Beta;
			object v4 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v4 is TFlxFormulaErrorValue) return v4;
			if (!ExtToDouble(v4, out Beta)) return TFlxFormulaErrorValue.ErrValue;
			if (Beta <= 0) return TFlxFormulaErrorValue.ErrNum;

			double Alpha;
			object v3 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v3 is TFlxFormulaErrorValue) return v3;
			if (!ExtToDouble(v3, out Alpha)) return TFlxFormulaErrorValue.ErrValue;
			if (Alpha <= 0) return TFlxFormulaErrorValue.ErrNum;

			double x;
			object v2 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v2 is TFlxFormulaErrorValue) return v2;
			if (!ExtToDouble(v2, out x)) return TFlxFormulaErrorValue.ErrValue;
			if (x < 0) return TFlxFormulaErrorValue.ErrNum;

			if (Cumulative)
			{
				return 1 - Math.Exp(-Math.Pow(x / Beta, Alpha));
			}
			return Alpha / Math.Pow(Beta, Alpha) * Math.Pow(x, Alpha - 1) * Math.Exp(-Math.Pow(x / Beta, Alpha));
		}
	}

	internal sealed class TPoissonToken : TStatDistToken
	{
		internal TPoissonToken(ptg aId, TCellFunctionData aFuncData) : base(3, aId, aFuncData) { }

		/// <summary>
		/// This is to consider numerical unstability when x is too big.
		/// see http://support.microsoft.com/default.aspx?scid=kb;en-us;828130
		/// </summary>
		/// <param name="x"></param>
		/// <param name="Mean"></param>
		/// <param name="Cumulative"></param>
		/// <returns></returns>
		private static double fx2(int x, double Mean, bool Cumulative)
		{
			const int MaxIter = 100000;
			double TotalUnscaledProbability = 0;
			double UnscaledResult = 0;
			double EssentiallyZero = 1e-12;

			int m = (int)Mean;
			TotalUnscaledProbability++;
			if (m == x) UnscaledResult++;
			if (Cumulative && m < x) UnscaledResult++;

			//k > m
			double PreviousValue = 1;
			bool Done = false;
			int k = m + 1;
			while (!Done && k <= MaxIter)
			{
				double CurrentValue = PreviousValue * Mean / k;
				TotalUnscaledProbability += CurrentValue;
				if (k == x) UnscaledResult += CurrentValue;
				if (Cumulative && k < x) UnscaledResult += CurrentValue;
				if (CurrentValue <= EssentiallyZero) Done = true;
				PreviousValue = CurrentValue;
				k++;
			}

			// k < m
			PreviousValue = 1;
			Done = false;
			k = m - 1;
			while (!Done && k >= 0)
			{
				double CurrentValue = PreviousValue * (k + 1) / Mean;
				TotalUnscaledProbability += CurrentValue;
				if (k == x) UnscaledResult += CurrentValue;
				if (Cumulative && k < x) UnscaledResult += CurrentValue;
				if (CurrentValue <= EssentiallyZero) Done = true;
				PreviousValue = CurrentValue;
				k--;
			}

			return UnscaledResult / TotalUnscaledProbability;
		}

		private static double fx(int x, double Mean, double xFactorial)
		{
			return Math.Exp(-Mean) * Math.Pow(Mean, x) / xFactorial;
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			bool Cumulative;
			object v5 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v5 is TFlxFormulaErrorValue) return v5;
			if (!ExtToBool(v5, out Cumulative)) return TFlxFormulaErrorValue.ErrValue;

			double Mean;
			object v4 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v4 is TFlxFormulaErrorValue) return v4;
			if (!ExtToDouble(v4, out Mean)) return TFlxFormulaErrorValue.ErrValue;
			if (Mean < 0) return TFlxFormulaErrorValue.ErrNum;

			int x;
			object v2 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v2 is TFlxFormulaErrorValue) return v2;
			TFlxFormulaErrorValue Err;
			if (!GetInt(v2, out x, out Err)) return Err;
			if (x < 0) return TFlxFormulaErrorValue.ErrNum;

			if (x * Math.Log10(Mean) >= 290 || x > 170) return fx2(x, Mean, Cumulative);

			if (Cumulative)
			{
				double Result = 0;
				double kFactorial = 1;
				for (int k = 0; k <= x; k++)
				{
					Result += fx(k, Mean, kFactorial);
					kFactorial *= (k + 1);
				}

				return Result;

			}
			return fx(x, Mean, TFactToken.Factorial(x));
		}
	}

	internal sealed class TBinomDistToken : TStatDistToken
	{
		internal TBinomDistToken(ptg aId, TCellFunctionData aFuncData) : base(4, aId, aFuncData) { }

		/// <summary>
		/// This is to consider numerical unstability when x is too big.
		/// see http://support.microsoft.com/default.aspx?scid=kb;en-us;827459
		/// </summary>
		/// <param name="x"></param>
		/// <param name="Cumulative"></param>
		/// <param name="n"></param>
		/// <param name="p"></param>
		/// <returns></returns>
		internal static double fx2(int x, int n, double p, bool Cumulative)
		{
			double TotalUnscaledProbability = 0;
			double UnscaledResult = 0;
			double EssentiallyZero = 1e-12;

			int m = (int)(n * p);
			TotalUnscaledProbability++;
			if (m == x) UnscaledResult++;
			if (Cumulative && m < x) UnscaledResult++;

			//k > m
			double PreviousValue = 1;
			bool Done = false;
			int k = m + 1;
			while (!Done && k <= n)
			{
				double CurrentValue = PreviousValue * (n - k + 1) * p / (k * (1 - p));
				TotalUnscaledProbability += CurrentValue;
				if (k == x) UnscaledResult += CurrentValue;
				if (Cumulative && k < x) UnscaledResult += CurrentValue;
				if (CurrentValue <= EssentiallyZero) Done = true;
				PreviousValue = CurrentValue;
				k++;
			}

			// k < m
			PreviousValue = 1;
			Done = false;
			k = m - 1;
			while (!Done && k >= 0)
			{
				double CurrentValue = PreviousValue * (k + 1) * (1 - p) / ((n - k) * p);
				TotalUnscaledProbability += CurrentValue;
				if (k == x) UnscaledResult += CurrentValue;
				if (Cumulative && k < x) UnscaledResult += CurrentValue;
				if (CurrentValue <= EssentiallyZero) Done = true;
				PreviousValue = CurrentValue;
				k--;
			}

			return UnscaledResult / TotalUnscaledProbability;
		}


		private static double fx(int x, int n, double p)
		{
			return TCombinToken.Combin(n, x) * Math.Pow(p, x) * Math.Pow((1 - p), (n - x));
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			bool Cumulative;
			object v5 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v5 is TFlxFormulaErrorValue) return v5;
			if (!ExtToBool(v5, out Cumulative)) return TFlxFormulaErrorValue.ErrValue;

			double p;
			object v4 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v4 is TFlxFormulaErrorValue) return v4;
			if (!ExtToDouble(v4, out p)) return TFlxFormulaErrorValue.ErrValue;
			if (p < 0 || p > 1) return TFlxFormulaErrorValue.ErrNum;

			int n;
			object v3 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v3 is TFlxFormulaErrorValue) return v3;
			TFlxFormulaErrorValue Err;
			if (!GetInt(v3, out n, out Err)) return Err;
			if (n < 0) return TFlxFormulaErrorValue.ErrNum;

			int x;
			object v2 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v2 is TFlxFormulaErrorValue) return v2;
			if (!GetInt(v2, out x, out Err)) return Err;
			if (x < 0 || x > n) return TFlxFormulaErrorValue.ErrNum;

			if (n >= 1030) return fx2(x, n, p, Cumulative);

			if (Cumulative)
			{
				double Result = 0;
				for (int k = 0; k <= x; k++)
				{
					Result += fx(k, n, p);
				}

				return Result;

			}
			return fx(x, n, p);
		}
	}

	internal sealed class TNegBinomDistToken : TStatDistToken
	{
		internal TNegBinomDistToken(int aArgCount, ptg aId, TCellFunctionData aFuncData) : base(aArgCount, aId, aFuncData) { }

		private static double fx(int x, int n, double p)
		{
			return TCombinToken.Combin(x + n - 1, n - 1) * Math.Pow(p, n) * Math.Pow((1 - p), x);
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			double p;
			object v4 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v4 is TFlxFormulaErrorValue) return v4;
			if (!ExtToDouble(v4, out p)) return TFlxFormulaErrorValue.ErrValue;
			if (p < 0 || p > 1) return TFlxFormulaErrorValue.ErrNum;

			int s;
			TFlxFormulaErrorValue Err;
			object v3 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v3 is TFlxFormulaErrorValue) return v3;
			if (!GetInt(v3, out s, out Err)) return Err;
			if (s < 1) return TFlxFormulaErrorValue.ErrNum;

			int fail;
			object v2 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v2 is TFlxFormulaErrorValue) return v2;
			if (!GetInt(v2, out fail, out Err)) return Err;
			if (fail < 0) return TFlxFormulaErrorValue.ErrNum;

			//see http://support.microsoft.com/default.aspx?kbid=828361
			if (fail + s - 1 >= 1030)
				return TBinomDistToken.fx2(s - 1, fail + s - 1, p, false) * p;

			return fx(fail, s, p);
		}
	}

	internal sealed class THypGeomDistToken : TStatDistToken
	{
		internal THypGeomDistToken(int aArgCount, ptg aId, TCellFunctionData aFuncData) : base(aArgCount, aId, aFuncData) { }

		// For those reading the code, the url here does not accept "&", so the real URL is:
		// http://support.microsoft.com/default.aspx?kbid=828515&product=xl2003

		/// <summary>
		/// This is to consider numerical unstability when x is too big.
		/// see http://support.microsoft.com/default.aspx?kbid=828515&amp;product=xl2003
		/// </summary>
		/// <param name="x"></param>
		/// <param name="MM"></param>
		/// <param name="NN"></param>
		/// <param name="n"></param>
		/// <returns></returns>
		internal static double fx2(int x, int n, int MM, int NN)
		{
			double TotalUnscaledProbability = 0;
			double UnscaledResult = 0;
			double EssentiallyZero = 1e-12;

			int m = (int)(MM * n / NN);
			TotalUnscaledProbability++;
			if (m == x) UnscaledResult++;

			//k > m
			double PreviousValue = 1;
			bool Done = false;
			int k = m + 1;
			while (!Done && k <= MM)
			{
				double CurrentValue = PreviousValue * ((double)n - k + 1) * ((double)MM - k + 1) / ((double)k * (NN - n - MM + k));
				TotalUnscaledProbability += CurrentValue;
				if (k == x) UnscaledResult += CurrentValue;
				if (CurrentValue <= EssentiallyZero) Done = true;
				PreviousValue = CurrentValue;
				k++;
			}

			// k < m
			PreviousValue = 1;
			Done = false;
			k = m - 1;
			while (!Done && k >= 0)
			{
				double CurrentValue = PreviousValue * ((double)k + 1) * ((double)NN - n - MM + k + 1) / ((double)(n - k) * (MM - k));
				TotalUnscaledProbability += CurrentValue;
				if (k == x) UnscaledResult += CurrentValue;
				if (CurrentValue <= EssentiallyZero) Done = true;
				PreviousValue = CurrentValue;
				k--;
			}

			return UnscaledResult / TotalUnscaledProbability;
		}


		private static double fx(int x, int n, int M, int NN)
		{
			return TCombinToken.Combin(M, x) * TCombinToken.Combin(NN - M, n - x) / TCombinToken.Combin(NN, n);
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			int NN;
			object v4 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v4 is TFlxFormulaErrorValue) return v4;
			TFlxFormulaErrorValue Err;
			if (!GetInt(v4, out NN, out Err)) return Err;
			if (NN < 0) return TFlxFormulaErrorValue.ErrNum;

			int M;
			object v3 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v3 is TFlxFormulaErrorValue) return v3;
			if (!GetInt(v3, out M, out Err)) return Err;
			if (M < 0 || M > NN) return TFlxFormulaErrorValue.ErrNum;

			int n;
			object v2 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v2 is TFlxFormulaErrorValue) return v2;
			if (!GetInt(v2, out n, out Err)) return Err;
			if (n < 0 || n > NN) return TFlxFormulaErrorValue.ErrNum;

			int x;
			object v1 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v1 is TFlxFormulaErrorValue) return v1;
			if (!GetInt(v1, out x, out Err)) return Err;
			if (x < 0 || x > Math.Min(n, M) || x < Math.Max(0, n - NN + M)) return TFlxFormulaErrorValue.ErrNum;

			if (NN >= 1030) return fx2(x, n, M, NN);

			return fx(x, n, M, NN);
		}
	}

	internal sealed class TGammaDistToken : TStatDistToken
	{
		internal TGammaDistToken(ptg aId, TCellFunctionData aFuncData) : base(4, aId, aFuncData) { }

		#region Complete Gamma function
		internal static double Gamma(double x)
		{

			double[] PP = {
							 1.60119522476751861407E-4,
							 1.19135147006586384913E-3,
							 1.04213797561761569935E-2,
							 4.76367800457137231464E-2,
							 2.07448227648435975150E-1,
							 4.94214826801497100753E-1,
							 9.99999999999999996796E-1
						 };
			double[] QQ = {
							 -2.31581873324120129819E-5,
							 5.39605580493303397842E-4,
							 -4.45641913851797240494E-3,
							 1.18139785222060435552E-2,
							 3.58236398605498653373E-2,
							 -2.34591795718243348568E-1,
							 7.14304917030273074085E-2,
							 1.00000000000000000320E0
						 };

			double q = Math.Abs(x);

			if (q > 33.0)
			{
				if (x < 0.0)
				{
					double p = Math.Floor(q);
					if (p == q) return Double.NaN;
					//int i = (int)p;
					double z = q - p;
					if (z > 0.5)
					{
						p += 1.0;
						z = q - p;
					}
					z = q * Math.Sin(Math.PI * z);
					if (z == 0) return Double.PositiveInfinity;
					z = Math.Abs(z);
					z = Math.PI / (z * Stirling(q));

					return -z;
				}
				else
				{
					return Stirling(x);
				}
			}

		{
			double z = 1.0;
			while (x >= 3.0)
			{
				x -= 1.0;
				z *= x;
			}

			while (x < 0.0)
			{
				if (x == 0.0)
				{
					return Double.PositiveInfinity;
				}
				else
					if (x > -1.0e-9)
				{
					return (z / ((1.0 + 0.5772156649015329 * x) * x));
				}
				z /= x;
				x += 1.0;
			}

			while (x < 2.0)
			{
				if (x == 0.0)
				{
					return Double.PositiveInfinity;
				}
				else
					if (x < 1.0e-9)
				{
					return (z / ((1.0 + 0.5772156649015329 * x) * x));
				}
				z /= x;
				x += 1.0;
			}

			if ((x == 2.0) || (x == 3.0)) return z;

			x -= 2.0;
			double p = eval_polR(PP, x);
			q = eval_polR(QQ, x);
			return z * p / q;
		}
		}

		private static double Stirling(double x)
		{
			double[] STIR = {
								7.87311395793093628397E-4,
								-2.29549961613378126380E-4,
								-2.68132617805781232825E-3,
								3.47222221605458667310E-3,
								8.33333333333482257126E-2,
			};
			const double MAXSTIR = 143.01608;
			const double SQTPI = 2.50662827463100050242E0;

			double w = 1.0 / x;
			double y = Math.Exp(x);

			w = 1.0 + w * eval_polR(STIR, w);

			if (x > MAXSTIR)
			{
				/* Avoid overflow in pow() */
				double v = Math.Pow(x, 0.5 * x - 0.25);
				y = v * (v / y);
			}
			else
			{
				y = Math.Pow(x, x - 0.5) / y;
			}
			y = SQTPI * y * w;
			return y;
		}

		#endregion
		#region Incomplete Gamma Function

		/// <summary>
		/// All Gamma code here is ported from Cephes Math Library
		/// </summary>
		/// <param name="a"></param>
		/// <param name="x"></param>
		internal static double IncGamma(double a, double x)
		{
			if (x <= 0 || a <= 0) return 0.0;
			if (x > 1.0 && x > a) return 1.0 - IncGammaC(a, x);

			/* Compute  x**a * exp(-x) / gamma(a)  */
			double ax = a * Math.Log(x) - x - LogGamma(a);
			//if( ax < -MAXLOG ) return( 0.0 );

			ax = Math.Exp(ax);

			/* power series */
			double r = a;
			double c = 1.0;
			double ans = 1.0;
			double eps = MachinePrecision();

			do
			{
				r += 1.0;
				c *= x / r;
				ans += c;
			}
			while (c / ans > eps);

			return (ans * ax / a);
		}

		static public double IncGammaC(double a, double x)
		{
			const double big = 4.503599627370496e15;
			const double biginv = 2.22044604925031308085e-16;

			if (x <= 0 || a <= 0) return 1.0;
			if (x < 1.0 || x < a) return 1.0 - IncGamma(a, x);

			double ax = a * Math.Log(x) - x - LogGamma(a);

			ax = Math.Exp(ax);

			/* continued fraction */
			double y = 1.0 - a;
			double z = x + y + 1.0;
			double c = 0.0;
			double pkm2 = 1.0;
			double qkm2 = x;
			double pkm1 = x + 1.0;
			double qkm1 = z * x;
			double ans = pkm1 / qkm1;
			double eps = MachinePrecision();

			double t, r;
			do
			{
				c += 1.0;
				y += 1.0;
				z += 2.0;
				double yc = y * c;
				double pk = pkm1 * z - pkm2 * yc;
				double qk = qkm1 * z - qkm2 * yc;
				if (qk != 0)
				{
					r = pk / qk;
					t = Math.Abs((ans - r) / r);
					ans = r;
				}
				else
				{
					t = 1.0;
				}

				pkm2 = pkm1;
				pkm1 = pk;
				qkm2 = qkm1;
				qkm1 = qk;
				if (Math.Abs(pk) > big)
				{
					pkm2 *= biginv;
					pkm1 *= biginv;
					qkm2 *= biginv;
					qkm1 *= biginv;
				}
			} while (t > eps);

			return ans * ax;
		}

		public static double LogGamma(double x)
		{
			double[] A = {
							 8.11614167470508450300E-4,
							 -5.95061904284301438324E-4,
							 7.93650340457716943945E-4,
							 -2.77777777730099687205E-3,
							 8.33333333333331927722E-2
						 };
			double[] B = {
							 -1.37825152569120859100E3,
							 -3.88016315134637840924E4,
							 -3.31612992738871184744E5,
							 -1.16237097492762307383E6,
							 -1.72173700820839662146E6,
							 -8.53555664245765465627E5
						 };
			double[] C = {
							 /* 1.00000000000000000000E0, */
							 -3.51815701436523470549E2,
							 -1.70642106651881159223E4,
							 -2.20528590553854454839E5,
							 -1.13933444367982507207E6,
							 -2.53252307177582951285E6,
							 -2.01889141433532773231E6
						 };

			if (x < -34.0)
			{
				double q = -x;
				double w = LogGamma(q);
				double p = Math.Floor(q);
				if (p == q) return Double.PositiveInfinity;
				double z = q - p;
				if (z > 0.5)
				{
					p += 1.0;
					z = p - q;
				}
				z = q * Math.Sin(Math.PI * z);
				if (z == 0.0) return Double.PositiveInfinity;
				z = Math.Log(Math.PI) - Math.Log(z) - w;
				return z;
			}

			if (x < 13.0)
			{
				double z = 1.0;
				while (x >= 3.0)
				{
					x -= 1.0;
					z *= x;
				}
				while (x < 2.0)
				{
					z /= x;
					if (Double.IsInfinity(z)) return z;
					x += 1.0;
				}
				if (z < 0.0) z = -z;
				if (x == 2.0) return Math.Log(z);
				x -= 2.0;
				double p = x * eval_polR(B, x) / eval_polR1(C, x);
				return (Math.Log(z) + p);
			}

		{
			double q = (x - 0.5) * Math.Log(x) - x + Math.Log(Math.Sqrt(2 * Math.PI));
			if (x > 1.0e8) return (q);

			double p = 1.0 / (x * x);
			if (x >= 1000.0)
			{
				q += ((7.9365079365079365079365e-4 * p
					- 2.7777777777777777777778e-3) * p
					+ 0.0833333333333333333333) / x;
			}
			else
			{
				q += eval_polR(A, p) / x;
			}
			return q;
		}
		}

		#endregion

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			bool Cumulative;
			object v5 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v5 is TFlxFormulaErrorValue) return v5;
			if (!ExtToBool(v5, out Cumulative)) return TFlxFormulaErrorValue.ErrValue;

			double Beta;
			object v4 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v4 is TFlxFormulaErrorValue) return v4;
			if (!ExtToDouble(v4, out Beta)) return TFlxFormulaErrorValue.ErrValue;
			if (Beta <= 0) return TFlxFormulaErrorValue.ErrNum;

			double Alpha;
			object v3 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v3 is TFlxFormulaErrorValue) return v3;
			if (!ExtToDouble(v3, out Alpha)) return TFlxFormulaErrorValue.ErrValue;
			if (Alpha <= 0) return TFlxFormulaErrorValue.ErrNum;

			double x;
			object v2 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v2 is TFlxFormulaErrorValue) return v2;
			if (!ExtToDouble(v2, out x)) return TFlxFormulaErrorValue.ErrValue;
			if (x < 0) return TFlxFormulaErrorValue.ErrNum;

			if (Cumulative)
			{
				double Result = IncGamma(Alpha, x / Beta);
				if (Double.IsInfinity(Result) || Double.IsNaN(Result)) return TFlxFormulaErrorValue.ErrNum;
				return Result;
			}
			else
			{
				double Result = Math.Pow(x, Alpha - 1) * Math.Exp(-x / Beta) / Math.Pow(Beta, Alpha) / Gamma(Alpha);
				if (Double.IsInfinity(Result) || Double.IsNaN(Result)) return TFlxFormulaErrorValue.ErrNum;
				return Result;
			}
		}

	}

	internal sealed class TGammaInvToken : TStatDistToken
	{
		internal TGammaInvToken(ptg aId, TCellFunctionData aFuncData) : base(3, aId, aFuncData) { }

		/// <summary>
		/// Note that the original function returned the *Complemented* result. For small p, 1-p = 0, and this function was not accurate.
		/// Here we have other method, that uses Newton when near and just halves the interval when far. 
		/// As the function is monotonous, we do not have convergence issues.
		/// </summary>
		/// <param name="a"></param>
		/// <param name="y0"></param>
		/// <param name="Beta"></param>
		/// <returns></returns>
		internal static double IncGammaInv(double a, double y0, double Beta)
		{

			const double MaxIter = 100;
			const double MaxLog = 7.09782712893383996843E2;        // log(2**1024)

			// bound the solution 
			double x0 = 0;
			double yl = 0;
			double x1 = Double.MaxValue;
			double yh = 1.0;

			double Eps = MachinePrecision();
			double Eps2 = 1e-14;

			// approximation to inverse function 
			double d = 1.0 / (9.0 * a);
			double y = (1 - d - CalcZ(y0) * Math.Sqrt(d));
			double x = a * y * y * y * Beta;

			double lgm = TGammaDistToken.LogGamma(a);

			for (int i = 0; i < MaxIter; i++)
			{
				if (x <= x0 || x >= x1)
				{
					if (x1 == Double.MaxValue) //Calculate a reasonable bound.
					{
						double dz = 0.0625;
						double z = x0 + 1;
						double yz = y0 - 1;
						while (yz < y0)
						{
							yz = TGammaDistToken.IncGamma(a, z / Beta);
							if (yz < y0)
							{
								x0 = z;
								z = (1.0 + dz) * z;
								dz += dz;
							}
						}
						x1 = z;
					}
					x = (x0 + x1) / 2;
				}

				if (x == x0 || x == x1) return x;  //We cannot "break" the interval anymore.

				y = TGammaDistToken.IncGamma(a, x / Beta);

				if (Math.Abs((y - y0) / y0) < Eps2) return x;
				if (y < yl || y > yh) return Double.NaN;

				if (y < y0)
				{
					x0 = x;
					yl = y;
				}
				else
				{
					x1 = x;
					yh = y;
				}

				// compute the derivative of the function at this point 
				d = (a - 1.0) * Math.Log(x / Beta) - x / Beta - lgm;
				if (d > -MaxLog)
				{

					d = -Math.Exp(d);
					// compute the step to the next approximation of x 
					double dx = (y - y0) / d * Beta;
					if (Math.Abs(dx / x) < Eps) return (x);
					x += dx;
				}
				else
					x = (x0 + x1) / 2;
			}

			return Double.NaN;
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			double Beta;
			object v4 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v4 is TFlxFormulaErrorValue) return v4;
			if (!ExtToDouble(v4, out Beta)) return TFlxFormulaErrorValue.ErrValue;
			if (Beta <= 0) return TFlxFormulaErrorValue.ErrNum;

			double Alpha;
			object v3 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v3 is TFlxFormulaErrorValue) return v3;
			if (!ExtToDouble(v3, out Alpha)) return TFlxFormulaErrorValue.ErrValue;
			if (Alpha <= 0) return TFlxFormulaErrorValue.ErrNum;

			double p;
			object v2 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v2 is TFlxFormulaErrorValue) return v2;
			if (!ExtToDouble(v2, out p)) return TFlxFormulaErrorValue.ErrValue;
			if (p < 0 || p > 1) return TFlxFormulaErrorValue.ErrNum;

			double Result = IncGammaInv(Alpha, p, Beta);
			if (Double.IsInfinity(Result) || Double.IsNaN(Result)) return TFlxFormulaErrorValue.ErrNum;
			return Result;
		}

	}

	internal sealed class TGammaLnToken : TOneDoubleArgToken
	{
		internal TGammaLnToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }

		protected override object Calc(double x)
		{
			if (x <= 0) return TFlxFormulaErrorValue.ErrNum;
			return TGammaDistToken.LogGamma(x);
		}
	}

	internal sealed class TChiDistToken : TStatDistToken
	{
		internal TChiDistToken(ptg aId, TCellFunctionData aFuncData) : base(2, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			int k;
			object v4 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v4 is TFlxFormulaErrorValue) return v4;
			TFlxFormulaErrorValue Err;
			if (!GetInt(v4, out k, out Err)) return Err;
			if (k < 1 || k > 1e10) return TFlxFormulaErrorValue.ErrNum;

			double x;
			object v2 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v2 is TFlxFormulaErrorValue) return v2;
			if (!ExtToDouble(v2, out x)) return TFlxFormulaErrorValue.ErrValue;
			if (x < 0) return TFlxFormulaErrorValue.ErrNum;

			double Result = TGammaDistToken.IncGammaC(k / 2.0, x / 2.0);
			if (Double.IsInfinity(Result) || Double.IsNaN(Result)) return TFlxFormulaErrorValue.ErrNum;
			return Result;
		}

	}

	internal sealed class TChiInvToken : TStatDistToken
	{
		internal TChiInvToken(ptg aId, TCellFunctionData aFuncData) : base(2, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			int k;
			object v4 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v4 is TFlxFormulaErrorValue) return v4;
			TFlxFormulaErrorValue Err;
			if (!GetInt(v4, out k, out Err)) return Err;
			if (k < 1 || k > 1e10) return TFlxFormulaErrorValue.ErrNum;

			double p;
			object v2 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v2 is TFlxFormulaErrorValue) return v2;
			if (!ExtToDouble(v2, out p)) return TFlxFormulaErrorValue.ErrValue;
			if (p <= 0 || p > 1) return TFlxFormulaErrorValue.ErrNum;

			double Result = TGammaInvToken.IncGammaInv(k / 2.0, 1 - p, 2);
			if (Double.IsInfinity(Result) || Double.IsNaN(Result)) return TFlxFormulaErrorValue.ErrNum;
			return Result;
		}

	}

	internal sealed class TChiTestToken : TBaseSumProductToken
	{
		internal TChiTestToken(ptg aId, TCellFunctionData aFuncData) : base(2, aId, aFuncData, true) { }
        
		protected override double DoY(double val)
		{
			return val;
		}

		protected override double DoX(double ResultValue, double val)
		{
			return (val - ResultValue) * (val - ResultValue) / ResultValue;
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			double[,] ScArr = null;
			bool[,] InvalidLines = null;

			object res = CalcResult(FTokenList, wi, CalcState, CalcStack, ref InvalidLines, ref ScArr); if (res != null) return res;
			double Sc = 0;

			int r = ScArr.GetLength(0);
			int c = ScArr.GetLength(1);

			for (int i = 0; i < r; i++)
			{
				for (int k = 0; k < c; k++)
				{
					if (!InvalidLines[i, k])
					{
						Sc += ScArr[i, k];
					}
				}
			}

			if (Double.IsInfinity(Sc)) return TFlxFormulaErrorValue.ErrDiv0;
			if (Sc == 0) return TFlxFormulaErrorValue.ErrDiv0;

			double df = (r - 1) * (c - 1);
			if (r <= 1)
			{
				if (c > 1) df = c - 1;
				else return TFlxFormulaErrorValue.ErrNA;
			}
			else
				if (c <= 1) df = r - 1;

			double Result = TGammaDistToken.IncGammaC(df / 2.0, (double)Sc / 2.0);
			if (Double.IsInfinity(Result) || Double.IsNaN(Result)) return TFlxFormulaErrorValue.ErrNum;
			return Result;

		}
	}

	internal sealed class TKurtToken : TStatToken
	{
		internal TKurtToken(int ArgCount, ptg aId, TCellFunctionData aFuncData) : base(ArgCount, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			int StartPosition = FTokenList.SavePosition();
			TAverageValue Average = ProcessRange(FArgCount, FTokenList, wi, TAverageAggregate.Instance0, CalcState, CalcStack, false);
			if (Average.HasErr) return Average.Err;
			if (Average.ValueCount <= 3) return TFlxFormulaErrorValue.ErrDiv0;
			double xBar = Average.Sum / Average.ValueCount;
			double n = Average.ValueCount;

			FTokenList.RestorePosition(StartPosition);
			TAverageValue StDev = ProcessRange(FArgCount, FTokenList, wi, new TSquaredDiffAggregate(xBar, false), CalcState, CalcStack, false);
			if (StDev.HasErr) return StDev.Err;
			if (StDev.ValueCount <= 1) return TFlxFormulaErrorValue.ErrDiv0;
			if (StDev.Sum == 0) return TFlxFormulaErrorValue.ErrDiv0;
			double Sigma2 = StDev.Sum / (StDev.ValueCount - 1);

			FTokenList.RestorePosition(StartPosition);
			TAverageValue Kurt = ProcessRange(FArgCount, FTokenList, wi, new TNSquaredDiffAggregate(xBar, false, 4), CalcState, CalcStack, false);
			if (Kurt.HasErr) return Kurt.Err;

			return (n) * (n + 1) / (n - 1) / (n - 2) / (n - 3) * Kurt.Sum / (Sigma2 * Sigma2) - 3 * (n - 1) * (n - 1) / (n - 2) / (n - 3);


		}
	}

	internal sealed class TSkewToken : TStatToken
	{
		internal TSkewToken(int ArgCount, ptg aId, TCellFunctionData aFuncData) : base(ArgCount, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			int StartPosition = FTokenList.SavePosition();
			TAverageValue Average = ProcessRange(FArgCount, FTokenList, wi, TAverageAggregate.Instance0, CalcState, CalcStack, false);
			if (Average.HasErr) return Average.Err;
			if (Average.ValueCount <= 2) return TFlxFormulaErrorValue.ErrDiv0;
			double xBar = Average.Sum / Average.ValueCount;
			double n = Average.ValueCount;

			FTokenList.RestorePosition(StartPosition);
			TAverageValue StDev = ProcessRange(FArgCount, FTokenList, wi, new TSquaredDiffAggregate(xBar, false), CalcState, CalcStack, false);
			if (StDev.HasErr) return StDev.Err;
			if (StDev.ValueCount <= 1) return TFlxFormulaErrorValue.ErrDiv0;
			if (StDev.Sum == 0) return TFlxFormulaErrorValue.ErrDiv0;
			double Sigma = Math.Sqrt(StDev.Sum / (StDev.ValueCount - 1));

			FTokenList.RestorePosition(StartPosition);
			TAverageValue Skew = ProcessRange(FArgCount, FTokenList, wi, new TNSquaredDiffAggregate(xBar, false, 3), CalcState, CalcStack, false);
			if (Skew.HasErr) return Skew.Err;

			return (n) / (n - 1) / (n - 2) * Skew.Sum / (Sigma * Sigma * Sigma);
		}
	}

	internal sealed class TConfidenceToken : TStatDistToken
	{
		internal TConfidenceToken(ptg aId, TCellFunctionData aFuncData) : base(3, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			int Size;
			object v3 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v3 is TFlxFormulaErrorValue) return v3;
			TFlxFormulaErrorValue Err;
			if (!GetInt(v3, out Size, out Err)) return Err;
			if (Size < 1) return TFlxFormulaErrorValue.ErrNum;

			double Sigma;
			object v4 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v4 is TFlxFormulaErrorValue) return v4;
			if (!GetDouble(v4, out Sigma)) return TFlxFormulaErrorValue.ErrValue;
			if (Sigma <= 0) return TFlxFormulaErrorValue.ErrNum;

			double Alpha;
			object v1 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v1 is TFlxFormulaErrorValue) return v1;
			if (!GetDouble(v1, out Alpha)) return TFlxFormulaErrorValue.ErrValue;
			if (Alpha <= 0 || Alpha >= 1) return TFlxFormulaErrorValue.ErrNum;

			return -CalcZ(Alpha / 2) * Sigma / Math.Sqrt((double)Size);
		}
	}

	internal sealed class TStandardizeToken : TNDoubleArgToken
	{
		internal TStandardizeToken(ptg aId, TCellFunctionData aFuncData) : base(3, aId, aFuncData) { }

		protected override object Calc(double[] x)
		{
			double Sigma = x[0];
			if (Sigma <= 0) return TFlxFormulaErrorValue.ErrNum;

			double Mean = x[1];

			double xx = x[2];

			return (xx - Mean) / Sigma;
		}
	}

	#endregion

	#region Statitical III
	internal sealed class TAveDevToken : TStatToken
	{
		internal TAveDevToken(int ArgCount, ptg aId, TCellFunctionData aFuncData) : base(ArgCount, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			int StartPosition = FTokenList.SavePosition();
			TAverageValue Average = ProcessRange(FArgCount, FTokenList, wi, TAverageAggregate.Instance0, CalcState, CalcStack, false);
			if (Average.HasErr) return Average.Err;
			if (Average.ValueCount <= 0) return TFlxFormulaErrorValue.ErrNum;
			double xBar = Average.Sum / Average.ValueCount;

			FTokenList.RestorePosition(StartPosition);
			TAverageValue ModSum = ProcessRange(FArgCount, FTokenList, wi, new TModDiffAggregate(xBar, false), CalcState, CalcStack, false);
			if (ModSum.HasErr) return ModSum.Err;
			if (ModSum.ValueCount <= 0) return TFlxFormulaErrorValue.ErrDiv0;

			return ModSum.Sum / ModSum.ValueCount;
		}
	}

	internal sealed class TDevSqToken : TStatToken
	{
		internal TDevSqToken(int ArgCount, ptg aId, TCellFunctionData aFuncData) : base(ArgCount, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			int StartPosition = FTokenList.SavePosition();
			TAverageValue Average = ProcessRange(FArgCount, FTokenList, wi, TAverageAggregate.Instance0, CalcState, CalcStack, false);
			if (Average.HasErr) return Average.Err;
			if (Average.ValueCount <= 0) return TFlxFormulaErrorValue.ErrNum;
			double xBar = Average.Sum / Average.ValueCount;

			FTokenList.RestorePosition(StartPosition);
			TAverageValue SqDiff = ProcessRange(FArgCount, FTokenList, wi, new TSquaredDiffAggregate(xBar, false), CalcState, CalcStack, false);
			if (SqDiff.HasErr) return SqDiff.Err;

			return SqDiff.Sum;
		}
	}

	internal abstract class TBaseXYStatToken : TBaseSumProductToken
	{
		protected TBaseXYStatToken(ptg aId, TCellFunctionData aFuncData)
			: base(2, aId, aFuncData, true)
		{
		}

		protected override double DoY(double val)
		{
			return val;
		}

		protected override double DoX(double ResultValue, double val)
		{
			return val;
		}

		internal object CalcParams(TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack, out double n, out double xSqDev, out double ySqDev, out double xyDev, out double xBar, out double yBar)
		{
			n = 0; xSqDev = 0; ySqDev = 0; xyDev = 0;
			xBar = 0; yBar = 0;
			//First Pass - Calculate X and Y Values.

			double[,] YValues = null;
			double[,] XValues = null;
			bool[,] InvalidLines = null;

			object res = CalcResult(FTokenList, wi, CalcState, CalcStack, ref InvalidLines, ref XValues, ref YValues, true); if (res != null) return res;

			int r = XValues.GetLength(0);
			int c = XValues.GetLength(1);
			n = 0;
			for (int i = 0; i < r; i++)
			{
				for (int k = 0; k < c; k++)
				{
					if (!InvalidLines[i, k])
					{
						yBar += YValues[i, k];
						xBar += XValues[i, k];
						n++;
					}
				}
			}

			if (n == 0) return TFlxFormulaErrorValue.ErrValue;

			xBar /= n;
			yBar /= n;

			//Now calculate the other values.
			for (int i = 0; i < r; i++)
			{
				for (int k = 0; k < c; k++)
				{
					if (!InvalidLines[i, k])
					{
						xyDev += (YValues[i, k] - yBar) * (XValues[i, k] - xBar);
						double d = (XValues[i, k] - xBar); xSqDev += d * d;
						d = (YValues[i, k] - yBar); ySqDev += d * d;
					}
				}
			}

			if (Double.IsInfinity(xyDev)) return TFlxFormulaErrorValue.ErrDiv0;

			return null;
		}
	}

	internal sealed class TSteyxToken : TBaseXYStatToken
	{
		internal TSteyxToken(ptg aId, TCellFunctionData aFuncData)
			: base(aId, aFuncData)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			double n; double xSqDev; double ySqDev; double xyDev; double xBar; double yBar;
			object res = CalcParams(FTokenList, wi, CalcState, CalcStack, out n, out xSqDev, out ySqDev, out xyDev, out xBar, out yBar);
			if (res != null) return res;
			if (xSqDev == 0) return TFlxFormulaErrorValue.ErrDiv0;

			if (n <= 2) return TFlxFormulaErrorValue.ErrDiv0;

			double Result = 1 / (n - 2) * (xSqDev - xyDev * xyDev / ySqDev);  //x and y are inverted here, since the y argument is the first.
			if (Double.IsInfinity(Result) || Double.IsNaN(Result)) return TFlxFormulaErrorValue.ErrNum;
			if (Result < 0) return TFlxFormulaErrorValue.ErrNum;
			return Math.Sqrt(Result);

		}
	}

	internal sealed class TRsqToken : TBaseXYStatToken
	{
		internal TRsqToken(ptg aId, TCellFunctionData aFuncData)
			: base(aId, aFuncData)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			double n; double xSqDev; double ySqDev; double xyDev; double xBar; double yBar;
			object res = CalcParams(FTokenList, wi, CalcState, CalcStack, out n, out xSqDev, out ySqDev, out xyDev, out xBar, out yBar);
			if (res != null) return res;
			if (xSqDev == 0 || ySqDev == 0) return TFlxFormulaErrorValue.ErrDiv0;


			double Result = xyDev * xyDev / xSqDev / ySqDev;  //x and y are inverted here, since the y argument is the first.
			if (Double.IsInfinity(Result) || Double.IsNaN(Result)) return TFlxFormulaErrorValue.ErrNum;
			return Result;

		}
	}

	internal sealed class TPearsonToken : TBaseXYStatToken
	{
		internal TPearsonToken(ptg aId, TCellFunctionData aFuncData)
			: base(aId, aFuncData)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			double n; double xSqDev; double ySqDev; double xyDev; double xBar; double yBar;
			object res = CalcParams(FTokenList, wi, CalcState, CalcStack, out n, out xSqDev, out ySqDev, out xyDev, out xBar, out yBar);
			if (res != null) return res;
			if (xSqDev == 0 || ySqDev == 0) return TFlxFormulaErrorValue.ErrDiv0;


			double Result = xyDev / Math.Sqrt(xSqDev * ySqDev);  //x and y are inverted here, since the y argument is the first.
			if (Double.IsInfinity(Result) || Double.IsNaN(Result)) return TFlxFormulaErrorValue.ErrNum;
			return Result;

		}
	}
   
	internal sealed class TSlopeToken : TBaseXYStatToken
	{
		internal TSlopeToken(ptg aId, TCellFunctionData aFuncData)
			: base(aId, aFuncData)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			double n; double xSqDev; double ySqDev; double xyDev; double xBar; double yBar;
			object res = CalcParams(FTokenList, wi, CalcState, CalcStack, out n, out xSqDev, out ySqDev, out xyDev, out xBar, out yBar);
			if (res != null) return res;
			if (ySqDev == 0) return TFlxFormulaErrorValue.ErrDiv0;


			double Result = xyDev / ySqDev;  //x and y are inverted here, since the y argument is the first.
			if (Double.IsInfinity(Result) || Double.IsNaN(Result)) return TFlxFormulaErrorValue.ErrNum;
			return Result;

		}
	}

	internal sealed class TInterceptToken : TBaseXYStatToken
	{
		internal TInterceptToken(ptg aId, TCellFunctionData aFuncData)
			: base(aId, aFuncData)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			double n; double xSqDev; double ySqDev; double xyDev; double xBar; double yBar;
			object res = CalcParams(FTokenList, wi, CalcState, CalcStack, out n, out xSqDev, out ySqDev, out xyDev, out xBar, out yBar);
			if (res != null) return res;
			if (ySqDev == 0) return TFlxFormulaErrorValue.ErrDiv0;


			double Slope = xyDev / ySqDev;  //x and y are inverted here, since the y argument is the first.
			if (Double.IsInfinity(Slope) || Double.IsNaN(Slope)) return TFlxFormulaErrorValue.ErrNum;
			return xBar - yBar * Slope;

		}
	}

	internal sealed class TFisherToken : TOneDoubleArgToken
	{
		internal TFisherToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }

		protected override object Calc(double x)
		{
			if (x <= -1 || x >= 1) return TFlxFormulaErrorValue.ErrNum;
			return Math.Log((1 + x) / (1 - x)) / 2;
		}
	}

	internal sealed class TFisherInvToken : TOneDoubleArgToken
	{
		internal TFisherInvToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) { }

		protected override object Calc(double x)
		{
			double e2x = Math.Exp(2 * x);
			return (e2x - 1) / (e2x + 1);
		}
	}

    internal sealed class TFrequencyToken : TBaseFunctionToken
    {
		internal TFrequencyToken(ptg aId, TCellFunctionData aFuncData) : base(2, aId, aFuncData) { }

        internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
        {
            object oBinList = FTokenList.EvaluateToken(wi, TArrayAggregate.Instance, CalcState, CalcStack);
            object[,] BinList = oBinList as object[,];
            if (BinList == null) BinList = new object[1,1];

            TFrequencyAggregate FreqAgg = TFrequencyAggregate.Create(BinList);
            if (FreqAgg == null) return TFlxFormulaErrorValue.ErrNA;

            object oResult = FTokenList.EvaluateToken(wi, FreqAgg, CalcState, CalcStack);
            if (oResult is TFlxFormulaErrorValue) return oResult;
            object[,] Result = oResult as object[,];
            if (Result == null) return TFlxFormulaErrorValue.ErrNA;

            return Result;
        }
    }


	#endregion

	#region Percentiles
	internal sealed class TMedianToken : TBaseFunctionToken
	{
		internal TMedianToken(int ArgCount, ptg aId, TCellFunctionData aFuncData) : base(ArgCount, aId, aFuncData) { }

		internal static object GetArray(TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack, int ArgCount)
		{
			TDoubleList Result = new TDoubleList();

			for (int i = 0; i < ArgCount; i++)
			{
				object Current = FTokenList.EvaluateToken(wi, new TMedianAggregate(), CalcState, CalcStack);
                TDoubleList Double1 = Current as TDoubleList;

                if (Double1 != null)
                {
                    Result.AddRange(Double1);
                    continue;
                }
             
				object[,] CArray = Current as Object[,];
				if (CArray != null)
				{
					for (int r = 0; r < CArray.GetLength(0); r++)
						for (int c = 0; c < CArray.GetLength(1); c++)
						{
							object o = ConvertToAllowedObject(CArray[r, c]);
							if (o is TFlxFormulaErrorValue) return o;
							if (o is double)
							{
								Result.Add((double)o);
							}
						}
                    continue;
				}
				
				object o1 = ConvertToAllowedObject(Current);
				if (o1 is TFlxFormulaErrorValue) return o1;
				if (o1 is double)
				{
					Result.Add((double)o1);
				}
			}

			if (Result.Count <= 0) return TFlxFormulaErrorValue.ErrNum;
			Result.Sort();

			return Result;
		}


		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
            return CalcMedian(FTokenList, wi, FArgCount, CalcState, CalcStack);
		}

        internal static object CalcMedian(TParsedTokenList FTokenList, TWorkbookInfo wi, int ArgCount, TCalcState CalcState, TCalcStack CalcStack)
        {
            object Result = GetArray(FTokenList, wi, CalcState, CalcStack, ArgCount);
            if (Result is TFlxFormulaErrorValue) return Result;

            TDoubleList Values = Result as TDoubleList;
            if (Values == null) return TFlxFormulaErrorValue.ErrNA; //Should not happen

            int i = (Values.Count - 1) / 2;
            if (Values.Count % 2 == 0)
                return (Values[i] + Values[i + 1]) / 2.0;

            return Values[i];
        }
	}

	internal sealed class TPercentileToken : TBaseFunctionToken
	{
		internal bool DoQuartile;
        internal bool Exclusive;

		internal TPercentileToken(ptg aId, TCellFunctionData aFuncData, bool aDoQuartile, bool aExclusive)
			: base(2, aId, aFuncData)
		{
			DoQuartile = aDoQuartile;
            Exclusive = aExclusive;
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
            return CalcPercentile(DoQuartile, Exclusive, FTokenList, wi, CalcState, CalcStack);
		}

        internal static object CalcPercentile(bool DoQuartile, bool Exclusive, TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack)
        {
            double k = 0; object[,] ArrK;
            object Ret = GetDoubleArgument(FTokenList, wi, CalcState, CalcStack, ref k, out ArrK); if (Ret != null) return Ret;
            object Result = TMedianToken.GetArray(FTokenList, wi, CalcState, CalcStack, 1);
            if (Result is TFlxFormulaErrorValue) return Result;

            TDoubleList Values = Result as TDoubleList;
            if (Values == null) return TFlxFormulaErrorValue.ErrNA; //Should not happen

            if (ArrK != null)
            {
                Object[,] ArrResult = new object[ArrK.GetLength(0), ArrK.GetLength(1)];
                for (int i = 0; i < ArrResult.GetLength(0); i++)
                    for (int j = 0; j < ArrResult.GetLength(1); j++)
                    {
                        if (!GetDoubleItem(ref ArrResult, ref ArrK, i, j, out k)) continue;
                        ArrResult[i, j] = CalcPercentile(DoQuartile, Exclusive, Values, k);
                    }

                return UnPack(ArrResult);
            }

            return CalcPercentile(DoQuartile, Exclusive, Values, k);
        }

		private static object CalcPercentile(bool DoQuartile, bool Exclusive, TDoubleList Values, double k)
		{
			if (DoQuartile)
			{
				int q;
				TFlxFormulaErrorValue Err;
				if (!GetUInt(k, out q, out Err)) return TFlxFormulaErrorValue.ErrNum;
				switch (q)
				{
                    case 0: k = 0; break;
					case 1: k = 0.25; break;
					case 2: k = 0.50; break;
					case 3: k = 0.75; break;
                    case 4: k = 1; break;
					default: return TFlxFormulaErrorValue.ErrNum;
				}
			}
            if (k < 0 || k > 1) return TFlxFormulaErrorValue.ErrNum;
            if (Exclusive && (k == 0 || k == 1)) return TFlxFormulaErrorValue.ErrNum;

			double n;
			int i;
            if (Exclusive)
            {
                if (k < 1d / (Values.Count + 1)) return TFlxFormulaErrorValue.ErrNum;
                if (k > 1d * (Values.Count) / (Values.Count + 1)) return TFlxFormulaErrorValue.ErrNum;

                n = (Values.Count + 1) * k - 1;
            }
            else
            {
                n = (Values.Count - 1) * k;
            }

    	    i = (int)n;
			if (i == Values.Count - 1) return Values[i]; //To avoid overflow
			return (Values[i] + (n - i) * (Values[i + 1] - Values[i]));
		}
	}

	internal sealed class TModeToken : TBaseFunctionToken
	{
		internal TModeToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
            return CalcMode(FArgCount, FTokenList, wi, CalcState, CalcStack);
		}

        internal static object CalcMode(int FArgCount, TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack)
        {
            object Result = TMedianToken.GetArray(FTokenList, wi, CalcState, CalcStack, FArgCount);
            if (Result is TFlxFormulaErrorValue) return Result;

            TDoubleList Values = Result as TDoubleList;
            if (Values == null) return TFlxFormulaErrorValue.ErrNA; //Should not happen

            return CalcMode(Values);
        }

		private static object CalcMode(TDoubleList Values)
		{
			if (Values.Count <= 1) return TFlxFormulaErrorValue.ErrNA;
			int MaxCount = 0;
			double Result = 0;

			int i = 0;
			while (i < Values.Count)
			{
				int k = 1;
				while (i + k < Values.Count && Values[i + k] == Values[i])
				{
					k++;
				}
				if (k > MaxCount)
				{
					MaxCount = k;
					Result = Values[i];
				}
				i += k;
			}

			if (MaxCount <= 1) return TFlxFormulaErrorValue.ErrNA;
			return Result;

		}
	}

	#endregion

	#region Financial
	internal abstract class TDeprBaseToken : TBaseFunctionToken
	{
		protected TDeprBaseToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData)
		{
		}

		protected object GetArguments(TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack, bool Check, out double dPeriod, out double Life, out double Salvage, out double Cost)
		{
			dPeriod = 0; Life = 0; Salvage = 0; Cost = 0;
			if (FArgCount >= 4)
			{
				object v4 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v4 is TFlxFormulaErrorValue) return v4;
				if (!ExtToDouble(v4, out dPeriod)) return TFlxFormulaErrorValue.ErrValue;
				if (Check && dPeriod <= 0) return TFlxFormulaErrorValue.ErrNum;
			}

			object v3 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v3 is TFlxFormulaErrorValue) return v3;
			if (!ExtToDouble(v3, out Life)) return TFlxFormulaErrorValue.ErrValue;
			if (Check && Life < 0) return TFlxFormulaErrorValue.ErrNum;

			object v2 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v2 is TFlxFormulaErrorValue) return v2;
			if (!ExtToDouble(v2, out Salvage)) return TFlxFormulaErrorValue.ErrValue;
			if (Check && Salvage < 0) return TFlxFormulaErrorValue.ErrNum;

			object v1 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v1 is TFlxFormulaErrorValue) return v1;
			if (!ExtToDouble(v1, out Cost)) return TFlxFormulaErrorValue.ErrValue;

			return null;
		}
	}
    
	internal sealed class TDBToken : TDeprBaseToken
	{
		internal TDBToken(int ArgCount, ptg aId, TCellFunctionData aFuncData) : base(ArgCount, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			int Month = 12;
			if (FArgCount > 4)
			{
				object v5 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v5 is TFlxFormulaErrorValue) return v5;
				TFlxFormulaErrorValue Err;
				if (!GetInt(v5, out Month, out Err)) return Err;
				if (Month < 1 || Month > 12) return TFlxFormulaErrorValue.ErrNum;
			}

			double dPeriod; double Life; double Salvage; double Cost;
			object ret = GetArguments(FTokenList, wi, CalcState, CalcStack, true, out dPeriod, out Life, out Salvage, out Cost);
			if (ret != null) return ret;
			if (Cost < 0) return TFlxFormulaErrorValue.ErrNum;
			int Period = (int)dPeriod; if (Period == 0) Period = 1; //weird, but it is this way. 0.1 is 1, 1.9 is 1

			if (Period > Life + 1 || (Month == 12 && Period > Life)) return TFlxFormulaErrorValue.ErrNum;
			if (Cost == 0) return 0;

			return CalcDB(Cost, Salvage, Life, Period, Month);
		}

		private static object CalcDB(double Cost, double Salvage, double Life, int Period, int Month)
		{
			Double Rate = Math.Round(1 - Math.Pow(Salvage / Cost, 1 / Life), 3);
			double Depreciation = Cost * Rate * Month / 12;
			double TotalDepreciation = Depreciation;

			for (int i = Math.Min(Period, (int)Life); i > 1; i--)
			{
				Depreciation = (Cost - TotalDepreciation) * Rate;
				TotalDepreciation += Depreciation;
			}

			if (Period > (int)Life)
			{
				Depreciation = ((Cost - TotalDepreciation) * Rate * (12 - Month)) / 12;
			}
			return Depreciation;
		}
	}
    
	internal sealed class TDDBToken : TDeprBaseToken
	{
		internal TDDBToken(int ArgCount, ptg aId, TCellFunctionData aFuncData) : base(ArgCount, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			double Factor = 2;
			if (FArgCount > 4)
			{
				object v5 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v5 is TFlxFormulaErrorValue) return v5;
				if (!ExtToDouble(v5, out Factor)) return TFlxFormulaErrorValue.ErrValue; ;
				if (Factor <= 0) return TFlxFormulaErrorValue.ErrNum;
			}

			double Period; double Life; double Salvage; double Cost;
			object ret = GetArguments(FTokenList, wi, CalcState, CalcStack, true, out Period, out Life, out Salvage, out Cost);
			if (ret != null) return ret;
			if (Cost < 0) return TFlxFormulaErrorValue.ErrNum;

			if (Period > Life) return TFlxFormulaErrorValue.ErrNum;
			if (Cost == 0) return 0;

			return CalcDDB(Cost, Salvage, Life, Period, Factor);
		}

		private static object CalcDDB(double Cost, double Salvage, double Life, double Period, double Factor)
		{
			if (Life < 2) return Cost - Salvage;

			if (Life == 2)
			{
				if (Period > 1) return 0;
				return (Cost - Salvage);
			}

			if (Period <= 1)
			{
				return Math.Min((Cost * Factor) / Life, Cost - Salvage);
			}

			double LF = (Life - Factor) / Life;
			double Result = ((Factor * Cost) / Life) * Math.Pow(LF, Period - 1);
			double Depr = (-Cost * Math.Pow(LF, Period)) + Salvage;
			if (Depr > 0)
				Result -= Depr;
			if (Result >= 0)
				return Result;
			return 0;
		}


	}

	internal sealed class TSydToken : TDeprBaseToken
	{
		internal TSydToken(ptg aId, TCellFunctionData aFuncData) : base(4, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			double Period; double Life; double Salvage; double Cost;
			object ret = GetArguments(FTokenList, wi, CalcState, CalcStack, true, out Period, out Life, out Salvage, out Cost);
			if (ret != null) return ret;

			if (Period > Life) return TFlxFormulaErrorValue.ErrNum;
			if (Cost == 0) return 0;

			return CalcSyd(Cost, Salvage, Life, Period);
		}

		private static object CalcSyd(double Cost, double Salvage, double Life, double Period)
		{
			return (Cost - Salvage) * (Life - Period + 1) * 2 / ((Life) * (Life + 1));
		}
	}

	internal sealed class TSlnToken : TDeprBaseToken
	{
		internal TSlnToken(ptg aId, TCellFunctionData aFuncData) : base(3, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			double Period; double Life; double Salvage; double Cost;
			object ret = GetArguments(FTokenList, wi, CalcState, CalcStack, false, out Period, out Life, out Salvage, out Cost);
			if (ret != null) return ret;

			if (Cost == 0) return 0;

			return CalcSln(Cost, Salvage, Life);
		}

		private static object CalcSln(double Cost, double Salvage, double Life)
		{
			if (Life == 0) return TFlxFormulaErrorValue.ErrDiv0;
			return (Cost - Salvage) / Life;
		}
	}

	internal abstract class TPVBaseToken : TBaseFunctionToken
	{
		protected TPVBaseToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData)
		{
		}

		protected object GetArguments(TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack, bool HasPer, out double Rate, out double Per, out  double nPer, out  double Pv, out  double Fv, out  int pType)
		{
			Rate = 0; nPer = 0; Pv = 0; Fv = 0; pType = 0; Per = 0;
			int PerOfs = HasPer ? 1 : 0;
			if (FArgCount >= 5 + PerOfs)
			{
				double dType;
				object v5 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v5 is TFlxFormulaErrorValue) return v5;
				if (!GetDouble(v5, out dType)) return TFlxFormulaErrorValue.ErrValue;
				if (dType != 0) pType = 1;
			}

			if (FArgCount >= 4 + PerOfs)
			{
				object v4 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v4 is TFlxFormulaErrorValue) return v4;
				if (!ExtToDouble(v4, out Fv)) return TFlxFormulaErrorValue.ErrValue;
			}

			object v3 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v3 is TFlxFormulaErrorValue) return v3;
			if (!ExtToDouble(v3, out Pv)) return TFlxFormulaErrorValue.ErrValue;

			object v2 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v2 is TFlxFormulaErrorValue) return v2;
			if (!ExtToDouble(v2, out nPer)) return TFlxFormulaErrorValue.ErrValue;

			if (HasPer)
			{
				object vPer = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (vPer is TFlxFormulaErrorValue) return vPer;
				if (!ExtToDouble(vPer, out Per)) return TFlxFormulaErrorValue.ErrValue;
				if (Per <= 0 || Per >= nPer + 1) return TFlxFormulaErrorValue.ErrNum;
			}

			object v1 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v1 is TFlxFormulaErrorValue) return v1;
			if (!ExtToDouble(v1, out Rate)) return TFlxFormulaErrorValue.ErrValue;

			return null;
		}
	}

	internal sealed class TPVToken : TPVBaseToken
	{
		internal TPVToken(int ArgCount, ptg aId, TCellFunctionData aFuncData) : base(ArgCount, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			double Rate; double Per; double nPer; double Pmt; double Fv; int pType;
			object ret = GetArguments(FTokenList, wi, CalcState, CalcStack, false, out Rate, out Per, out nPer, out Pmt, out Fv, out pType);
			if (ret != null) return ret;

			return CalcPV(Rate, nPer, Pmt, Fv, pType);
		}

		internal static object CalcPV(double Rate, double nPer, double Pmt, double Fv, int pType)
		{
			if (Rate == 0) return -Fv - Pmt * nPer;

			double FracPer = 1;
			if (pType != 0) FracPer += Rate;
			double tv = Math.Pow(1 + Rate, nPer);
			return (-Fv - Pmt * FracPer * (tv - 1) / Rate) / tv;
		}
	}

	internal sealed class TFVToken : TPVBaseToken
	{
		internal TFVToken(int ArgCount, ptg aId, TCellFunctionData aFuncData) : base(ArgCount, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			double Rate; double Per; double nPer; double Pmt; double Pv; int pType;
			object ret = GetArguments(FTokenList, wi, CalcState, CalcStack, false, out Rate, out Per, out nPer, out Pmt, out Pv, out pType);
			if (ret != null) return ret;

			return CalcFV(Rate, nPer, Pmt, Pv, pType);
		}

		internal static object CalcFV(double Rate, double nPer, double Pmt, double Pv, int pType)
		{
			if (Rate == 0) return -Pv - Pmt * nPer;

			double FracPer = 1;
			if (pType != 0) FracPer += Rate;
			double tv = Math.Pow(1 + Rate, nPer);
			return -(Pv * tv) - (FracPer * Pmt / Rate) * (tv - 1);
		}
	}

	internal enum TValueSign
	{
		Negative,
		All,
		Positive
	}

	internal sealed class TNPVToken : TPVBaseToken
	{
		internal TNPVToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			//We cannot use a simple aggregate here because we cannot aggregate until 
			//we know the rate, and in RPN this is the last parameter.
			//So we will have to look at the objects by hand.

			object[] Values = new object[FArgCount - 1];
			for (int i = FArgCount - 1; i > 0; i--)
			{
				Values[i - 1] = GetValues(FTokenList, wi, CalcState, CalcStack);
				if (Values[i - 1] is TFlxFormulaErrorValue) return Values[i - 1];
			}
			double Rate;
			object v1 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v1 is TFlxFormulaErrorValue) return v1;
			if (!ExtToDouble(v1, out Rate)) return TFlxFormulaErrorValue.ErrValue;


			long ValueCount;
			return CalcNPV(wi.Xls, Rate, Values, CalcState, CalcStack, TValueSign.All, out ValueCount);
		}

		internal static object GetValues(TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack)
		{
			//Here we can have a reference or a value. As we don't know this before, we must try both.
			//int OriginalPos = FTokenList.SavePosition();

			TAddressList CellRefs = null;
			object[,] values = null;
			object ret = GetRangeOrArray(FTokenList, wi, CalcState, CalcStack, out CellRefs, out values, TFlxFormulaErrorValue.ErrValue, true);
			if (ret != null) return ret;
			//{
			//    FTokenList.RestorePosition(OriginalPos);
			//    object v1= FTokenList.EvaluateToken(wi, CalcState, CalcStack);if (v1 is TFlxFormulaErrorValue) return v1;
			//    return v1;
			//}

			if (values != null) return values;

			if (CellRefs.Has3dRef()) return TFlxFormulaErrorValue.ErrValue;

			return CellRefs;

		}

		internal static object CalcNPV(ExcelFile Xls, double Rate, object[] Values, TCalcState CalcState, TCalcStack CalcStack, TValueSign ValueSign, out long ValueCount)
		{
			ValueCount = 0;
			if (Rate == -1) return TFlxFormulaErrorValue.ErrDiv0;
			double Result = 0;
			double AcumRate = 1 + Rate;
			for (int i = 0; i < Values.Length; i++)
			{
				object ret = AddCashFlows(Xls, Values[i], Rate, ref Result, ref AcumRate, ref ValueCount, CalcState, CalcStack, ValueSign);
				if (ret != null) return ret;
			}

			if (ValueCount <= 0) return TFlxFormulaErrorValue.ErrValue;  //No data points.
			return Result;
		}

		private static object AddCashFlows(ExcelFile Xls, object Value, double Rate, ref double ResultValue, ref double AcumRate, ref long ValueCount, TCalcState CalcState, TCalcStack CalcStack, TValueSign ValueSign)
		{
			object[,] a1 = Value as object[,];
			if (a1 != null)
			{
				for (int i = 0; i < a1.GetLength(0); i++)
					for (int k = 0; k < a1.GetLength(1); k++)
					{
						if (a1[i, k] is TFlxFormulaErrorValue) return a1[i, k];
						if (a1[i, k] is double)  //no exttodouble here.
						{
							AddCashFlow((double)a1[i, k], Rate, ref ResultValue, ref AcumRate, ref ValueCount, ValueSign);
						}
					}
			}

			else
			{
				TAddressList Addr = Value as TAddressList;
				if (Addr != null)
				{
					for (int i = Addr.Count - 1; i >= 0; i--) //Reversed order
					{
						object res = ProcessAddr(Xls, Addr[i], Rate, ref ResultValue, ref AcumRate, ref ValueCount, CalcState, CalcStack, ValueSign);
						if (res != null) return res;
					}
				}
				else
				{
					double v1;
					if (ExtToDouble(Value, out v1))
					{
						AddCashFlow(v1, Rate, ref ResultValue, ref AcumRate, ref ValueCount, ValueSign);
					}
				}
			}

			return null;
		}

		private static object ProcessAddr(ExcelFile Xls, TAddress[] Addr, double Rate, ref double ResultValue, ref double AcumRate, ref long ValueCount, TCalcState CalcState, TCalcStack CalcStack, TValueSign ValueSign)
		{
			if (Addr.Length < 1 || Addr.Length > 2) return TFlxFormulaErrorValue.ErrValue;
			TAddress a1 = Addr[0];
			TAddress a2 = a1;
			if (Addr.Length > 1)
				a2 = Addr[1];

			if (a1 == null || a2 == null) return TFlxFormulaErrorValue.ErrNA;

			if (a1.wi.Xls != a2.wi.Xls || a1.Sheet != a2.Sheet)
				return TFlxFormulaErrorValue.ErrRef;

			int MinRow = Math.Min(a1.Row, a2.Row);
			int MinCol = Math.Min(a1.Col, a2.Col);
			int MaxRow = Math.Max(a1.Row, a2.Row);
			int MaxCol = Math.Max(a1.Col, a2.Col);

			for (int r = MinRow; r <= MaxRow; r++)
				for (int c = MinCol; c <= MaxCol; c++)
				{
					object o = ConvertToAllowedObject(a1.wi.Xls.GetCellValueAndRecalc(a1.Sheet, r, c, CalcState, CalcStack));
					if (o is TFlxFormulaErrorValue) return o;

					if (o is double)   //no exttodouble here.
					{
						AddCashFlow((double)o, Rate, ref ResultValue, ref AcumRate, ref ValueCount, ValueSign);
					}

				}
			return null;
		}

		private static void AddCashFlow(double val, double Rate, ref double ResultValue, ref double AcumRate, ref long ValueCount, TValueSign ValueSign)
		{
			if ((ValueSign == TValueSign.Negative && val <= 0) || ValueSign == TValueSign.All || (ValueSign == TValueSign.Positive && val >= 0))
			{
				ResultValue += val / AcumRate;
			}
			AcumRate *= (1 + Rate);
			ValueCount++;
		}
	}

	internal sealed class TIRRToken : TPVBaseToken
	{
		internal TIRRToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			double Guess = 0.1;
			if (FArgCount > 1)
			{
				object v1 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v1 is TFlxFormulaErrorValue) return v1;
				if (!ExtToDouble(v1, out Guess)) return TFlxFormulaErrorValue.ErrValue;
			}
			if (Guess <= -1) return TFlxFormulaErrorValue.ErrValue;

			double[] Values;
			object ret = GetValues(FTokenList, wi, out Values, CalcState, CalcStack);
			if (ret != null) return ret;
			if (Values == null || Values.Length < 2) return TFlxFormulaErrorValue.ErrNum;

			return CalcIRR(Guess, Values, CalcState, CalcStack);
		}

		private static object GetValues(TParsedTokenList FTokenList, TWorkbookInfo wi, out double[] ResultValue, TCalcState CalcState, TCalcStack CalcStack)
		{
			ResultValue = null;
			TAddressList CellRefs = null;
			object[,] values = null;
			object ret = GetRangeOrArray(FTokenList, wi, CalcState, CalcStack, out CellRefs, out values, TFlxFormulaErrorValue.ErrValue, true);
			if (ret != null) return ret;

			if (values != null)
			{
				double[] DoubleValues;
				object ret2 = ExtractDoubles(values, out DoubleValues);
				if (ret2 != null) return ret2;
				ResultValue = DoubleValues;
				return null;
			}
			if (CellRefs.Has3dRef()) return TFlxFormulaErrorValue.ErrValue;

		{
			double[] DoubleValues;
			object ret2 = ExtractDoubles(wi.Xls, CellRefs, out DoubleValues, CalcState, CalcStack);
			if (ret2 != null) return ret2;

			ResultValue = DoubleValues;
			return null;
		}
		}

		internal static object CalcIRR(double Guess, double[] Values, TCalcState CalcState, TCalcStack CalcStack)
		{
			double Max = CalcMaxAbsValue(Values);

			if (Values.Length <= 1) return TFlxFormulaErrorValue.ErrNum;

			double MaxEps = Max * 1e-9;
			double Guess1 = Guess;
			double Pv1 = CalcOptPv2(Values, Guess1);

			double Guess2 = Pv1 > 0 ? Guess1 + 1e-5 : Guess1 - 1e-5;
			if (Guess2 <= -1) return TFlxFormulaErrorValue.ErrNum;
			double Pv2 = CalcOptPv2(Values, Guess2);

			for (int i = 0; i < 150; i++)
			{
				if (Pv2 == Pv1) //Try moving the guess a little more.
				{
					if (Guess2 > Guess1)
					{
						Guess1 -= 1E-05;
					}
					else
					{
						Guess1 += 1E-05;
					}

					Pv1 = CalcOptPv2(Values, Guess1);
					if (Pv2 == Pv1) //Still the same.
					{
						return TFlxFormulaErrorValue.ErrDiv0;
					}
				}

				Guess1 = Guess2 - (((Guess2 - Guess1) * Pv2) / (Pv2 - Pv1)); //derivate

				if (Guess1 <= -1) //Bound the guess
				{
					Guess1 = (Guess2 - 1) * 0.5;
				}

				Pv1 = CalcOptPv2(Values, Guess1);
				double dGuess = Math.Abs(Guess2 - Guess1);

				if ((Math.Abs(Pv1) < MaxEps) && (dGuess < 1E-07))
				{
					return Guess1;
				}

				double Tmp = Pv1; Pv1 = Pv2; Pv2 = Tmp;  //Swap Pv
				Tmp = Guess1; Guess1 = Guess2; Guess2 = Tmp;  //Swap guess.
			}

			//Not found.
			return TFlxFormulaErrorValue.ErrNum;
		}

		private static double CalcMaxAbsValue(double[] Values)
		{
			double Result = 0;
			for (int i = 0; i < Values.Length; i++)
			{
				if (Math.Abs(Values[i]) > Result) Result = Math.Abs(Values[i]);
			}
			return Result;
		}

		private static double CalcOptPv2(double[] Values, double Guess)
		{
			int i = 0;
			while (i < Values.Length && Values[i] == 0) i++;

			double Result = 0;
			double GuessPlus1 = 1 + Guess;

			for (int j = Values.Length - 1; j >= i; j--)
			{
				Result /= GuessPlus1;
				Result += Values[j];
			}
			return Result;
		}


		private static object ExtractDoubles(object[,] a1, out double[] ResultValue)
		{
			ResultValue = null;
			List<double> ResultA = new List<double>();
			for (int i = 0; i < a1.GetLength(0); i++)
			{
				for (int k = 0; k < a1.GetLength(1); k++)
				{
					if (a1[i, k] is TFlxFormulaErrorValue) return a1[i, k];
					if (a1[i, k] is double)  //no exttodouble here.
					{
						ResultA.Add((double)a1[i, k]);
					}
				}
			}
			ResultValue = ResultA.ToArray();
			return null;
		}


		private static object ExtractDoubles(ExcelFile Xls, TAddressList Addr, out double[] ResultList, TCalcState CalcState, TCalcStack CalcStack)
		{
			ResultList = null;
			TDoubleList ResultA = new TDoubleList();
			for (int i = Addr.Count - 1; i >= 0; i--) //Reversed order
			{
				object res = ProcessAddr(Xls, Addr[i], ResultA, CalcState, CalcStack);
				if (res != null) return res;
			}

			ResultList = ResultA.ToArray();
			return null;
		}

		private static object ProcessAddr(ExcelFile Xls, TAddress[] Addr, TDoubleList ResultList, TCalcState CalcState, TCalcStack CalcStack)
		{
			if (Addr.Length < 1 || Addr.Length > 2) return TFlxFormulaErrorValue.ErrValue;
			TAddress a1 = Addr[0];
			TAddress a2 = a1;
			if (Addr.Length > 1)
				a2 = Addr[1];

			if (a1 == null || a2 == null) return TFlxFormulaErrorValue.ErrNA;

			if (a1.wi.Xls != a2.wi.Xls || a1.Sheet != a2.Sheet)
				return TFlxFormulaErrorValue.ErrRef;

			int MinRow = Math.Min(a1.Row, a2.Row);
			int MinCol = Math.Min(a1.Col, a2.Col);
			int MaxRow = Math.Max(a1.Row, a2.Row);
			int MaxCol = Math.Max(a1.Col, a2.Col);

			for (int r = MinRow; r <= MaxRow; r++)
			{
				for (int c = MinCol; c <= MaxCol; c++)
				{
					object o = ConvertToAllowedObject(a1.wi.Xls.GetCellValueAndRecalc(a1.Sheet, r, c, CalcState, CalcStack));
					if (o is TFlxFormulaErrorValue) return o;

					if (o is double)   //no exttodouble here.
					{
						ResultList.Add((double)o);
					}
				}
			}
			return null;
		}
	}

	internal sealed class TMIRRToken : TPVBaseToken
	{
		internal TMIRRToken(int ArgCount, ptg aId, TCellFunctionData aFuncData)
			: base(ArgCount, aId, aFuncData)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			double ReinvestRate;
			object v2 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v2 is TFlxFormulaErrorValue) return v2;
			if (!ExtToDouble(v2, out ReinvestRate)) return TFlxFormulaErrorValue.ErrValue;

			double FinanceRate;
			object v1 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v1 is TFlxFormulaErrorValue) return v1;
			if (!ExtToDouble(v1, out FinanceRate)) return TFlxFormulaErrorValue.ErrValue;

			object[] Values = new object[1];
			Values[0] = TNPVToken.GetValues(FTokenList, wi, CalcState, CalcStack);
			if (Values[0] is TFlxFormulaErrorValue) return Values[0];


			return CalcMIRR(wi.Xls, FinanceRate, ReinvestRate, Values, CalcState, CalcStack);
		}

		internal static object CalcMIRR(ExcelFile Xls, double FinanceRate, double ReinvestRate, object[] Values, TCalcState CalcState, TCalcStack CalcStack)
		{
			long n;
			object oFNvp = TNPVToken.CalcNPV(Xls, FinanceRate, Values, CalcState, CalcStack, TValueSign.Negative, out n);
			if (oFNvp is TFlxFormulaErrorValue) return oFNvp;
			if (!(oFNvp is Double)) return TFlxFormulaErrorValue.ErrNum;
			double FNVp = (double)oFNvp;

			long n2;
			object oRNvp = TNPVToken.CalcNPV(Xls, ReinvestRate, Values, CalcState, CalcStack, TValueSign.Positive, out n2);
			if (n2 == 0 || n == 0) return TFlxFormulaErrorValue.ErrDiv0;
			if (n2 != n) return TFlxFormulaErrorValue.ErrNA;
			if (oRNvp is TFlxFormulaErrorValue) return oRNvp;
			if (!(oRNvp is Double)) return TFlxFormulaErrorValue.ErrNum;
			double RNVp = (double)oRNvp;

			double num = -RNVp * Math.Pow(1 + ReinvestRate, n);
			double den = FNVp * (1 + FinanceRate);
			if (den == 0) return TFlxFormulaErrorValue.ErrDiv0;

			return Math.Pow(num / den, 1.0 / (n - 1.0)) - 1.0;


		}

	}

	internal sealed class TNPerToken : TPVBaseToken
	{
		internal TNPerToken(int ArgCount, ptg aId, TCellFunctionData aFuncData) : base(ArgCount, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			double Rate; double Per; double Pmt; double Pv; double Fv; int pType;
			object ret = GetArguments(FTokenList, wi, CalcState, CalcStack, false, out Rate, out Per, out Pmt, out Pv, out Fv, out pType);
			if (ret != null) return ret;
			if (Rate <= -1) return TFlxFormulaErrorValue.ErrNum;

			return CalcNPer(Rate, Pmt, Pv, Fv, pType);
		}

		internal static object CalcNPer(double Rate, double Pmt, double Pv, double Fv, int pType)
		{
			if (Rate == 0)
			{
				if (Pmt == 0) return TFlxFormulaErrorValue.ErrDiv0;
				return (-Pv - Fv) / Pmt;
			}

			double TmpPer = pType != 0 ? (Pmt * (Rate + 1)) / Rate : Pmt / Rate;

			double TotFv = -Fv + TmpPer;
			double TotPv = Pv + TmpPer;
			if ((TotFv < 0) && (TotPv < 0))
			{
				TotFv = -TotFv;
				TotPv = -TotPv;
			}
			else if ((TotPv <= 0) || (TotFv <= 0)) return TFlxFormulaErrorValue.ErrNum;

			return (Math.Log(TotFv) - Math.Log(TotPv)) / Math.Log(1 + Rate);
		}
	}

	internal sealed class TPMTToken : TPVBaseToken
	{
		internal TPMTToken(int ArgCount, ptg aId, TCellFunctionData aFuncData) : base(ArgCount, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			double Rate; double Per; double nPer; double Pv; double Fv; int pType;
			object ret = GetArguments(FTokenList, wi, CalcState, CalcStack, false, out Rate, out Per, out nPer, out Pv, out Fv, out pType);
			if (ret != null) return ret;
			if (nPer == 0) return TFlxFormulaErrorValue.ErrDiv0;

			return CalcPMT(Rate, nPer, Pv, Fv, pType);
		}

		internal static object CalcPMT(double Rate, double nPer, double Pv, double Fv, int pType)
		{
			if (Rate == 0) return -(Fv + Pv) / nPer;

			double FracPer = 1;
			if (pType != 0) FracPer += Rate;
			double tv = Math.Pow(1 + Rate, nPer);
			return -(Fv + tv * Pv) / FracPer / (tv - 1) * Rate;
		}
	}

	internal sealed class TIPMTToken : TPVBaseToken
	{
		internal TIPMTToken(int ArgCount, ptg aId, TCellFunctionData aFuncData) : base(ArgCount, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			double Rate; double Per; double nPer; double Pv; double Fv; int pType;
			object ret = GetArguments(FTokenList, wi, CalcState, CalcStack, true, out Rate, out Per, out nPer, out Pv, out Fv, out pType);
			if (ret != null) return ret;

			return CalcIPMT(Rate, Per, nPer, Pv, Fv, pType);
		}

		internal static object CalcIPMT(double Rate, double Per, double nPer, double Pv, double Fv, int pType)
		{
			if (pType != 0 && Per == 1) return 0;

			object oPmt = TPMTToken.CalcPMT(Rate, nPer, Pv, Fv, pType);
			if (!(oPmt is double)) return oPmt;
			double Pmt = (double)oPmt;

			if (pType != 0) Pv += Pmt;

			double PerRemain = pType == 0 ? 1 : 2;
			object NewFv = TFVToken.CalcFV(Rate, Per - PerRemain, Pmt, Pv, 0);
			if (!(NewFv is double)) return NewFv;
			return Rate * (double)NewFv;
		}
	}

	internal sealed class TPPMTToken : TPVBaseToken
	{
		internal TPPMTToken(int ArgCount, ptg aId, TCellFunctionData aFuncData) : base(ArgCount, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			double Rate; double Per; double nPer; double Pv; double Fv; int pType;
			object ret = GetArguments(FTokenList, wi, CalcState, CalcStack, true, out Rate, out Per, out nPer, out Pv, out Fv, out pType);
			if (ret != null) return ret;

			return CalcPPMT(Rate, Per, nPer, Pv, Fv, pType);
		}

		private static object CalcPPMT(double Rate, double Per, double nPer, double Pv, double Fv, int pType)
		{
			object oPmt = ConvertToAllowedObject(TPMTToken.CalcPMT(Rate, nPer, Pv, Fv, pType));
			if (!(oPmt is double)) return oPmt;
			double Pmt = (double)oPmt;

			object oIPmt = ConvertToAllowedObject(TIPMTToken.CalcIPMT(Rate, Per, nPer, Pv, Fv, pType));
			if (!(oIPmt is double)) return oIPmt;
			double IPmt = (double)oIPmt;

			return Pmt - IPmt;
		}
	}

	internal sealed class TRateToken : TPVBaseToken
	{
		internal TRateToken(int ArgCount, ptg aId, TCellFunctionData aFuncData) : base(ArgCount, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			double Guess = 0.1;
			if (FArgCount >= 6)
			{
				object v6 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v6 is TFlxFormulaErrorValue) return v6;
				if (!GetDouble(v6, out Guess)) return TFlxFormulaErrorValue.ErrValue;
			}

			int pType = 0;
			if (FArgCount >= 5)
			{
				double dType;
				object v5 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v5 is TFlxFormulaErrorValue) return v5;
				if (!GetDouble(v5, out dType)) return TFlxFormulaErrorValue.ErrValue;
				if (dType != 0) pType = 1;
			}

			double Fv = 0;
			if (FArgCount >= 4)
			{
				object v4 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v4 is TFlxFormulaErrorValue) return v4;
				if (!ExtToDouble(v4, out Fv)) return TFlxFormulaErrorValue.ErrValue;
			}

			double Pv;
			object v3 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v3 is TFlxFormulaErrorValue) return v3;
			if (!ExtToDouble(v3, out Pv)) return TFlxFormulaErrorValue.ErrValue;

			double Pmt;
			object v1 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v1 is TFlxFormulaErrorValue) return v1;
			if (!ExtToDouble(v1, out Pmt)) return TFlxFormulaErrorValue.ErrValue;

			double nPer;
			object v2 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v2 is TFlxFormulaErrorValue) return v2;
			if (!ExtToDouble(v2, out nPer)) return TFlxFormulaErrorValue.ErrValue;
			if (nPer <= 0) return TFlxFormulaErrorValue.ErrNum;

			return CalcRate(nPer, Pmt, Pv, Fv, pType == 0, Guess);
		}

		internal static object CalcRate(double nPer, double Pmt, double Pv, double Fv, bool EndOfPeriod, double Guess)
		{
			double Guess1 = Guess;
			double Rt1 = EvalRate(Guess1, nPer, Pmt, Pv, Fv, EndOfPeriod);

			double Guess2 = Rt1 > 0 ? Guess1 / 2 : Guess1 * 2;
			double Rt2 = EvalRate(Guess2, nPer, Pmt, Pv, Fv, EndOfPeriod);

			for (int i = 0; i < 150; i++)
			{
				if (Rt2 == Rt1) //Try moving the guess a little more.
				{
					if (Guess2 > Guess1)
					{
						Guess1 -= 1E-05;
					}
					else
					{
						Guess1 += 1E-05;
					}

					Rt1 = EvalRate(Guess1, nPer, Pmt, Pv, Fv, EndOfPeriod);
					if (Rt2 == Rt1) //Still the same.
					{
						return TFlxFormulaErrorValue.ErrDiv0;
					}
				}

				Guess1 = Guess2 - (((Guess2 - Guess1) * Rt2) / (Rt2 - Rt1)); //derivate
				Rt1 = EvalRate(Guess1, nPer, Pmt, Pv, Fv, EndOfPeriod);

				if ((Math.Abs(Rt1) < 1E-07))
				{
					return Guess1;
				}

				double Tmp = Rt1; Rt1 = Rt2; Rt2 = Tmp;  //Swap Pv
				Tmp = Guess1; Guess1 = Guess2; Guess2 = Tmp;  //Swap guess.
			}

			//Not found.
			return TFlxFormulaErrorValue.ErrNum;
		}

		private static double EvalRate(double Rate, double nPer, double Pmt, double Pv, double Fv, bool EndOfPeriod)
		{
			if (Rate == 0)
			{
				return ((Pv + (Pmt * nPer)) + Fv);
			}
			double Rp = Math.Pow(Rate + 1, nPer);

			double Rr = EndOfPeriod ? 1 : 1 + Rate;
			return (((Pv * Rp) + (((Pmt * Rr) * (Rp - 1)) / Rate)) + Fv);
		}
	}


	#endregion

	#region Database
	#region Criteria classes
	internal abstract class TCriteria
	{
		internal abstract bool Evaluate(ExcelFile Xls, TCalcState CalcState, TCalcStack CalcStack, int Sheet, int Row);
	}

	#region Value criteria
	internal abstract class TValueCriteria : TCriteria
	{
		TCriteriaType ct;
		int Column;

		internal TValueCriteria(TCriteriaType act, int aColumn)
		{
			ct = act;
			Column = aColumn;
		}

		internal static TValueCriteria Create(object Criteria, int aColumn, bool AcceptWildcards)
		{
			//See if it begins on "<", ">", etc, and modify the criteria if it does.
			TCriteriaType tempct = TCriteriaType.EQ;
			string StrCriteria = Criteria as String;
			if (StrCriteria != null)
			{
				CalcCriteria(ref StrCriteria, out tempct);
				Criteria = StrCriteria;
			}

			return FindCorrectType(tempct, aColumn, Criteria, AcceptWildcards);
		}

		internal override bool Evaluate(ExcelFile Xls, TCalcState CalcState, TCalcStack CalcStack, int Sheet, int Row)
		{
			object o = TExcelTypes.ConvertToAllowedObject(Xls.GetCellValueAndRecalc(Sheet, Row, Column, CalcState, CalcStack), Xls.OptionsDates1904);
			return MeetsCriteria(o);
		}

		internal static TValueCriteria FindCorrectType(TCriteriaType tempct, int aColumn, object Criteria, bool AcceptWildcards)
		{
			if (Criteria == null) return new TNullCriteria(tempct, aColumn);
			string sCriteria = FlxConvert.ToString(Criteria);
			double dCriteria = 0;

			if (Criteria is Double)
			{
				return new TDoubleCriteria(tempct, aColumn, (double)Criteria);
			}
			else
                if (TCompactFramework.ConvertToNumber(sCriteria.Trim(), CultureInfo.CurrentCulture, out dCriteria))
			{
				return new TDoubleCriteria(tempct, aColumn, dCriteria);
			}

            if (Criteria is bool)
            {
                return new TBoolCriteria(tempct, aColumn, (bool)Criteria);
            }
            else
            {
                    if (String.Equals(sCriteria, TFormulaMessages.TokenString(TFormulaToken.fmTrue), StringComparison.CurrentCultureIgnoreCase))
                    {
                        return new TBoolCriteria(tempct, aColumn, true);
                    }
                    else
                        if (String.Equals(sCriteria, TFormulaMessages.TokenString(TFormulaToken.fmFalse), StringComparison.CurrentCultureIgnoreCase))
                        {
                            return new TBoolCriteria(tempct, aColumn, false);
                        }
            }

			if (AcceptWildcards)
			{
				return new TStringWithWildcardsCriteria(tempct, aColumn, sCriteria);
			}
			else
			{
				return new TStringCriteria(tempct, aColumn, sCriteria);
			}
		}


		internal static void CalcCriteria(ref string Criteria, out TCriteriaType ct)
		{
			ct = TCriteriaType.EQ;
			string s = (string)Criteria;
			if (s == null || s.Length <= 0) return;

			if (s[0] == TFormulaMessages.TokenChar(TFormulaToken.fmEQ))
			{
				ct = TCriteriaType.EQ;
				Criteria = s.Substring(1);
				return;
			}
			if (s[0] == TFormulaMessages.TokenChar(TFormulaToken.fmLT))
			{
				if (s.Length > 1)
				{
					if (s[1] == TFormulaMessages.TokenChar(TFormulaToken.fmGT))
					{
						ct = TCriteriaType.NE;
						Criteria = s.Substring(2);
						return;
					}
					if (s[1] == TFormulaMessages.TokenChar(TFormulaToken.fmEQ))
					{
						ct = TCriteriaType.LE;
						Criteria = s.Substring(2);
						return;
					}
				}
				ct = TCriteriaType.LT;
				Criteria = s.Substring(1);
				return;
			}

			if (s[0] == TFormulaMessages.TokenChar(TFormulaToken.fmGT))
			{
				if (s.Length > 1)
				{
					if (s[1] == TFormulaMessages.TokenChar(TFormulaToken.fmEQ))
					{
						ct = TCriteriaType.GE;
						Criteria = s.Substring(2);
						return;
					}
				}
				ct = TCriteriaType.GT;
				Criteria = s.Substring(1);
				return;
			}
		}


		internal abstract bool Compare(object v1, ref int ResultValue);
		internal abstract bool CompareEqual(object v1);
		internal virtual bool CompareNotEqual(object v1)
		{
			int Cmp = 0;
			return !Compare(v1, ref Cmp) || Cmp != 0;
		}

		internal bool MeetsCriteria(object v1)
		{
			int Cmp = 0;

			switch (ct)
			{
					//!!!!!!!!!!!!!!!! Equal behaves different than the rest. < !!!!!!!!!!!!!!!!!!!!!!!
					//On Equal, sumif(xx,"=3") will match a string "3"; sumif(xx:"<3") will not use the string.
					// Also, sumif(xx,"<>3") will consider the string to be <>3.
				case TCriteriaType.EQ:
					return CompareEqual(v1);

				case TCriteriaType.NE:
					return CompareNotEqual(v1);

				case TCriteriaType.GT:
					if (Compare(v1, ref Cmp)
						&& Cmp < 0)
					{
						return true;
					}
					break;
				case TCriteriaType.GE:
					if (Compare(v1, ref Cmp)
						&& Cmp <= 0)
					{
						return true;
					}
					break;
				case TCriteriaType.LT:
					if (Compare(v1, ref Cmp)
						&& Cmp > 0)
					{
						return true;
					}
					break;
				case TCriteriaType.LE:
					if (Compare(v1, ref Cmp)
						&& Cmp >= 0)
					{
						return true;
					}
					break;
			}
			return false;
		}
	}

	internal class TBoolCriteria : TValueCriteria
	{
		bool Value;

		internal TBoolCriteria(TCriteriaType acf, int aColumn, bool aValue)
			: base(acf, aColumn)
		{
			Value = aValue;
		}

		internal override bool Compare(object v1, ref int ResultValue)
		{
			if (v1 is bool)
			{
				ResultValue = Value.CompareTo(v1);
				return true;
			}
			return false;
		}

		internal override bool CompareEqual(object v1)
		{
			if (v1 is bool)
			{
				return Value.CompareTo(v1) == 0;
			}
			return false;
		}
	}

	internal class TNullCriteria : TValueCriteria
	{
		internal TNullCriteria(TCriteriaType acf, int aColumn)
			: base(acf, aColumn)
		{
		}

		internal override bool Compare(object v1, ref int ResultValue)
		{
			return false;
		}

		internal override bool CompareEqual(object v1)
		{
			return v1 is double && (double)v1 == 0;
		}
	}

	internal class TDoubleCriteria : TValueCriteria
	{
		double Value;

		internal TDoubleCriteria(TCriteriaType acf, int aColumn, double aValue)
			: base(acf, aColumn)
		{
			Value = aValue;
		}

		internal override bool Compare(object v1, ref int ResultValue)
		{
			if (v1 is Double)
			{
				ResultValue = Value.CompareTo(v1);
				return true;
			}
			return false;
		}

		internal override bool CompareEqual(object v1)
		{
			double temp = 0;
			if (v1 is Double)
			{
				if ((double)v1 == Value)
				{
					return true;
				}
			}
			else
			{
                if (TCompactFramework.ConvertToNumber(FlxConvert.ToString(v1), CultureInfo.CurrentCulture, out temp))
					if (temp == Value)
					{
						return true;
					}
			}

			return false;
		}
	}

	internal class TStringCriteria : TValueCriteria
	{
		string Value;

		internal TStringCriteria(TCriteriaType acf, int aColumn, string aValue)
			: base(acf, aColumn)
		{
			Value = aValue;
		}

		internal override bool Compare(object v1, ref int ResultValue)
		{
			string sv1 = v1 as string;
			if (sv1 != null)
			{
				ResultValue = String.Compare(Value, sv1, StringComparison.CurrentCultureIgnoreCase);
				return true;
			}
			return false;
		}

		internal override bool CompareEqual(object v1)
		{
			if (String.Equals(FlxConvert.ToString(v1), Value, StringComparison.CurrentCultureIgnoreCase))
			{
				return true;
			}

			return false;
		}
	}

	internal class TStringWithWildcardsCriteria : TValueCriteria
	{
		string Value;

		internal TStringWithWildcardsCriteria(TCriteriaType acf, int aColumn, string aValue)
			: base(acf, aColumn)
		{
			Value = aValue;
		}

		internal override bool Compare(object v1, ref int ResultValue)
		{
			string sv1 = v1 as string;
			if (sv1 != null)
			{
				ResultValue = String.Compare(Value, sv1, StringComparison.CurrentCultureIgnoreCase);
				return true;
			}
			return false;
		}

		internal override bool CompareEqual(object v1)
		{
			object res = TBaseParsedToken.CompareWithWildcards(FlxConvert.ToString(v1), Value, true);
			return res is int && (int)res == 0;
		}

		internal override bool CompareNotEqual(object v1)
		{
			return !CompareEqual(v1);
		}

	}
	#endregion

	internal class TFalseCriteria : TCriteria
	{
		internal override bool Evaluate(ExcelFile Xls, TCalcState CalcState, TCalcStack CalcStack, int Sheet, int Row)
		{
			return false;
		}
	}

	internal sealed class TFormulaCriteria : TCriteria
	{
		TParsedTokenList Formula;  
		TWorkbookInfo wi;
		int Row0;

		internal TFormulaCriteria(TParsedTokenList FormulaData, ExcelFile aXls, int SheetIndex, TRowAndCols RowsAndCol, int DbRow)
		{
			Formula = FormulaData;
			wi = new TWorkbookInfo(aXls, SheetIndex, RowsAndCol.Row, RowsAndCol.Col, RowsAndCol.RowCount, RowsAndCol.ColCount, 0, 0, false);
			Row0 = DbRow + 1;
		}

		internal override bool Evaluate(ExcelFile Xls, TCalcState CalcState, TCalcStack CalcStack, int Sheet, int Row)
		{
			wi.RowOfs = Row - Row0;
			object ok = Formula.EvaluateAll(wi, CalcState, CalcStack);
			bool Result;
			if (!TBaseParsedToken.ExtToBool(ok, out Result)) return false;
			return Result;
		}
	}
	#endregion

	#region State Classes
	internal class TDbState
	{
		internal TCriteria[][] Criteria;
		internal int Count;
		internal double Sum;
		internal bool First = true;
		internal object Value;
	}

	internal class TDbStDevState : TDbState
	{
		internal double Avg;
		internal double AvgSum;
		internal bool CalculatingAverage = true;
	}

	#endregion

	internal abstract class TDatabaseBaseToken : TBaseFunctionToken
	{
		protected TDatabaseBaseToken(ptg aId, TCellFunctionData aFuncData)
			: base(3, aId, aFuncData)
		{
		}

		private static object ProcessArray(ref double[,] ResultValue, object[,] arr)
		{
			if (ResultValue == null)
			{
				ResultValue = new double[arr.GetLength(0), arr.GetLength(1)];
				for (int i = 0; i < ResultValue.GetLength(0); i++)
					for (int k = 0; k < ResultValue.GetLength(1); k++)
					{
						object o = ConvertToAllowedObject(arr[i, k]);
						if (o is TFlxFormulaErrorValue) return o;
						if (o is double)
							ResultValue[i, k] = (double)o;
					}
			}
			else
			{
				if (ResultValue.GetLength(0) != arr.GetLength(0) || ResultValue.GetLength(1) != arr.GetLength(1))
					return TFlxFormulaErrorValue.ErrValue;

				for (int i = 0; i < ResultValue.GetLength(0); i++)
					for (int k = 0; k < ResultValue.GetLength(1); k++)
					{
						object o = ConvertToAllowedObject(arr[i, k]);
						if (o is TFlxFormulaErrorValue) return o;
						if (o is double)
							ResultValue[i, k] *= (double)o;
						else
							ResultValue[i, k] = 0;
					}
			}

			return null;
		}

		private static object ProcessRef(ExcelFile Xls, ref double[,] ResultValue, TAddressList adr, TCalcState CalcState, TCalcStack CalcStack)
		{
			if (adr.Count != 1) return TFlxFormulaErrorValue.ErrValue;
			if (adr[0].Length < 1 || adr[0].Length > 2) return TFlxFormulaErrorValue.ErrValue;
			TAddress a1 = adr[0][0];
			TAddress a2 = a1;
			if (adr[0].Length > 1)
				a2 = adr[0][1];

			if (a1 == null || a2 == null) return TFlxFormulaErrorValue.ErrNA;

			if (a1.wi.Xls != a2.wi.Xls || a1.Sheet != a2.Sheet)
				return TFlxFormulaErrorValue.ErrRef;

			int MinRow = Math.Min(a1.Row, a2.Row);
			int MinCol = Math.Min(a1.Col, a2.Col);

			if (ResultValue == null)
			{
				ResultValue = new double[Math.Abs(a1.Row - a2.Row) + 1, Math.Abs(a1.Col - a2.Col) + 1];
				for (int i = 0; i < ResultValue.GetLength(0); i++)
					for (int k = 0; k < ResultValue.GetLength(1); k++)
					{
						object o = ConvertToAllowedObject(a1.wi.Xls.GetCellValueAndRecalc(a1.Sheet, i + MinRow, k + MinCol, CalcState, CalcStack));
						if (o is TFlxFormulaErrorValue) return o;
						if (o is double)
							ResultValue[i, k] = (double)o; //here strings and booleans are 0.

					}
			}
			else
			{
				if (ResultValue.GetLength(0) != Math.Abs(a1.Row - a2.Row) + 1 || ResultValue.GetLength(1) != Math.Abs(a1.Col - a2.Col) + 1)
					return TFlxFormulaErrorValue.ErrValue;

				for (int i = 0; i < ResultValue.GetLength(0); i++)
					for (int k = 0; k < ResultValue.GetLength(1); k++)
					{
						object o = ConvertToAllowedObject(a1.wi.Xls.GetCellValueAndRecalc(a1.Sheet, i + MinRow, k + MinCol, CalcState, CalcStack));
						if (o is TFlxFormulaErrorValue) return o;
						if (o is double)
							ResultValue[i, k] *= (double)o;
						else
							ResultValue[i, k] = 0;
					}
			}

			return null;
		}

		internal object LoopRecords(TParsedTokenList FTokenList, TDbState State, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack)
		{
			object CriteriaRef = FTokenList.EvaluateTokenRef(wi, CalcState, CalcStack);
			if (CriteriaRef is TFlxFormulaErrorValue) return TFlxFormulaErrorValue.ErrValue;
			TAddress CriteriaCellRef1 = null, CriteriaCellRef2 = null;
			object ret = GetRange(CriteriaRef, out CriteriaCellRef1, out CriteriaCellRef2, TFlxFormulaErrorValue.ErrValue);
			if (ret != null) return ret;
			if (CriteriaCellRef1.wi.Xls != CriteriaCellRef2.wi.Xls || CriteriaCellRef1.Sheet != CriteriaCellRef2.Sheet) return TFlxFormulaErrorValue.ErrValue;
			if (CriteriaCellRef1.Row >= CriteriaCellRef2.Row) return TFlxFormulaErrorValue.ErrValue;

			object oField = FTokenList.EvaluateToken(wi, TArrayAggregate.Instance, CalcState, CalcStack); if (oField is TFlxFormulaErrorValue) return oField;

			object DbRef = FTokenList.EvaluateTokenRef(wi, CalcState, CalcStack);
			if (DbRef is TFlxFormulaErrorValue) return TFlxFormulaErrorValue.ErrValue;
			TAddress DbCellRef1 = null, DbCellRef2 = null;
			ret = GetRange(DbRef, out DbCellRef1, out DbCellRef2, TFlxFormulaErrorValue.ErrValue);
			if (ret != null) return ret;
			if (DbCellRef1.wi.Xls != DbCellRef2.wi.Xls || DbCellRef1.Sheet != DbCellRef2.Sheet) return TFlxFormulaErrorValue.ErrValue;
			if (DbCellRef1.Row >= DbCellRef2.Row) return TFlxFormulaErrorValue.ErrValue;

			string FieldStr = oField as String;
			int FieldIndex = -1;
			if (FieldStr != null)
			{
				for (int c = DbCellRef1.Col; c <= DbCellRef2.Col; c++)
				{
					if (String.Equals(
						FieldStr,
						Convert.ToString(DbCellRef1.wi.Xls.GetCellValueAndRecalc(DbCellRef1.Sheet, DbCellRef1.Row, c, CalcState, CalcStack)),
						StringComparison.CurrentCultureIgnoreCase))
					{
						FieldIndex = c - DbCellRef1.Col + 1;
						break;
					}
				}
			}
			else
			{
				double d; if (!ExtToDouble(oField, out d)) return TFlxFormulaErrorValue.ErrValue;
				FieldIndex = (int)Math.Floor(d);
			}

			if (FieldIndex <= 0 || FieldIndex > DbCellRef2.Col - DbCellRef1.Col + 1) return TFlxFormulaErrorValue.ErrValue;

			TCaseInsensitiveHashtableStrInt HeaderCaptions = ParseHeaders(wi.Xls, DbCellRef1, DbCellRef2, CalcState, CalcStack);
			FillCriteria(wi.Xls, State, CriteriaCellRef1, CriteriaCellRef2, HeaderCaptions, CalcState, CalcStack, DbCellRef1.Row);

			return CalcValue(wi.Xls, State, DbCellRef1, DbCellRef2, DbCellRef1.Col + FieldIndex - 1, CalcState, CalcStack);
		}

		private static TCaseInsensitiveHashtableStrInt ParseHeaders(ExcelFile Xls, TAddress Cell1, TAddress Cell2, TCalcState CalcState, TCalcStack CalcStack)
		{
			TCaseInsensitiveHashtableStrInt Result = new TCaseInsensitiveHashtableStrInt();
			for (int c = Cell1.Col; c <= Cell2.Col; c++)
			{
				object o = ConvertToAllowedObject(Cell1.wi.Xls.GetCellValueAndRecalc(Cell1.Sheet, Cell1.Row, c, CalcState, CalcStack));
				if (o == null) continue;
				string Caption = Convert.ToString(o);
				if (Caption == null || Caption.Length == 0) continue;
				if (Result.ContainsKey(Caption)) continue;
				Result[Caption] = c;
			}
			return Result;
		}

		private static int[] ParseCriteriaHeaders(ExcelFile Xls, TAddress Cell1, TAddress Cell2, TCaseInsensitiveHashtableStrInt HeaderCaptions, TCalcState CalcState, TCalcStack CalcStack)
		{
			int[] Result = new int[Cell2.Col - Cell1.Col + 1];

			for (int c = Cell1.Col; c <= Cell2.Col; c++)
			{
				object o = ConvertToAllowedObject(Cell1.wi.Xls.GetCellValueAndRecalc(Cell1.Sheet, Cell1.Row, c, CalcState, CalcStack));
				string Caption = Convert.ToString(o);

				int Cap;
				if (HeaderCaptions.TryGetValue(Caption, out Cap))
				{
					Result[c - Cell1.Col] = Cap;
				}
				else
				{
					Result[c - Cell1.Col] = -1; //This might be a formula criteria.
				}
			}
			return Result;
		}

		private static void FillCriteria(ExcelFile Xls, TDbState State, TAddress Cell1, TAddress Cell2, TCaseInsensitiveHashtableStrInt HeaderCaptions, TCalcState CalcState, TCalcStack CalcStack, int FirstDbRow)
		{
			int[] HeaderColumns = ParseCriteriaHeaders(Xls, Cell1, Cell2, HeaderCaptions, CalcState, CalcStack);

			State.Criteria = new TCriteria[Cell2.Row - Cell1.Row][];
			int RowCriteriaIndex = 0;

			for (int r = Cell1.Row + 1; r <= Cell2.Row; r++)
			{
				int CriteriaIndex = 0;
				State.Criteria[RowCriteriaIndex] = new TCriteria[Cell2.Col - Cell1.Col + 1];
				for (int c = Cell1.Col; c <= Cell2.Col; c++)
				{
					int DataIndex = HeaderColumns[c - Cell1.Col];
					if (DataIndex < 0) //Formula
					{
                        int XF = -1;
						object o = Cell1.wi.Xls.GetCellValue(Cell1.Sheet, r, c, ref XF); //Here we need the real formula object, even if not calculated.
						if (o != null)  //Empty values mean no condition.
						{
							TFormula Fmla = o as TFormula; //it if is not a formula, since the caption does not exist, it will have to be a "true/false" value.
							if (Fmla != null)
							{
								State.Criteria[RowCriteriaIndex][CriteriaIndex] = new TFormulaCriteria(Fmla.Data, Cell1.wi.Xls, Cell1.Sheet, new TRowAndCols(r, 0, c, 0), FirstDbRow);
								CriteriaIndex++;
							}
							else
							{
								bool Cnd;
								if (!ExtToBool(o, out Cnd) || !Cnd)
								{
									State.Criteria[RowCriteriaIndex][0] = new TFalseCriteria();
									break; //no need for more criterias on this row.
								}
							}
						}
					}
					else
					{
						object o = ConvertToAllowedObject(Cell1.wi.Xls.GetCellValueAndRecalc(Cell1.Sheet, r, c, CalcState, CalcStack));
						//o might be a formulaerrorvalue here.
						if (o == null) continue;
						State.Criteria[RowCriteriaIndex][CriteriaIndex] = TValueCriteria.Create(o, DataIndex, true);
						CriteriaIndex++;
					}
				}
				RowCriteriaIndex++;
			}
		}

		private static bool PassesOneRowCriteria(ExcelFile Xls, TCalcState CalcState, TCalcStack CalcStack, int DataSheet, int DataRow, TCriteria[] CritRow)
		{
			for (int i = 0; i < CritRow.Length; i++)
			{
				TCriteria Cr = CritRow[i];
				if (Cr == null) return true; //At the end.
				if (!Cr.Evaluate(Xls, CalcState, CalcStack, DataSheet, DataRow)) return false;
			}
			return true;
		}

		private static bool PassesCriteria(ExcelFile Xls, TCalcState CalcState, TCalcStack CalcStack, TDbState State, int DataSheet, int DataRow)
		{
			for (int i = 0; i < State.Criteria.Length; i++)
			{
				if (PassesOneRowCriteria(Xls, CalcState, CalcStack, DataSheet, DataRow, State.Criteria[i])) return true;
			}
			return false;
		}

		private object CalcValue(ExcelFile Xls, TDbState State, TAddress Cell1, TAddress Cell2, int AggColumn, TCalcState CalcState, TCalcStack CalcStack)
		{
			for (int r = Cell1.Row + 1; r <= Cell2.Row; r++)
			{
				if (PassesCriteria(Cell1.wi.Xls, CalcState, CalcStack, State, Cell1.Sheet, r))
				{
					object o = ConvertToAllowedObject(Cell1.wi.Xls.GetCellValueAndRecalc(Cell1.Sheet, r, AggColumn, CalcState, CalcStack));
					object res = Aggregate(o, State);
					if (res != null) return res;
				}
			}

			return null;
		}

		protected abstract object Aggregate(object o, TDbState State);
	}

	internal sealed class TDCountToken : TDatabaseBaseToken
	{
		internal TDCountToken(ptg aId, TCellFunctionData aFuncData) : base(aId, aFuncData) {}

		protected override object Aggregate(object o, TDbState State)
		{
			if (o is double) State.Count++;
			return null;
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			TDbState State = new TDbState();
			object Res = LoopRecords(FTokenList, State, wi, CalcState, CalcStack);
			if (Res != null) return Res;
			return State.Count;
		}
	}

	internal sealed class TDCountAToken : TDatabaseBaseToken
	{
		internal TDCountAToken(ptg aId, TCellFunctionData aFuncData)
			: base(aId, aFuncData)
		{
		}

		protected override object Aggregate(object o, TDbState State)
		{
			if (o != null) State.Count++;
			return null;
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			TDbState State = new TDbState();
			object Res = LoopRecords(FTokenList, State, wi, CalcState, CalcStack);
			if (Res != null) return Res;
			return State.Count;
		}
	}

	internal sealed class TDSumToken : TDatabaseBaseToken
	{
		internal TDSumToken(ptg aId, TCellFunctionData aFuncData)
			: base(aId, aFuncData)
		{
		}

		protected override object Aggregate(object o, TDbState State)
		{
			if (o is TFlxFormulaErrorValue) return o;
			if (o is double) State.Sum += (double)o;
			return null;
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			TDbState State = new TDbState();
			object Res = LoopRecords(FTokenList, State, wi, CalcState, CalcStack);
			if (Res != null) return Res;
			return State.Sum;
		}

	}

	internal sealed class TDMinToken : TDatabaseBaseToken
	{
		internal TDMinToken(ptg aId, TCellFunctionData aFuncData)
			: base(aId, aFuncData)
		{
		}

		protected override object Aggregate(object o, TDbState State)
		{
			if (o is TFlxFormulaErrorValue) return o;
			if (o is double)
			{
				if (State.First || (double)o < State.Sum)
				{
					State.Sum = (double)o;
					State.First = false;
				}
			}
			return null;
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			TDbState State = new TDbState();
			object Res = LoopRecords(FTokenList, State, wi, CalcState, CalcStack);
			if (Res != null) return Res;
			return State.Sum;
		}

	}

	internal sealed class TDMaxToken : TDatabaseBaseToken
	{
		internal TDMaxToken(ptg aId, TCellFunctionData aFuncData)
			: base(aId, aFuncData)
		{
		}

		protected override object Aggregate(object o, TDbState State)
		{
			if (o is TFlxFormulaErrorValue) return o;
			if (o is double)
			{
				if (State.First || (double)o > State.Sum)
				{
					State.Sum = (double)o;
					State.First = false;
				}
			}
			return null;
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			TDbState State = new TDbState();
			object Res = LoopRecords(FTokenList, State, wi, CalcState, CalcStack);
			if (Res != null) return Res;
			return State.Sum;
		}
	}

	internal sealed class TDAverageToken : TDatabaseBaseToken
	{
		internal TDAverageToken(ptg aId, TCellFunctionData aFuncData)
			: base(aId, aFuncData)
		{
		}

		protected override object Aggregate(object o, TDbState State)
		{
			if (o is TFlxFormulaErrorValue) return o;
			if (o is double)
			{
				State.Count++;
				State.Sum += (double)o;
			}
			return null;
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			TDbState State = new TDbState();
			object Res = LoopRecords(FTokenList, State, wi, CalcState, CalcStack);
			if (Res != null) return Res;
			if (State.Count == 0) return TFlxFormulaErrorValue.ErrDiv0;
			return State.Sum / State.Count;
		}

	}

	internal sealed class TDProductToken : TDatabaseBaseToken
	{
		internal TDProductToken(ptg aId, TCellFunctionData aFuncData)
			: base(aId, aFuncData)
		{
		}

		protected override object Aggregate(object o, TDbState State)
		{
			if (o is TFlxFormulaErrorValue) return o;
			if (o is double)
			{
				State.Sum *= (double)o;
				State.First = false;
			}
			return null;
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			TDbState State = new TDbState();
			State.Sum = 1;
			object Res = LoopRecords(FTokenList, State, wi, CalcState, CalcStack);
			if (Res != null) return Res;
			if (State.First) return 0;
			return State.Sum;
		}

	}

	internal sealed class TDGetToken : TDatabaseBaseToken
	{
		internal TDGetToken(ptg aId, TCellFunctionData aFuncData)
			: base(aId, aFuncData)
		{
		}

		protected override object Aggregate(object o, TDbState State)
		{
			if (State.Count > 0) return TFlxFormulaErrorValue.ErrNum;
			State.Count++;
			State.Value = o;
			return null;
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			TDbState State = new TDbState();
			object Res = LoopRecords(FTokenList, State, wi, CalcState, CalcStack);
			if (Res != null) return Res;
			if (State.Count != 1) return TFlxFormulaErrorValue.ErrValue;
			return State.Value;
		}

	}

	internal class TDVarStDevToken : TDatabaseBaseToken
	{
		int P; //This variable is inmutable, so we can have it here.
		protected TDVarStDevToken(ptg aId, TCellFunctionData aFuncData, int aP)
			: base(aId, aFuncData)
		{
			P = aP;
		}


		protected override object Aggregate(object o, TDbState State)
		{
			if (o is TFlxFormulaErrorValue) return o;
            
			TDbStDevState s = State as TDbStDevState;
			if (o is double)
			{
				if (s.CalculatingAverage)
				{
					s.Count++;
					s.Sum += (double)o;
				}
				else
				{
					double x = (double)o - s.Avg;
					s.AvgSum += x * x;
				}

			}
			return null;
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			TDbStDevState State = new TDbStDevState();

			int SavePosition = FTokenList.SavePosition();
			object Res = LoopRecords(FTokenList, State, wi, CalcState, CalcStack);
			if (Res != null) return Res;

			if (State.Count - P <= 0) return TFlxFormulaErrorValue.ErrDiv0;

			State.Avg = State.Sum / State.Count;

			State.AvgSum = 0;
			State.CalculatingAverage = false;
			FTokenList.RestorePosition(SavePosition);
			Res = LoopRecords(FTokenList, State, wi, CalcState, CalcStack);
			if (Res != null) return Res;

			return State.AvgSum / (State.Count - P);
		}
	}

	internal sealed class TDVarToken : TDVarStDevToken
	{
		internal TDVarToken(ptg aId, TCellFunctionData aFuncData, int aP)
			: base(aId, aFuncData, aP)
		{
		}
	}

	internal sealed class TDStDevToken : TDVarStDevToken
	{
		internal TDStDevToken(ptg aId, TCellFunctionData aFuncData, int aP)
			: base(aId, aFuncData, aP)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			object Result = base.Evaluate(FTokenList, wi, f, CalcState, CalcStack);
			if (Result is double)
			{
				return Math.Sqrt((double)Result);
			}
			return Result;
		}

	}

	#endregion

	#region Misc
	internal sealed class THyperlinkToken : TBaseFunctionToken
	{
		internal THyperlinkToken(int ArgCount, ptg aId, TCellFunctionData aFuncData) : base(ArgCount, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			object FriendlyName = null;
			if (FArgCount > 1)
			{
				object v2 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v2 is TFlxFormulaErrorValue) return v2;
				FriendlyName = v2;
			}

			object v1 = FTokenList.EvaluateToken(wi, CalcState, CalcStack); if (v1 is TFlxFormulaErrorValue) return v1;

			if (FArgCount > 1) return FriendlyName;
			return v1;
		}
	}

	#endregion

	#region User Defined

	internal sealed class TUserDefinedToken : TBaseFunctionToken
	{
		internal TUserDefinedToken(int ArgCount, ptg aId, TCellFunctionData aFuncData) : base(ArgCount, aId, aFuncData) { }

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			if (FArgCount < 1)
			{
				return TFlxFormulaErrorValue.ErrNA;
			}

			object[] Parameters = new object[FArgCount - 1];
			for (int i = FArgCount - 2; i >= 0; i--)
			{
				Parameters[i] = ConvertToAllowedParam(FTokenList.EvaluateToken(wi, TUdfAggregate.Instance, CalcState, CalcStack));
			}

			string FunctionName = FlxConvert.ToString(FTokenList.EvaluateToken(wi, CalcState, CalcStack));
			int ActiveSheet = wi.Xls.ActiveSheet;
			try
			{
				return ConvertToAllowedParam(
					wi.Xls.EvaluateUserDefinedFunction(
					FunctionName, new TUdfEventArgs(wi.Xls, wi.SheetIndexBase1, wi.Row + 1 + wi.RowOfs, wi.Col + 1 + wi.ColOfs), Parameters));
			}
			finally
			{
				wi.Xls.ActiveSheet = ActiveSheet;
			}
		}
	}

	#endregion
	#endregion

	#region Name tokens
	internal class TNameToken : TBaseParsedToken
	{
		// Name in xlsx must be an index too, we can't store the name string itself in the tparsettokenlist(or someone might rename the range and it will fail)
		// but as we want to save back name strings anyway even if they don't exist, we need to first store the string, and inmmediately after loading, replace those by ids.

		internal int NameIndex;
		private ptg FId;
        private bool Recalculating; //prevent infinite recursion in names

		internal TNameToken(ptg aId, int aNameIndex)
			: base(0)
		{
			FId = aId;
			NameIndex = aNameIndex;
		}

		internal override ptg GetId
		{
			get { return FId; }
		}

		internal override TBaseParsedToken SetId(ptg aId)
		{
			FId = aId;
			return this;
		}

		protected object DoError(TWorkbookInfo wi, TUnsupportedFormulaErrorType ErrorType, string ExternalName)
		{
			wi.AddUnsupported(ErrorType, ExternalName);

			return TFlxFormulaErrorValue.ErrNA;
		}

        protected virtual bool DoEvaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, out object error, out TParsedTokenList NameData, TCalcState CalcState, TCalcStack CalcStack)
		{
			string ExternalName;
			bool IsAddin;
			bool NameError;
			NameData = wi.Xls.GetNamedRangeData(NameIndex, out ExternalName, out IsAddin, out NameError);
			if (NameError)
			{
				error = TFlxFormulaErrorValue.ErrNA;
				return false;
			}

			if (ExternalName != null)
			{
				if (IsAddin)
				{
					if (!wi.Xls.IsDefinedFunction(ExternalName))
					{
						error = DoError(wi, TUnsupportedFormulaErrorType.MissingFunction, ExternalName);
						return false;
					}
                    
					error = ExternalName;
					return false;
				}
				else
				{
					error = DoError(wi, TUnsupportedFormulaErrorType.ExternalReference, ExternalName);
					return false;
				}
			}
			else
			{
				if (NameData == null || NameData.IsEmpty)
				{
					error = TFlxFormulaErrorValue.ErrName;
					return false;
				}
				else
				{
					error = null;
					return true;
				}
			}
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
            if (Recalculating)
            {
                Recalculating = false;
                wi.AddUnsupported(TUnsupportedFormulaErrorType.CircularReference, string.Empty);
                return TFlxFormulaErrorValue.ErrName;
            }
            Recalculating = true;
            try
            {
                TParsedTokenList NameData;
                object error;
                bool NeedsEval = DoEvaluate(FTokenList, wi, f, out error, out NameData, CalcState, CalcStack);
                if (!NeedsEval) return error;

                return NameData.EvaluateAll(wi, f, CalcState, CalcStack);
            }
            finally
            {
                Recalculating = false;
            }
		}


		internal override object EvaluateRef(TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack)
		{
			TParsedTokenList NameData;
			object error;
			bool NeedsEval = DoEvaluate(FTokenList, wi, TErr2Aggregate.Instance, out error, out NameData, CalcState, CalcStack);
			if (!NeedsEval) return error;

			return NameData.EvaluateAllRef(wi, CalcState, CalcStack);
		}

		internal override bool Same(TBaseParsedToken aBaseParsedToken)
		{
			if (!base.Same(aBaseParsedToken)) return false;
			TNameToken tk = aBaseParsedToken as TNameToken;
			if (tk == null) return false;
			if (tk.NameIndex != NameIndex || tk.FId != FId) return false;
			return true;
		}

	}

	internal class TNameXToken : TNameToken
	{
		internal int FExternSheet;

		internal TNameXToken(ptg aId, int aExternSheet, int aNameIndex)
			: base(aId, aNameIndex)
		{
			FExternSheet = aExternSheet;
		}

        protected override bool DoEvaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, out object error, out TParsedTokenList NameData, TCalcState CalcState, TCalcStack CalcStack)
		{
			string ExternalBook;
			string ExternalName;
			int SheetIndexInOtherFile;
			bool IsAddin;
			bool NameError;

			NameData = wi.Xls.GetNamedRangeData(FExternSheet, NameIndex, out ExternalBook, out ExternalName, out SheetIndexInOtherFile, out IsAddin, out NameError);
			if (NameError)
			{
				error = TFlxFormulaErrorValue.ErrNA;
				return false;
			}

			if (ExternalName != null)
			{
				if (IsAddin)
				{
					if (!wi.Xls.IsDefinedFunction(ExternalName))
					{
						error = DoError(wi, TUnsupportedFormulaErrorType.MissingFunction, ExternalName);
						return false;
					}
                    
					error = ExternalName;
					return false;
				}
				else
				{
					ExcelFile ExternalXls = wi.Xls.GetSupportingFile(ExternalBook);
					if (ExternalXls != null)
					{
						int ExternalNameIndex = ExternalXls.FindNamedRange(ExternalName, SheetIndexInOtherFile);
						if (ExternalNameIndex < 0)
						{
							ExternalBook = null; //range not found.
							SheetIndexInOtherFile = -1;
						}

						if (ExternalBook == null)
						{
							error = TFlxFormulaErrorValue.ErrRef;
							return false;
						}

						TWorkbookInfo wi2 = wi.ShallowClone();
						wi2.Xls = wi.Xls.GetSupportingFile(ExternalBook);
						if (wi2.Xls == null)
						{
							error = TFlxFormulaErrorValue.ErrRef;
							return false;
						}

						error = wi2.Xls.EvaluateNamedRange(ExternalNameIndex, SheetIndexInOtherFile, f, CalcState, CalcStack);
						return false;
					}
					else
					{
						error = DoError(wi, TUnsupportedFormulaErrorType.ExternalReference, ExternalBook);
						return false;
					}
				}
			}
			else
			{
				if (NameData == null || NameData.IsEmpty)
				{
					error = TFlxFormulaErrorValue.ErrName;
					return false;
				}
				else
				{
					error = true;
					return true;
				}
			}
		}

		internal override bool Same(TBaseParsedToken aBaseParsedToken)
		{
			if (!base.Same(aBaseParsedToken)) return false;
			TNameXToken tk = aBaseParsedToken as TNameXToken;
			if (tk == null) return false;
			if (tk.FExternSheet != FExternSheet) return false;
			return true;
		}

		internal override int ExternSheet
		{
			get
			{
				return FExternSheet;
			}
			set
			{
				FExternSheet = value;
			}
		}


	}
	#endregion

	#region What-if Table
	internal class TTableToken : TBaseParsedToken
	{
		internal int Row;
		internal int Col;

		internal TTableToken(int aRow, int aCol) : base(0) { Row = aRow; Col = aCol; }

		internal override ptg GetId
		{
			get { return ptg.Tbl; }
		}

		internal override bool Same(TBaseParsedToken aBaseParsedToken)
		{
			if (!base.Same(aBaseParsedToken)) return false;
			TTableToken tk = aBaseParsedToken as TTableToken;
			if (tk == null) return false;
			if (tk.Row != Row || tk.Col != Col) return false;
			return true;
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			TCalcState NewState = CalcState.Clone();
			NewState.WhatIfRow = wi.Row + 1;
			NewState.WhatIfCol = wi.Col + 1;
			NewState.WhatIfSheet = wi.SheetIndexBase1;

			TCellAddress InputRow, InputCol;
			TXlsCellRange TableRange = wi.Xls.GetWhatIfTable(wi.SheetIndexBase1, Row + 1, Col + 1, out InputRow, out InputCol);
			if (TableRange == null) return TFlxFormulaErrorValue.ErrNA;

			if (InputRow != null)
			{
				NewState.TableRowCell.Row = InputRow.Row;
				NewState.TableRowCell.Col = InputRow.Col;
			}
			else
			{
				NewState.TableRowCell.Row = 0;
				NewState.TableRowCell.Col = 0;
			}

			if (InputCol != null)
			{
				NewState.TableColCell.Row = InputCol.Row;
				NewState.TableColCell.Col = InputCol.Col;
			}
			else
			{
				NewState.TableColCell.Row = 0;
				NewState.TableColCell.Col = 0;
			}

			if (TableRange.Top <= 0) return TFlxFormulaErrorValue.ErrRef;

			int Row0 = TableRange.Top - 1;
			int Col0 = TableRange.Left - 1;

			int RowEval, ColEval;
			if (NewState.TableRowCell.IsEmpty())
			{
				ColEval = wi.Col + 1;
				NewState.TableRowValue = null;
			}
			else
			{
				ColEval = Col0;
				NewState.TableRowValue = wi.Xls.GetCellValueAndRecalc(wi.SheetIndexBase1, Row0, wi.Col + 1, CalcState, CalcStack); //Use old calcstate here.
			}

			if (NewState.TableColCell.IsEmpty())
			{
				RowEval = wi.Row + 1;
				NewState.TableColValue = null;
			}
			else
			{
				RowEval = Row0;
				NewState.TableColValue = wi.Xls.GetCellValueAndRecalc(wi.SheetIndexBase1, wi.Row + 1, Col0, CalcState, CalcStack); //Use old calcstate here.
			}

            return wi.Xls.GetCellValueAndRecalc(wi.SheetIndexBase1, RowEval, ColEval, NewState, CalcStack);

		}

	}

	/// <summary>
	/// A special ptg table for formulas, where ptg.tbl doesn't mean a table :(
	/// </summary>
	internal class TTableObjToken : TBaseParsedToken
	{
		internal int Row;
		internal int Col;

		internal TTableObjToken(int aRow, int aCol) : base(0) { Row = aRow; Col = aCol; }

		internal override ptg GetId
		{
			get { return ptg.Tbl; }
		}

		internal override bool Same(TBaseParsedToken aBaseParsedToken)
		{
			if (!base.Same(aBaseParsedToken)) return false;
			TTableObjToken tk = aBaseParsedToken as TTableObjToken;
			if (tk == null) return false;
			if (tk.Row != Row || tk.Col != Col) return false;
			return true;
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			return TFlxFormulaErrorValue.ErrNA; //This token shouldn't be evaluated.
		}

	}
	#endregion

	#region Non Calculating Tokens
	internal abstract class TIgnoreInCalcToken : TBaseParsedToken
	{
		protected TIgnoreInCalcToken(int ArgCount)
			: base(ArgCount)
		{
		}

		internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState CalcState, TCalcStack CalcStack)
		{
			return FTokenList.EvaluateToken(wi, f, CalcState, CalcStack);
		}

		internal override object EvaluateRef(TParsedTokenList FTokenList, TWorkbookInfo wi, TCalcState CalcState, TCalcStack CalcStack)
		{
			return FTokenList.EvaluateTokenRef(wi, CalcState, CalcStack);
		}

		internal override void Flush(TParsedTokenList FTokenList)
		{
			FTokenList.LightPop().Flush(FTokenList);
		}

	}

	internal class TParenToken : TIgnoreInCalcToken
	{
		private TParenToken()
			: base(0)
		{
		}
		internal static readonly TParenToken Instance = new TParenToken();

		internal override ptg GetId
		{
			get { return ptg.Paren; }
		}
        
		internal override TBaseParsedToken Clone()
		{
			return this;
		}

	}

	internal class TExp_Token : TIgnoreInCalcToken
	{
		internal int Row;
		internal int Col;

		internal TExp_Token(int aRow, int aCol) : base(0) { Row = aRow; Col = aCol; }

		internal override ptg GetId
		{
			get { return ptg.Exp; }
		}

		internal override bool Same(TBaseParsedToken aBaseParsedToken)
		{
			if (!base.Same(aBaseParsedToken)) return false;
			TExp_Token tk = aBaseParsedToken as TExp_Token;
			if (tk == null) return false;
			if (tk.Row != Row || tk.Col != Col) return false;
			return true;
		}

	}


	internal struct TRefRange
	{
		internal int FirstRow;
		internal int FirstCol;
		internal int LastRow;
		internal int LastCol;

		internal TRefRange Clone()
		{
			return (TRefRange)MemberwiseClone();
		}
	}

	#endregion

	#region Mem tokens
	internal class TSimpleMemToken: TIgnoreInCalcToken
	{
		private ptg FId;
		internal int PositionOfNextPtg;

		internal TSimpleMemToken(ptg aId, int aPositionOfNextPtg): base(0)
		{
			FId = aId;
			PositionOfNextPtg = aPositionOfNextPtg;
		}

		internal override ptg GetId
		{
			get { return FId; }
		}

		internal override TBaseParsedToken SetId(ptg aId)
		{
			FId = aId;
			return this;
		}

		internal override bool Same(TBaseParsedToken aBaseParsedToken)
		{
			if (!base.Same(aBaseParsedToken)) return false;
			TSimpleMemToken tk = aBaseParsedToken as TSimpleMemToken;
			if (tk == null) return false;
			if (tk.FId != FId || PositionOfNextPtg != tk.PositionOfNextPtg) return false;
			return true;
		}
	}

	/// <summary>
	/// An array of constant references used when intersecting ranges.
	/// Can be safely ignored when calculating.
	/// </summary>
	internal class TMemAreaToken : TSimpleMemToken
	{
		internal TRefRange[] Data;

		internal TMemAreaToken(ptg aId, TRefRange[] aData, int aPositionOfNextPtg): base(aId, aPositionOfNextPtg)
		{
			Data = aData;
		}

		internal override TBaseParsedToken Clone()
		{
			TMemAreaToken Result = (TMemAreaToken)MemberwiseClone();
			Result.Data = (TRefRange[])Data.Clone(); //even when the array clone is a shallow copy, as trefrange are structs, we are ok.
			return Result;
		}

		internal override bool Same(TBaseParsedToken aBaseParsedToken)
		{
			if (!base.Same(aBaseParsedToken)) return false;
			TMemAreaToken tk = aBaseParsedToken as TMemAreaToken;
			if (tk == null) return false;
			if (tk.Data.Length != Data.Length) return false;
			//We could also check for the data itself, but it is not worth, since this token is not used anyway.
			return true;
		}
	}

	internal class TMemErrToken: TSimpleMemToken
	{
		internal TFlxFormulaErrorValue ErrorValue;

		internal TMemErrToken(ptg aId, TFlxFormulaErrorValue aErrorValue, int aPositionOfNextPtg): base(aId, aPositionOfNextPtg)
		{
			ErrorValue = aErrorValue;
		}

		internal override bool Same(TBaseParsedToken aBaseParsedToken)
		{
			if (!base.Same(aBaseParsedToken)) return false;
			TMemErrToken tk = aBaseParsedToken as TMemErrToken;
			if (tk == null) return false;
			if (tk.ErrorValue != ErrorValue) return false;
			return true;
		}

	}

	internal class TMemNoMemToken: TSimpleMemToken
	{
		internal TMemNoMemToken(ptg aId, int aPositionOfNextPtg): base(aId, aPositionOfNextPtg)
		{
		}
	}

	internal class TMemFuncToken: TSimpleMemToken
	{
		internal TMemFuncToken(ptg aId, int aPositionOfNextPtg): base(aId, aPositionOfNextPtg)
		{
		}
	}
	

	/// <summary>
	/// This class doesn't exist anymore in biff8, and differently from memarea, doesn't add an array.
	/// </summary>
	internal class TMemAreaNToken: TSimpleMemToken
	{
		internal TMemAreaNToken(ptg aId, int aPositionOfNextPtg): base(aId, aPositionOfNextPtg)
		{
		}
	}

	internal class TMemNoMemNToken: TSimpleMemToken
	{
		internal TMemNoMemNToken(ptg aId, int aPositionOfNextPtg): base(aId, aPositionOfNextPtg)
		{
		}
	}
	

	#endregion

	#region Attr
	internal abstract class TAttrToken : TIgnoreInCalcToken
	{
		internal TAttrToken(int ArgCount) : base(0) {}

		internal override ptg GetId
		{
			get { return ptg.Attr; }
		}

		internal override TBaseParsedToken Clone()
		{
			return this;
		}

	}

	internal sealed class TAttrVolatileToken : TAttrToken
	{
		private TAttrVolatileToken() : base(0) { }
		internal static readonly TAttrVolatileToken Instance = new TAttrVolatileToken();
	}

	internal class TAttrSpaceToken : TAttrToken
	{
		internal bool Volatile;
		internal FormulaAttr SpaceType;
		internal int SpaceCount;

		internal TAttrSpaceToken(FormulaAttr aSpaceType, int aSpaceCount, bool aVolatile)
			: base(0)
		{
			SpaceType = aSpaceType;
			SpaceCount = aSpaceCount;
			Volatile = aVolatile;
		}

		internal override bool Same(TBaseParsedToken aBaseParsedToken)
		{
			if (!base.Same(aBaseParsedToken)) return false;
			TAttrSpaceToken tk = aBaseParsedToken as TAttrSpaceToken;
			if (tk == null) return false;
			if (tk.Volatile != Volatile || tk.SpaceCount != SpaceCount || tk.SpaceType != SpaceType) return false;
			return true;
		}

	}

	internal class TAttrGotoToken : TAttrToken
	{
		/// <summary>
		/// Stores the absolute position where to jump in the List.
		/// </summary>
        internal int PositionOfNextPtg;

        internal TAttrGotoToken(int aPositionOfNextPtg)
			: base(0)
		{
            PositionOfNextPtg = aPositionOfNextPtg;
		}

		internal override bool Same(TBaseParsedToken aBaseParsedToken)
		{
			if (!base.Same(aBaseParsedToken)) return false;
			TAttrGotoToken tk = aBaseParsedToken as TAttrGotoToken;
			if (tk == null) return false;
			if (tk.PositionOfNextPtg != PositionOfNextPtg) return false;
			return true;
		}
	}

	internal class TAttrOptIfToken : TAttrGotoToken
	{
        internal TAttrOptIfToken(int aPositionOfNextPtg)
            : base(aPositionOfNextPtg)
		{
		}
	}

	internal class TAttrOptChooseToken : TAttrToken
	{
		internal int[] PositionOfNextPtg;
		internal TAttrOptChooseToken(int[] aPositionOfNextPtg)
			: base(0)
		{
			PositionOfNextPtg = aPositionOfNextPtg;
		}

		internal override bool Same(TBaseParsedToken aBaseParsedToken)
		{
			if (!base.Same(aBaseParsedToken)) return false;
			TAttrOptChooseToken tk = aBaseParsedToken as TAttrOptChooseToken;
			if (tk == null) return false;
			if (tk.PositionOfNextPtg == null) return PositionOfNextPtg == null;
			if (PositionOfNextPtg == null) return false;
			if (PositionOfNextPtg.Length != tk.PositionOfNextPtg.Length) return false;
			for (int i = 0; i < PositionOfNextPtg.Length; i++)
			{
				if (PositionOfNextPtg[i] != tk.PositionOfNextPtg[i]) return false;
			}
			return true;
		}
	}


	internal class TAttrSumToken : TSumToken  //THIS TOKEN IS NOT IGNOREINCALC!!
	{
		private TAttrSumToken()
			: base(1, ptg.Attr, TXlsFunction.GetData(4))
		{
		}

		internal static readonly TAttrSumToken Instance = new TAttrSumToken();

		internal override TBaseParsedToken SetId(ptg aId)
		{
			return this; //we can never change this id from 0x19
		}
	}

	#endregion

	#endregion

	#region Functions Implemented
	/// <summary>
	/// Data of an implemented formula function.
	/// </summary>
	public class TImplementedFunction
	{
		private int FId;
		private string FFunctionName;
		private int FMinArgCount;
		private int FMaxArgCount;

		/// <summary>
		/// Creates a new TImplementedFunction with supplied values.
		/// </summary>
		/// <param name="aId">Formula ID</param>
		/// <param name="aFunctionName">Formula Name</param>
		/// <param name="aMinArgCount">Minimum argument count for the function.</param>
		/// <param name="aMaxArgCount">Maximum argument count for the function.</param>
		public TImplementedFunction(int aId, string aFunctionName, int aMinArgCount, int aMaxArgCount)
		{
			Id = aId;
			FunctionName = aFunctionName;
			FMinArgCount = aMinArgCount;
			FMaxArgCount = aMaxArgCount;
		}

		/// <summary>
		/// Formula ID. (Excel specific)
		/// </summary>
		public int Id { get { return FId; } set { FId = value; } }

		/// <summary>
		/// Formula Name.
		/// </summary>
		public string FunctionName { get { return FFunctionName; } set { FFunctionName = value; } }

		/// <summary>
		/// Minimum argument count for the function.
		/// </summary>
		public int MinArgCount { get { return FMinArgCount; } set { FMinArgCount = value; } }

		/// <summary>
		/// Maximum argument count for the function.
		/// </summary>
		public int MaxArgCount { get { return FMaxArgCount; } set { FMaxArgCount = value; } }
	}

	/// <summary>
	/// Holds a list of currently implemented formula functions on FlexCel.
	/// </summary>
	[Serializable]
	public class TImplementedFunctionList : 
        Dictionary<int, TImplementedFunction>
	{
		/// <summary>
		/// Creates and initializes a list with all implemented functions.
		/// </summary>
		public TImplementedFunctionList()
		{
			Fill();
		}

#if (!COMPACTFRAMEWORK && !SILVERLIGHT)
		/// <summary>
		/// Creates a serialized object.
		/// </summary>
		/// <param name="info"></param>
		/// <param name="context"></param>
		protected TImplementedFunctionList(SerializationInfo info, StreamingContext context)
			: base(info, context)
		{
			Fill();
		}
#endif

		private void Fill()
		{
			//When adding a new function here, search for AddFunctionHere and add it on all those places

			AddFunction(0);
			AddFunction(169);
			AddFunction(1);
			AddFunction(347); //countblank
			AddFunction(4);
			AddFunction(183); //Product
			AddFunction(321); //sumsq
			AddFunction(5);
			AddFunction(361); //AverageA
			AddFunction(31);
			AddFunction(32);
			AddFunction(115);
			AddFunction(116);
			AddFunction(112);
			AddFunction(113);
			AddFunction(118);
			AddFunction(111);
			AddFunction(214);
			AddFunction(119);
			AddFunction(124);
			AddFunction(336);
			AddFunction(117);
			AddFunction(114);
			AddFunction(30);  //rept
			AddFunction(82);
			AddFunction(162);
			AddFunction(120);
			AddFunction(48);
			AddFunction(33);
			AddFunction(121);
			AddFunction(130);
			AddFunction(131);

			AddFunction(34);
			AddFunction(35);
			AddFunction(36);
			AddFunction(37);
			AddFunction(38);

			AddFunction(13); //dollar
			AddFunction(14); //fixed

			AddFunction(359); //Hyperlink

			AddFunction(6);
			AddFunction(7);
			AddFunction(8);
			AddFunction(9);
			AddFunction(75); //AREAS
			AddFunction(76); //ROWS
			AddFunction(77); //COLUMNS
			AddFunction(83); //TRANSPOSE

			AddFunction(39);
			AddFunction(27);
			AddFunction(24);
			AddFunction(288);
			AddFunction(212);  //ROUNDUP
			AddFunction(213);

			AddFunction(279);
			AddFunction(298);

			AddFunction(285);
			AddFunction(21);
			AddFunction(25);
			AddFunction(22);
			AddFunction(109);
			AddFunction(23);
			AddFunction(19);
			AddFunction(337);
			AddFunction(63);
			AddFunction(26);
			AddFunction(20);
			AddFunction(197);

			AddFunction(184); //fact
			AddFunction(276); //combin
			AddFunction(299); //permut

			AddFunction(344);

			AddFunction(148);

			AddFunction(342);
			AddFunction(343);
			AddFunction(15);  //sin
			AddFunction(16);
			AddFunction(17);
			AddFunction(18);
			AddFunction(97);
			AddFunction(98);
			AddFunction(99);

			AddFunction(229);  //sinh
			AddFunction(230);  //cosh
			AddFunction(231);  //tanh
			AddFunction(232);  //asinh
			AddFunction(233);  //acosh
			AddFunction(234);  //atanh


			AddFunction(345);
			AddFunction(346);

			AddFunction(65);
			AddFunction(140);
			AddFunction(67);
			AddFunction(68);
			AddFunction(69);
			AddFunction(220); //Days360

			AddFunction(66);
			AddFunction(141);
			AddFunction(71);
			AddFunction(72);
			AddFunction(73);
			AddFunction(70);  //WeekDay

			AddFunction(74);
			AddFunction(221);

			AddFunction(261);
			AddFunction(2);
			AddFunction(3);
			AddFunction(126);
			AddFunction(127);
			AddFunction(128);
			AddFunction(129);
			AddFunction(190);
			AddFunction(198);
			AddFunction(105);

			AddFunction(86);
			AddFunction(10);

			AddFunction(78);
			AddFunction(100);
			AddFunction(101);
			AddFunction(102);
			AddFunction(28);  //lookup
			AddFunction(29);
			AddFunction(64);
			AddFunction(219); //address

			AddFunction(228);

			AddFunction(46); //var
			AddFunction(194); //varp
			AddFunction(365); //vara
			AddFunction(367); //varpa
			AddFunction(12); //stdev
			AddFunction(193); //stdevp
			AddFunction(366); //stdeva
			AddFunction(364); //stdevpa
			AddFunction(307); //correl
			AddFunction(308); //covar
			AddFunction(322); //kurt
			AddFunction(323); //skew

			AddFunction(125); //cell

			AddFunction(325); //Large
			AddFunction(326); //Small
			AddFunction(362); //maxa
			AddFunction(363); //mina

			AddFunction(40); //DCount
			AddFunction(41); //DSum
			AddFunction(42); //DAverage
			AddFunction(43); //DMin
			AddFunction(44); //DMax
			AddFunction(189); //DProduct
			AddFunction(199); //DCountA
			AddFunction(235); //DGet
			AddFunction(47); //Dvar
			AddFunction(196); //DvarP
			AddFunction(45); //DStDev
			AddFunction(195); //DStDev

			AddFunction(354); //Roman

			AddFunction(247); //db
			AddFunction(142); //sln
			AddFunction(143); //syd
			AddFunction(144); //ddb

			AddFunction(11); //NPV
			AddFunction(56); //PV
			AddFunction(57); //FV
			AddFunction(58); //NPer
			AddFunction(60); //Rate
			AddFunction(61); //MIRR
			AddFunction(62); //IRR
			AddFunction(167); //IPMT
			AddFunction(168); //PPMT

			AddFunction(271); //GAMMALN
			AddFunction(273); //BINOMDIST
			AddFunction(274); //CHIDIST
			AddFunction(275); //CHIINV
			AddFunction(277); //CONFIDENCE
			AddFunction(280); //EXPONDIST
			AddFunction(286); //GAMMADIST
			AddFunction(287); //GAMMAINV
			AddFunction(289); //HYPGEOMDIST
			AddFunction(290); //LOGNORMDIST
			AddFunction(291); //LOGINV
			AddFunction(292); //NEGBINOMDIST
			AddFunction(293); //NORMDIST
			AddFunction(294); //NORMSDIST
			AddFunction(295); //NORMINV
			AddFunction(296); //NORMSINV
			AddFunction(297); //STANDARDIZE
			AddFunction(300); //POISSON
			AddFunction(302); //WEIBULL
			AddFunction(306); //CHITEST
			AddFunction(324); //ZTEST

			AddFunction(319); //GEOMEAN
			AddFunction(320); //HARMEAN
			AddFunction(216); //RANK

			AddFunction(303); //SUMXMY2
			AddFunction(304); //SUMX2MY2
			AddFunction(305); //SUMX2PY2

			AddFunction(165); //MMult

			AddFunction(269);//AVEDEV
			AddFunction(318);//DEVSQ
			AddFunction(311);//INTERCEPT
			AddFunction(312);//PEARSON
			AddFunction(313);//RSQ
			AddFunction(314);//STEYX
			AddFunction(315);//SLOPE

			AddFunction(283); //FISHER
			AddFunction(284); //FISHERINV

			AddFunction(227); //MEDIAN
			AddFunction(327); //QUARTILE
			AddFunction(328); //PERCENTILE
			AddFunction(330); //MODE

			AddFunction(255); //USER DEFINED

            AddFunction(0x1E0); //IfError
            AddFunction(0x1E1); //CountIfs
            AddFunction(0x1E2); //SumIfs
            AddFunction(0x1E3); //AverageIf
            AddFunction(0x1E4); //AverageIfs

            AddFunction((int)TFutureFunctions.CeilingPrecise); //Ceiling.Precise
            AddFunction((int)TFutureFunctions.IsoCeiling); //Iso.Ceiling
            AddFunction((int)TFutureFunctions.FloorPrecise); //Floor.Precise
            AddFunction((int)TFutureFunctions.Aggregate); //Aggregate
            AddFunction((int)TFutureFunctions.PercentileExc);
            AddFunction((int)TFutureFunctions.QuartileExc);

            AddFunction((int)TFutureFunctions.BinomDist);
            AddFunction((int)TFutureFunctions.ChisqDistRt);
            AddFunction((int)TFutureFunctions.ChisqInvRt);
            AddFunction((int)TFutureFunctions.ChisqTest);
            AddFunction((int)TFutureFunctions.ConfidenceNorm);
            
            AddFunction((int)TFutureFunctions.CovarianceP);
            AddFunction((int)TFutureFunctions.ExponDist);
            AddFunction((int)TFutureFunctions.GammaDist);
            AddFunction((int)TFutureFunctions.GammaInv);
            AddFunction((int)TFutureFunctions.HypGeomDist);

            AddFunction((int)TFutureFunctions.LogNormDist);
            AddFunction((int)TFutureFunctions.LogNormInv);
            AddFunction((int)TFutureFunctions.ModeSngl);
            AddFunction((int)TFutureFunctions.NegBinom);
            AddFunction((int)TFutureFunctions.NormDist);
            
            AddFunction((int)TFutureFunctions.NormInv);
            AddFunction((int)TFutureFunctions.NormSDist);
            AddFunction((int)TFutureFunctions.NormSInv);
            AddFunction((int)TFutureFunctions.PercentileInc);
            AddFunction((int)TFutureFunctions.QuartileInc);
           
            AddFunction((int)TFutureFunctions.PercentRankInc);
            AddFunction((int)TFutureFunctions.PoissonDist);
            AddFunction((int)TFutureFunctions.RankEq);
            AddFunction((int)TFutureFunctions.StDevP);
            AddFunction((int)TFutureFunctions.StDevS);
            
            AddFunction((int)TFutureFunctions.VarP);
            AddFunction((int)TFutureFunctions.VarS);
            AddFunction((int)TFutureFunctions.WeibullDist);
            AddFunction((int)TFutureFunctions.ZTest); 

        }

		private void AddFunction(int id)
		{
			TCellFunctionData fd = null;
			fd = TXlsFunction.GetData(id);
			base.Add(fd.Index, new TImplementedFunction(fd.Index, fd.Name, fd.MinArgCount, fd.MaxArgCount));
		}        
	}

	#endregion
}
