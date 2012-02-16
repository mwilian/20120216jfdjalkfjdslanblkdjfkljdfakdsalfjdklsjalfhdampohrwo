using System;

using FlexCel.Core;
using System.Collections.Generic;
using FlexCel.Render;

#if (WPF)
using RectangleF = System.Windows.Rect;
using SizeF = System.Windows.Size;
using real = System.Double;
#else
using System.Drawing;
using real = System.Single;
#endif


namespace FlexCel.XlsAdapter
{
	/// <summary>
	/// Class for calculating the automatic row heights. 
	/// This is a tricky thing because we are coupling GDI+ calls with 
	/// non-graphic code, but there is no other way to do it.
	/// </summary>
	internal class TRowHeightCalc: IDisposable
	{
		#region Privates
		int[] XFHeight;
		TFlxFont[] XFFonts;
		TWorkbookGlobals Wg;
		TGraphicCanvas gr;


		#endregion

		#region Constructor
		internal TRowHeightCalc(TWorkbookGlobals aWg)
		{
			Wg = aWg;
			gr = new TGraphicCanvas();

			InitXF();
		}

		private void InitXF()
		{
			XFHeight = new int[Wg.CellXF.Count];
			XFFonts = new TFlxFont[XFHeight.Length];
			for (int i = 0; i< XFHeight.Length; i++)
			{
				TXFRecord xf = Wg.CellXF[i];
				int FontIndex = xf.GetActualFontIndex(Wg.Fonts);
				TFontRecord Fr = Wg.Fonts[FontIndex];
				XFFonts[i] = Fr.FlxFont();
				XFHeight[i] = (int)ExcelMetrics.GetRowHeightInPixels(XFFonts[i]);

			}
		}
		#endregion

		#region Public methods
		internal int CalcCellHeight(int Row, int Col, TRichString val, int CellXF, ExcelFile Workbook, real RowMultDisplay, real ColMultDisplay, TMultipleCellAutofitList MultipleRowAutofits)
		{
			if (CellXF < 0) return 0xFF;
			if (CellXF >= XFHeight.Length) return 0xFF;  //Just to make sure. We don't want a wrong file to blow here.
			int Result = XFHeight[CellXF];
			int Result0 = Result;

			if (val == null) return Result;

			TXFRecord XFRec = Wg.CellXF[CellXF];
			bool Vertical;
			real Alpha = FlexCelRender.CalcAngle(XFRec.Rotation, out Vertical);


			if (!XFRec.WrapText && !Vertical && Alpha == 0) return Result;

			Font AFont = gr.FontCache.GetFont(XFFonts[CellXF], 1);
			
			RectangleF CellRect = new RectangleF();
			TXlsCellRange rg = Workbook.CellMergedBounds(Row, Col);
			
			for (int c = rg.Left; c <= rg.Right; c++)
			{
				CellRect.Width += Workbook.GetColWidth(c, true) / ColMultDisplay;
			}

			SizeF TextExtent;
			TXRichStringList TextLines = new TXRichStringList();
			TFloatList MaxDescent = new TFloatList();

			real Clp = 1 * FlexCelRender.DispMul / 100f;
			real Wr = 0;

			if (Alpha == 0) Wr = CellRect.Right - CellRect.Left - 2 * Clp;
			else
			{
				Wr = 1e6f; //When we have an angle, it means "make the row as big as needed to fit". So SplitText needs no limits.
			}


			RenderMetrics.SplitText(gr.Canvas, gr.FontCache, 1, val, AFont, Wr, TextLines, out TextExtent, Vertical, MaxDescent, null);
			
			if (TextLines.Count <= 0) return Result;
			
			real H = 0;
			real W = 0;
			for (int i = 0; i < TextLines.Count; i++)
			{
				H += TextLines[i].YExtent;
				if (TextLines[i].XExtent > W) W = TextLines[i].XExtent;
			}

			if (Alpha != 0)
			{
				real SinAlpha = (real)Math.Sin(Alpha * Math.PI / 180); real CosAlpha = (real)Math.Cos(Alpha * Math.PI / 180);
				H = H * CosAlpha + W * Math.Abs(SinAlpha);
			}
			Result = (int) Math.Ceiling(RowMultDisplay * (H + 2*Clp));

			if (rg.RowCount > 1) 
			{
				if (MultipleRowAutofits != null) MultipleRowAutofits.Add(new TMultipleCellAutofit(rg, Result));
				return Result0; //We will autofit this later.
			}
			if (Result < 0) Result = 0;
			
			return Result;
		}
		#endregion

		#region IDisposable Members

		public void Dispose()
		{
			gr.Dispose();
            GC.SuppressFinalize(this);
        }

		#endregion
	}

	internal class TColWidthCalc: IDisposable
	{
		#region Privates
		TFlxFont[] XFFonts;
		TWorkbookGlobals Wg;

		TGraphicCanvas gr;

		#endregion

		#region Constructor
		internal TColWidthCalc(TWorkbookGlobals aWg)
		{
			Wg = aWg;
			gr = new TGraphicCanvas();

			InitXF();
		}

		private void InitXF()
		{
			XFFonts = new TFlxFont[Wg.CellXF.Count];
			for (int i = 0; i< XFFonts.Length; i++)
			{
				TXFRecord xf = Wg.CellXF[i];
				int FontIndex = xf.GetActualFontIndex(Wg.Fonts);
				TFontRecord Fr = Wg.Fonts[FontIndex];
				XFFonts[i] = Fr.FlxFont();
			}
		}
		#endregion

		#region Public methods
		internal int CalcCellWidth(int Row, int Col, TRichString val, int CellXF, ExcelFile Workbook, real RowMultDisplay, real ColMultDisplay, TMultipleCellAutofitList MultipleColAutofits)
		{
			if (val == null || val.Value == null || val.Value.Length == 0) return 0;

			TXFRecord XFRec = Wg.CellXF[CellXF];
			bool Vertical;
			real Alpha = FlexCelRender.CalcAngle(XFRec.Rotation, out Vertical);

			Font AFont = gr.FontCache.GetFont(XFFonts[CellXF], 1); //dont dispose
			
			TXlsCellRange rg = Workbook.CellMergedBounds(Row, Col);

			SizeF TextExtent;
			TXRichStringList TextLines = new TXRichStringList();
			TFloatList MaxDescent = new TFloatList();

			real Clp = 1 * FlexCelRender.DispMul / 100f;
			real Wr = 0;

			if (Alpha != 90 && Alpha != -90) Wr = 1e6f; //this means "make the column as big as needed to fit". So SplitText needs no limits.
			else
			{
				RectangleF CellRect = new RectangleF();
				for (int r = rg.Top; r <= rg.Bottom; r++)
				{
					CellRect.Height += Workbook.GetRowHeight(r, true) / RowMultDisplay;
				}
				Wr = CellRect.Height - 2 * Clp; 
			}


			RenderMetrics.SplitText(gr.Canvas, gr.FontCache, 1, val, AFont, Wr, TextLines, out TextExtent, Vertical, MaxDescent, null);
			
			if (TextLines.Count <= 0) return 0;
			
			real Rr = 0;
			real Ww = 0;
			for (int i = 0; i < TextLines.Count; i++)
			{
				Rr += TextLines[i].YExtent;
				if (TextLines[i].XExtent > Ww) Ww = TextLines[i].XExtent;
			}
			if (Alpha != 0)
			{
				real SinAlpha = (real)Math.Sin(Alpha * Math.PI / 180); real CosAlpha = (real)Math.Cos(Alpha * Math.PI / 180);
				Ww = Ww * CosAlpha + Rr * Math.Abs(SinAlpha);
			}
			int Result = (int) Math.Ceiling(ColMultDisplay * (Ww + 4*Clp));

			if (rg.ColCount > 1) 
			{
				if (MultipleColAutofits != null) MultipleColAutofits.Add(new TMultipleCellAutofit(rg, Result));
				return 0; //We will autofit this later.
			}

			if (Result < 0) Result = 0;
			return Result;
			
		}
		#endregion

		#region IDisposable Members

		public void Dispose()
		{
			gr.Dispose();
            GC.SuppressFinalize(this);
		}

		#endregion
	}


	#region Multiple Cell Autofits
	/// <summary>
	/// We will use this to store cells for autofitting later. We can't really autofit rows when cells have merged rows until
	/// we have autofitted all rows in the merged range. Just then it makes sense to autofit the big cell, if height is still not enough.
	/// </summary>
	internal class TMultipleCellAutofit
	{
		internal TXlsCellRange Cell;
		internal int NeededSize;

		internal TMultipleCellAutofit(TXlsCellRange aCell, int aNeededSize)
		{
			Cell = aCell;
			NeededSize = aNeededSize;
		}
	}

	internal class TMultipleCellAutofitList: List<TMultipleCellAutofit>
	{
	}
	#endregion
}
