using System;
using System.Collections.Generic;

using FlexCel.Core;

#if (MONOTOUCH)
using System.Drawing;
using real = System.Single;
using Color = MonoTouch.UIKit.UIColor;
using Font = MonoTouch.UIKit.UIFont;
#else
	#if (WPF)
	using RectangleF = System.Windows.Rect;
	using SizeF = System.Windows.Size;
	using PointF = System.Windows.Point;
	using real = System.Double;
	using System.Windows.Media;
	#else
	using real = System.Single;
	using Colors = System.Drawing.Color;
	using DashStyles = System.Drawing.Drawing2D.DashStyle;
	using System.Drawing;
	using System.Drawing.Drawing2D;
	using System.Drawing.Imaging;
	#endif
#endif
	using System.Globalization;

namespace FlexCel.Render
{

	internal sealed class TChartCanvas
	{        
		internal IFlxGraphics Canvas;
		internal RectangleF Coords;
		internal RectangleF ChartCoords;

		internal TChartCanvas(IFlxGraphics aCanvas, RectangleF aCoords, RectangleF aChartCoords)
		{
			Canvas = aCanvas;
			Coords = aCoords;
			ChartCoords = aChartCoords;
		}

		internal SizeF MeasureString(string text, Font aFont)
		{
			return Canvas.MeasureString(text, aFont);
		}

		internal bool tx(TAxisInfo Ax, ref real x, real w)
		{
			if (Ax != null && Ax.Horizontal)
			{
				if (Ax.ReverseValues) x = Coords.Left + Coords.Right - x - w;
				return true;
			}
			return false;

		}
		internal real tx(TAxisInfo Ax1, TAxisInfo Ax2, real x, real w)
		{
			if (!tx(Ax1, ref x, w)) 			
			{
				tx(Ax2, ref x, w);
			}
			return x;

		}

		internal bool ty(TAxisInfo Ax, ref real y, real h)
		{
			if (Ax != null && !Ax.Horizontal)
			{
				if (Ax.ReverseValues) y = Coords.Top + Coords.Bottom - y - h;
				return true;
			}
			return false;

		}

		internal real ty(TAxisInfo Ax1, TAxisInfo Ax2, real y, real h)
		{
			if (!ty(Ax1, ref y, h)) 			
			{
				ty(Ax2, ref y, h);
			}
			return y;

		}

		internal void Mirror(TAxisInfo Ax1, TAxisInfo Ax2, ref RectangleF Rect)
		{
			Rect.X = tx(Ax1, Ax2, Rect.Left, Rect.Width);
			Rect.Y = ty(Ax1, Ax2, Rect.Top, Rect.Height);
		}

		internal void MirrorXY(TAxisInfo Ax1, TAxisInfo Ax2, ref RectangleF ContainingRect, ref real[] X, ref real[] Y, TXRichStringList TextLines, real SinAlpha, real CosAlpha)
		{
				real OldY = ContainingRect.Y;
				ContainingRect.Y = ty(Ax1, Ax2, ContainingRect.Y, ContainingRect.Height);

				for (int n = 0; n < Y.Length; n++)
				{
					Y[n] += ContainingRect.Y - OldY;
				}

				for (int n = 0; n < X.Length; n++)
				{
					X[n] = tx(Ax1, Ax2, X[n], TextLines[n].XExtent * CosAlpha - TextLines[n].YExtent * SinAlpha);
				}
				ContainingRect.X = tx(Ax1, Ax2, ContainingRect.X, ContainingRect.Width);
}


		internal void FillRectangle(TAxisInfo Ax1, TAxisInfo Ax2, Brush aBrush, RectangleF Rect)
		{
			Canvas.FillRectangle(aBrush, tx(Ax1, Ax2, Rect.Left, Rect.Width), ty(Ax1, Ax2, Rect.Top, Rect.Height), Rect.Width, Rect.Height);
		}

		internal void FillRectangle(TAxisInfo Ax1, TAxisInfo Ax2, Brush aBrush, real x1, real y1, real w, real h)
		{
			Canvas.FillRectangle(aBrush, tx(Ax1, Ax2, x1, w), ty(Ax1, Ax2, y1, h), w, h);
		}

		internal void DrawAndFillRectangle(TAxisInfo Ax1, TAxisInfo Ax2, Pen aPen, Brush aBrush, real x1, real y1, real w, real h)
		{
			Canvas.DrawAndFillRectangle(aPen, aBrush, tx(Ax1, Ax2, x1, w), ty(Ax1, Ax2, y1, h), w, h);
		}

		internal void DrawImage(TAxisInfo Ax1, TAxisInfo Ax2, Image aImage, real x1, real y1, real w, real h)
		{
			if (aImage == null) return;
			Canvas.DrawImage(aImage, new RectangleF(tx(Ax1, Ax2, x1, w), ty(Ax1, Ax2, y1, h), w, h), new RectangleF(0,0,aImage.Width, aImage.Height), 
				FlxConsts.NoTransparentColor, FlxConsts.DefaultBrightness, FlxConsts.DefaultContrast, FlxConsts.DefaultGamma, ColorUtil.Empty, null);
		}

		internal void DrawLine(TAxisInfo Ax1, TAxisInfo Ax2, Pen aPen, real x1, real y1, real x2, real y2)
		{
			Canvas.DrawLine(aPen, tx(Ax1, Ax2, x1, 0), ty(Ax1, Ax2, y1, 0), tx(Ax1, Ax2, x2, 0), ty(Ax1, Ax2, y2, 0));
		}

		internal void DrawString(TAxisInfo Ax1, TAxisInfo Ax2, string Text, Font aFont, Brush aBrush, real x, real y, real tw, real th)
		{
			Canvas.DrawString(Text, aFont, aBrush, tx(Ax1, Ax2, x, tw), ty(Ax1, Ax2, y, -th));
		}

		internal void DrawAndFillBeziers(TAxisInfo Ax1, TAxisInfo Ax2, Pen aPen, Brush aBrush, TPointF[] Points)
		{
			TPointF[] NewPoints = Points;

			if ((Ax1 != null && Ax1.ReverseValues) //This is going to be a very rare thing.
				||
				(Ax2 != null && Ax2.ReverseValues))
			{
				NewPoints = new TPointF[Points.Length];
				for (int i = 0; i < Points.Length; i++)
				{
					NewPoints[i] = new TPointF(tx(Ax1, Ax2, Points[i].X, 0), ty(Ax1, Ax2, Points[i].Y, 0));
				}
			}
			Canvas.DrawAndFillBeziers(aPen, aBrush, NewPoints);
		}

		internal void DrawAndFillPolygon(TAxisInfo Ax1, TAxisInfo Ax2, Pen aPen, Brush aBrush, TPointF[] Points)
		{
			TPointF[] NewPoints = Points;

			if ((Ax1 != null && Ax1.ReverseValues) //This is going to be a very rare thing.
				||
				(Ax2 != null && Ax2.ReverseValues))
			{
				NewPoints = new TPointF[Points.Length];
				for (int i = 0; i < Points.Length; i++)
				{
					NewPoints[i] = new TPointF(tx(Ax1, Ax2, Points[i].X, 0), ty(Ax1, Ax2, Points[i].Y, 0));
				}
			}
			Canvas.DrawAndFillPolygon(aPen, aBrush, NewPoints);
		}

		internal void DrawLines(TAxisInfo Ax1, TAxisInfo Ax2, Pen aPen, TPointF[] Points)
		{
			TPointF[] NewPoints = Points;

			if ((Ax1 != null && Ax1.ReverseValues) //This is going to be a very rare thing.
				||
				(Ax2 != null && Ax2.ReverseValues))
			{
				NewPoints = new TPointF[Points.Length];
				for (int i = 0; i < Points.Length; i++)
				{
					NewPoints[i] = new TPointF(tx(Ax1, Ax2, Points[i].X, 0), ty(Ax1, Ax2, Points[i].Y, 0));
				}
			}
			Canvas.DrawLines(aPen, NewPoints);
		}


	}

	/// <summary>
	/// Draws a chart.
	/// </summary>
	internal sealed class DrawChart
    {
        #region Variables
        const real BarLabelOfs = 2.5f;
		const real LineLabelOfs = 5f;
		const real HorizTextLabelMargin = 2.5f; //This value is different in Excel 2007 and 2003. 2 is valid for Excel 2007, and it considers only the "gray box". Excel 2003 uses the gray box + width of the numbers in y-axis to calculate this number.
		internal const real LeaderLinesBreakOfs = 5f;
		private static real OnePix = 0.0005f; //Just approx. Resolution on pdf is 4.
        #endregion

        #region Constructor
        private DrawChart() { }
        #endregion

        #region Brush and Pen Utilities
        internal static Brush GetBrush(RectangleF Coords, TShapeProperties ShProp, ExcelFile Workbook, TShadowInfo ShadowInfo, real Zoom100)
		{
			return DrawShape.GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100);
		}

		private static ChartSeriesOptions GetSeriesOptionsPie(ChartSeriesOptions GlobalOpt, TSeriesOptionsList Opt, int i)
		{
			if (Opt[i] != null && Opt[i].PieOptions != null) return Opt[i];
			if (Opt[-1] != null && Opt[-1].PieOptions != null) return Opt[-1];
			if (GlobalOpt != null && GlobalOpt.PieOptions != null) return GlobalOpt;
			return null;
		}

		private static ChartSeriesOptions GetSeriesOptionsMisc(ChartSeriesOptions GlobalOpt, TSeriesOptionsList Opt, int i)
		{
			if (Opt[i] != null && Opt[i].MiscOptions != null) return Opt[i];
			if (Opt[-1] != null && Opt[-1].MiscOptions != null) return Opt[-1];
			if (GlobalOpt != null && GlobalOpt.MiscOptions != null) return GlobalOpt;
			return null;
		}

		internal static bool PointNeedsCustomBrush(ChartSeriesOptions SeriesOptions)
		{
			return (SeriesOptions != null && SeriesOptions.FillOptions != null && !SeriesOptions.FillOptions.AutomaticColors);
		}

		internal static bool PointNeedsCustomPen(ChartSeriesOptions SeriesOptions)
		{
			return (SeriesOptions != null && SeriesOptions.LineOptions != null); //Here even with automatic colors, we might have a different linestyle. && !SeriesOptions.LineOptions.AutomaticColors);
		}

		internal static bool PointNeedsCustomMarker(ChartSeriesOptions SeriesOptions)
		{
			return (SeriesOptions != null && SeriesOptions.MarkerOptions != null); //Here even with automatic colors, we might have a different MarkerType. && !SeriesOptions.LineOptions.AutomaticColors);
		}


		private static Brush GetCustomBrush(ExcelFile Workbook, RectangleF Coords, ChartFillOptions FillOptions, TShapeOptionList ExtraOptions)
		{
			if (ExtraOptions != null)
			{
				TShapeProperties ExtraProps = new TShapeProperties();
				ExtraProps.ShapeOptions = ExtraOptions;
				Brush R1 = GetBrush(Coords, ExtraProps, Workbook, new TShadowInfo(TShadowStyle.None, 0), 1);
				if (R1 != null) 
				{
					return R1;
				}
			}

			if (FillOptions.Pattern == TChartPatternStyle.None) return null;
			return new SolidBrush(FillOptions.FgColor);
		}

		private static Brush AutoBrush(TChartOptions Options, ExcelFile Workbook, int i)
		{
			if (Options.ChangeColorsOnEachSeries || Options.ChartType != TChartType.Pie)
				return AutomaticBrush(Workbook, i);
			return AutomaticBrush(Workbook, 0);
		}

		internal static Brush GetSeriesBrush(ExcelFile Workbook, RectangleF Coords, TChartOptions Options, ChartSeriesOptions GlobalOptions, ChartSeriesOptions GlobalSeriesOptions, ChartSeriesOptions SeriesOptions, int i)
		{
			if (SeriesOptions != null && SeriesOptions.FillOptions != null)
			{
				if (SeriesOptions.FillOptions.AutomaticColors) return AutoBrush(Options, Workbook, i);
				return GetCustomBrush(Workbook, Coords, SeriesOptions.FillOptions, SeriesOptions.ExtraOptions);
			}

			if (GlobalSeriesOptions != null && GlobalSeriesOptions.FillOptions != null)
			{
				if (GlobalSeriesOptions.FillOptions.AutomaticColors) return AutoBrush(Options, Workbook, i);
				return GetCustomBrush(Workbook, Coords, GlobalSeriesOptions.FillOptions, GlobalSeriesOptions.ExtraOptions);
			}

			if (GlobalOptions != null && GlobalOptions.FillOptions != null)
			{
				if (GlobalOptions.FillOptions.AutomaticColors) return AutoBrush(Options, Workbook, i);
				return GetCustomBrush(Workbook, Coords, GlobalOptions.FillOptions, GlobalOptions.ExtraOptions);
			}

			return AutoBrush(Options, Workbook, i);
		}

		internal static Brush GetFrameBrush(ExcelFile Workbook, RectangleF Coords, TChartFrameOptions FrameOptions, Color DefaultColor)
		{
			if (FrameOptions != null && FrameOptions.FillOptions != null)
			{
				return GetCustomBrush(Workbook, Coords, FrameOptions.FillOptions, FrameOptions.ExtraOptions);
			}

			if (DefaultColor == ColorUtil.Empty) return null;
			return new SolidBrush(DefaultColor);
		}

		internal static Pen GetCustomPen(ExcelFile Workbook, ChartLineOptions LineOptions, TShapeOptionList ExtraOptions)
		{	 
			return GetCustomPen(Workbook, LineOptions, ExtraOptions, Colors.Black);
		}


		private static Pen GetCustomPen(ExcelFile Workbook, ChartLineOptions LineOptions, TShapeOptionList ExtraOptions, Color DefaultColor)
		{	 
			if (ExtraOptions != null)
			{
				TShapeProperties ExtraProps = new TShapeProperties();
				ExtraProps.ShapeOptions = ExtraOptions;
				Pen P1 = GetPen(ExtraProps, Workbook, new TShadowInfo(TShadowStyle.None, 0));
				if (P1 != null) 
				{
					return P1;
				}
			}

			if (LineOptions == null) return new Pen(DefaultColor);
			if (LineOptions.Style == TChartLineStyle.None) 
			{
				return null;
			}

			real width = 1;
			int iWidth = (int) LineOptions.LineWeight + 1;
			if (iWidth > 0 && iWidth < 4) width = iWidth;
			Pen Result = new Pen(LineOptions.LineColor, width);
			switch (LineOptions.Style)
			{
				case TChartLineStyle.Dot:  Result.DashStyle = DashStyles.Dot;break;
				case TChartLineStyle.Dash:  Result.DashStyle = DashStyles.Dash;break;
				case TChartLineStyle.DashDot:  Result.DashStyle = DashStyles.DashDot;break;
				case TChartLineStyle.DashDotDot:  Result.DashStyle = DashStyles.DashDotDot;break;
			}
			return Result;
		}


		private static Pen AutoPen(TChartOptions Options, ExcelFile Workbook, int i, bool IsBorder)
		{
			if (IsBorder)
			{
				return new Pen(Colors.Black);
			}

			if (Options.ChangeColorsOnEachSeries || Options.ChartType != TChartType.Pie)
				return AutomaticPen(Workbook, i);
			return AutomaticPen(Workbook, 0);
		}

		internal static Pen GetSeriesPen(ExcelFile Workbook, TChartOptions Options, ChartSeriesOptions GlobalOptions, ChartSeriesOptions GlobalSeriesOptions, ChartSeriesOptions SeriesOptions, int i, bool IsBorder)
		{
			if (SeriesOptions != null && SeriesOptions.LineOptions != null)
			{
				if (SeriesOptions.LineOptions.AutomaticColors) return AutoPen(Options, Workbook, i, IsBorder);
				return GetCustomPen(Workbook, SeriesOptions.LineOptions, SeriesOptions.ExtraOptions);
			}

			if (GlobalSeriesOptions != null && GlobalSeriesOptions.LineOptions != null)
			{
				if (GlobalSeriesOptions.LineOptions.AutomaticColors) return AutoPen(Options, Workbook, i, IsBorder);
				return GetCustomPen(Workbook, GlobalSeriesOptions.LineOptions, GlobalSeriesOptions.ExtraOptions);
			}

			if (GlobalOptions != null && GlobalOptions.LineOptions != null)
			{
				if (GlobalOptions.LineOptions.AutomaticColors) return AutoPen(Options, Workbook, i, IsBorder);
				return GetCustomPen(Workbook, GlobalOptions.LineOptions, GlobalOptions.ExtraOptions);
			}

			return AutoPen(Options, Workbook, i, IsBorder);
		}


		internal static Pen GetPen(TShapeProperties ShProp, ExcelFile Workbook, TShadowInfo ShadowInfo)
		{
			return DrawShape.GetPen(ShProp, Workbook, ShadowInfo, false);
		}

		internal static Pen GetFramePen(ExcelFile Workbook, TChartFrameOptions FrameOptions, Color DefaultColor)
		{
			if (FrameOptions != null && FrameOptions.LineOptions != null)
			{
				return GetCustomPen(Workbook, FrameOptions.LineOptions, FrameOptions.ExtraOptions);
			}

			if (DefaultColor == ColorUtil.Empty) return null;
			return new Pen(DefaultColor);
		}

		internal static Pen GetAxisMainPen(TAxisLineOptions LineOptions)
		{
			if (LineOptions == null || LineOptions.MainAxis == null) return new Pen(Colors.Black);
			return GetCustomPen(null, LineOptions.MainAxis, null); 
		}

		internal static Pen GetAxisMajorPen(TAxisLineOptions LineOptions)
		{
			if (LineOptions == null || LineOptions.MajorGridLines == null) return new Pen(Colors.Black);
			return GetCustomPen(null, LineOptions.MajorGridLines, null); 
		}
		internal static Pen GetAxisMinorPen(TAxisLineOptions LineOptions)
		{
			if (LineOptions == null || LineOptions.MinorGridLines == null) return new Pen(Colors.Black);
			return GetCustomPen(null, LineOptions.MinorGridLines, null); 
		}

		#endregion

		#region Automatic Colors
		internal static Color AutomaticColor(ExcelFile xls, int index, int offset)
		{
			index = ((index +16+ offset) % xls.ColorPaletteCount) + 1;
			return xls.GetColorPalette(index, Colors.White);
		}

		internal static Brush AutomaticBrush(ExcelFile xls, int index)
		{
			if (index >= xls.ColorPaletteCount * 5) return new HatchBrush(HatchStyle.SmallCheckerBoard, Colors.White, AutomaticColor(xls, index, 0));
			if (index >= xls.ColorPaletteCount * 4) return new HatchBrush(HatchStyle.DarkVertical, Colors.White, AutomaticColor(xls, index, 0));
			if (index >= xls.ColorPaletteCount * 3) return new HatchBrush(HatchStyle.DarkHorizontal, Colors.White, AutomaticColor(xls, index, 0));
			if (index >= xls.ColorPaletteCount * 2) return new HatchBrush(HatchStyle.DarkDownwardDiagonal, Colors.White, AutomaticColor(xls, index, 0));
			if (index >= xls.ColorPaletteCount * 1) return new HatchBrush(HatchStyle.Percent50, Colors.White, AutomaticColor(xls, index, 0));
			return new SolidBrush(AutomaticColor(xls, index, 0));
		}

		internal static Pen AutomaticPen(ExcelFile xls, int index)
		{
			return new Pen(AutomaticColor(xls, index, 8));
		}

		internal static TChartMarkerType AutomaticMarker(int index)
		{
			switch (index % 9)
			{
				case 0: return TChartMarkerType.Diamond;
				case 1: return TChartMarkerType.Square;
				case 2: return TChartMarkerType.Triangle;
				case 3: return TChartMarkerType.X;
				case 4: return TChartMarkerType.Star;
				case 5: return TChartMarkerType.Circle;
				case 6: return TChartMarkerType.Plus;
				case 7: return TChartMarkerType.DowJones;
				default: return TChartMarkerType.StandardDeviation;
			}
		}
		internal static bool MarkerNeedsBackground(TChartMarkerType MarkerType)
		{
			if (MarkerType == TChartMarkerType.X ||
				MarkerType == TChartMarkerType.Star ||
				MarkerType == TChartMarkerType.Plus) return false;

			return true;
		}

		#endregion

		#region Misc Utilities
		internal static string GetNumberFormat(string MainNumberFormat, string BackupNumberFormat, TAxisInfo Axis)
		{
			string NumberFormat = MainNumberFormat;
			if (NumberFormat == null && (Axis == null || !Axis.IsDate)) NumberFormat = BackupNumberFormat;
			if (Axis == null || !Axis.IsDate || (NumberFormat != null && NumberFormat.Length > 0)) return NumberFormat;
			switch (Axis.DateUnits)
			{
				case TDateUnits.Month: return "mmm-yy";
				case TDateUnits.Year: return "yyyy";
				default: return TFlxNumberFormat.RegionalDateString;
			}
		}
		#endregion

		#region Draw
		internal static void Draw(ExcelChart Chart, IFlxGraphics Canvas, ExcelFile Workbook, TFontCache FontCache, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
		{
			Canvas.SaveState();
			try
			{
				DrawBackgroundAndClip(Canvas, Workbook, Chart, ShProp, Coords, ShadowInfo, Zoom100);
				TChartOptions[] ChartOptions = Chart.ChartOptions;
				TChartAxis[] Axis = Chart.GetChartAxis();
				TDataLabel[] DataLabels = Chart.GetDataLabels();

				RectangleF PlotCoords;
				RectangleF ChartCoords = Coords;
				ChartCoords.Inflate(-TDrawObjects.ChartMargin, -TDrawObjects.ChartMargin);
				
				if (Axis != null && Axis.Length > 0)
				{
					Rectangle a = Axis[0].AxisLocation;
					PlotCoords = new RectangleF(ChartCoords.Left + ChartCoords.Width * a.Left / 4000f,
						ChartCoords.Top + ChartCoords.Height * a.Top / 4000f,
						ChartCoords.Width * a.Width / 4000f, ChartCoords.Height * a.Height / 4000f);
				}
				else
				{
					PlotCoords = ChartCoords;
				}

				if (PlotCoords.Height <= 0 || PlotCoords.Width <= 0 || ChartCoords.Height <= 0 || ChartCoords.Width <= 0) return;

				ChartSeries[] SeriesValues = new ChartSeries[Chart.SeriesCount];
				for (int i = 1; i <= SeriesValues.Length; i++)
				{
					SeriesValues[i - 1] = Chart.GetSeries(i, false, true, true);
				}

				int[] MaxX;
				TAxisInfo[] YAxis, XAxis;
				int[] GroupSeriesCount;

				CalcSeriesValues(Canvas, PlotCoords, FontCache, Zoom100, Chart, SeriesValues, ChartOptions, Axis, out MaxX, out YAxis, out GroupSeriesCount);
				CalcCategoriesAxis(Workbook, Canvas, PlotCoords, FontCache, Zoom100, Chart, SeriesValues, ChartOptions, Axis, MaxX, YAxis, out XAxis);
				TLegend Legend = new TLegend(Workbook, Canvas, Chart, ChartCoords, PlotCoords, SeriesValues, FontCache, Zoom100, ChartOptions);
				PlotCoords = Legend.CalcLegend();

				TChartCanvas ChartCanvas = new TChartCanvas(Canvas, PlotCoords, Coords); 

				int an0 = 0;
				while (an0 < ChartOptions.Length && ChartOptions[an0].PlotArea == null && ChartOptions[an0].ChartType != TChartType.Pie) an0++;

				TChartPlotArea PlotArea = null;
				bool FirstIsPie = false;
				if (an0 < ChartOptions.Length)
				{
					PlotArea = ChartOptions[an0].PlotArea;
					FirstIsPie = ChartOptions[an0].ChartType == TChartType.Pie;
				}

				DrawPlotArea1(Workbook, ChartCanvas, PlotArea);  //only the first series draws the plot area
				int an = 0; //ChartOptions[an0].AxisNumber;  //Gridlines always are drawn on primary axis
				if (!FirstIsPie) DrawPlotArea2(Workbook, ChartCanvas, XAxis[an], YAxis[an]);  
				DrawPlotArea3(Workbook, ChartCanvas, PlotArea);  //only the first series draws the plot area

				TLabelDescriptionList LabelDescriptions = new TLabelDescriptionList(DataLabels, ChartOptions);

				Canvas.SaveState();
				try
				{
					Canvas.SetClipReplace(PlotCoords); 
					DrawChartData(ChartCanvas, Workbook, FontCache, ShadowInfo, Clipping, Zoom100, Chart, ChartOptions, PlotCoords, SeriesValues, MaxX, YAxis, XAxis, GroupSeriesCount, LabelDescriptions);
				}
				finally
				{
					Canvas.RestoreState();
				}

				DrawAxis(Workbook, ChartCanvas, ChartCoords, PlotCoords, FontCache, XAxis, YAxis, PlotArea, Zoom100);
				TDataLabels LabelRenderer = new TDataLabels(Workbook, Canvas, Chart, ChartCoords, PlotCoords, FontCache, Zoom100, ChartOptions, SeriesValues, ChartCanvas);
				LabelRenderer.Draw(DataLabels,  LabelDescriptions, XAxis, YAxis);
				for (int i = 0; i < YAxis.Length; i++)
				{
					LabelRenderer.DrawSimpleLabel(XAxis[i].Caption);
					LabelRenderer.DrawSimpleLabel(YAxis[i].Caption);
				}

				Legend.Draw();
			}
			finally
			{
				Canvas.RestoreState();
			}
		}

		#endregion

		#region Background
		private static void DrawBackgroundAndClip(IFlxGraphics Canvas, ExcelFile Workbook, ExcelChart Chart, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, real Zoom100)
		{
			if (Chart.Background != null)
			{
				using (Brush ABrush = GetFrameBrush(Workbook, Coords, Chart.Background, Colors.White))
				{
					Canvas.FillRectangle(ABrush, Coords);
				}
				using (Pen APen = GetFramePen(Workbook, Chart.Background, Colors.Black))
				{
					Canvas.DrawRectangle(APen, Coords.Left, Coords.Top, Coords.Width, Coords.Height);
				}
			}
			else
			{
				using (Brush ABrush = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
				{
					Canvas.FillRectangle(ABrush, Coords);
				}
				using (Pen APen = GetPen(ShProp, Workbook, ShadowInfo))
				{
					Canvas.DrawRectangle(APen, Coords.Left, Coords.Top, Coords.Width, Coords.Height);
				}
			}


			Canvas.SetClipReplace(Coords);
		}
	

		#endregion

		#region Draw Plot Area
		private static void DrawPlotArea1(ExcelFile Workbook, TChartCanvas Canvas, TChartPlotArea BoxStyle)
		{
			if (BoxStyle != null) 
			{
				using (Brush aBrush = GetFrameBrush(Workbook, Canvas.Coords, BoxStyle.ChartFrameOptions, Colors.Silver))
				{
					Canvas.FillRectangle(null, null, aBrush, Canvas.Coords);
				}
			}
		}

		private static void DrawPlotArea2(ExcelFile Workbook, TChartCanvas Canvas, TAxisInfo XAxis, TAxisInfo YAxis)
		{
			List<real> yused = new List<real>();

			DrawGridLines(Canvas, YAxis, XAxis, YAxis.StartOffset, yused, Workbook.OptionsDates1904);

			List<real> xused = new List<real>();
			DrawGridLines(Canvas, XAxis, YAxis, XAxis.StartOffset, xused, Workbook.OptionsDates1904);
		}

		private static void DrawPlotArea3(ExcelFile Workbook, TChartCanvas Canvas, TChartPlotArea BoxStyle)
		{
			if (BoxStyle != null) 
			{
				using (Pen aPen = GetFramePen(Workbook, BoxStyle.ChartFrameOptions, Colors.Black))
				{
					Canvas.Canvas.DrawRectangle(aPen, Canvas.Coords.Left, Canvas.Coords.Top, Canvas.Coords.Width, Canvas.Coords.Height);
				}
			}
		}

		private static void DrawGridLines(TChartCanvas Canvas, TAxisInfo Axis, TAxisInfo OtherAxis, real StartOffset, List<real> yused, bool Dates1904)
		{
			if (Axis.LineOptions != null && Axis.LineOptions.MajorGridLines != null)
			{
				using (Pen majorPen = GetAxisMajorPen(Axis.LineOptions))
				{
                    DrawGrid(Canvas, majorPen, Canvas.Coords, Axis, OtherAxis, Axis.MajorScale, StartOffset, yused, Dates1904, Axis.DateUnitsMajor);
				}
			}

			if (Axis.LineOptions != null && Axis.LineOptions.MinorGridLines != null)
			{
				using (Pen minorPen = GetAxisMinorPen(Axis.LineOptions))
				{	
					DrawGrid(Canvas, minorPen, Canvas.Coords, Axis, OtherAxis, Axis.MinorScale, StartOffset, yused, Dates1904, Axis.DateUnitsMinor);
				}
			}	
		}

		private static void DrawGrid(TChartCanvas Canvas, Pen aPen, RectangleF Coords, TAxisInfo Axis, 
			TAxisInfo OtherAxis, double scale, real StartOffset, List<real> yused, bool Dates1904, TDateUnits TickDateUnits)
		{
			if (scale <= 0) return;

			real Max;
			int sign;
			real dx;
			real x0;
			GetXAndDx(Coords, Axis, scale, StartOffset, out Max, out sign, out dx, out x0);
			if (dx == 0 || Math.Sign(dx) != Math.Sign(sign)) return; //do not allow infinite loop

			real x = x0;
			int i = 0;
			while (x * sign <= Max + OnePix)
			{
				real x2 = (real)(i * scale); //store a more precise representation.

				int index = yused.BinarySearch(x2);
				if (index < 0) 
				{
					yused.Insert(~index, x2);

					Draw90DegreeLine(Canvas, aPen, Coords, x, Axis, OtherAxis);
				}
				i++;

                x = NextTick(ref Coords, Axis, scale, StartOffset, Dates1904, sign, dx, x0, i, TickDateUnits);
			}
		}

        private static real NextTick(ref RectangleF Coords, TAxisInfo Axis, double scale, real StartOffset, bool Dates1904, int sign, real dx, real x0, int i, TDateUnits TickDateUnits)
        {
            if (Axis.IsDate)
            {
                if (Axis.Max > Axis.Min)
                {
                    real DateToPlot = (real)TDateAxisTransform.AddDateUnit(Axis.Min, TickDateUnits, (int)scale * i, Dates1904);
                    real Width = Axis.Horizontal ? Coords.Width : Coords.Height;
                    Width -= 2 * StartOffset;
                    return (real)(x0 + sign * (Width * (DateToPlot - Axis.Min) / (Axis.Max - Axis.Min)));
                }
            }

			return x0 + dx * i; //we do not just accumulate one dx here, to avoid rounding errors.
        }

		private static void GetXAndDx(RectangleF Coords, TAxisInfo Axis, double scale, real StartOffset, out real Max, out int sign, out real dx, out real x0)
		{
			real Width = Axis.Horizontal ? Coords.Width : Coords.Height;
			Max = Axis.Horizontal ? Coords.Right : -Coords.Top;
			sign = Axis.Horizontal ? 1 : -1;
			x0 = Axis.Horizontal ? Coords.Left : Coords.Bottom;

			if (Axis.Max == Axis.Min) 
			{
				dx = (real)(sign * Width / 2);
				x0 += dx;
				dx *= 4;
				return;
			}

			if (Axis.Max < Axis.Min || scale == 0) 
			{
				dx = 0;
				return;
			}

			dx = (real)(sign * (Width - 2 * StartOffset) / (Axis.Max - Axis.Min) * scale);  //y grows from top to bottom.

		}
	
		private static void Draw90DegreeLine(TChartCanvas Canvas, Pen aPen, RectangleF Coords, real y, TAxisInfo Axis, TAxisInfo OtherAxis)
		{
			if (Axis.Horizontal)
			{
				Canvas.DrawLine(Axis, OtherAxis, aPen, y, Coords.Top, y, Coords.Bottom);
			}
			else
			{
				Canvas.DrawLine(Axis, OtherAxis, aPen, Coords.Left, y, Coords.Right, y);
			}
		}
		#endregion

		#region Draw Axis
		private static Brush GetBackBrush(ExcelFile Workbook, RectangleF Coords, TChartPlotArea BoxStyle, TAxisInfo x1, TAxisInfo x2, TAxisInfo x3, TAxisInfo x4)
		{
			if (
				(x1.TickOptions == null || x1.TickOptions.BackgroundMode != TBackgroundMode.Opaque) &&
				(x2.TickOptions == null || x2.TickOptions.BackgroundMode != TBackgroundMode.Opaque) &&
				(x3.TickOptions == null || x3.TickOptions.BackgroundMode != TBackgroundMode.Opaque) &&
				(x4.TickOptions == null || x4.TickOptions.BackgroundMode != TBackgroundMode.Opaque)
				) return null; 

			return GetFrameBrush(Workbook, Coords, BoxStyle.ChartFrameOptions, Colors.Silver);
		}

		private static void DrawAxis(ExcelFile Workbook, TChartCanvas Canvas, RectangleF ChartCoords, RectangleF PlotCoords, TFontCache FontCache, TAxisInfo[] XAxis, TAxisInfo[] YAxis, TChartPlotArea BoxStyle, real Zoom100)
		{
			using (Brush BackBrush = GetBackBrush(Workbook, PlotCoords, BoxStyle, XAxis[0], XAxis[1], YAxis[0], YAxis[1]))
			{			
				DrawAxis(Workbook, Canvas, ChartCoords, PlotCoords, FontCache, YAxis, XAxis, BackBrush, Zoom100);
				DrawAxis(Workbook, Canvas, ChartCoords, PlotCoords, FontCache, XAxis, YAxis, BackBrush, Zoom100);
			}
		}

		private static real CalcAxisPos(RectangleF PlotCoords, TAxisInfo Axis, TAxisInfo OtherAxis, out bool LabelAtLeft)
		{
			LabelAtLeft = true;
			if (OtherAxis.MaxCross)
			{
				LabelAtLeft = false;
				return Axis.Horizontal? PlotCoords.Top: PlotCoords.Right;
			}

			double d = OtherAxis.Min;
			if (OtherAxis.RangeOptions != null && OtherAxis.RangeOptions.ValueAxisBetweenCategories) d-= 0.5;
			real a = OtherAxis.Max == OtherAxis.Min?
				0:  
				(real)((OtherAxis.CrossPoint - d) / (OtherAxis.Max - OtherAxis.Min));

			real Min = Axis.Horizontal? PlotCoords.Top: PlotCoords.Left;
			real Result = Axis.Horizontal? PlotCoords.Bottom - (PlotCoords.Height - OtherAxis.StartOffset * 2) * a + OtherAxis.StartOffset:
				PlotCoords.Left + (PlotCoords.Width - OtherAxis.StartOffset * 2) * a - OtherAxis.StartOffset;
			if (Result < Min) Result = Min;
			return Result;
		}

		private static void DrawAxis(ExcelFile Workbook, TChartCanvas Canvas, RectangleF ChartCoords, RectangleF PlotCoords, TFontCache FontCache, TAxisInfo[] Axis, TAxisInfo[] OtherAxis, Brush BackBrush, real Zoom100)
		{
			for (int an = 0; an < Axis.Length; an++)
			{
				real xy; 
				bool LabelAtLeft;
				
				xy = CalcAxisPos(PlotCoords, Axis[an], OtherAxis[an], out LabelAtLeft);				

				Font aFont = FontCache.GetFont(Axis[an].Font.Font, Zoom100);
				SizeF tzero = Canvas.MeasureString("M", aFont);

				DrawAxisLine(Canvas, PlotCoords, Axis[an], OtherAxis[an], xy);
				DrawAxisTicks(Canvas, PlotCoords, Axis[an], OtherAxis[an], xy, tzero, LabelAtLeft, Axis[an].StartOffset, Workbook.OptionsDates1904);

                double Scale = Axis[an].MajorScale;
                if (Axis[an].IsCategory && !Axis[an].IsDate) Scale = 1;

				if (Scale == 0) continue;

				real Max; int sign; real dx; real x0;
				GetXAndDx(PlotCoords, Axis[an], Scale, Axis[an].StartOffset, out Max, out sign, out dx, out x0);
				if (dx == 0 || Math.Sign(dx) != Math.Sign(sign)) return; //do not allow infinite recursion

				real yStartOfsPix = 0;
				if (Axis[an].Max > Axis[an].Min) //if not, this is a single point.
				{
					if (Axis[an].IsDate)
					{
						yStartOfsPix = Axis[an].StartOffset * sign;
					}
					else
					{
						if ((Axis[an].RangeOptions != null && Axis[an].RangeOptions.ValueAxisBetweenCategories)) 
							yStartOfsPix = Axis[an].StartOffset * sign;
					}
				}

				real y = x0 + yStartOfsPix;
				
				Canvas.Canvas.SaveState();
				int i = 0;
				
                while (y * sign <= Max + OnePix)
				{
					real y1 = NextTick(ref PlotCoords, Axis[an], Scale, Axis[an].StartOffset, Workbook.OptionsDates1904, sign, dx, x0 + yStartOfsPix, i + 1, Axis[an].DateUnitsMajor);

					DrawOneLabel(Workbook, Canvas, FontCache, Zoom100, ChartCoords, ref PlotCoords, Axis[an], OtherAxis[an], BackBrush, xy, LabelAtLeft, aFont, ref tzero, y, i, Scale, y1 - y);
					i++;
					y = y1;
				}

				Canvas.Canvas.RestoreState();
					
			}
		}

		private static void DrawOneLabel(ExcelFile Workbook, TChartCanvas Canvas, TFontCache FontCache, real Zoom100, RectangleF ChartCoords, ref RectangleF PlotCoords, TAxisInfo Axis, TAxisInfo OtherAxis, Brush BackBrush, real xy, bool LabelAtLeft, Font aFont, ref SizeF tzero, real y, int i, double Scale, real dx)
		{
			if (Axis.TickOptions == null || Axis.TickOptions.LabelPosition == TAxisLabelPosition.None) return; //do not draw this axis.
			if (Axis.LineOptions != null && Axis.LineOptions.DoNotDrawLabelsIfNotDrawingAxis && Axis.LineOptions.MainAxis.Style == TChartLineStyle.None) return;

			bool LabelReallyAtLeft = LabelAtLeft ^ OtherAxis.ReverseValues;
			int lf = 1;
			if (Axis.RangeOptions != null) lf = Axis.RangeOptions.LabelFrequency;
			if (lf > 0 && (i % lf) != 0) return;

            Color FontColor;
            string ylabel;
            GetCatAxisLabel(Workbook, Axis, i, Scale, out FontColor, out ylabel);

			int Rotation = 0;
			THFlxAlignment HJustify = THFlxAlignment.center;
			TVFlxAlignment VJustify = LabelAtLeft? TVFlxAlignment.top: TVFlxAlignment.bottom; 
			if (!Axis.Horizontal)
			{
				HJustify = LabelAtLeft? THFlxAlignment.right: THFlxAlignment.left;
				VJustify = TVFlxAlignment.center;
			}

			bool HAlignGeneral;

			if (Axis.TickOptions != null)
			{
				Rotation = Axis.TickOptions.Rotation;
			}
			
			bool Vertical;
			real ddx = Math.Abs(dx);
			real lx = ddx / 2;
			real Alpha = FlexCelRender.CalcAngle(Rotation, out Vertical);
			
			if (Axis.Horizontal)
			{
				if (Alpha != 90 && Alpha != -90)
				{
					bool GoesRight = (Alpha > 0 && LabelReallyAtLeft) || (Alpha < 0 && !LabelReallyAtLeft);
					bool GoesLeft = (Alpha > 0 && !LabelReallyAtLeft) || (Alpha < 0 && LabelReallyAtLeft);
					if ((GoesRight && !Axis.ReverseValues) || (GoesLeft && Axis.ReverseValues))
					{
						HJustify = THFlxAlignment.right;
						lx = ddx;
					}
					else if (((GoesRight && Axis.ReverseValues) || (GoesLeft && !Axis.ReverseValues)))
					{
						HJustify = THFlxAlignment.left;
						lx = 0;
					}
				}
				else
				{
					if (!LabelAtLeft) VJustify = TVFlxAlignment.bottom;
				}
			}

			else

			{
				if (Alpha != 90 && Alpha != -90)
				{
					bool GoesUp = (Alpha > 0 && LabelReallyAtLeft) || (Alpha < 0 && !LabelReallyAtLeft);
					bool GoesDown = (Alpha > 0 && !LabelReallyAtLeft) || (Alpha < 0 && LabelReallyAtLeft);
					if ((GoesUp && !Axis.ReverseValues) || (GoesDown && Axis.ReverseValues))
					{
						VJustify = TVFlxAlignment.top;
						lx = 0;
					}
					else if (((GoesUp && Axis.ReverseValues) || (GoesDown && !Axis.ReverseValues)))
					{
						VJustify = TVFlxAlignment.bottom;
						lx = ddx;
					}
				}
				else
				{
					if (!LabelAtLeft) HJustify = THFlxAlignment.right;
				}
			}

			real SinAlpha = (real)Math.Sin(Alpha * Math.PI / 180); real CosAlpha = (real)Math.Cos(Alpha * Math.PI / 180);
			SizeF TextExtent;
			TXRichStringList TextLines;
			TFloatList MaxDescent;

			THAlign HAlign = THAlign.Center ; TVAlign VAlign = TVAlign.Center;
			FlexCelRender.GetHJustify(HJustify, ref HAlign, out HAlignGeneral);
			FlexCelRender.GetVJustify(VJustify, Alpha, ref VAlign);
			
			RectangleF CellRect;
			
			if (Axis.Horizontal)
			{
				real xxy = xy;
				Canvas.ty(OtherAxis, ref xxy, 0);
				real Top = xy + tzero.Width;
				real Height = LabelReallyAtLeft? ChartCoords.Bottom - (xxy + tzero.Width): xxy - tzero.Width - ChartCoords.Top;
				CellRect = new RectangleF(y - lf * lx, Top, lf * ddx, Height);
			}
			else
			{
				real xxy = xy;
				Canvas.tx(OtherAxis, ref xxy, 0);
				real Width = LabelReallyAtLeft? xxy - tzero.Width - ChartCoords.Left: ChartCoords.Right - xxy - tzero.Width;
				CellRect = new RectangleF(xy + tzero.Width , y - lf * lx, Width, lf * ddx);
			}

			TRichString Text = new TRichString(ylabel);

			TextPainter.CalcTextBox(Canvas.Canvas, FontCache, Zoom100, CellRect, 0, true, Alpha, Vertical, Text, aFont, null, out TextExtent, out TextLines, out MaxDescent);
			if (TextLines.Count <= 0) return;

		
			real[] XX;
			real[] YY;
			RectangleF ContainingRect = TextPainter.CalcTextCoords(out XX, out YY, Text, VAlign, ref HAlign, 0, Alpha, CellRect, 0, TextExtent, HAlignGeneral, Vertical, SinAlpha, CosAlpha, TextLines, Workbook.Linespacing, VJustify);

			if (Axis.Horizontal && !(LabelAtLeft))
			{
				TextPainter.RelocateBox(ref ContainingRect, XX, YY, 0, -CellRect.Height - 2 * tzero.Height);
			}

			if (!Axis.Horizontal && (LabelAtLeft))
			{
				TextPainter.RelocateBox(ref ContainingRect, XX, YY, -CellRect.Width - 2 * tzero.Width, 0);
			}

			if (OtherAxis.ReverseValues || Axis.ReverseValues) 
			{
				Canvas.MirrorXY(Axis, OtherAxis, ref ContainingRect, ref XX, ref YY, TextLines, SinAlpha, CosAlpha);
			}
			
			if (Axis.TickOptions != null && Axis.TickOptions.BackgroundMode == TBackgroundMode.Opaque)
			{
				if (ContainingRect.Right <= PlotCoords.Left || ContainingRect.Left >= PlotCoords.Right || ContainingRect.Top >= PlotCoords.Bottom || ContainingRect.Bottom <= PlotCoords.Top)
				{
				}
				else
				{
					Canvas.Canvas.FillRectangle(BackBrush, ContainingRect);
				}
			}

			TextPainter.DrawRichText(Workbook, Canvas.Canvas, FontCache, Zoom100, true, ref ContainingRect, ref ChartCoords, ref ContainingRect, ref ContainingRect, 0, HJustify, VJustify, Alpha,
				FontColor, new TSubscriptData(TFlxFontStyles.None), Text, TextExtent, TextLines, aFont, MaxDescent, XX, YY);
			
		}

        internal static void GetCatAxisLabel(ExcelFile Workbook, TAxisInfo Axis, int i, double Scale, out Color FontColor, out string ylabel)
        {
            ylabel = null;
            FontColor = Colors.Black;
            if (Axis == null) Axis = new TAxisInfo(1, 1, 1, 1, false, false, false);
            if (Axis.TickOptions != null) FontColor = Axis.TickOptions.LabelColor;


            string NumberFormat = GetNumberFormat(Axis.NumberFormat, Axis.CellNumberFormat, Axis);

            if (Axis.Labels == null)
            {
                double val = Axis.Min + i * Scale;
                if (Axis.IsDate)
                {
                    val = (real)TDateAxisTransform.AddDateUnit(Axis.Min, Axis.DateUnitsMajor, (int)Scale * i, Workbook.OptionsDates1904);
                }
                TRichString Richylabel = TFlxNumberFormat.FormatValue(val, NumberFormat, ref FontColor, Workbook);
                ylabel = Richylabel == null ? String.Empty : Richylabel.ToString();
                if (Axis.NeedsPercent) ylabel += "%";
            }
            else
            {
                int iLabel = (int)(i * Scale);
                if (iLabel >= Axis.Labels.Length) ylabel = String.Empty; else ylabel = FlxConvert.ToString(TFlxNumberFormat.FormatValue(Axis.Labels[iLabel], NumberFormat, ref FontColor, Workbook));
            }
        }		

		private static void DrawAxisLine(TChartCanvas Canvas, RectangleF Coords, TAxisInfo Axis, TAxisInfo OtherAxis, real xy)
		{
			if (Axis.LineOptions != null) //here null means default color.
			{
				using (Pen mainPen = GetAxisMainPen(Axis.LineOptions))
				{
					Draw90DegreeLine(Canvas, mainPen, Coords, xy, OtherAxis, Axis);
				}
			}

		}

		private static RectangleF CalcTicksRect(TTickType TickType, RectangleF Coords, bool Horizontal, real ypos, real Height, bool LabelAtLeft)
		{
			real h1 = 0;
			real h2 = 0;
			switch (TickType)
			{
				case TTickType.Inside: h1 = -Height;break;
				case TTickType.Outside: h2 = Height;break;
				case TTickType.Cross: h1 = -Height; h2 = Height; break;
			}

			if (!LabelAtLeft)
			{
				real h3 = h2;
				h2 = -h1;
				h1 = -h3;
			}

			if (Horizontal)
			{
				return RectangleF.FromLTRB(Coords.Left, ypos + h1, Coords.Right, ypos + h2);
			}
			else				
			{
				return RectangleF.FromLTRB(ypos - h1, Coords.Top, ypos - h2, Coords.Bottom);
			}

		}
		private static void DrawAxisTicks(TChartCanvas Canvas, RectangleF Coords, TAxisInfo Axis, TAxisInfo OtherAxis, real xy, SizeF tzero, bool LabelAtLeft, real StartOffset, bool Dates1904)
		{
			if (Axis.TickOptions == null) return;

			List<real> yused = new List<real>();
			if (Axis.TickOptions.MajorTickType != TTickType.None || Axis.TickOptions.MinorTickType != TTickType.None)
			{
				RectangleF MajorCoords = CalcTicksRect(Axis.TickOptions.MajorTickType, Coords, Axis.Horizontal, xy, (real) (tzero.Height / 3.0), LabelAtLeft);
				RectangleF MinorCoords = CalcTicksRect(Axis.TickOptions.MinorTickType, Coords, Axis.Horizontal, xy, (real) (tzero.Height / 5.0), LabelAtLeft);
				
				using (Pen mainPen = GetAxisMainPen(Axis.LineOptions))
				{
					if (Axis.TickOptions.MajorTickType != TTickType.None) DrawGrid(Canvas, mainPen, MajorCoords, Axis, OtherAxis, Axis.MajorScale, StartOffset, yused, Dates1904, Axis.DateUnitsMajor);
                    if (Axis.TickOptions.MinorTickType != TTickType.None) DrawGrid(Canvas, mainPen, MinorCoords, Axis, OtherAxis, Axis.MinorScale, StartOffset, yused, Dates1904, Axis.DateUnitsMinor);
				}
			}
		}

		#endregion

		#region Draw Chart Data
		private static void DrawChartData(TChartCanvas Canvas, ExcelFile Workbook, TFontCache FontCache, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100, ExcelChart Chart, TChartOptions[] ChartOptions, RectangleF PlotCoords, ChartSeries[] SeriesValues, int[] MaxX, TAxisInfo[] YAxis, TAxisInfo[] XAxis, int[] GroupSeriesCount, TLabelDescriptionList LabelDescriptions)
		{
			THiLoData[] HiLoData = new THiLoData[YAxis.Length];
			TMarkerCache Markers = new TMarkerCache();
			for (int i = 0; i < HiLoData.Length; i++) HiLoData[i] = new THiLoData();
			
			DrawChartPoints(Canvas, Workbook, FontCache, ShadowInfo, Clipping, Zoom100, Chart, ChartOptions, PlotCoords, SeriesValues, MaxX, YAxis, XAxis, GroupSeriesCount, LabelDescriptions, HiLoData, Markers);

			for (int i = 0; i < YAxis.Length; i++)
			{
				DrawHiLoLines(Workbook, Canvas, HiLoData[i], XAxis[i], YAxis[i]);
			}

			for (int i = 0; i < YAxis.Length; i++)
			{
				DrawUpDownBars(Workbook, Canvas, HiLoData[i], XAxis[i], YAxis[i]);
			}

			for (int i = 0; i < Markers.Count; i++)
			{
				DrawMarkers(Workbook, Canvas, Markers[i].SeriesValues, Markers[i].Points, Markers[i].SeriesOptions, Markers[i].XAxis, Markers[i].YAxis, null, Zoom100);
			}

		}
		
		private static void DrawChartPoints(TChartCanvas Canvas, ExcelFile Workbook, TFontCache FontCache, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100, ExcelChart Chart, TChartOptions[] ChartOptions, RectangleF PlotCoords, ChartSeries[] SeriesValues, int[] MaxX, TAxisInfo[] YAxis, TAxisInfo[] XAxis, 
			int[] GroupSeriesCount, TLabelDescriptionList LabelDescriptions, THiLoData[] HiLoData, TMarkerCache Markers)
		{
			int[] GroupSeriesIndex = new int[GroupSeriesCount.Length];
			object[] FirstCategory = null;
			for (int Pass = 0; Pass < GroupSeriesCount.Length; Pass++)
			{
				double[] StackedOfsPos = null;
				double[] StackedOfsNeg = null;
				TPointF[] LastPoints = null;
				double PieR1 = 0;
                //object[] FirstCategory = null; doesn't go here. Even if we have 2 x axis, Excel uses the first.
				for (int k = 0; k < SeriesValues.Length; k++)
				{
					int ib = SeriesValues[k].ChartOptionsIndex;
					int an = ChartOptions[ib].AxisNumber;

					if (ib != Pass) continue;
					//int xIndex = 0;
					//while (xIndex < MaxX.Length - 1 && MaxX[xIndex] == 0) xIndex++;
					int xIndex = an;

					switch (ChartOptions[ib].ChartType)
					{
						case TChartType.Area:
							if (YAxis[an].MajorScale <= 0) continue;
							if (((TAreaLineChartOptions)ChartOptions[ib]).StackedMode != TStackedMode.None)
							{
								if (StackedOfsPos == null) StackedOfsPos = new double[MaxX[xIndex]];
							}

							DrawLineChart(Workbook, Chart, SeriesValues[k], Workbook, Canvas, PlotCoords,
								(TAreaLineChartOptions)ChartOptions[ib], XAxis[xIndex], YAxis[an], Zoom100, true, MaxX[xIndex], StackedOfsPos, ref LastPoints, LabelDescriptions.GetLabel(k, ib), null, 
                                k == 0, k == SeriesValues.Length - 1, Markers, ref FirstCategory);
							break;
						case TChartType.Bar:
							if (YAxis[an].MajorScale <= 0) continue;
							if (((TBarChartOptions)ChartOptions[ib]).StackedMode != TStackedMode.None)
							{
								if (StackedOfsPos == null) StackedOfsPos = new double[MaxX[xIndex]];
								if (StackedOfsNeg == null) StackedOfsNeg = new double[MaxX[xIndex]];
							}

							DrawBarChart(Chart, SeriesValues[k], GroupSeriesIndex[ib], GroupSeriesCount[ib], Workbook, Canvas, PlotCoords,
								(TBarChartOptions)ChartOptions[ib], XAxis[xIndex], YAxis[an], Zoom100, MaxX[xIndex], StackedOfsPos, StackedOfsNeg, LabelDescriptions.GetLabel(k, ib));
							break;
						case TChartType.Line:
							if (YAxis[an].MajorScale <= 0) continue;
							if (((TAreaLineChartOptions)ChartOptions[ib]).StackedMode != TStackedMode.None)
							{
								if (StackedOfsPos == null) StackedOfsPos = new double[MaxX[xIndex]];
							}
							TPointF[] LastPoints2 = null;

                            DrawLineChart(Workbook, Chart, SeriesValues[k], Workbook, Canvas, PlotCoords,
								(TAreaLineChartOptions)ChartOptions[ib], XAxis[xIndex], YAxis[an], Zoom100, false, 
                                MaxX[xIndex], StackedOfsPos, ref LastPoints2, LabelDescriptions.GetLabel(k, ib), 
                                HiLoData[an], k == 0, k == SeriesValues.Length - 1, Markers, ref FirstCategory);
							break;
						case TChartType.Pie:
							Canvas.Canvas.RestoreState(); //Pies are not clipped.
							DrawPieChart(SeriesValues[k], GroupSeriesIndex[ib], GroupSeriesCount[ib], Workbook, Chart, Canvas, FontCache, PlotCoords, ShadowInfo, Clipping, (TPieChartOptions)ChartOptions[ib], Zoom100, ref PieR1, LabelDescriptions.GetLabel(k, ib));
							Canvas.Canvas.SaveState();
							Canvas.Canvas.SetClipReplace(PlotCoords);
							break;
						case TChartType.Radar:
							break;
						case TChartType.Scatter:
							if (YAxis[an].MajorScale <= 0) continue;
							if (XAxis[an].MajorScale <= 0) continue;

							DrawScatterChart(Chart, SeriesValues[k], Workbook, Canvas, PlotCoords,
								(TScatterChartOptions)ChartOptions[ib], XAxis[an], YAxis[an], Zoom100, LabelDescriptions.GetLabel(k, ib), Markers);

							break;
						case TChartType.Surface:
							break;
					}

					GroupSeriesIndex[ib]++;
				}
			}
		}

		#endregion

		#region Calculate Axis
		private static bool IsDoubleStacked(TChartType ChartType)
		{
			return ChartType == TChartType.Bar;
		}

		private static void StackSeries(ChartSeries[] SeriesValues, TChartOptions[] ChartOptions, int[] MaxX, TAxisInfo[] YAxis, bool[] FirstMaxY, bool[] FirstMinY, double[][] MinVal, double[][] MaxVal)
		{
			double[][] maxposarray = new double[ChartOptions.Length][];

			for (int i = 1; i <= SeriesValues.Length; i++)
			{
				int OpIn = SeriesValues[i - 1].ChartOptionsIndex;
				TChartOptions Options = ChartOptions[OpIn];
				if (Options.ChartType == TChartType.Pie) continue; //Pie charts do not have axis.

				IStackedOptions Io = Options as IStackedOptions;
				int an = Options.AxisNumber;
				if (Io != null && Io.StackedMode == TStackedMode.Stacked100)
				{
					if (!IsDoubleStacked(Options.ChartType) && maxposarray[OpIn] == null) maxposarray[OpIn] = new double[MaxX[an]];
					for (int k = 0; k < MaxX[an]; k++)
					{
						if (k >= SeriesValues[i - 1].DataValues.Length) break;

						double neg = MinVal[OpIn] == null? 0: MinVal[OpIn][k];
						double pos = MaxVal[OpIn] == null? 0: MaxVal[OpIn][k];
						double GroupMax = pos - neg;
						double d;
						if (!TBaseParsedToken.GetDouble(SeriesValues[i - 1].DataValues[k], out d)) d = 0;
						if (GroupMax > 0)
						{
							double newd = 100 * d /GroupMax;
							SeriesValues[i - 1].DataValues[k] = newd;

							double maxpos; double minpos;
							if (IsDoubleStacked(Options.ChartType))
							{
								maxpos = 100 * pos / GroupMax;
								minpos = 100 * neg / GroupMax;
							}
							else
							{
								maxposarray[OpIn][k] += newd;
								maxpos = maxposarray[OpIn][k];
								minpos = maxpos;
							}

							if (FirstMaxY[an] || maxpos > YAxis[an].Max)
							{
								YAxis[an].Max = maxpos;
								FirstMaxY[an] = false;
							}
							if (FirstMinY[an] || minpos < YAxis[an].Min)
							{
								YAxis[an].Min = minpos;
								FirstMinY[an] = false;
							}
						}

					}
				}
			}
		}

		private static void CalcSeriesValues(IFlxGraphics Canvas, RectangleF PlotCoords, TFontCache FontCache, real Zoom100, ExcelChart Chart, ChartSeries[] SeriesValues, TChartOptions[] ChartOptions, TChartAxis[] Axis, out int[] MaxX, out TAxisInfo[] YAxis, out int[] GroupSeriesCount)
		{
			MaxX = new int[2];
			YAxis = new TAxisInfo[2];
			bool[] FirstMaxY = new bool[YAxis.Length];
			bool[] FirstMinY = new bool[YAxis.Length];

			for (int i = 0; i < YAxis.Length; i++)
			{
				YAxis[i] = new TAxisInfo(0,0,0,0, false, false, false);
				FirstMaxY[i] = true;
				FirstMinY[i] = true;
			}
			GroupSeriesCount = new int[ChartOptions.Length];
			double[][] StackedPos = new double[ChartOptions.Length][];
			double[][] StackedNeg = new double[ChartOptions.Length][];
			double[][] Stacked;
			double[][] MinVal = new double[ChartOptions.Length][];
			double[][] MaxVal = new double[ChartOptions.Length][];


			//First calculate max and min for each axis. First pass calculate maxx, second maxy.
			//Stacked go by chart group. Each chart group stacks independently.
			for (int Pass = 0; Pass < 2; Pass++)
			{
				for (int i = 1; i <= SeriesValues.Length; i++)
				{
					int OpIn = SeriesValues[i - 1].ChartOptionsIndex;
					if (Pass == 0) GroupSeriesCount[OpIn]++;
					TChartOptions Options = ChartOptions[OpIn];
					if (Options.ChartType == TChartType.Pie) continue; //Pie charts do not have axis.
				
					int an = Options.AxisNumber;
				
					IStackedOptions Io = Options as IStackedOptions;
					TStackedMode StackedMode = Io==null? TStackedMode.None: Io.StackedMode;
					if (StackedMode == TStackedMode.Stacked100) YAxis[an].NeedsPercent = true; else YAxis[an].OneNotPercent = true;

					if (IsDoubleStacked(Options.ChartType)) 
					{
						TBarChartOptions bc = Options as TBarChartOptions;
						if (bc.Horizontal) YAxis[an].Horizontal = true;
					}

					if (SeriesValues[i - 1].DataValues != null)
					{
						if (Pass == 0)
						{
							if (SeriesValues[i - 1].DataValues.Length > MaxX[an]) MaxX[an] = SeriesValues[i - 1].DataValues.Length;
						}
						else
						{
							for (int k = 0; k < SeriesValues[i - 1].DataValues.Length; k++)
							{
								if (SeriesValues[i - 1].DataValues[k] is double)
								{
									double d = Convert.ToDouble(SeriesValues[i - 1].DataValues[k], CultureInfo.CurrentCulture);
									if (d < 0 && IsDoubleStacked(Options.ChartType)) Stacked = StackedNeg; else Stacked = StackedPos; //Lines have only one stack.
									if (IsDoubleStacked(Options.ChartType))
									{
										MaxVal[OpIn] = StackedPos[OpIn];
										MinVal[OpIn] = StackedNeg[OpIn];
									}
									else if (StackedMode == TStackedMode.Stacked100) //we only need minval and maxval in line/area charts stacked at 100%
									{
										if (MaxVal[OpIn] == null) MaxVal[OpIn] = new double[MaxX[an]];
										if (MinVal[OpIn] == null) MinVal[OpIn] = new double[MaxX[an]];
										if (d>0) MaxVal[OpIn][k]+=d; else MinVal[OpIn][k]+=d;
									}

									if (Stacked[OpIn] == null) Stacked[OpIn] = new double[MaxX[an]];
									d += Stacked[OpIn][k];
									if  (StackedMode != TStackedMode.None) Stacked[OpIn][k] = d;
									if (StackedMode != TStackedMode.Stacked100)
									{
										if (FirstMaxY[an] || d > YAxis[an].Max)
										{
											YAxis[an].Max = d;
											FirstMaxY[an] = false;
										}
										if (FirstMinY[an] || d < YAxis[an].Min)
										{
											YAxis[an].Min = d;
											FirstMinY[an] = false;
										}
									}
								}
							}
						}
					}
				}
			}

			//Now scale the 100% stacked.
			//100% is equivalent with 100, not with 1.

			StackSeries(SeriesValues, ChartOptions, MaxX, YAxis, FirstMaxY, FirstMinY, MinVal, MaxVal);
	
			//Calculate the real axis values.
			for (int i = 0; i < YAxis.Length; i++)
			{
				bool AutoMin = true;
				bool AutoMax = true;
				bool AutoMajor = true;
				bool AutoMinor = true;

				if (i < Axis.Length)
				{
					TChartAxis Ax = Axis[i];
					CalcOneValueAxis(false, SeriesValues, ChartOptions, i, YAxis[i], ref AutoMin, ref AutoMax, ref AutoMinor, ref AutoMajor, Ax.ValueAxis);
				}
				
				if (YAxis[i].Font == null)
				{
					YAxis[i].Font = (TFlxChartFont)Chart.DefaultAxisFont.Clone();
					YAxis[i].Font.Font.Size20 = (int)Math.Round(Chart.DefaultAxisFont.Font.Size20 * Chart.DefaultAxisFont.Scale);
				}

				CalcValueAxisDefaults(Canvas, PlotCoords, FontCache, Zoom100, YAxis[i], AutoMin, AutoMax, AutoMinor, AutoMajor);
			}
		}


		private static void CalcCategoriesAxis(ExcelFile Workbook, IFlxGraphics Canvas, RectangleF PlotCoords, TFontCache FontCache, real Zoom100, ExcelChart Chart, ChartSeries[] SeriesValues, TChartOptions[] ChartOptions, TChartAxis[] Axis, int[] MaxX, TAxisInfo[] YAxis, out TAxisInfo[] XAxis)
		{
			XAxis = new TAxisInfo[2];
			//Calculate the real axis values.
			for (int i = 0; i < XAxis.Length; i++)
			{
				XAxis[i] = new TAxisInfo(0,0,0,0, false, false, true);
				XAxis[i].IsCategory = true; //might change later.

				bool AutoMin = true;
				bool AutoMax = true;
				bool AutoMajor = true;
				bool AutoMinor = true;

				if (i < Axis.Length)
				{
					TValueAxis ValAxis = Axis[i].CategoryAxis as TValueAxis;
					if (ValAxis != null)
					{
						CalcOneValueAxis(true, SeriesValues, ChartOptions, i, XAxis[i], ref AutoMin, ref AutoMax, ref AutoMinor, ref AutoMajor, ValAxis);
					}
					
					TCategoryAxis CatAxis = Axis[i].CategoryAxis as TCategoryAxis;
					if (CatAxis != null)
					{
						CalcOneCategoryAxis(Workbook, SeriesValues, ChartOptions, i, XAxis[i], ref AutoMin, ref AutoMax, ref AutoMinor, ref AutoMajor, CatAxis, Workbook.OptionsDates1904);
					}
				
				}
				else
				{
					AutoMin = false;
					AutoMax = false;
					AutoMinor = false;
					AutoMajor = false;
				}
				
				if (XAxis[i].Font == null)
				{
					XAxis[i].Font = (TFlxChartFont) Chart.DefaultAxisFont.Clone();
					XAxis[i].Font.Font.Size20 = (int)Math.Round(Chart.DefaultAxisFont.Font.Size20 * Chart.DefaultAxisFont.Scale);
				}

				XAxis[i].Horizontal = !YAxis[i].Horizontal;

				if (XAxis[i].IsCategory && !XAxis[i].IsDate)
				{
					CalcCategoryAxisDefaults(PlotCoords, XAxis[i], i, SeriesValues, ChartOptions, MaxX, AutoMin, AutoMax, AutoMinor, AutoMajor);
				}
				else
				{
                    TDateAxisTransform DateAxisTransform = new TDateAxisTransform(Workbook);
                    CalcValueAxisLimits(XAxis[i], SeriesValues, ChartOptions, i, AutoMin, AutoMax, DateAxisTransform);
                    if (XAxis[i].IsDate)
                    {
                        CalcDateAxisDefaults(Workbook, Canvas, FontCache, Zoom100, XAxis[i], SeriesValues, PlotCoords, AutoMin, AutoMax, AutoMinor, AutoMajor);
                    }
                    else
                    {
                        CalcValueAxisDefaults(Canvas, PlotCoords, FontCache, Zoom100, XAxis[i], AutoMin, AutoMax, AutoMinor, AutoMajor);
                    }
				}

			}
			
		}

		private static void CalcValueAxisLimits(TAxisInfo XAxis, ChartSeries[] SeriesValues, TChartOptions[] ChartOptions, 
            int AxisNumber, bool AutoMin, bool AutoMax, TDateAxisTransform DateAxisTransform)
		{
			bool FirstMaxX = true;
			bool FirstMinX = true;

			for (int i = 1; i <= SeriesValues.Length; i++)
			{
				int OpIn = SeriesValues[i - 1].ChartOptionsIndex;
				TChartOptions Options = ChartOptions[OpIn];
				if (Options.ChartType == TChartType.Pie) continue; //Pie charts do not have axis.

				int an = Options.AxisNumber;
				if (an != AxisNumber) continue;

				object[] vals = SeriesValues[i - 1].CategoriesValues;
				if (vals != null)
				{
					for (int k = 0; k < vals.Length; k++)
					{
						if (vals[k] is double)
						{
                            double d = Convert.ToDouble(vals[k], CultureInfo.CurrentCulture);
                            if (XAxis.IsDate) d = DateAxisTransform.ConvertDateUnit(d, XAxis.DateUnits);

							if (AutoMax && (FirstMaxX || d > XAxis.Max))
							{
								XAxis.Max = d;
								FirstMaxX = false;
							}
							if ((AutoMin && FirstMinX || d < XAxis.Min))
							{
								XAxis.Min = d;
								FirstMinX = false;
							}
						}
					}
				}
			}
		}

		private static void CalcOneValueAxis(bool FormatsFromCategories, ChartSeries[] SeriesValues, TChartOptions[] ChartOptions, int AxisNumber, TAxisInfo YAxis, ref bool AutoMin, ref bool AutoMax, ref bool AutoMinor, ref bool AutoMajor, TValueAxis ValueAxis)
		{
			YAxis.IsCategory = false;
			YAxis.Font = ValueAxis.Font;
			YAxis.NumberFormat = ValueAxis.NumberFormat;

			int an0 = GetFirstSeries(SeriesValues, ChartOptions, AxisNumber);
			if (an0 >= 0)
			{
				string[] CatFmt = FormatsFromCategories? SeriesValues[an0].CategoriesFormats: SeriesValues[an0].DataFormats;
				if (CatFmt != null && CatFmt.Length > 0) YAxis.CellNumberFormat = CatFmt[0];
			}

			YAxis.LineOptions = ValueAxis.AxisLineOptions;
			YAxis.TickOptions = ValueAxis.TickOptions;
			YAxis.RangeOptions = ValueAxis.RangeOptions;
			YAxis.MaxCross = (ValueAxis.AxisOptions & TValueAxisOptions.MaxCross) != 0;
			YAxis.Caption = ValueAxis.Caption;

			TValueAxisOptions AxOp = ValueAxis.AxisOptions;
			if ((AxOp & TValueAxisOptions.AutoMin) == 0) { AutoMin = false; YAxis.Min = ValueAxis.Min; }
			if ((AxOp & TValueAxisOptions.AutoMax) == 0) { AutoMax = false; YAxis.Max = ValueAxis.Max; }
			if ((AxOp & TValueAxisOptions.AutoMajor) == 0) { AutoMajor = false; YAxis.MajorScale = ValueAxis.Major; }
			if ((AxOp & TValueAxisOptions.AutoMinor) == 0) { AutoMinor = false; YAxis.MinorScale = ValueAxis.Minor; }

			if ((AxOp & TValueAxisOptions.AutoCross) == 0) { YAxis.CrossPoint = ValueAxis.CrossValue; }

			if ((AxOp & TValueAxisOptions.Reverse) != 0) YAxis.ReverseValues = true;
			if ((AxOp & TValueAxisOptions.LogScale) != 0) YAxis.Logarithmic = true;

		}

		private static TDateUnits GetDateUnits(object[] avalues, bool Dates1904)
		{
			if (avalues == null) return TDateUnits.Day;
            List<double> values = new List<double>();
            foreach (object o in avalues)
            {
                if (o is double) values.Add((double)o);
            }

            values.Sort();
			TDateUnits Result = TDateUnits.Year;
            for (int i = 1; i < values.Count; i++)
            {
                DateTime d1;
                DateTime d2;
                if (!FlxDateTime.TryFromOADate(values[i - 1], Dates1904, out d1)) continue;
                if (!FlxDateTime.TryFromOADate(values[i], Dates1904, out d2)) continue;
                if (d1.AddMonths(1) > d2) return TDateUnits.Day;
                if (d1.AddYears(1) > d2) Result = TDateUnits.Month;
            }
			return Result;
		}

		private static int GetFirstSeries(ChartSeries[] SeriesValues, TChartOptions[] ChartOptions, int AxisNumber)
		{
			for (int k = 0; k < SeriesValues.Length; k++)
			{
				int ib = SeriesValues[k].ChartOptionsIndex;
				int an = ChartOptions[ib].AxisNumber;
				if (an == AxisNumber) return k;
			}
			return -1;
		}

        private static bool GetAutoDate(ExcelFile Workbook, ChartSeries[] SeriesValues, TChartOptions[] ChartOptions, int AxisNumber)
        {
            for (int i = 0; i < SeriesValues.Length; i++)
            {
                int ib = SeriesValues[i].ChartOptionsIndex;
                int an = ChartOptions[ib].AxisNumber;

                if (an != AxisNumber) continue;
                if (SeriesValues[i].CategoriesValues == null || SeriesValues[i].CategoriesFormats == null) return false; //the first series defines the type of axis.

                for (int k = 0; k < SeriesValues[i].CategoriesValues.Length; k++)
                {
                    if (SeriesValues[i].CategoriesValues[k] == null || SeriesValues[i].CategoriesFormats[k] == null) continue;
                    bool HasDate; bool HasTime; Color aColor = ColorUtil.Empty;
                    TFlxNumberFormat.FormatValue(SeriesValues[i].CategoriesValues[k], SeriesValues[i].CategoriesFormats[k], ref aColor, Workbook, out HasDate, out HasTime);
                    return HasDate;
                }
            }
            return false;
        }


		private static void CalcOneCategoryAxis(ExcelFile Workbook, ChartSeries[] SeriesValues, TChartOptions[] ChartOptions, int AxisNumber, TAxisInfo XAxis, ref bool AutoMin, ref bool AutoMax, ref bool AutoMinor, ref bool AutoMajor, TCategoryAxis CatAxis, bool Dates1904)
		{
			XAxis.IsCategory = true;
			XAxis.Font = CatAxis.Font;
			XAxis.NumberFormat = CatAxis.NumberFormat;

			int an0 = GetFirstSeries(SeriesValues, ChartOptions, AxisNumber);
			if (an0 >= 0)
			{
				string[] CatFmt = SeriesValues[an0].CategoriesFormats;
				if (CatFmt != null && CatFmt.Length > 0) XAxis.CellNumberFormat = CatFmt[0];
			}

			XAxis.LineOptions = CatAxis.AxisLineOptions;
			XAxis.TickOptions = CatAxis.TickOptions;
			XAxis.RangeOptions = CatAxis.RangeOptions;
			XAxis.Caption = CatAxis.Caption;


			if (CatAxis.RangeOptions != null)
			{
				XAxis.MaxCross = CatAxis.RangeOptions.ValueAxisAtMaxCategory;
			}

			TCategoryAxisOptions AxOp = CatAxis.AxisOptions;
			if ((AxOp & TCategoryAxisOptions.AutoMin) == 0) { AutoMin = false; XAxis.Min = CatAxis.Min; }
			if ((AxOp & TCategoryAxisOptions.AutoMax) == 0) { AutoMax = false; XAxis.Max = CatAxis.Max; }
			if ((AxOp & TCategoryAxisOptions.AutoMajor) == 0) { AutoMajor = false; XAxis.MajorScale = CatAxis.MajorValue; }
			if ((AxOp & TCategoryAxisOptions.AutoMinor) == 0) { AutoMinor = false; XAxis.MinorScale = CatAxis.MinorValue; }

            if ((AxOp & TCategoryAxisOptions.AutoDate) != 0)
            {
                XAxis.IsDate = GetAutoDate(Workbook, SeriesValues, ChartOptions, AxisNumber);
            }
            else
            {
                if ((AxOp & TCategoryAxisOptions.DateAxis) != 0) { XAxis.IsDate = true; XAxis.DateAxisForced = true; }
            }
			if (XAxis.IsDate)
			{
				TDateUnits AutoValue = TDateUnits.Day;
				bool AutoUnits = (AxOp & TCategoryAxisOptions.AutoBase) != 0;

				if (AutoMajor || AutoMinor || AutoUnits)
				{
					int k = GetFirstSeries(SeriesValues, ChartOptions, AxisNumber);
					if (k >= 0) 
					{
						AutoValue = GetDateUnits(SeriesValues[k].CategoriesValues, Workbook.OptionsDates1904);
					}
				}

				if (!AutoUnits)
				{
					XAxis.DateUnits = (TDateUnits) CatAxis.BaseUnit;
				}
				else
				{
					XAxis.DateUnits = AutoValue;
				}

				if (AutoMajor) XAxis.DateUnitsMajor = (TDateUnits) Math.Max((int)AutoValue, (int)XAxis.DateUnits); else XAxis.DateUnitsMajor = (TDateUnits) CatAxis.MajorUnit;
				if (AutoMinor) XAxis.DateUnitsMinor = (TDateUnits) Math.Max((int)AutoValue, (int)XAxis.DateUnits); else XAxis.DateUnitsMinor = (TDateUnits) CatAxis.MinorUnit;

				if (XAxis.DateUnits > XAxis.DateUnitsMinor) XAxis.DateUnits = XAxis.DateUnitsMinor;
				if (XAxis.DateUnits > XAxis.DateUnitsMajor) XAxis.DateUnits = XAxis.DateUnitsMajor;

			}

			XAxis.CrossPoint = CatAxis.CrossValue;

			if (CatAxis.RangeOptions != null) XAxis.ReverseValues = CatAxis.RangeOptions.ReverseCategories;
		}

		private static double CalcAutoScale(double y)
		{
			//Y scale might be 1, 2 or 5 (since those are the only divisors of 10)
			//We need to fit at least 5 divisions on the chart, and also have a minumum gap at the top of y*0.05.
			//so yscale must be the biggest of (1,2,5) that is less than (y*1.05)/5.
			//The cut is made this way: let z = y*1.05/5
			//z<2 ->1, z <4 ->2, else 5

			//On the other side, if the line gap between 2 lines is less than the height of text, we should try a bigger scale.
			
			double yToFit = Math.Abs(y);
			double z = yToFit/5;
			if (z < Single.Epsilon * 10)
			{
				return 10;
			}

			double ZScale = Math.Pow(10, Math.Floor(Math.Log10(z)));
			double ZScaled = z / ZScale;

			double yscale;
			if (ZScaled < 2) yscale = 1; else
				if (ZScaled <4) yscale = 2; else
				yscale = 5;

			yscale = yscale * ZScale;
			return yscale;
		}


		private static double TryNextScale(double z)
		{
			double ZScale = Math.Pow(10, Math.Floor(Math.Log10(z)));
			double ZScaled = z / ZScale;

			double yscale;
			if (ZScaled < 2) yscale = 2; else
				if (ZScaled <4) yscale = 5; else
				yscale = 10;

			yscale = yscale * ZScale;
			return yscale;
		}

		private static void CalcZeroCross(TAxisInfo Axis, bool AutoMin, bool AutoMax)
		{
			double dx = Math.Abs(Axis.Max - Axis.Min);
			if (AutoMin && Axis.Min > 0)
			{
				if (dx == 0 || dx * 6 > Axis.Max)
				{
					Axis.Min = 0;
				}
			}

			if (AutoMax && Axis.Max < 0)
			{
				if (dx == 0 || dx * 6 > Math.Abs(Axis.Min))
				{
					Axis.Max = 0;
				}
			}
		}

		private static void CalcValueAxisDefaults(IFlxGraphics Canvas, RectangleF PlotCoords, TFontCache FontCache, real Zoom100, TAxisInfo YAxis, bool AutoMin, bool AutoMax, bool AutoMinor, bool AutoMajor)
		{
			CalcZeroCross(YAxis, AutoMin, AutoMax);
			double w1 = 1.05;
			double w2 = w1 - 1;
			
			if (AutoMajor)
			{

				double yscalePos = 0;
				double yscaleNeg = 0;
				if (YAxis.Max > 0) 
				{
					if (YAxis.Min > 0) 
						yscalePos = CalcAutoScale((YAxis.Max - YAxis.Min)*(w1 + w2));
					else
						yscalePos = CalcAutoScale(YAxis.Max*w1);
				}
				if (YAxis.Min < 0) 
					if (YAxis.Max < 0)
						yscaleNeg = CalcAutoScale((-YAxis.Min + YAxis.Max)*(w1 + w2));
				    else
						yscaleNeg = CalcAutoScale(YAxis.Min * w1);

				YAxis.MajorScale = Math.Max(yscaleNeg, yscalePos);
				if (YAxis.MajorScale > 0 && YAxis.Max - YAxis.Min > 0)
				{
					real AxisHeight = YAxis.Horizontal? PlotCoords.Width: PlotCoords.Height;		
					Font aFont = FontCache.GetFont(YAxis.Font.Font, Zoom100);

					real TextHeight = Canvas.MeasureString("M", aFont).Height * 1.1f;
					TextHeight *= HorizLabelMargin(YAxis);

					
					real ValueRange = (real) (AxisHeight / (YAxis.Max - YAxis.Min));
					while (ValueRange * YAxis.MajorScale < TextHeight)
					{
						YAxis.MajorScale = TryNextScale(YAxis.MajorScale);
					}
				}
			}

			if (YAxis.MajorScale == 0) return;

			if (AutoMinor) YAxis.MinorScale = YAxis.MajorScale / 5f;

			if (AutoMax && YAxis.Max != 0)
			{ 
				double posdy = YAxis.Max - Math.Max(0, YAxis.Min);
				double negdy = -YAxis.Min + Math.Min(0, YAxis.Max);
				if (YAxis.Max > 0)
					YAxis.Max =  Math.Ceiling((Math.Max(0, YAxis.Min) + posdy * w1) / YAxis.MajorScale) * YAxis.MajorScale;
				else
				{
					YAxis.Max = Math.Floor((-YAxis.Max - negdy * w2) / YAxis.MajorScale) * -YAxis.MajorScale;
					if (YAxis.Max > 0) YAxis.Max = 0;
				}
			}
			if (AutoMin && YAxis.Min != 0) 
			{
				double posdy = YAxis.Max - Math.Max(0, YAxis.Min); //recalculate them.
				double negdy = -YAxis.Min + Math.Min(0, YAxis.Max);
				if (YAxis.Min < 0)
					YAxis.Min = Math.Ceiling((-Math.Min(0, YAxis.Max) + negdy * w1) / YAxis.MajorScale) * -YAxis.MajorScale;
				else
				{
					YAxis.Min = Math.Floor((YAxis.Min - posdy * w2) / YAxis.MajorScale) * YAxis.MajorScale;
                    if (YAxis.Min < 0) YAxis.Min = 0;
				}
			}


			if (!YAxis.OneNotPercent)
			{
				if (AutoMax && YAxis.Max > 100) YAxis.Max = 100;
				if (AutoMin && YAxis.Min < -100) YAxis.Min = -100;
			}

			if (YAxis.Max - YAxis.Min <= 0) return;

            CheckCrossPoint(YAxis);

			//if (MaxX[an] <= 0) return;
			//if (SeriesValues.Length <= 0) return;
		}

		private static void CalcCategoryAxisDefaults(RectangleF PlotCoords, TAxisInfo XAxis, int i, ChartSeries[] SeriesValues, TChartOptions[] ChartOptions, int[] MaxX, bool AutoMin, bool AutoMax, bool AutoMinor, bool AutoMajor)
		{
			if (AutoMin) XAxis.Min = 1;
			if (AutoMax) XAxis.Max = MaxX[i];
			if (AutoMajor)
			{
				XAxis.MajorScale = 1;
				if (XAxis.RangeOptions != null) XAxis.MajorScale = XAxis.RangeOptions.TickFrequency;
			}
			if (AutoMinor) XAxis.MinorScale = XAxis.MajorScale / 2f;

			for (int n = 1; n <= SeriesValues.Length; n++)
			{
				int OpIn = SeriesValues[n - 1].ChartOptionsIndex;
				TChartOptions Options = ChartOptions[OpIn];
				if (Options.ChartType == TChartType.Pie) continue; //Pie charts do not have axis.

				if (Options.AxisNumber == i)
				{
					XAxis.Labels = SeriesValues[n - 1].CategoriesValues;
					break;
				}
			}

            CheckCrossPoint(XAxis);

            int WallOfs = GetWallOffs(XAxis);

			if (MaxX[i] - (1 - WallOfs) > 0)
			{
				double CWidth = XAxis.Horizontal? PlotCoords.Width: PlotCoords.Height;
				double SerieWidth = CWidth / (MaxX[i] - (1 - WallOfs));
			
				XAxis.StartOffset = (real)(WallOfs * SerieWidth / 2f);
			}
		}

		private static real HorizLabelMargin(TAxisInfo Axis)
		{
			int Rotation = 0;
			if (Axis.TickOptions != null)
			{
				Rotation = Axis.TickOptions.Rotation;
			}

			bool Vertical;
			real alpha = FlexCelRender.CalcAngle(Rotation, out Vertical);

			if (Vertical)
			{
				if (Axis.Horizontal) return 1;
				return HorizTextLabelMargin;
			}

			if (!Axis.Horizontal) alpha-=90;
			return (real)(1 + ((HorizTextLabelMargin - 1) * Math.Cos(Math.Abs(alpha)* Math.PI/180)));

		}

		private static void CalcDateAxisDefaults(ExcelFile Workbook, IFlxGraphics Canvas, TFontCache FontCache, real Zoom100, TAxisInfo Axis, ChartSeries[] SeriesValues, RectangleF PlotCoords, bool AutoMin, bool AutoMax, bool AutoMinor, bool AutoMajor)
		{
			if (AutoMajor)
			{
				Axis.MajorScale = 1;
				if (Axis.Max - Axis.Min > 0)
				{
					real AxisHeight = Axis.Horizontal? PlotCoords.Width: PlotCoords.Height;		
					Font aFont = FontCache.GetFont(Axis.Font.Font, Zoom100);

					real TextHeight = Canvas.MeasureString("M", aFont).Height * 1.1f;
					TextHeight *= HorizLabelMargin(Axis);

					
					real ValueRange = (real) (AxisHeight / (Axis.Max - Axis.Min) * TDateAxisTransform.MinDateUnit(Axis.DateUnitsMajor));
					while (ValueRange * Axis.MajorScale < TextHeight)
					{
						Axis.MajorScale = TryNextScale(Axis.MajorScale);
					}
				}
			}

			Axis.CrossPoint = TDateAxisTransform.AddDateUnit(Axis.Min, Axis.DateUnits, (int)Axis.CrossPoint - 1, Workbook.OptionsDates1904);
            CheckCrossPoint(Axis);

			if (Axis.Max <= Axis.Min)
			{
				Axis.StartOffset = 0;
				return;
			}

			int WallOfs = GetWallOffs(Axis);
			double CWidth = Axis.Horizontal ? PlotCoords.Width : PlotCoords.Height;
			double WOfs = WallOfs * TDateAxisTransform.MinDateUnit(Axis.DateUnits);
			double SWidth = CWidth / ((Axis.Max - Axis.Min) + WOfs);
			Axis.StartOffset = (real)(SWidth * WOfs) / 2;
		}

        private static void CheckCrossPoint(TAxisInfo Axis)
        {
            if (Axis.MaxCross) Axis.CrossPoint = Axis.Max;
            if (Axis.CrossPoint > Axis.Max) Axis.CrossPoint = Axis.Max;
            if (Axis.CrossPoint < Axis.Min) Axis.CrossPoint = Axis.Min;
        }


		#endregion

		#region Bar Draw
		private static bool GetBarDrawParams(RectangleF Coords, TBarChartOptions Options, bool Horizontal, int MaxX, int GroupSeriesCount, int WallOfs,
			out double CHeight, out double CWidth, out double SerieWidth, out double BarWidth, out double XOffset)
		{
			CWidth = Horizontal? Coords.Height: Coords.Width;
			CHeight = Horizontal? Coords.Width: Coords.Height;

			SerieWidth = CWidth / (MaxX - (1 - WallOfs));

			double z = Options.CategoriesGap + GroupSeriesCount * (1 + Options.BarOverlap) - Options.BarOverlap;
			if (z <= 0) 
			{
				BarWidth = 0;
				XOffset = 0;
				return false;
			}
			
			BarWidth = SerieWidth / z;

			XOffset = BarWidth * Options.CategoriesGap / 2;
			if (WallOfs == 0)
				XOffset -= SerieWidth /2f;
			return true;

		}

		private static void GetBarCoords(TLabelDescription aLabel, RectangleF Bar, RectangleF ChartCoords, RectangleF PlotCoords, bool Horizontal, bool Stacked, bool XReversed, bool YReversed)
		{
            TDataLabelPosition dlp = aLabel.FLabel.LabelOptions.Position;
            if (dlp == TDataLabelPosition.Automatic && Stacked) dlp = TDataLabelPosition.Center;
			switch (dlp)
			{
				case TDataLabelPosition.Any:
                    SetDefaultBarLabelPos(aLabel, Horizontal, Bar);
                    float XDir = XReversed ? -1 : 1;
                    float YDir = YReversed ? -1 : 1;
                    if (TDataLabels.SetDefaultLabelBoxPos(ChartCoords, PlotCoords, aLabel, false, XDir, YDir)) return;
                    
                    aLabel.Position = TDataLabels.DefaultLabelBoxXY(ChartCoords, aLabel.FLabel);
                    break;

				case TDataLabelPosition.Center:
					aLabel.Position = new RectangleF((Bar.Left + Bar.Right) / 2, (Bar.Top + Bar.Bottom) / 2, 0, 0);
					aLabel.HLabelPosition = THLabelPosition.Center;
					aLabel.VLabelPosition = TVLabelPosition.Center;
					break;

				case TDataLabelPosition.Inside:
					if (Horizontal) 
					{
						aLabel.Position = new RectangleF(Bar.Right - BarLabelOfs, (Bar.Top + Bar.Bottom) / 2, 0, 0);
						aLabel.HLabelPosition = THLabelPosition.Right;
						aLabel.VLabelPosition = TVLabelPosition.Center;
					}
					else
					{
						aLabel.Position = new RectangleF((Bar.Left + Bar.Right) / 2, Bar.Top - BarLabelOfs, 0, 0);
						aLabel.HLabelPosition = THLabelPosition.Center;
						aLabel.VLabelPosition = TVLabelPosition.Up;
					}
					break;

				case TDataLabelPosition.Axis:
					if (Horizontal) 
					{
						aLabel.Position = new RectangleF(Bar.Left + BarLabelOfs, (Bar.Top + Bar.Bottom) / 2, 0, 0);
						aLabel.HLabelPosition = THLabelPosition.Left;
						aLabel.VLabelPosition = TVLabelPosition.Center;
					}
					else
					{
						aLabel.Position = new RectangleF((Bar.Left + Bar.Right) / 2, Bar.Bottom - BarLabelOfs, 0, 0);
						aLabel.HLabelPosition = THLabelPosition.Center;
						aLabel.VLabelPosition = TVLabelPosition.Down;
					}
					break;

				default:
                    SetDefaultBarLabelPos(aLabel, Horizontal, Bar);
					break;
			}
		}

        private static void SetDefaultBarLabelPos(TLabelDescription aLabel, bool Horizontal, RectangleF Bar)
        {
            if (Horizontal)
            {
                aLabel.Position = new RectangleF(Bar.Right + BarLabelOfs, (Bar.Top + Bar.Bottom) / 2, 0, 0);
                aLabel.HLabelPosition = THLabelPosition.Left;
                aLabel.VLabelPosition = TVLabelPosition.Center;
            }
            else
            {
                aLabel.Position = new RectangleF((Bar.Left + Bar.Right) / 2, Bar.Top - BarLabelOfs, 0, 0);
                aLabel.HLabelPosition = THLabelPosition.Center;
                aLabel.VLabelPosition = TVLabelPosition.Down;
            }
        }


		private static void DrawBarChart(ExcelChart Chart, ChartSeries SeriesValues, int GroupSeriesIndex, int GroupSeriesCount, ExcelFile xls, TChartCanvas Canvas, 
			RectangleF Coords, TBarChartOptions Options, 
			TAxisInfo XAxis, TAxisInfo YAxis, real Zoom100, int MaxX, double[] StackedOfsPos, double[] StackedOfsNeg, 
			TOneSeriesLabelDescription Labels)
		{
			if (MaxX <= 0) return;
			if (SeriesValues == null || SeriesValues.DataValues == null) return;

            int WallOfs = GetWallOffs(XAxis);

			double CWidth, CHeight, SerieWidth, BarWidth, XOffset;
			if (!GetBarDrawParams(Coords, Options, YAxis.Horizontal, MaxX, GroupSeriesCount, WallOfs, out CHeight, out CWidth, out SerieWidth, out BarWidth, out XOffset)) return;

			double Sw = XOffset;
			
			double BarOffsetX = GroupSeriesIndex * BarWidth * (1 + Options.BarOverlap);
			PointF LastBarPoint = PointF.Empty;

			using (Brush ABrush = GetSeriesBrush(xls, new RectangleF(0,0,1,1), Options, Options.SeriesOptions, SeriesValues.Options[-1], null, SeriesValues.SeriesNumber))
			{
				using (Pen aPen = GetSeriesPen(xls, Options, Options.SeriesOptions, SeriesValues.Options[-1], null, SeriesValues.SeriesNumber, true))
				{
					for (int k = 0; k < SeriesValues.DataValues.Length; k++)
					{
						double CrossPoint = 0;
						if (Options.StackedMode == TStackedMode.None && YAxis.CrossPoint != 0)
						{
							CrossPoint = YAxis.CrossPoint;
						}

						double BarHeight = 0;
						if (SeriesValues.DataValues[k] == null && Chart.PlotEmptyCells == TPlotEmptyCells.Zero) BarHeight = -CrossPoint;
						else
						{
							if (SeriesValues.DataValues[k] is double)
							{
                                BarHeight = Convert.ToDouble(SeriesValues.DataValues[k], CultureInfo.CurrentCulture) - CrossPoint;
							}
						}
			
						double[] StackedOfs = BarHeight < 0? StackedOfsNeg: StackedOfsPos;
						double Stacked = StackedOfs==null? 0: StackedOfs[k];
						double BarBottomOfs = (YAxis.Max - CrossPoint) * CHeight / (YAxis.Max - YAxis.Min) - Stacked;
						double BarTopOfs =  (YAxis.Max - BarHeight - CrossPoint) * CHeight /(YAxis.Max - YAxis.Min) - Stacked;


						PointF BarPoint1;
						PointF BarPoint2;
						if (YAxis.Horizontal)
						{
							BarPoint1 = new PointF(Coords.Right - (real)BarTopOfs, 
								Coords.Bottom - (real)(Sw + BarOffsetX + BarWidth));
							BarPoint2 = new PointF(BarPoint1.X, 
								(real)(BarPoint1.Y + BarWidth));
						}
						else
						{
							BarPoint1 = new PointF(Coords.Left + (real)(Sw + BarOffsetX + BarWidth), 
								(real)(Coords.Top + BarTopOfs));
							BarPoint2 = new PointF(Coords.Left + (real)(Sw + BarOffsetX), 
								BarPoint1.Y);
						}

						
						if (BarBottomOfs < BarTopOfs)
						{
							double tmp = BarBottomOfs;
							BarBottomOfs = BarTopOfs;
							BarTopOfs = tmp;
						}


						RectangleF Bar;
						if (YAxis.Horizontal)
						{
							Bar = new RectangleF(Coords.Right - (real)BarBottomOfs, 
								Coords.Bottom - (real)(Sw + BarOffsetX + BarWidth), 
								(real)(BarBottomOfs - BarTopOfs),
								(real)BarWidth);
						}
						else
						{
							Bar = new RectangleF(Coords.Left + (real)(Sw + BarOffsetX), 
								(real)(Coords.Top + BarTopOfs), 
								(real)BarWidth, 
								(real)(BarBottomOfs - BarTopOfs));
						}

						if (Labels != null)
						{
							TLabelDescription Lbl = Labels[k];
							bool AlreadyThere = Lbl != null;
							if (!AlreadyThere) Lbl = Labels[-1];
							if (Lbl != null && !Lbl.FLabel.LabelOptions.Deleted)
							{
								if (AlreadyThere) 
								{
									GetBarCoords(Lbl, Bar, Canvas.ChartCoords, Coords, YAxis.Horizontal, Options.StackedMode != TStackedMode.None, XAxis.ReverseValues, YAxis.ReverseValues);
									Lbl.Ax1 = XAxis;
									Lbl.Ax2 = YAxis;
								}
								else 
								{
									TLabelDescription tmp2 = new TLabelDescription(RectangleF.Empty, Lbl.FLabel, 0);
                                    GetBarCoords(tmp2, Bar, Canvas.ChartCoords, Coords, YAxis.Horizontal, Options.StackedMode != TStackedMode.None, XAxis.ReverseValues, YAxis.ReverseValues);
                                    Labels.Add(k, tmp2);
									tmp2.Ax1 = XAxis;
									tmp2.Ax2 = YAxis;
								}
							}
						}

						Brush ABrush2 = ABrush;
						if (!(ABrush is SolidBrush) || PointNeedsCustomBrush(SeriesValues.Options[k]))
							ABrush2 = GetSeriesBrush(xls, Bar, Options, Options.SeriesOptions, SeriesValues.Options[-1], SeriesValues.Options[k], SeriesValues.SeriesNumber);
						try
						{
							Pen aPen2 = aPen;
							if (PointNeedsCustomPen(SeriesValues.Options[k]))
								aPen2 = GetSeriesPen(xls, Options, Options.SeriesOptions, SeriesValues.Options[-1], SeriesValues.Options[k], SeriesValues.SeriesNumber, true);
							try
							{
								Canvas.DrawAndFillRectangle(YAxis, XAxis, aPen2, ABrush2, Bar.X, Bar.Y, Bar.Width, Bar.Height);
							}
							finally
							{
								if (aPen2 != aPen && aPen2 != null) aPen2.Dispose();
							}
						}
						finally
						{
							if (ABrush2 != ABrush && ABrush2 != null) ABrush2.Dispose();
						}

						if (Options != null && Options.SeriesLines != null && k > 0 && Options.StackedMode != TStackedMode.None) 
						{
							DrawSeriesLines(xls, Canvas, LastBarPoint, BarPoint2, Options.SeriesLines, XAxis, YAxis);
						}
						LastBarPoint = BarPoint1;

						if (StackedOfs != null) 
						{
							if (BarHeight < 0)
								StackedOfs[k] -=BarBottomOfs - BarTopOfs;
							else
								StackedOfs[k] +=BarBottomOfs - BarTopOfs;
						}
						Sw += SerieWidth;
					}
				}
			}
		}
		

		#endregion

		#region Line And Area Draw
		#region Line markers
		internal static void DrawOneMarker(TChartCanvas Canvas, TPointF Point, double MarkerSize, TMarkerImgInfo MarkerImage, TChartMarkerType MarkerType, Pen aPen, Brush aBrush, TAxisInfo Ax1, TAxisInfo Ax2)
		{
			if (MarkerImage != null) 
			{
				real w = MarkerImage.Width;
				real h = MarkerImage.Height;
				Canvas.DrawImage(Ax1, Ax2, MarkerImage.Img, Point.X - w / 2, Point.Y - h / 2, w, h);
				return;
			}

			if (MarkerType == TChartMarkerType.None) return;

			double r = MarkerSize /40.0;
			switch (MarkerType)
			{
				case TChartMarkerType.None: break;
				case TChartMarkerType.Circle: 
					TPointF[] Circle = TEllipticalArc.GetPoints(Point.X, Point.Y, r, r, 0, 0, 2*Math.PI);
					Canvas.DrawAndFillBeziers(Ax1, Ax2, aPen, aBrush, Circle);					
					break;
				case TChartMarkerType.Diamond: 
					TPointF[] Diamond = new TPointF[4];
					Diamond[0] = new TPointF((real)(Point.X - r), (real)(Point.Y));
					Diamond[1] = new TPointF((real)(Point.X), (real)(Point.Y + r));
					Diamond[2] = new TPointF((real)(Point.X + r), (real)(Point.Y));
					Diamond[3] = new TPointF((real)(Point.X), (real)(Point.Y - r));
					Canvas.DrawAndFillPolygon(Ax1, Ax2, aPen, aBrush, Diamond);					
					break;
				case TChartMarkerType.DowJones: 
					Canvas.DrawLine(Ax1, Ax2, aPen, Point.X, Point.Y, (real)(Point.X + r), Point.Y);					
					break;
				case TChartMarkerType.Plus: 
					Canvas.FillRectangle(Ax1, Ax2, aBrush, (real)(Point.X - r), (real)(Point.Y - r), (real) (2 * r), (real) (2 * r));					
					Canvas.DrawLine(Ax1, Ax2, aPen, (real)(Point.X), (real)(Point.Y - r), (real)(Point.X), (real)(Point.Y + r));					
					Canvas.DrawLine(Ax1, Ax2, aPen, (real)(Point.X - r), (real)(Point.Y), (real)(Point.X + r), (real)(Point.Y));					
					break;
				case TChartMarkerType.Square: 
					Canvas.DrawAndFillRectangle(Ax1, Ax2, aPen, aBrush, (real)(Point.X - r), (real)(Point.Y - r), (real) (2 * r), (real) (2 * r));					
					break;
				case TChartMarkerType.StandardDeviation: 
					Canvas.DrawLine(Ax1, Ax2, aPen, (real)(Point.X - r), Point.Y, (real)(Point.X + r), Point.Y);					
					break;
				case TChartMarkerType.Star: 
					Canvas.FillRectangle(Ax1, Ax2, aBrush, (real)(Point.X - r), (real)(Point.Y - r), (real) (2 * r), (real) (2 * r));					
					Canvas.DrawLine(Ax1, Ax2, aPen, (real)(Point.X), (real)(Point.Y - r), (real)(Point.X), (real)(Point.Y + r));					
					Canvas.DrawLine(Ax1, Ax2, aPen, (real)(Point.X - r), (real)(Point.Y - r), (real)(Point.X + r), (real)(Point.Y + r));					
					Canvas.DrawLine(Ax1, Ax2, aPen, (real)(Point.X + r), (real)(Point.Y - r), (real)(Point.X - r), (real)(Point.Y + r));					
					break;
				case TChartMarkerType.X: 
					Canvas.FillRectangle(Ax1, Ax2, aBrush, (real)(Point.X - r), (real)(Point.Y - r), (real) (2 * r), (real) (2 * r));					
					Canvas.DrawLine(Ax1, Ax2, aPen, (real)(Point.X - r), (real)(Point.Y - r), (real)(Point.X + r), (real)(Point.Y + r));					
					Canvas.DrawLine(Ax1, Ax2, aPen, (real)(Point.X + r), (real)(Point.Y - r), (real)(Point.X - r), (real)(Point.Y + r));					
					break;
				case TChartMarkerType.Triangle: 
					TPointF[] Triangle = new TPointF[3];
					Triangle[0] = new TPointF((real)(Point.X - r), (real)(Point.Y + r));
					Triangle[1] = new TPointF((real)(Point.X), (real)(Point.Y - r));
					Triangle[2] = new TPointF((real)(Point.X + r), (real)(Point.Y + r));
					Canvas.DrawAndFillPolygon(Ax1, Ax2, aPen, aBrush, Triangle);					
					break;
			}
		}

		private static void GetMarkerPenAndBrush(ExcelFile xls, int SeriesNumber, ChartSeriesOptions Opt, ref Pen aPen, ref Brush aBrush)
		{
			if (Opt != null && Opt.MarkerOptions != null)
			{
				if (!Opt.MarkerOptions.NoForeground)
				{
					if (Opt.MarkerOptions.AutomaticColors) aPen = AutomaticPen(xls, SeriesNumber);
					else
					{
						aPen = new Pen(Opt.MarkerOptions.FgColor);
					}
				}
				else aPen = null;

				if (!Opt.MarkerOptions.NoBackground)
				{
					if (Opt.MarkerOptions.AutomaticColors) aBrush = AutomaticBrush(xls, SeriesNumber + 8);
					else
					{
						aBrush = new SolidBrush(Opt.MarkerOptions.BgColor);
					}
				}
				else aBrush = null;

			}
		}

		private static void DrawMarkers(ExcelFile xls, TChartCanvas Canvas, ChartSeries SeriesValues, TPointF[] Points, ChartSeriesOptions SeriesOptions, TAxisInfo Ax1, TAxisInfo Ax2, bool[] Exclude, real Zoom100)
		{
			Pen aPen = null;
			Brush aBrush = null;
			TMarkerImgInfo MarkerImage = null;
			double MarkerSize;
			TChartMarkerType MarkerType;

			Canvas.Canvas.SaveState();
			try
			{
				Canvas.Canvas.SetClipReplace(Canvas.ChartCoords);
				try
				{
					GetMarkerOptions(xls, SeriesValues, SeriesOptions, out aPen, out aBrush, out MarkerSize, out MarkerType, out MarkerImage, Zoom100);

					for (int k = 0; k < Points.Length; k++)
					{
						if (Exclude != null && Exclude[k]) continue;
						Pen aPen2 = aPen;
						Brush aBrush2 = aBrush;
						double MarkerSize2 = MarkerSize;
						TChartMarkerType MarkerType2 = MarkerType;
						if (PointNeedsCustomMarker(SeriesValues.Options[k]))
						{
							GetMarkerPenAndBrush(xls, SeriesValues.SeriesNumber, SeriesValues.Options[k], ref aPen2, ref aBrush2);
							MarkerSize2 = SeriesValues.Options[k].MarkerOptions.MarkerSize;
							MarkerType2 = SeriesValues.Options[k].MarkerOptions.MarkerType;
						}
						try
						{
							DrawOneMarker(Canvas, Points[k], MarkerSize2, MarkerImage, MarkerType2, aPen2, aBrush2, Ax1, Ax2);
						}
						finally
						{
							if (aPen2 != null && aPen2 != aPen) aPen2.Dispose();
							if (aBrush2 != null && aBrush2 != aBrush) aBrush2.Dispose();
						}
					}
				}
				finally
				{
					if (aPen != null) aPen.Dispose();
					if (aBrush != null) aBrush.Dispose();
					if (MarkerImage != null) MarkerImage.Dispose();
				}
			}
			finally
			{
				Canvas.Canvas.RestoreState();
			}
		}

		
		internal static void GetMarkerOptions(ExcelFile xls, ChartSeries SeriesValues, ChartSeriesOptions SeriesOptions, out Pen aPen, out Brush aBrush, out double MarkerSize, out TChartMarkerType MarkerType, out TMarkerImgInfo MarkerImage, real Zoom100)
		{
			MarkerImage = null;
			ChartSeriesOptions Opt = SeriesValues.Options[-1];
			bool IsSeriesValues = true;
			if (Opt == null || Opt.MarkerOptions == null) 
			{
				Opt = SeriesOptions;
				IsSeriesValues = false;
			}

			aPen = null;
			aBrush = null;
			MarkerSize = 100;
			MarkerType = AutomaticMarker(SeriesValues.SeriesNumber);


			if (Opt == null || Opt.MarkerOptions == null)
			{
				aPen = AutomaticPen(xls, SeriesValues.SeriesNumber);
				if (MarkerNeedsBackground(MarkerType)) aBrush = AutomaticBrush(xls, SeriesValues.SeriesNumber + 8);
			}
			else
			{
				if (Opt.ExtraOptions != null) //There is an image attached to the datapoint.
				{
					MarkerImage = new TMarkerImgInfo();
					ImageAttributes Attr = null;
					try
					{
						TShapeProperties ExtraProps = new TShapeProperties();
						ExtraProps.ShapeOptions = Opt.ExtraOptions;
						MarkerImage.Img = DrawShape.GetTextureImage(ExtraProps, xls, TFillType.Picture, out Attr, out MarkerImage.Width, out MarkerImage.Height, 1);
						if (MarkerImage.Img != null)
						{
							MarkerImage.Height *= Zoom100;
							MarkerImage.Width *= Zoom100;
							return;
						}
						else
						{
							MarkerImage.Dispose();
							MarkerImage = null;
						}
					}
					finally
					{
						if (Attr != null) Attr.Dispose();
					}
				}
			
				MarkerSize = Opt.MarkerOptions.MarkerSize;
				if (IsSeriesValues || Opt.MarkerOptions.MarkerType == TChartMarkerType.None) MarkerType = Opt.MarkerOptions.MarkerType;
				GetMarkerPenAndBrush(xls, SeriesValues.SeriesNumber, Opt, ref aPen, ref aBrush);
			}
		}

		private static TPointF[] Slice(TPointF[] Points, int First, int Last)
		{
			if (First == 0 && Last == Points.Length) return Points;
			TPointF[] Result = new TPointF[Last - First + 1];
			Array.Copy(Points, First, Result, 0, Result.Length);
			return Result;
		}

		private static void DrawLinePoints(TChartCanvas Canvas, Pen aPen, TPointF[] Points, bool[] Exclude, int First, int Last, ChartSeriesOptions SeriesOptions, TAxisInfo Ax1, TAxisInfo Ax2)
		{
			int PLength = Last - First + 1;
			if (SeriesOptions == null || SeriesOptions.MiscOptions == null || 
				!SeriesOptions.MiscOptions.SmoothedLines || Points.Length <= 2)
			{
				if (First != 0 || Last != Points.Length - 1) Points = Slice(Points, First, Last);
				Canvas.DrawLines(Ax1, Ax2, aPen, Points);
				return;
			}

			TPointF[] bPoints = new TPointF[3 * PLength - 2];

			for (int i = First; i <= Last; i++)
			{
				CalcBezierPoints(bPoints, Points, Exclude, i, i - First);
			}

			Canvas.DrawAndFillBeziers(Ax1, Ax2, aPen, null, bPoints);
		}

		private static void CalcBezierPoints(TPointF[] bPoints, TPointF[] Points, bool[] Exclude, int i, int k)
		{
			//Bezier algorithm adapted from http://www.xlrotor.com/resources/files.shtml
			//double d01 = Distance(Points, Exclude, i - 1, i);
			double d12 = Distance(Points, Exclude, i, i+1);
			//double d23 = Distance(Points, Exclude, i+1, i+2);
			double d02 = Distance(Points, Exclude, i-1, i+1);
			double d13 = Distance(Points, Exclude, i, i+2);
    
			bPoints[k * 3] = Points[i];


			double f1 = 1;
			double f2 = 1;

			if (d02 == 0 || d13 == 0)  //degenerated cases.
			{
				f1 = 0;
				f2 = 0;
			}

			else
			{
				if (
					((d02 / 6 <= d12 / 2) && (d13 / 6 <= d12 / 2)) //Normal case. Both vectors are smaller than d12/2
					)
				{
					f1 = i > 0? 1 / 6.0: 1 / 3.0;   
					if( i < Points.Length -1) f2 = 1 / 6.0; else f2 = 1 / 3.0;
				}
				else
				{
					if ((d02 / 6.0 > d12 / 2.0) && (d13 / 6.0 > d12 / 2.0))  //Both vectors bigger.
					{
						f1 = d12 / 2 / d02;
						f2 = d12 / 2 / d13;
					}
					else 
					{
						if (d02 / 6.0 > d12 / 2.0)  //only d02 is bigger.
						{
							f1 = d12 / 2 / d02;
							f2 = d12 / 2 / d13 * (d13 / d02);
						}
						else  //only d13 is bigger.
						{
							f1 = d12 / 2 / d02 * (d02 / d13);
							f2 = d12 / 2 / d13;
						}
					}
				}
			}
			SetPoint(bPoints, k*3 + 1, AddMulSubPoint(Points, i, i+1, i-1, f1));
			SetPoint(bPoints, k*3 + 2,  AddMulSubPoint(Points, i + 1, i, i + 2, f2));
		}

		private static void SetPoint(TPointF[] bPoints, int ib, TPointF Data)
		{
			if (ib < 0) return;
			if (ib > bPoints.Length -1) return;
			bPoints[ib] = Data;
		}

		private static TPointF AddMulSubPoint(TPointF[] Points, int a, int b, int c, double f)
		{
			TPointF Result = GetPoint(Points, a);
			TPointF pb = GetPoint(Points, b);
			TPointF pc = GetPoint(Points, c);
			double X = (pb.X - pc.X) * f;
			double Y = (pb.Y - pc.Y) * f;
			Result.X += (real)X;
			Result.Y += (real)Y;
			return Result;
		}

		private static TPointF GetPoint(TPointF[] Points, int i)
		{
			if (i < 0) i = 0;
			if (i > Points.Length - 1) i = Points.Length - 1;
			return Points[i];
		}

		private static double Distance(TPointF[] Points, bool[] Exclude, int i1, int i2)
		{
			if (i1 < 0) i1 = 0;
			if (i2 < 0) i2 = 0;
			if (i1 > Points.Length - 1) i1 = Points.Length - 1;
			if (i2 > Points.Length - 1) i2 = Points.Length - 1;
			if (Exclude != null && (Exclude[i1] || Exclude[i2])) return 0;
			return Math.Sqrt( 
				Math.Pow((Points[i1].X - Points[i2].X), 2) + 
				Math.Pow((Points[i1].Y - Points[i2].Y), 2)
				);
		}

		#endregion

		#region Drop Lines
		private static void DrawDropLines(ExcelFile xls, TChartCanvas Canvas, TPointF[] Points, TChartDropBars DropBars, TAxisInfo Ax1, TAxisInfo Ax2, real YPos)
		{
			if (DropBars.DropLines == null) return;
			using (Pen aPen = GetCustomPen(xls, DropBars.DropLines, null, Colors.Black))
			{
				for (int i = 0; i < Points.Length; i++)
				{
					Canvas.DrawLine(Ax1, Ax2, aPen, Points[i].X, Points[i].Y, Points[i].X, YPos);
				}
			}
		}

		private static void DrawSeriesLines(ExcelFile xls, TChartCanvas Canvas, PointF LastBar, PointF Bar, ChartLineOptions SeriesLines, TAxisInfo Ax1, TAxisInfo Ax2)
		{
			if (SeriesLines == null) return;
			using (Pen aPen = GetCustomPen(xls, SeriesLines, null, Colors.Black))
			{
				Canvas.DrawLine(Ax1, Ax2, aPen, LastBar.X, LastBar.Y, Bar.X, Bar.Y);
			}
		}

		private static void CopArray(ref TPointF[] dest, TPointF[] source)
		{
			if (dest == null)  {dest = (TPointF[])source.Clone(); return;}

			if (dest.Length < source.Length)
			{
				TPointF[] res = new TPointF[source.Length];
				Array.Copy(dest, 0, res, 0, dest.Length);
				Array.Copy(source, 0, res, dest.Length, source.Length - dest.Length);
				dest = res;
			}

		}

		private static void CalcHiLoLines(ExcelFile xls, TPointF[] Points, THiLoData HiLoData, bool[] Exclude, TChartDropBars DropBars)
		{
			if (Points == null || Points.Length == 0) return;
			if (HiLoData.Exclude == null) 
			{
				if (Exclude == null) HiLoData.Exclude = new bool[Points.Length];
				else HiLoData.Exclude = (bool[]) Exclude.Clone();
			}

			CopArray (ref HiLoData.Hi, Points);
			CopArray (ref HiLoData.Lo, Points);

			if (Exclude != null)
			{
				if (HiLoData.Exclude.Length < Exclude.Length)
				{
					bool[] res = new bool[Exclude.Length];
					Array.Copy(HiLoData.Exclude, 0, res, 0, HiLoData.Exclude.Length);
					HiLoData.Exclude = res;
				}
			}

			if (HiLoData.Exclude.Length < Points.Length)
			{
				bool[] res = new bool[Points.Length];
				Array.Copy(HiLoData.Exclude, 0, res, 0, HiLoData.Exclude.Length);
				HiLoData.Exclude = res;
			}


			HiLoData.DropBars = DropBars;
			for (int i = 0; i < Points.Length; i++)
			{
				if (Exclude == null || !Exclude[i])
				{
					if (HiLoData.Exclude[i] || Points[i].Y < HiLoData.Lo[i].Y) {HiLoData.Lo[i].Y = Points[i].Y; HiLoData.Lo[i].X = Points[i].X;}
					if (HiLoData.Exclude[i] || Points[i].Y > HiLoData.Hi[i].Y) {HiLoData.Hi[i].Y = Points[i].Y; HiLoData.Hi[i].X = Points[i].X;}
					HiLoData.Exclude[i] = false;

				}
			}
		}

		private static void DrawHiLoLines(ExcelFile xls, TChartCanvas Canvas, THiLoData HiLoData, TAxisInfo Ax1, TAxisInfo Ax2)
		{
			if (HiLoData.DropBars == null || HiLoData.DropBars.HiLoLines == null) return;
			if (HiLoData.Lo.Length != HiLoData.Hi.Length) return;
			if (HiLoData.Lo.Length != HiLoData.Exclude.Length) return;
			using (Pen aPen = GetCustomPen(xls, HiLoData.DropBars.HiLoLines, null, Colors.Black))
			{
				for (int i = 0; i < HiLoData.Hi.Length; i++)
				{
					if (!HiLoData.Exclude[i])Canvas.DrawLine(Ax1, Ax2, aPen, HiLoData.Lo[i].X, HiLoData.Lo[i].Y, HiLoData.Hi[i].X, HiLoData.Hi[i].Y);
				}
			}
		}

		#endregion

		#region Up/Down bars
		private static void DrawUpDownBars(ExcelFile xls, TChartCanvas Canvas, THiLoData HiLoData, TAxisInfo Ax1, TAxisInfo Ax2)
		{
			if (HiLoData.DropBars == null) return;
			if (HiLoData.First == null || HiLoData.Last == null) return;

			for (int i = 0; i < HiLoData.First.Length; i++)
			{
				if (HiLoData.ExcludeBars != null && HiLoData.ExcludeBars[i]) continue;
				if (i >= HiLoData.Last.Length) return;

				bool UpBar = HiLoData.First[i].Y > HiLoData.Last[i].Y;
				TChartOneDropBar DropBar = UpBar ? HiLoData.DropBars.UpBar: HiLoData.DropBars.DownBar;
				if (DropBar == null) continue;

				real BarW2 = (real) (HiLoData.SerieWidth / (1 + DropBar.GapWidth / 100f) / 2);

				RectangleF BarCoords = RectangleF.FromLTRB(
					HiLoData.First[i].X - BarW2, 
					Math.Min(HiLoData.First[i].Y,HiLoData.Last[i].Y), 
					HiLoData.Last[i].X + BarW2, 
					Math.Max(HiLoData.First[i].Y,HiLoData.Last[i].Y));

				using (Pen aPen = GetCustomPen(xls, DropBar.Frame.LineOptions, DropBar.Frame.ExtraOptions, Colors.Black))
				{			
					using (Brush aBrush = GetCustomBrush(xls, BarCoords, DropBar.Frame.FillOptions, DropBar.Frame.ExtraOptions))
						Canvas.DrawAndFillRectangle(Ax1, Ax2, aPen, aBrush, BarCoords.X, BarCoords.Y, BarCoords.Width, BarCoords.Height);
				}
			}
		}
		



		#endregion

		private static void GetLineCoords(TLabelDescription aLabel, TPointF Pos, RectangleF ChartCoords, RectangleF PlotCoords, bool IsArea, double Height, bool XReversed, bool YReversed)
		{
            TDataLabelPosition dlp = aLabel.FLabel.LabelOptions.Position;
            if (IsArea && dlp == TDataLabelPosition.Automatic) dlp = TDataLabelPosition.Center;
            float XDir = XReversed ? -1 : 1;
            float YDir = YReversed ? -1 : 1;

			switch (dlp)
			{
				case TDataLabelPosition.Any:
                    SetDefaultLineLabelPos(aLabel, Pos, XReversed, XDir);
                    if (TDataLabels.SetDefaultLabelBoxPos(ChartCoords, PlotCoords, aLabel, false, XDir, YDir)) return;
                    
                    aLabel.Position = TDataLabels.DefaultLabelBoxXY(ChartCoords, aLabel.FLabel);

					break;

				case TDataLabelPosition.Center:
                    double h = IsArea ? Height : 0;
					aLabel.Position = new RectangleF(Pos.X, Pos.Y + (float)(h /2.0), 0, 0);
					aLabel.HLabelPosition = THLabelPosition.Center;
					aLabel.VLabelPosition = TVLabelPosition.Center;
					break;

				case TDataLabelPosition.Above:
					aLabel.Position = new RectangleF(Pos.X, Pos.Y - LineLabelOfs, 0, 0);
					aLabel.HLabelPosition = THLabelPosition.Center;
					aLabel.VLabelPosition = TVLabelPosition.Down;
					break;

				case TDataLabelPosition.Below:
					aLabel.Position = new RectangleF(Pos.X, Pos.Y + LineLabelOfs, 0, 0);
					aLabel.HLabelPosition = THLabelPosition.Center;
					aLabel.VLabelPosition = TVLabelPosition.Up;
					break;

				case TDataLabelPosition.Left:
					aLabel.Position = new RectangleF(Pos.X - LineLabelOfs, Pos.Y, 0, 0);
					aLabel.HLabelPosition = THLabelPosition.Right;
					aLabel.VLabelPosition = TVLabelPosition.Center;
					break;

                case TDataLabelPosition.Right:
                    aLabel.Position = new RectangleF(Pos.X + LineLabelOfs, Pos.Y, 0, 0);
                    aLabel.HLabelPosition = THLabelPosition.Left;
                    aLabel.VLabelPosition = TVLabelPosition.Center;
                    break;

				default:
                    SetDefaultLineLabelPos(aLabel, Pos, XReversed, XDir);
					break;

			}
		}

        private static void SetDefaultLineLabelPos(TLabelDescription aLabel, TPointF Pos, bool XReversed, float XDir)
        {          
            aLabel.Position = new RectangleF(Pos.X + LineLabelOfs * XDir, Pos.Y, 0, 0);
            if (XReversed) aLabel.HLabelPosition = THLabelPosition.Right; else aLabel.HLabelPosition = THLabelPosition.Left;
            aLabel.VLabelPosition = TVLabelPosition.Center;
        }

        private static void DrawLineChart(ExcelFile Workbook, ExcelChart Chart, ChartSeries SeriesValues, ExcelFile xls, TChartCanvas Canvas, RectangleF Coords, TAreaLineChartOptions Options, TAxisInfo XAxis, TAxisInfo YAxis, real Zoom100, bool IsArea,
            int MaxX, double[] StackedOfs, ref TPointF[] LastPoints, TOneSeriesLabelDescription Labels, 
            THiLoData HiLoData, bool FirstSeries, bool LastSeries, TMarkerCache Markers, ref object[] FirstCategory)
        {
            if (SeriesValues == null || SeriesValues.DataValues == null) return;

            TDateAxisTransform DateAxisTransform = null;
            if (XAxis.IsDate)
            {
                if (FirstCategory == null) FirstCategory = SeriesValues.CategoriesValues; //This will only work the first time for the group of series. If the first was null, DateAxis will be false and it won't enter here no more.
                
                DateAxisTransform = new TDateAxisTransform(Workbook);
                if (FirstCategory == null || !DateAxisTransform.Fill(SeriesValues, XAxis.DateUnits, FirstCategory) && XAxis.DateAxisForced)
                {
                    XAxis.IsDate = false;
                    DateAxisTransform = null;
                }
            }

            real CWidth = Coords.Width;
            real CHeight = Coords.Height;

            int WallOfs = GetWallOffs(XAxis);

            int AreaOfsInt = IsArea ? 1 : 0;

            int StackedPointCount = 0;
            if (IsArea && Options.StackedMode != TStackedMode.None && LastPoints != null)
            {
                StackedPointCount = MaxX - 2;
            }

            bool[] Exclude = null;
            TPointF[] Points = new TPointF[MaxX + AreaOfsInt * 2 + StackedPointCount];
            int MaxPoint = 0;

            real MaxSeriesy = Coords.Top;
            real MinSeriesy = Coords.Bottom;
            bool HasCustomLineStyles = false;

            double FirstSerieWidth = CWidth / (MaxX - (1 - WallOfs));

            for (int k = 0; k < MaxX; k++)
            {
                int PointPos = DateAxisTransform == null? MaxPoint: DateAxisTransform.TransformSeriesToPoints(k);

                double BarHeight = 0;
                if (k < SeriesValues.DataValues.Length)
                {
                    if (!IsArea && Options.StackedMode == TStackedMode.None && SeriesValues.DataValues[k] == null)
                    {
                        switch (Chart.PlotEmptyCells)
                        {
                            case TPlotEmptyCells.Interpolated:
                                continue;
                            case TPlotEmptyCells.NotPlotted:
                                if (Exclude == null) Exclude = new bool[Points.Length];
                                Exclude[PointPos] = true;
                                MaxPoint++;
                                continue;
                        }
                    }
                    else
                    {
                        if (SeriesValues.DataValues[k] is double)
                        {
                            BarHeight = Convert.ToDouble(SeriesValues.DataValues[k], CultureInfo.CurrentCulture);
                        }
                    }
                }
                else
                {
                    if (!IsArea && Options.StackedMode == TStackedMode.None)
                    {
                        switch (Chart.PlotEmptyCells)
                        {
                            case TPlotEmptyCells.Interpolated:
                                continue;
                            case TPlotEmptyCells.NotPlotted:
                                if (Exclude == null) Exclude = new bool[Points.Length];
                                Exclude[PointPos] = true;
                                MaxPoint++;
                                continue;
                        }
                    }
                }
                double Stacked = StackedOfs == null ? 0 : StackedOfs[k];
                double BarTopOfs = (YAxis.Max - BarHeight) * CHeight / (YAxis.Max - YAxis.Min) - Stacked;

                if (XAxis.IsDate)
                {
					if (FirstCategory == null || !(FirstCategory[k] is double))
					{
						continue;
					}
					double xPos = DateAxisTransform.ConvertDateUnit((double)FirstCategory[k], XAxis.DateUnits);

                    if (DateAxisTransform.DateRange > 0)
                    {
                        Points[PointPos].X = GetXDate(Coords.Left, XAxis, DateAxisTransform, CWidth, WallOfs, ref FirstSerieWidth, xPos);
                    }
                    else
                    {
                        Points[PointPos].X = (Coords.Left + Coords.Right) / 2;
                    }
                }
                else
                {
                    Points[PointPos].X = Coords.Left + (real)((k + WallOfs / 2f) * FirstSerieWidth);
                }
                Points[PointPos].Y = Coords.Top + (real)(BarTopOfs);

                if (Labels != null)
                {
                    TLabelDescription Lbl = Labels[k];
                    bool AlreadyThere = Lbl != null;
                    if (!AlreadyThere) Lbl = Labels[-1];
                    if (Lbl != null && !Lbl.FLabel.LabelOptions.Deleted)
                    {
                        if (AlreadyThere)
                        {
                            GetLineCoords(Lbl, Points[PointPos], Canvas.ChartCoords, Coords, IsArea, BarTopOfs, XAxis.ReverseValues, YAxis.ReverseValues);
                            Lbl.Ax1 = XAxis;
                            Lbl.Ax2 = YAxis;
                        }
                        else
                        {
                            TLabelDescription tmp = new TLabelDescription(RectangleF.Empty, Lbl.FLabel, 0);
                            GetLineCoords(tmp, Points[PointPos], Canvas.ChartCoords, Coords, IsArea, BarHeight * CHeight / (YAxis.Max - YAxis.Min), XAxis.ReverseValues, YAxis.ReverseValues);
                            Labels.Add(k, tmp);
                            tmp.Ax1 = XAxis;
                            tmp.Ax2 = YAxis;
                        }
                    }
                }



                if (Points[PointPos].Y < MinSeriesy) MinSeriesy = Points[PointPos].Y;
                if (Points[PointPos].Y > MaxSeriesy) MaxSeriesy = Points[PointPos].Y;
                MaxPoint++;

                if (StackedOfs != null)
                {
                    StackedOfs[k] += BarHeight * CHeight / (YAxis.Max - YAxis.Min);
                }

                if (PointNeedsCustomPen(SeriesValues.Options[k]))
                {
                    HasCustomLineStyles = true;
                }

            }

            if (MaxPoint < Points.Length)
            {
                if (!IsArea)
                {
                    TPointF[] NewPoints = new TPointF[MaxPoint];
                    Array.Copy(Points, 0, NewPoints, 0, NewPoints.Length);
                    Points = NewPoints;
                }
                else
                {
                    for (int i = MaxPoint; i < MaxX; i++)
                    {
                        int z = MaxPoint > 0 ? MaxPoint - 1 : 0;
                        Points[i] = Points[z];
                    }
                }
            }

            using (Pen aPen = GetSeriesPen(xls, Options, Options.SeriesOptions, SeriesValues.Options[-1], null, SeriesValues.SeriesNumber, IsArea))
            {
                if (IsArea)
                {
                    if (Options.StackedMode != TStackedMode.None && LastPoints != null)
                    {
                        for (int n = MaxX - 1; n >= 0; n--)
                        {
                            Points[MaxX + MaxX - 1 - n].X = LastPoints[n].X;
                            Points[MaxX + MaxX - 1 - n].Y = LastPoints[n].Y;

                            if (LastPoints[n].Y < MinSeriesy) MinSeriesy = LastPoints[n].Y;
                            if (LastPoints[n].Y > MaxSeriesy) MaxSeriesy = LastPoints[n].Y;
                        }
                    }
                    else
                    {
                        double CrossPoint = YAxis.CrossPoint;

                        Points[MaxX].X = Coords.Right - (real)(WallOfs * FirstSerieWidth / 2f);
                        Points[MaxX].Y = (real)(Coords.Top + (YAxis.Max - CrossPoint) * CHeight / (YAxis.Max - YAxis.Min));
                        Points[MaxX + 1].X = Coords.Left + (real)(WallOfs * FirstSerieWidth / 2f);
                        Points[MaxX + 1].Y = Points[MaxX].Y;

                        if (Points[MaxX].Y < MinSeriesy) MinSeriesy = Points[MaxX].Y;
                        if (Points[MaxX].Y > MaxSeriesy) MaxSeriesy = Points[MaxX].Y;
                    }

                    LastPoints = Points;
                    Brush aBrush = null;
                    RectangleF Coords2 = Coords;
                    if (MinSeriesy < MaxSeriesy) Coords2 = new RectangleF(Coords.Left, MinSeriesy, Coords.Width, MaxSeriesy - MinSeriesy);

                    aBrush = GetSeriesBrush(xls, Coords2, Options, Options.SeriesOptions, SeriesValues.Options[-1], null, SeriesValues.SeriesNumber);
                    try
                    {
                        Canvas.DrawAndFillPolygon(XAxis, YAxis, aPen, aBrush, Points);
                    }
                    finally
                    {
                        if (aBrush != null) aBrush.Dispose();
                    }
                }
                else
                {
                    DrawLines(SeriesValues, xls, Canvas, Options, XAxis, YAxis, IsArea, aPen, Points, Exclude, HasCustomLineStyles, DateAxisTransform);
                    //Markers need to be drawn at the end.
                    //DrawMarkers(xls, Canvas, SeriesValues, Points, Options.SeriesOptions, YAxis, XAxis, Exclude, Zoom100);
                    Markers.Add(new TMarkerData(SeriesValues, Points, Options.SeriesOptions, XAxis, YAxis));
                }
            }

            if (Options != null && Options.DropBars != null)
            {
                DrawDropLines(xls, Canvas, Points, Options.DropBars, XAxis, YAxis, (real)(Coords.Top + (YAxis.Max - YAxis.CrossPoint) * CHeight / (YAxis.Max - YAxis.Min)));
                if (HiLoData != null)
                {
                    CalcHiLoLines(xls, Points, HiLoData, Exclude, Options.DropBars);
                }

                if (HiLoData != null)
                {
                    HiLoData.SerieWidth = FirstSerieWidth;
                    if (FirstSeries) HiLoData.First = (TPointF[])Points.Clone();
                    if (LastSeries) HiLoData.Last = (TPointF[])Points.Clone();
                    if (Exclude != null && (FirstSeries || LastSeries))
                    {
                        if (HiLoData.ExcludeBars == null) HiLoData.ExcludeBars = (bool[])Exclude.Clone();
                        else
                            for (int i = 0; i < Math.Min(HiLoData.ExcludeBars.Length, Exclude.Length); i++) if (Exclude[i]) HiLoData.ExcludeBars[i] = true;
                    }
                }
            }
        }

        private static float GetXDate(real CoordsLeft, TAxisInfo XAxis, TDateAxisTransform DateAxisTransform, float CWidth, int WallOfs, ref double FirstSerieWidth, double xPos)
        {
            double WOfs = WallOfs * TDateAxisTransform.MinDateUnit(XAxis.DateUnits);
            double SWidth = CWidth / (DateAxisTransform.DateRange + WOfs);
            FirstSerieWidth = SWidth * WOfs;
            return CoordsLeft + (real)((xPos - DateAxisTransform.MinDate + WOfs / 2) * SWidth);
        }

        private static int GetWallOffs(TAxisInfo XAxis)
        {
            bool NextToWall = false;
            if (XAxis.RangeOptions != null && !XAxis.RangeOptions.ValueAxisBetweenCategories) NextToWall = true;

            int WallOfs = NextToWall ? 0 : 1;
            return WallOfs;
        }

		private static void DrawLines(ChartSeries SeriesValues, ExcelFile xls, TChartCanvas Canvas, TChartOptions Options, TAxisInfo XAxis, TAxisInfo YAxis, bool IsArea, Pen aPen, TPointF[] Points, 
			bool[] Exclude, bool HasCustomLineStyles, TDateAxisTransform DateAxisTransform)
		{
			ChartSeriesOptions So = GetSeriesOptionsMisc(Options.SeriesOptions, SeriesValues.Options, SeriesValues.SeriesNumber);

			if (!HasCustomLineStyles && Exclude == null)  //Most common case
			{
				DrawLinePoints(Canvas, aPen, Points, Exclude, 0, Points.Length - 1, So, YAxis, XAxis);
			}
			else  //We will have to split the polyline
			{
				int k = 0;
				while (k < Points.Length - 1)
				{
					int Last = k + 1;
					/*while (!PointNeedsCustomPen(SeriesValues.Options[Last]))
					{
						Last++;
					}*/

					if (Exclude != null && (Exclude[k] || Exclude[Last])) 
					{
						k++;
						continue;
					}

					Pen aPen2 = aPen;
					//int z = DateAxisTransform.TransformPointsToSeries(k);  //We don't have to transform points to series here. SeriesValues.Options is already in "points".
					if (PointNeedsCustomPen(SeriesValues.Options[k + 1])) 
						aPen2 = GetSeriesPen(xls, Options, Options.SeriesOptions, SeriesValues.Options[-1], SeriesValues.Options[k + 1], SeriesValues.SeriesNumber, IsArea);
					try
					{
						DrawLinePoints(Canvas, aPen2, Points, Exclude, k, Last, So, YAxis, XAxis);
					}
					finally
					{
						if (aPen2 != null && aPen2 != aPen) aPen2.Dispose();
					}
					k = Last;
				}
			}
		}
        
        
		#endregion

		#region Scatter
		private static void DrawScatterChart(ExcelChart Chart, ChartSeries SeriesValues, ExcelFile xls, TChartCanvas Canvas, RectangleF Coords, TScatterChartOptions Options, TAxisInfo XAxis, TAxisInfo YAxis, real Zoom100, 
			TOneSeriesLabelDescription Labels, TMarkerCache Markers)
		{
			if (SeriesValues == null || SeriesValues.DataValues == null || SeriesValues.CategoriesValues == null) return;
			if (YAxis.Max <= YAxis.Min) return;
			if (XAxis.Max <= XAxis.Min) return;

			real CWidth = Coords.Width;
			real CHeight = Coords.Height;
			real CTop = Coords.Top;
			real CLeft = Coords.Left;

			using (Pen aPen = GetSeriesPen(xls, Options, Options.SeriesOptions, SeriesValues.Options[-1], null, SeriesValues.SeriesNumber, false))
			{
				TPointF[] Points = new TPointF[Math.Min(SeriesValues.DataValues.Length, SeriesValues.CategoriesValues.Length)];

				bool HasCustomLineStyles = false;
				int MaxPoint = 0;
				bool[] Exclude = null;

				for (int k = 0; k < Points.Length; k++)
				{
					if (SeriesValues.DataValues[k] == null || SeriesValues.CategoriesValues[k] == null)
					{
						switch (Chart.PlotEmptyCells)
						{
							case TPlotEmptyCells.Interpolated:
								continue;								
							case TPlotEmptyCells.NotPlotted:
								if (Exclude == null) Exclude = new bool[Points.Length];
								Exclude[MaxPoint] = true;
								MaxPoint++;
								continue;
						}
					}

					
					double x = 0;
					if (SeriesValues.CategoriesValues[k] is double)
					{
                        x = Convert.ToDouble(SeriesValues.CategoriesValues[k], CultureInfo.CurrentCulture);
					}

					double y = 0;
					if (SeriesValues.DataValues[k] is double)
					{
						y = Convert.ToDouble(SeriesValues.DataValues[k]);
					}

					Points[MaxPoint].X = (real) (CLeft + (x - XAxis.Min) * CWidth / (XAxis.Max - XAxis.Min));
					Points[MaxPoint].Y = (real) (CTop + (YAxis.Max - y) * CHeight / (YAxis.Max - YAxis.Min));
					
					if (Labels != null)
					{
						TLabelDescription Lbl = Labels[k];
						bool AlreadyThere = Lbl != null;
						if (!AlreadyThere) Lbl = Labels[-1];
						if (Lbl != null && !Lbl.FLabel.LabelOptions.Deleted)
						{
							if (AlreadyThere) GetLineCoords(Lbl, Points[MaxPoint], Canvas.ChartCoords, Coords, false, 0, XAxis.ReverseValues, YAxis.ReverseValues);
							else 
							{
								TLabelDescription tmp = new TLabelDescription(RectangleF.Empty, Lbl.FLabel, 0);
                                GetLineCoords(tmp, Points[MaxPoint], Canvas.ChartCoords, Coords, false, 0, XAxis.ReverseValues, YAxis.ReverseValues);
								Labels.Add(k, tmp);
							}
						}
					}

					
					MaxPoint++;

					if (PointNeedsCustomPen(SeriesValues.Options[k]))
					{
						HasCustomLineStyles = true;
					}

				}

				if (MaxPoint < Points.Length)
				{
					TPointF[] NewPoints = new TPointF[MaxPoint];
					Array.Copy(Points, 0, NewPoints, 0, NewPoints.Length);
					Points = NewPoints;
				}

				DrawLines(SeriesValues, xls, Canvas, Options, XAxis, YAxis, false, aPen, Points, Exclude, HasCustomLineStyles, new TDateAxisTransform(null));
				//Markers need to be drawn at the end.
				//DrawMarkers(xls, Canvas, SeriesValues, Points, Options.SeriesOptions, YAxis, XAxis, Exclude, Zoom100);
				Markers.Add(new TMarkerData(SeriesValues, Points, Options.SeriesOptions, XAxis, YAxis));
			}
		}			
        
		#endregion

		#region Pie Draw

		internal static TPointF[] GetSlice(double x0, double y0, double rx, double ry, double DonutRadiusx, double DonutRadiusy, double Alpha1, double Alpha2)
		{
			TPointF[] UpperArc = TEllipticalArc.GetPoints(x0, y0, rx, ry, 0, Alpha1, Alpha2);
			if (DonutRadiusx == 0 && DonutRadiusy == 0) //pie slice
			{
				TPointF[] Result = new TPointF[UpperArc.Length + 6];
				Array.Copy(UpperArc,0, Result, 3, UpperArc.Length);
				Result[0] = new TPointF((real)x0, (real)y0);
				Result[1] = Result[0];
				Result[2] = Result[3];
				
				int n = Result.Length - 3;
				Result[n] = Result[n-1];
				Result[n+1] = Result[0];
				Result[n+2] = Result[0];

				return Result;
			}
			else
			{

				TPointF[] LowerArc = TEllipticalArc.GetPoints(x0, y0, DonutRadiusx, DonutRadiusy, 0, Alpha1, Alpha2);
				TPointF[] Result = new TPointF[UpperArc.Length + LowerArc.Length + 5];
				Array.Copy(UpperArc,0, Result, 3, UpperArc.Length);
				for (int i = 0; i < LowerArc.Length; i++)
				{
					Result[Result.Length - 1 - i] = LowerArc[i];
				}
				Result[0] = LowerArc[0];
				Result[1] = Result[0];
				Result[2] = Result[3];

				int n = 3 + UpperArc.Length;
				Result[n] = Result[n-1];
				Result[n+1] = Result[n+2];

				return Result;
			}
		}

		private static double GetSliceWidth(ChartSeries value, int i)
		{
			double Result = 0;
			if (value.DataValues[i] is double)
			{
                Result = Convert.ToDouble(value.DataValues[i], CultureInfo.CurrentCulture);
				if (Result < 0) Result = -Result; //yep, negatives are not 0.
			}	
			return Result;
		}

		private static void DrawPieChart(ChartSeries SeriesValues, int GroupSeriesIndex, int GroupSeriesCount, ExcelFile xls, ExcelChart Chart, TChartCanvas Canvas, TFontCache FontCache, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, TPieChartOptions Options, real Zoom100, ref double r1, TOneSeriesLabelDescription Labels)
		{		
			if (SeriesValues == null || SeriesValues.DataValues == null) return;
			if (Options.DonutRadius == 0) //pie
			{
				if (GroupSeriesIndex > 0) return;
				DrawPieOrDonutChart(SeriesValues, xls, Chart, Canvas, Coords, ShadowInfo, Clipping, Options, Zoom100, Options.DonutRadius, 100, true, Labels);
			}
			else
			{
				double r = Options.DonutRadius;
				if (GroupSeriesIndex == 0) r1 = r;

				double r2 = r + (GroupSeriesIndex+1) / (double) GroupSeriesCount * (100 - r);
				DrawPieOrDonutChart(SeriesValues, xls, Chart, Canvas, Coords, ShadowInfo, Clipping, Options, Zoom100, r1, r2, GroupSeriesIndex == GroupSeriesCount - 1, Labels);
				r1 = r2;
			}
		}

		private static void GetSliceCenter(double x0, double y0, ChartSeriesOptions Options, double Angle, double r, out double x0a, out double y0a)
		{
			x0a = x0;
			y0a = y0;
			if (Options == null || Options.PieOptions == null) return;

			double r1 = Options.PieOptions.SliceDistance * r / 100;
			
			x0a += r1 * Math.Cos(Angle);
			y0a += r1 * Math.Sin(Angle);

		}

		private static void GetPieCoords(TLabelDescription aLabel, 
			double x0, double y0, double r, double DonutRadius, double Alpha,
			RectangleF ChartCoords)
		{
			aLabel.LeaderPoint = new TPointF((real)(x0 + r * Math.Cos(Alpha)), (real)(y0 + r * Math.Sin(Alpha)));
			TDataLabelPosition Pos = DonutRadius != 0? TDataLabelPosition.Center: aLabel.FLabel.LabelOptions.Position;

			if (Pos == TDataLabelPosition.Any)
			{
				aLabel.Position = TDataLabels.DefaultLabelBoxXY(ChartCoords, aLabel.FLabel);
				return;
			}

			double r0;
			switch (Pos)
			{
				case TDataLabelPosition.Center:
				{
					r0 = (DonutRadius + r) / 2;
					aLabel.Position= new RectangleF((real)(x0 + r0 * Math.Cos(Alpha)), (real)(y0 + r0 * Math.Sin(Alpha)), 0, 0);
					aLabel.HLabelPosition = THLabelPosition.Center;
					aLabel.VLabelPosition = TVLabelPosition.Center;
					break;
				}

				case TDataLabelPosition.Inside:
				{
					r0 = r * 4 / 5;
					aLabel.Position= new RectangleF((real)(x0 + r0 * Math.Cos(Alpha)), (real)(y0 + r0 * Math.Sin(Alpha)), 0, 0);
					aLabel.HLabelPosition = THLabelPosition.Center;
					aLabel.VLabelPosition = TVLabelPosition.Center;
					break;
				}


				default: 
				{
					r0 = r * 1.05;
					double SinAlpha = Math.Sin(Alpha);
					double CosAlpha = Math.Cos(Alpha);
					double Sin45 = Math.Sin(Math.PI / 4);
					aLabel.Position= new RectangleF((real)(x0 + r0 * CosAlpha), (real)(y0 + r0 * SinAlpha), 0, 0);
					if (Math.Abs(SinAlpha) < Sin45)
					{
						if (CosAlpha > 0) 
						{
							aLabel.HLabelPosition = THLabelPosition.Left;
						}
						else
						{
							aLabel.HLabelPosition = THLabelPosition.Right;
						}
					}
					else
					{
						aLabel.HLabelPosition = THLabelPosition.Left;
						aLabel.XOfs = (1 - CosAlpha / Sin45)/2;
					}
					if (Math.Abs(CosAlpha) < Sin45)
					{
						if (SinAlpha > 0) 
						{
							aLabel.VLabelPosition = TVLabelPosition.Up;
						}
						else
						{
							aLabel.VLabelPosition = TVLabelPosition.Down;
						}
					}
					else
					{
						aLabel.VLabelPosition = TVLabelPosition.Up;
						aLabel.YOfs = (1 - SinAlpha / Sin45)/2;
					}

					break;
				}
			}
		}


		private static void DrawPieOrDonutChart(ChartSeries SeriesValues, ExcelFile xls, ExcelChart Chart, TChartCanvas Canvas, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, TPieChartOptions Options, real Zoom100, double DonutPercent1, double DonutPercent2, bool LastInSeries, TOneSeriesLabelDescription Labels)
		{
			if (SeriesValues == null || SeriesValues.DataValues == null || SeriesValues.DataValues.Length == 0) return;
			
			double TotalWidth = 0;
			for (int i = 0; i < SeriesValues.DataValues.Length; i++)
			{
				double SliceWidth = GetSliceWidth(SeriesValues, i);
				TotalWidth += SliceWidth;
			}

			if (TotalWidth == 0) return;

			double Angle1 = (Options.FirstSliceAngle - 90) * Math.PI / 180;
			double x0 = Coords.Left + Coords.Width / 2;
			double y0 = Coords.Top + Coords.Height / 2;

			double rx0 = Math.Min(Coords.Width / 2, Coords.Height / 2); 
			double GlobalSliceDistance = 0;
			if (Options.SeriesOptions != null && Options.SeriesOptions.PieOptions != null) GlobalSliceDistance = Options.SeriesOptions.PieOptions.SliceDistance;
			double rx = rx0 / (1.0 + GlobalSliceDistance / 100.0);
			double ry = rx;

			double DonutRadius = DonutPercent1 * rx / 100;
			
			for (int i = 0; i < SeriesValues.DataValues.Length; i++)
			{
				using (Pen aPen = GetSeriesPen(xls, Options, Options.SeriesOptions, SeriesValues.Options[-1], SeriesValues.Options[i], i, true))
				{
					double SliceWidth = GetSliceWidth(SeriesValues, i);
					double Angle2 = Angle1 + (SliceWidth / TotalWidth) * 2 * Math.PI;
					double x0a; double y0a;
					ChartSeriesOptions GlobalOptions = LastInSeries? GetSeriesOptionsPie(Options.SeriesOptions, SeriesValues.Options, i): null; //The slice applies only to the last series.
					GetSliceCenter(x0, y0, GlobalOptions, (Angle1 + Angle2) / 2, rx* DonutPercent2 / 100, out x0a, out y0a);

					if (Labels != null)
					{
						TLabelDescription Lbl = Labels[i];
						bool AlreadyThere = Lbl != null;
						if (!AlreadyThere) Lbl = Labels[-1];
						if (Lbl != null && !Lbl.FLabel.LabelOptions.Deleted)
						{
							if (AlreadyThere) 
							{
								GetPieCoords(Lbl, x0a, y0a, rx * DonutPercent2 / 100, DonutRadius, (Angle1 + Angle2) / 2, Canvas.ChartCoords);
								Lbl.Percent = SliceWidth / TotalWidth;
								if (Lbl.FLabel.LabelOptions.Position == TDataLabelPosition.Any) //do not draw leader lines otherwise
								{
									Lbl.LeaderLines = Options.LeaderLines;
									Lbl.LeaderLineStyle = Options.LeaderLineStyle;
								}
							}
							else 
							{
								TLabelDescription tmp = new TLabelDescription(RectangleF.Empty, Lbl.FLabel, 0);
								GetPieCoords(tmp, x0a, y0a, rx * DonutPercent2 / 100, DonutRadius, (Angle1 + Angle2) / 2, Canvas.ChartCoords);
								Labels.Add(i, tmp);
								tmp.Percent = SliceWidth / TotalWidth;
								Lbl.LeaderLines = Options.LeaderLines;
								Lbl.LeaderLineStyle = Options.LeaderLineStyle;
							}

						}
					}


					using (Brush aBrush = GetSeriesBrush(xls, Coords, Options, Options.SeriesOptions, SeriesValues.Options[-1], SeriesValues.Options[i], i))
					{
						Canvas.DrawAndFillBeziers(null, null, aPen, aBrush, GetSlice(x0a, y0a, rx * DonutPercent2 / 100, ry * DonutPercent2 / 100, DonutRadius, DonutRadius, Angle1, Angle2));
						Angle1 = Angle2;
					}
				}
			}
		}
        
		#endregion
	}

	#region AxisInfo

	internal enum TDateUnits
	{
		Day = 0,
		Month = 1, 
		Year = 2
	}

	internal class TAxisInfo
	{
		internal double Min;
		internal double Max;
		internal double MinorScale;
		internal double MajorScale;
		internal bool NeedsPercent;
		internal bool OneNotPercent;
		internal bool Horizontal;

		internal double CrossPoint;
		internal real StartOffset;
		internal bool MaxCross;

		internal TFlxChartFont Font;
		internal string NumberFormat;
		internal string CellNumberFormat;
		internal TAxisLineOptions LineOptions;
		internal TAxisTickOptions TickOptions;
		internal TAxisRangeOptions RangeOptions;
		internal bool IsCategory;
		internal bool ReverseValues;
		internal bool Logarithmic;

		internal bool IsDate;
		internal TDateUnits DateUnits;
		internal TDateUnits DateUnitsMajor;
		internal TDateUnits DateUnitsMinor;
        internal bool DateAxisForced;

		internal object[] Labels;

		internal TDataLabel Caption;

		internal TAxisInfo(double aMin, double aMax, double aMinorScale, double aMajorScale, bool aNeedsPercent, bool aHorizontal, bool aOneNotPercent)
		{
			Min = aMin;
			Max = aMax;
			MinorScale = aMinorScale;
			MajorScale = aMajorScale;
			NeedsPercent = aNeedsPercent;
			Horizontal = aHorizontal;
			OneNotPercent = aOneNotPercent;

		}
	}
	#endregion

	#region TChartElement
	internal class TChartElement
	{
		protected ExcelFile Workbook;
		protected IFlxGraphics Canvas;
		protected ExcelChart Chart;
		protected RectangleF PlotCoords;
		protected RectangleF ChartCoords;

		protected TChartOptions[] ChartOptions;
		protected TFontCache FontCache;
		protected real Zoom100;

		internal TChartElement(ExcelFile aWorkbook, IFlxGraphics aCanvas, ExcelChart aChart, RectangleF aChartCoords, RectangleF aPlotCoords, 
			TFontCache aFontCache, real aZoom100, TChartOptions[] aChartOptions)
		{
			Canvas = aCanvas;
			Chart = aChart;
			PlotCoords = aPlotCoords;
			ChartCoords = aChartCoords;
			Workbook = aWorkbook;
			FontCache = aFontCache;
			Zoom100 = aZoom100;
			ChartOptions = aChartOptions;
		}


	}
	#endregion

	#region Legend

	internal class TItemCaption
	{
		internal string FLabel;
		internal int Series;

		internal TItemCaption(string aLabel, int aSeries)
		{
			FLabel = aLabel;
			Series = aSeries;
		}
	}

#if (FRAMEWORK20)
    internal class TItemCaptionList : List<TItemCaption>
    {
    }
#else 
	internal class TItemCaptionList: ArrayList
	{
		internal new TItemCaption this[int index]
		{
			get
			{
				return (TItemCaption) base[index];
			}
		}
	}
#endif

	internal class TLegend: TChartElement
	{
		const real wTriple = 4.5f;
		const real wSimple = 2.5f;

		private TChartLegend Legend;
		private ChartSeries[] SeriesValues;
		private TItemCaptionList ItemCaptions;

		private RectangleF Box;
		private bool TripleWidth;

		int SeriesPerCol;
		SizeF Size0;
		bool UseFirstSeries;

		internal TLegend(ExcelFile aWorkbook, IFlxGraphics aCanvas, ExcelChart aChart, RectangleF aChartCoords, RectangleF aPlotCoords, ChartSeries[] aSeriesValues,
			TFontCache aFontCache, real aZoom100, TChartOptions[] aChartOptions):
			base(aWorkbook, aCanvas, aChart, aChartCoords, aPlotCoords, aFontCache, aZoom100, aChartOptions)
		{
			Legend = Chart.GetChartLegend();
			SeriesValues = aSeriesValues;
		}

		internal RectangleF CalcLegend()
		{
			if (Legend == null) return PlotCoords;
			Box = new RectangleF(ChartCoords.Left + ChartCoords.Width * Legend.X/ 4000f, 
				ChartCoords.Top + ChartCoords.Height * Legend.Y / 4000f, 
				ChartCoords.Width * Legend.Width / 4000f, 
				ChartCoords.Height * Legend.Height / 4000f);
			

			ItemCaptions = GetItemCaptions(ChartOptions);

			switch (Legend.Placement)
			{
				case TChartLegendPos.Bottom:
					break;
			}

			CalcTextDimensions();

			return PlotCoords;
		}

		private bool IsEntryDeleted(int s, int p)
		{
			if (s < 0 || s >= SeriesValues.Length || SeriesValues[s].DataValues == null) return true; //This happens no "virtual" series like ErrorBars. We do not want legend entries for them.
			return SeriesValues[s].LegendOptions[p] != null &&  SeriesValues[s].LegendOptions[p].EntryDeleted;
		}

		private static bool IsLineType(TChartType ChartType)
		{
			return ChartType == TChartType.Line || ChartType == TChartType.Scatter;
		}

		private bool HasRealPen(TChartOptions Options, int PosInSeries)
		{
			using (Pen aPen = GetPen(PosInSeries, Options, false))
			{
				return aPen != null;
			}
		}

		private TItemCaptionList GetItemCaptions(TChartOptions[] ChartOptions)
		{
			UseFirstSeries = true;
			TripleWidth = false;
			for (int i = 0; i < SeriesValues.Length; i++)
			{
				int ib = SeriesValues[i].ChartOptionsIndex;

				//If there are only pies or dougnuths on this chart, use a legend with the first series only.
				if (ChartOptions[ib].ChartType != TChartType.Pie)
				{
					UseFirstSeries = false;
				}

				//If there is more that one chart group, use first series.
				if (ib > 0)
				{
					UseFirstSeries = false;
				}

				if (IsLineType(ChartOptions[ib].ChartType) && !IsEntryDeleted(ib, -1) && HasRealPen(ChartOptions[ib], i)) //Here UseFirstSeries is always false, since this is a line chart.
				{
					TripleWidth = true;
				}
			}

			if (UseFirstSeries)
			{
				TItemCaptionList Result = new TItemCaptionList();
				if (SeriesValues.Length > 0)
				{
					for (int i = 0; i < SeriesValues[0].DataValues.Length; i++)
					{
						if (IsEntryDeleted(0, i)) continue;
						if (SeriesValues[0].CategoriesValues == null)
						{
							Result.Add(new TItemCaption((i + 1).ToString(), i));
						}
						else
						{
							if (i < SeriesValues[0].CategoriesValues.Length)
							{
								Result.Add(new TItemCaption(FlxConvert.ToString(SeriesValues[0].CategoriesValues[i]), i));
							}
						}
					}
				}
				return Result;
			}

			//Use all series.
		{
			TItemCaptionList Result = new TItemCaptionList();
			for (int i = 0; i < SeriesValues.Length; i++)
			{
				int ib = i;
				if (IsEntryDeleted(ib, -1)) continue;

				Result.Add(new TItemCaption(GetLegendCaption(SeriesValues[ib].TitleValue, SeriesValues[ib].SeriesNumber), i));
			}
			return Result;
		}
		}

		internal static string GetLegendCaption(object TitleValue, int SeriesNumber)
		{
			if (TitleValue == null)
			{
				return FlxMessages.GetString(FlxMessage.TxtSeries) + (SeriesNumber + 1).ToString();
			}
			else
			{
				return FlxConvert.ToString(TitleValue);
			}
		}

		private void CalcTextDimensions()
		{
			SeriesPerCol = ItemCaptions.Count;
			Font aFont = GetFont();
			Size0 = Canvas.MeasureString("a", aFont);
			real w0;
			if (TripleWidth) w0 = Size0.Width * wTriple; else w0 = Size0.Width * wSimple;
			for (int i = 0; i < ItemCaptions.Count; i++)
			{
				SizeF sz = Canvas.MeasureStringEmptyHasHeight(ItemCaptions[i].FLabel, aFont);
				if (sz.Width + w0> 0)
				{
					int Spc = (int)(Box.Width / (sz.Width + w0));
					if (Spc < 1) Spc = 1;
					if (Spc < SeriesPerCol) SeriesPerCol = Spc;
				}
				if (SeriesPerCol <= 1) break;
			}

		}

		internal void Draw()
		{
			if (Legend == null) return;
			DrawFramingBox();
			if (ItemCaptions.Count <= 0) return;
			DrawTexts(FontCache, Zoom100);
		}

		internal void DrawFramingBox()
		{
			using (Pen aPen = DrawChart.GetFramePen(Workbook, Legend.Frame, Colors.Black))
			{
				using (Brush aBrush = DrawChart.GetFrameBrush(Workbook, Box, Legend.Frame, Colors.White))
				{
					Canvas.DrawAndFillRectangle(aPen, aBrush, Box.X, Box.Y, Box.Width, Box.Height);
				}
			}
		}

		internal void DrawTexts(TFontCache FontCache, real Zoom100)
		{
			if (ItemCaptions == null) return;
			if (SeriesPerCol < 1 || ItemCaptions.Count <= 0) return;

			Font aFont = GetFont();
			Color FontColor = Colors.Black;
			if (Legend.TextOptions != null && Legend.TextOptions.TextColor != null) FontColor = Legend.TextOptions.TextColor.FgColor;


			real dx = Box.Width / SeriesPerCol;
			int rows = (ItemCaptions.Count - 1) / SeriesPerCol + 1;
			real dy = Box.Height / rows;

			int row = 0;
			int col = 0;
			real MarkerWidth = TripleWidth? Size0.Width * wTriple: Size0.Width * wSimple;
			if (dx <= MarkerWidth) return;
			
			for (int i = 0; i < ItemCaptions.Count; i++)
			{
				if (ItemCaptions[i].FLabel != null) 
				{
					TXRichStringList TextLines = new TXRichStringList();
					TFloatList MaxDescent = new TFloatList();
					SizeF tz;
					RenderMetrics.SplitText(Canvas, FontCache, Zoom100, new TRichString(ItemCaptions[i].FLabel), aFont, dx - MarkerWidth, TextLines, out tz, false, MaxDescent, null);
					
					if (TextLines.Count > 0)
					{
					
						real CenterYOfs = (Box.Height / (real) rows - tz.Height) / 2;
						if (CenterYOfs < 0) CenterYOfs = 0;

						DrawMarker(Box.X + Size0.Width / 2f + dx * col, Box.Y + CenterYOfs + dy*row + (TextLines[0].YExtent - Size0.Width) / 2f, 
							Size0.Width, ItemCaptions[i].Series);

						TLegendEntryOptions LegendOptions = null;
						if (UseFirstSeries) 
						{
							LegendOptions = SeriesValues[0].LegendOptions[ItemCaptions[i].Series];
						}
						else
						{
							LegendOptions = SeriesValues[ItemCaptions[i].Series].LegendOptions[-1];
						}
						using (Brush aBrush = GetBrush(FontColor, LegendOptions))
						{
							Font bFont = GetFont(aFont, LegendOptions);
							real ddy = 0;
							for (int z = 0; z < TextLines.Count; z++)
							{
								ddy = ddy + TextLines[z].YExtent;

								if (ddy > dy)
								{
									Canvas.SaveState();
									Canvas.SetClipIntersect(new RectangleF(Box.X + col * dx, Box.Y + row * dy, Box.X + (col + 1) * dx, Box.Y + (row + 1) * dy));
								}

								Canvas.DrawString(FlxConvert.ToString(TextLines[z].s), bFont, aBrush, Box.X + MarkerWidth + dx * col, 
									Box.Y + CenterYOfs + ddy + dy * row);

								if (ddy > dy)
								{
									Canvas.RestoreState();
									break;
								}

							}
						}
					}
				}

				col++;
				if (col >= SeriesPerCol)
				{
					col = 0;
					row++;
				}
			}
		}

		private Font GetFont()
		{
			Font aFont;
			if (Legend.TextOptions != null && Legend.TextOptions.Font != null)
			{
				aFont = FontCache.GetFont(Legend.TextOptions.Font.Font, Zoom100);
			}
			else
			{				
				aFont = FontCache.GetFont(Chart.DefaultFont.Font, Zoom100);
			}
			return aFont;
		}

		private void DrawMarker(real x, real y, real w, int PosInSeries)
		{
			int Series = UseFirstSeries? 0: PosInSeries;
			int ib = SeriesValues[Series].ChartOptionsIndex;

			if (IsLineType(ChartOptions[ib].ChartType))
			{
				DrawLineMarker(x, y, w, PosInSeries, ChartOptions[ib]);
			}
			else
			{
				DrawBoxMarker(x, y, w, PosInSeries, ChartOptions[ib]);
			}
		}


		private Pen GetPen(int PosInSeries, TChartOptions Options, bool IsBorder)
		{
			int p = PosInSeries;
			ChartSeriesOptions PointOptions = null;
			int id = PosInSeries;
			if (UseFirstSeries)
			{
				p = 0;
				PointOptions = SeriesValues[0].Options[PosInSeries];
			}
			else
			{
				id = SeriesValues[p].SeriesNumber;
			}

			return DrawChart.GetSeriesPen(Workbook, Options, Options.SeriesOptions, SeriesValues[p].Options[-1], PointOptions, id, IsBorder);
		}

		private Brush GetBrush(int PosInSeries, TChartOptions Options, RectangleF Coords)
		{
			int p = PosInSeries;
			ChartSeriesOptions PointOptions = null;
			int id = PosInSeries;
			if (UseFirstSeries)
			{
				p = 0;
				PointOptions = SeriesValues[0].Options[PosInSeries];
			}
			else
			{
				id = SeriesValues[p].SeriesNumber;
			}

			return DrawChart.GetSeriesBrush(Workbook, Coords, Options, Options.SeriesOptions, SeriesValues[p].Options[-1], PointOptions, id);
		}

		private static Brush GetBrush(Color Defaultcolor, TLegendEntryOptions Options)
		{
			if (Options == null || Options.TextFormat == null || Options.TextFormat.TextColor == null) return new SolidBrush(Defaultcolor);
			if (Options.TextFormat.TextColor.Pattern == TChartPatternStyle.None) return null;
			return new SolidBrush(Options.TextFormat.TextColor.FgColor);
		}

		private Font GetFont(Font DefaultFont, TLegendEntryOptions Options)
		{
			if (Options == null || Options.TextFormat == null || Options.TextFormat.Font == null) return DefaultFont;
			TFlxChartFont f = Options.TextFormat.Font;
			return FontCache.GetFont(f.Font, Zoom100);
		}

		private void DrawLineMarker(real x, real y, real w, int PosInSeries, TChartOptions Options)
		{
			using (Pen aPen = GetPen(PosInSeries, Options, false))
			{
				Canvas.DrawLine(aPen, x + w* 0.5f, y + w / 2f, x + w * 3.5f, y + w / 2f);

				if (!UseFirstSeries)
				{
					Pen MarkerPen = null;
					Brush MarkerBrush = null;
					TMarkerImgInfo MarkerImage = null;
					double MarkerSize;
					TChartMarkerType MarkerType;

					real xoffs = TripleWidth? 2f: 1f;

					try
					{
						DrawChart.GetMarkerOptions(Workbook, SeriesValues[PosInSeries], Options.SeriesOptions, out MarkerPen, out MarkerBrush, out MarkerSize, out MarkerType, out MarkerImage, Zoom100);
						if (MarkerImage != null)
						{
							MarkerImage.Width = w;
							MarkerImage.Height = w;
						}
						DrawChart.DrawOneMarker(new TChartCanvas(Canvas, Rectangle.Empty, ChartCoords), new TPointF(x + w*xoffs, y + w / 2f), w * 20f, MarkerImage, MarkerType,
							MarkerPen, MarkerBrush, null, null);
					}
					finally
					{
						if (MarkerPen != null) MarkerPen.Dispose();
						if (MarkerBrush != null) MarkerBrush.Dispose();
						if (MarkerImage != null) MarkerImage.Dispose();
					}
				}
			}
		}

		private void DrawBoxMarker(real x, real y, real w, int PosInSeries, TChartOptions Options)
		{
			using (Pen aPen = GetPen(PosInSeries, Options, true))
			{
				real w1 = TripleWidth? w* 3f: w;
				RectangleF r = new RectangleF(x + w * 0.5f, y, w1, w);
				using (Brush aBrush = GetBrush(PosInSeries, Options, r))
				{
					Canvas.DrawAndFillRectangle(aPen, aBrush, r.X, r.Y, r.Width, r.Height);
				}
			}
		}

	}
	#endregion

	#region DataLabels
	internal class TDataLabels: TChartElement
	{
		private ChartSeries[] SeriesValues;
		private TChartCanvas ChartCanvas;

		internal TDataLabels(ExcelFile aWorkbook, IFlxGraphics aCanvas, ExcelChart aChart, RectangleF aChartCoords, RectangleF aPlotCoords,
			TFontCache aFontCache, real aZoom100, TChartOptions[] aChartOptions, ChartSeries[] aSeriesValues, TChartCanvas aChartCanvas):
			base(aWorkbook, aCanvas, aChart, aChartCoords, aPlotCoords, aFontCache, aZoom100, aChartOptions)
		{
			SeriesValues = aSeriesValues;
			ChartCanvas = aChartCanvas;
		}

		internal void Draw(TDataLabel[] Labels, TLabelDescriptionList LabelDescriptions, TAxisInfo[] XAxis, TAxisInfo[] YAxis)
		{
			if (Labels == null) return;
			for (int i = 0; i < Labels.Length; i++)
			{
				if (Labels[i].LinkedTo != TLinkOption.DataLabel)
				{
					DrawSimpleLabel(Labels[i]);
				}
			}

			for (int i = 0; i < SeriesValues.Length; i++)
			{
				TOneSeriesLabelDescription LabelArray = LabelDescriptions.GetLabel(i, SeriesValues[i].ChartOptionsIndex);
				if (LabelArray != null)
				{
					for (int p = 0; p <= LabelArray.Max; p++) 
					{
						TLabelDescription Lbl = LabelArray[p];
                        if (Lbl != null && !Lbl.FLabel.LabelOptions.Deleted && Lbl.FLabel.SeriesIndex < SeriesValues.Length)
                        {
                            TFlxChartFont DefaultLabelFont = Chart.DefaultLabelFont;
                            int si = Lbl.FLabel.SeriesIndex;
                            if (si < 0) si = i;
                            DrawLabel(Lbl.FLabel, si, Lbl.Position, p, Lbl.Percent, Lbl.HLabelPosition, Lbl.VLabelPosition,
                                Lbl.XOfs, Lbl.YOfs, Lbl, Lbl.Ax1, Lbl.Ax2, DefaultLabelFont);
                        }

					}
				}
			}
		}
		

		internal TRichString GetLabelTxt(TAxisInfo Axis, TDataLabel aLabel, int SeriesIndex, int Position, double Percent)
		{
			Color FontColor = ColorUtil.Empty;
			switch (aLabel.LabelOptions.DataType)
			{
				case TLabelDataValue.Manual:
				{
					object[] vals = aLabel.LabelValues;
					if (vals == null || vals.Length == 0) return null;
					TRichString Result = vals[0] as TRichString;
					if (Result != null) return Result;
					return TFlxNumberFormat.FormatValue(vals[0], aLabel.NumberFormat, ref FontColor, Workbook);
				}

				case TLabelDataValue.SeriesInfo:
				{
					if (aLabel.LinkedTo == TLinkOption.ChartTitle)
					{
						object val = SeriesValues[SeriesIndex].TitleValue;
						if (val != null)
						{
							string NumberFormat = String.Empty;  //This value is already formatted.
							return TFlxNumberFormat.FormatValue(val, NumberFormat, ref FontColor, Workbook);
						}
						return null;
					}

					TRichString Separator = new TRichString("; ");
                    if (aLabel.LabelOptions.Separator != null) Separator = new TRichString(aLabel.LabelOptions.Separator);
					TRichString Result = null;
					if (aLabel.LabelOptions.ShowSeriesName)
					{
						if (Result == null) Result = new TRichString(); else Result += Separator;

						object val = TLegend.GetLegendCaption(SeriesValues[SeriesIndex].TitleValue, SeriesValues[SeriesIndex].SeriesNumber);
						if (val != null)
						{
							string NumberFormat = DrawChart.GetNumberFormat(String.Empty, null, Axis);  //This value does not use the label number format.
							Result += TFlxNumberFormat.FormatValue(val, NumberFormat, ref FontColor, Workbook);
						}
					}
					if (aLabel.LabelOptions.ShowCategories)
					{
						if (Result == null) Result = new TRichString(); else Result += Separator;
						if (SeriesIndex >= 0 && SeriesIndex < SeriesValues.Length)
						{
							object[] vals = SeriesValues[SeriesIndex].CategoriesValues;
                            object FinalVal = null;
                            if (vals != null)
                            {
                                if (Position >= 0 && Position < vals.Length) FinalVal = vals[Position];
                            }
                            else
                            {
                                string ylabel;
                                DrawChart.GetCatAxisLabel(Workbook, Axis, Position, 1, out FontColor, out ylabel);
                                FinalVal = ylabel;
                            }

                            if (FinalVal != null)
                            {
								string BackupFormat = null;
								string[] Formats = SeriesValues[SeriesIndex].CategoriesFormats;
								if (Formats != null && Position >= 0 && Position < Formats.Length) BackupFormat = Formats[Position];

                                string NumberFormat = null; //weird, but values and categs are not formatted if we show percents. At least in office 2007/10
                                if (!aLabel.LabelOptions.ShowPercents) NumberFormat = DrawChart.GetNumberFormat(aLabel.NumberFormat, BackupFormat, Axis);  //This value DOES use the label number format.

								Result += TFlxNumberFormat.FormatValue(FinalVal, NumberFormat, ref FontColor, Workbook);
							}
						}
					}
					if (aLabel.LabelOptions.ShowValues)
					{
						if (Result == null) Result = new TRichString(); else Result += Separator;
						if (SeriesIndex >= 0 && SeriesIndex < SeriesValues.Length)
						{
							object[] vals = SeriesValues[SeriesIndex].DataValues;
							if (vals != null && Position >= 0 && Position < vals.Length) 
							{
								string BackupFormat = null;
								string[] Formats = SeriesValues[SeriesIndex].DataFormats;
								if (Formats != null && Position >= 0 && Position < Formats.Length) BackupFormat = Formats[Position];

                                string NumberFormat = null;
                                if (!aLabel.LabelOptions.ShowPercents) NumberFormat = DrawChart.GetNumberFormat(aLabel.NumberFormat, BackupFormat, Axis);  //This value DOES use the label number format.
					
								Result += TFlxNumberFormat.FormatValue(vals[Position], NumberFormat, ref FontColor, Workbook);
							}
						}
					}

					if (aLabel.LabelOptions.ShowPercents)
					{
						if (Result == null) Result = new TRichString(); 
						else 
						{
							if (aLabel.LabelOptions.ShowCategories && !aLabel.LabelOptions.ShowValues) Result += new TRichString("\n");
							else Result += Separator;
						}
						string NumberFormat = DrawChart.GetNumberFormat(aLabel.NumberFormat, null, Axis);
						if (NumberFormat == null || NumberFormat.Length == 0) NumberFormat = "0%";
						Result += TFlxNumberFormat.FormatValue(Percent, NumberFormat, ref FontColor, Workbook);
					}
					return Result;
				}

			}
			return null;
		}

		private RectangleF GetLabelBoundingBox(TDataLabel aLabel, real dx, real dy)
		{
			switch (aLabel.LabelOptions.Position)
			{
				default:
					return new RectangleF(ChartCoords.Left + ChartCoords.Width * aLabel.TextOptions.X/ 4000f, 
						ChartCoords.Top + ChartCoords.Height * aLabel.TextOptions.Y / 4000f, 
						dx, 
						dy);
			}

		}

		internal static real MaxLabelWidth(RectangleF ChartCoords)
		{
			return ChartCoords.Width / 5f;
		}

        internal static RectangleF DefaultLabelBoxXY(RectangleF ChartCoords, TDataLabel aLabel)
        {
            RectangleF r = new RectangleF(ChartCoords.Left + ChartCoords.Width * aLabel.TextOptions.X / 4000f,
                ChartCoords.Top + ChartCoords.Height * aLabel.TextOptions.Y / 4000f,
                MaxLabelWidth(ChartCoords),
                ChartCoords.Height); //Label box resizes with data. We need to recalculate it. 

            return r;
        }

        internal static bool SetDefaultLabelBoxPos(RectangleF ChartCoords, RectangleF PlotCoords, TLabelDescription aLabel, bool IsPie, float XDir, float YDir)
        {
            TChartLabelPosition lp = aLabel.FLabel.TextOptions.Position;

            switch (lp.TopLeftMode)
            {
                case TChartLabelPositionMode.MDFX:
                    break;
                case TChartLabelPositionMode.MDABS:
                    break;
                case TChartLabelPositionMode.MDPARENT:
                    if (IsPie)
                    {
                    }
                    else
                    {
                        aLabel.Position.X += XDir * PlotCoords.Width * lp.X1 / 1000f;
                        if (aLabel.Position.X < ChartCoords.X) aLabel.Position.X = ChartCoords.X;
                        if (aLabel.Position.X > ChartCoords.Right) aLabel.Position.X = ChartCoords.Right;
                        aLabel.Position.Y += YDir * PlotCoords.Height * lp.Y1 / 1000f;
                        if (aLabel.Position.Y < ChartCoords.Y) aLabel.Position.Y = ChartCoords.Y;
                        if (aLabel.Position.Y > ChartCoords.Bottom) aLabel.Position.Y = ChartCoords.Bottom;
                        return true;
                    }
                    break;
                case TChartLabelPositionMode.MDKTH:
                    break;
                case TChartLabelPositionMode.MDCHART:
                    break;
                default:
                    break;
            }

            return false;
        }
		

		internal const real BoxMargin = 5;


		internal void DrawSimpleLabel(TDataLabel aLabel)
		{
			if (aLabel == null) return;
			if (aLabel.LabelOptions != null && aLabel.LabelOptions.Deleted) return;

			RectangleF CellRect = GetLabelBoundingBox(aLabel, ChartCoords.Width, ChartCoords.Height);
			CellRect.Width = ChartCoords.Right - CellRect.Left - 2 * BoxMargin;
			DrawLabel(aLabel, aLabel.SeriesIndex, CellRect, 0, 0, THLabelPosition.Fixed, TVLabelPosition.Fixed, 0, 0, null, null, null, null);
		}

		internal void DrawLabel(TDataLabel aLabel, int SeriesIndex, RectangleF CellRect, int PointPosition, double Percent, THLabelPosition HPos, TVLabelPosition VPos, 
			double XMul, double YMul, TLabelDescription LeaderLineInfo, TAxisInfo Ax1, TAxisInfo Ax2, TFlxChartFont DefaultFont)
		{
			if (aLabel == null) return;
			int Rotation = 0;

			if (HPos != THLabelPosition.Fixed) CellRect.Width = MaxLabelWidth(ChartCoords);
			if (VPos != TVLabelPosition.Fixed) CellRect.Height = ChartCoords.Height;

			THFlxAlignment HJustify = THFlxAlignment.center;
			TVFlxAlignment VJustify = TVFlxAlignment.center;
			bool HAlignGeneral;

			if (aLabel.TextOptions != null)
			{
				HJustify = aLabel.TextOptions.HAlign;
				VJustify = aLabel.TextOptions.VAlign;
				Rotation = aLabel.TextOptions.Rotation;
			}

			Font aFont;
			if (aLabel.TextOptions != null && aLabel.TextOptions.Font != null)
			{
				aFont = FontCache.GetFont(aLabel.TextOptions.Font.Font, Zoom100);
			}
			else
			{
				TFlxChartFont Fnt = Chart.DefaultLabelFont;
				if (DefaultFont != null) Fnt = DefaultFont;				
				aFont = FontCache.GetFont(Fnt.Font, (real)Fnt.Scale * Zoom100);
			}

			TRichString Text = GetLabelTxt(null, aLabel, SeriesIndex, PointPosition, Percent);
			if (Text == null) return;

			bool Vertical;
			real Alpha = FlexCelRender.CalcAngle(Rotation, out Vertical);
			real SinAlpha = (real)Math.Sin(Alpha * Math.PI / 180); real CosAlpha = (real)Math.Cos(Alpha * Math.PI / 180);
			SizeF TextExtent;
			TXRichStringList TextLines;
			TFloatList MaxDescent;

			THAlign HAlign = THAlign.Center ; TVAlign VAlign = TVAlign.Center;
			FlexCelRender.GetHJustify(HJustify, ref HAlign, out HAlignGeneral);
			FlexCelRender.GetVJustify(VJustify, Alpha, ref VAlign);
			

			TextPainter.CalcTextBox(Canvas, FontCache, Zoom100, CellRect, 0, true, Alpha, Vertical, Text, aFont, null, out TextExtent, out TextLines, out MaxDescent);
			if (TextLines.Count <= 0) return;
		
			real[] X;
			real[] Y;
			RectangleF ContainingRect = TextPainter.CalcTextCoords(out X, out Y, Text, VAlign, ref HAlign, 0, Alpha, CellRect, 0, TextExtent, HAlignGeneral, Vertical, SinAlpha, CosAlpha, 
                TextLines, Workbook.Linespacing, VJustify);
			ContainingRect.Inflate(BoxMargin / 2, BoxMargin);

			//Relocate the box at the expected location.

			real XOfs = 0;
			switch (HPos)
			{
				case THLabelPosition.Right:
					XOfs = CellRect.Left - ContainingRect.Left - ContainingRect.Width; 
					CellRect.X -= ContainingRect.Width;
					break;
				case THLabelPosition.Center:
					XOfs = CellRect.Left - ContainingRect.Left - ContainingRect.Width / 2; 
					CellRect.X -= ContainingRect.Width / 2;
					break;
				default:
					XOfs = (real)(CellRect.Left - ContainingRect.Left - ContainingRect.Width * XMul); 
					CellRect.X -= (real)(ContainingRect.Width * XMul);
					break;
			}

			real YOfs = 0;
			switch (VPos)
			{
				case TVLabelPosition.Down:
					YOfs = CellRect.Top - ContainingRect.Top - ContainingRect.Height;
					CellRect.Y -= ContainingRect.Height; 
					break;
				case TVLabelPosition.Center:
					YOfs = CellRect.Top - ContainingRect.Top - ContainingRect.Height / 2; 
					CellRect.Y -= ContainingRect.Height / 2;
					break;
				default:
					YOfs = (real)(CellRect.Top - ContainingRect.Top - ContainingRect.Height * YMul) ; 
					CellRect.Y -= (real)(ContainingRect.Height * YMul);
					break;
			}

			TextPainter.RelocateBox(ref ContainingRect, X, Y, XOfs, YOfs);

			if ((Ax1 != null && Ax1.ReverseValues) || (Ax2 != null && Ax2.ReverseValues)) 
			{
				ChartCanvas.Mirror(Ax1, Ax2, ref CellRect);
				ChartCanvas.MirrorXY(Ax1, Ax2, ref ContainingRect, ref X, ref Y, TextLines, SinAlpha, CosAlpha);
			}



			DrawLabelBkg(ContainingRect, aLabel);

			Color FontColor = Colors.Black;
			if (aLabel.TextOptions != null && aLabel.TextOptions.TextColor != null) FontColor = aLabel.TextOptions.TextColor.FgColor;

			Canvas.SaveState();
			TextPainter.DrawRichText(Workbook, Canvas, FontCache, Zoom100, true, ref CellRect, ref ChartCoords, ref CellRect, ref ContainingRect, 0, HJustify, VJustify, Alpha,
				FontColor, new TSubscriptData(TFlxFontStyles.None), Text, TextExtent, TextLines, aFont, MaxDescent, X, Y);
			Canvas.RestoreState();

			if (LeaderLineInfo != null && LeaderLineInfo.LeaderLines) DrawLeaderLine(LeaderLineInfo, ContainingRect);

		}

		private void DrawLabelBkg(RectangleF R, TDataLabel aLabel)
		{
			using (Brush aBrush = DrawChart.GetFrameBrush(Workbook, R, aLabel.Frame, ColorUtil.Empty))
			{
				using (Pen aPen = DrawChart.GetFramePen(Workbook, aLabel.Frame, ColorUtil.Empty))
				{
					Canvas.DrawAndFillRectangle(aPen, aBrush, R.X , R.Y , R.Width , R.Height);
				}
			}
		}

		private static double DistanceSquared(TPointF p1, TPointF p2)
		{
			return (p1.X - p2.X) * (p1.X - p2.X) + (p1.Y - p2.Y) * (p1.Y - p2.Y);
		}

		private void DrawLeaderLine(TLabelDescription Lbl, RectangleF TextRect)
		{
			TPointF NearestPoint = new TPointF(TextRect.Left, (TextRect.Top + TextRect.Bottom) / 2);
			TPointF BreakPoint = new TPointF(NearestPoint.X - DrawChart.LeaderLinesBreakOfs, NearestPoint.Y);

			TPointF NewPoint = new TPointF(TextRect.Right, (TextRect.Top + TextRect.Bottom) / 2);
			if (DistanceSquared(NewPoint, Lbl.LeaderPoint) < DistanceSquared(NearestPoint, Lbl.LeaderPoint)) 
			{
				NearestPoint = NewPoint;
				BreakPoint = new TPointF(NearestPoint.X + DrawChart.LeaderLinesBreakOfs, NearestPoint.Y);
			}

			NewPoint = new TPointF((TextRect.Right + TextRect.Left) / 2, TextRect.Top);
			if (DistanceSquared(NewPoint, Lbl.LeaderPoint) < DistanceSquared(NearestPoint, Lbl.LeaderPoint)) 
			{
				NearestPoint = NewPoint;
				BreakPoint = new TPointF(NearestPoint.X, NearestPoint.Y - DrawChart.LeaderLinesBreakOfs);
			}

			NewPoint = new TPointF((TextRect.Right + TextRect.Left) / 2, TextRect.Bottom);
			if (DistanceSquared(NewPoint, Lbl.LeaderPoint) < DistanceSquared(NearestPoint, Lbl.LeaderPoint))
			{
				NearestPoint = NewPoint;
				BreakPoint = new TPointF(NearestPoint.X, NearestPoint.Y + DrawChart.LeaderLinesBreakOfs);
			}


			//if (a) return;
			using (Pen aPen = DrawChart.GetCustomPen(Workbook, Lbl.LeaderLineStyle, null))
			{
				TPointF[] Points = new TPointF[3];
				Points[0] = Lbl.LeaderPoint;
				Points[1] = BreakPoint;
				Points[2] = NearestPoint;
				Canvas.DrawLines(aPen, Points);
			}
		}
	}

    internal struct TLabelPos
    {
        int SeriesIndex;
        int ChartOptionIndex;

        public TLabelPos(int aSeriesIndex, int aChartOptionIndex)
        {
            SeriesIndex = aSeriesIndex;
            ChartOptionIndex = aChartOptionIndex;
        }
    }

	internal class TLabelDescriptionList:
#if(FRAMEWORK20)
		Dictionary<TLabelPos, TOneSeriesLabelDescription>
#else
		Hashtable
#endif
	{
		internal TLabelDescriptionList(TDataLabel[] aSeriesLabels, TChartOptions[] ChartOptions)
		{
			for (int i = 0; i < ChartOptions.Length; i++)
			{
				if (ChartOptions[i].DefaultLabel != null) // && ChartOptions[i].DefaultLabel.DataPointIndex == -1) 
				{
					ChartOptions[i].DefaultLabel.DataPointIndex = -1;
                    ChartOptions[i].DefaultLabel.SeriesIndex = -1;
					AddLabel(-1, new TLabelDescription(RectangleF.Empty, ChartOptions[i].DefaultLabel, 0), i);
				}
			}

			foreach (TDataLabel lbl in aSeriesLabels)
			{
				if (lbl.LinkedTo == TLinkOption.DataLabel)
				{
					AddLabel(lbl.SeriesIndex, new TLabelDescription(RectangleF.Empty, lbl, 0), -1);
				}
			}
		}

		private void AddLabel(int SeriesIndex, TLabelDescription Lbl, int ChartOptions)
		{
			TOneSeriesLabelDescription Series = GetLabel(SeriesIndex, ChartOptions);
			if (Series == null)
			{
				Series = new TOneSeriesLabelDescription();
				Add(new TLabelPos(SeriesIndex, ChartOptions), Series);
			}

			Series.Add(Lbl.FLabel.DataPointIndex, Lbl);
		}

#if (FRAMEWORK20)
		internal TOneSeriesLabelDescription GetLabel(int SeriesNumber, int ib)
		{
			TOneSeriesLabelDescription Result;
            if (TryGetValue(new TLabelPos(SeriesNumber, -1), out Result)) return Result;
            if (TryGetValue(new TLabelPos(-1, ib), out Result) && Result != null)
            {
                TOneSeriesLabelDescription R2 = new TOneSeriesLabelDescription();
                foreach (int k in Result.Keys)
                {
                    TLabelDescription ld = (TLabelDescription)Result[k].Clone();
                    if (ld.FLabel != null) ld.FLabel.SeriesIndex = SeriesNumber;
                    R2.Add(k, ld);                    
                }
                Result = R2;
                Add(new TLabelPos(SeriesNumber, -1), Result);
                return Result;
            }

		    return null;
		}
#else
		internal TOneSeriesLabelDescription GetLabel(int SeriesNumber, int ib)
		{
			TOneSeriesLabelDescription Result = (TOneSeriesLabelDescription) this[SeriesNumber];
			if (Result != null) return Result;
			
			return (TOneSeriesLabelDescription) this[-10 -ib];
		}
#endif
	}

	internal class TOneSeriesLabelDescription:
#if(FRAMEWORK20)
		Dictionary<int, TLabelDescription>
#else
		Hashtable
#endif
	{
		internal int Max;

#if (FRAMEWORK20)
		internal new TLabelDescription this[int PointNumber]
		{
			get
			{
				TLabelDescription Result;
				if (!TryGetValue(PointNumber, out Result)) return null;
				return Result;
			}
		}

		internal new void Add(int Point, TLabelDescription d)

#else
		internal TLabelDescription this[int PointNumber]
		{
			get
			{
				return (TLabelDescription) base[PointNumber];
			}
		}

		internal void Add(int Point, TLabelDescription d)
#endif

		{
			if (Point > Max) Max = Point;
			base.Add(Point, d);
		}
	}

	internal enum THLabelPosition
	{
		Fixed,
		Left,
		Right,
		Center
	}

	internal enum TVLabelPosition
	{
		Fixed,
		Up,
		Down,
		Center
	}

	internal class TLabelDescription: ICloneable
	{
		internal RectangleF Position;
		internal TDataLabel FLabel;
		internal double Percent;
		internal bool LeaderLines;
		internal ChartLineOptions LeaderLineStyle;
		internal THLabelPosition HLabelPosition;
		internal TVLabelPosition VLabelPosition;
		internal double XOfs;
		internal double YOfs;
		internal TPointF LeaderPoint;
		internal TAxisInfo Ax1;
		internal TAxisInfo Ax2;

		internal TLabelDescription(RectangleF aPosition, TDataLabel aLabel, double aPercent)
		{
			Position = aPosition;
			FLabel = aLabel;
			Percent = aPercent;
			HLabelPosition = THLabelPosition.Fixed;
			VLabelPosition = TVLabelPosition.Fixed;
		}


        #region ICloneable Members

        public object Clone()
        {
            TLabelDescription Result = (TLabelDescription)MemberwiseClone();
            if (FLabel != null) Result.FLabel = (TDataLabel)FLabel.Clone();
            if (LeaderLineStyle != null) Result.LeaderLineStyle = (ChartLineOptions)LeaderLineStyle.Clone();
            //Axis won't be cloned.
            return Result;
        }

        #endregion
    }

	#endregion

	#region TMarkerImgInfo

	internal class TMarkerImgInfo: IDisposable
	{
		internal Image Img;
		internal real Height;
		internal real Width;


		#region IDisposable Members

		public void Dispose()
		{
			if (Img != null) Img.Dispose();
            GC.SuppressFinalize(this);
        }

		#endregion

	}
	#endregion

	#region THiLoData
	class THiLoData
	{
		internal TPointF[] Hi;
		internal TPointF[] Lo;

		internal TPointF[] First;
		internal TPointF[] Last;
		internal bool[] Exclude;
		internal bool[] ExcludeBars;

		internal double SerieWidth;

		internal TChartDropBars DropBars;
	}
	#endregion

	#region MarkerCache
	internal struct TMarkerData
	{
		internal ChartSeries SeriesValues; 
		internal TPointF[] Points;
		internal ChartSeriesOptions SeriesOptions;
		internal TAxisInfo XAxis;
		internal TAxisInfo YAxis;

		internal TMarkerData(ChartSeries aSeriesValues, TPointF[] aPoints, ChartSeriesOptions aSeriesOptions, TAxisInfo aXAxis, TAxisInfo aYAxis)
		{
			SeriesValues = aSeriesValues;
			Points = aPoints;
			SeriesOptions = aSeriesOptions;
			XAxis =aXAxis;
			YAxis = aYAxis;
		}
	}

	internal class TMarkerCache
	{
		private List<TMarkerData> FList;

		internal TMarkerCache()
		{
			FList = new List<TMarkerData>();
		}
		
		public void Add(TMarkerData Marker)
		{
			FList.Add(Marker);
		}

		public TMarkerData this[int index] {get {return FList[index];}}

		public int Count {get {return FList.Count;}}
	}
	#endregion

    #region DateAxisAux
    internal struct TDateAxisPos: IComparable
    {
        internal int Position;
        internal double Value;

        internal TDateAxisPos(int aPosition, double aValue)
        {
            Position = aPosition;
            Value = aValue;
        }

        #region IComparable Members

        public int CompareTo(object obj)
        {
            return Value.CompareTo(((TDateAxisPos)obj).Value);
        }

        #endregion
    }

    internal class TDateAxisTransform
    {
        int[] SeriesToPoints;
        int[] PointsToSeries;
        ExcelFile Workbook;
        internal double MinDate;
        internal double MaxDate;

        internal TDateAxisTransform(ExcelFile aWorkbook)
        {
            Workbook = aWorkbook;
        }

        internal bool Fill(ChartSeries Series, TDateUnits DateUnits, object[] FirstCategoryValues)
        {
            if (FirstCategoryValues == null) return false;
            bool Result = true;

#if (FRAMEWORK20)
            List<TDateAxisPos> Points = new List<TDateAxisPos>(FirstCategoryValues.Length);
#else
			ArrayList Points = new ArrayList(FirstCategoryValues.Length);
#endif
            bool First = true;
            for (int i = 0; i < FirstCategoryValues.Length; i++)
            {
                if (FirstCategoryValues[i] != null)
                {
                    if (!(FirstCategoryValues[i] is double))
                    {
                        Result = false;
                        continue;
                    }
                    double d = (double)FirstCategoryValues[i];
                    Points.Add(new TDateAxisPos(i, d));

                    if (First)
                    {
                        MinDate = ConvertDateUnit(d, DateUnits);
                        MaxDate = MinDate;
						First = false;
                    }
                    else
                    {
                        double z = ConvertDateUnit(d, DateUnits);
                        if (z < MinDate) MinDate = z;
                        if (z > MaxDate) MaxDate = z;
                    }
                }
            }

            Points.Sort();
			SeriesToPoints = new int[FirstCategoryValues.Length];
			PointsToSeries = new int[FirstCategoryValues.Length];
            for (int i = 0; i < Points.Count; i++)
            {
                PointsToSeries[i] = ((TDateAxisPos)Points[i]).Position;
                SeriesToPoints[((TDateAxisPos)Points[i]).Position] = i;
            }

            return Result;
        }

        internal int TransformPointsToSeries(int position)
        {
            if (PointsToSeries == null) return position;
            if (position >= PointsToSeries.Length) return position;
            return PointsToSeries[position];
        }

        internal int TransformSeriesToPoints(int position)
        {
            if (SeriesToPoints == null) return position;
            if (position >= SeriesToPoints.Length) return position;
            return SeriesToPoints[position];
        }

        internal double DateRange { get { return MaxDate - MinDate; } }

        internal double ConvertDateUnit(double d, TDateUnits DateUnits)
        {
            switch (DateUnits)
            {
                case TDateUnits.Day:
                    return Math.Floor(d);

                case TDateUnits.Month:
                    {
                        DateTime dt = FlxDateTime.FromOADate(d, Workbook.OptionsDates1904);
                        return FlxDateTime.ToOADate(new DateTime(dt.Year, dt.Month, 1), Workbook.OptionsDates1904);
                    }

                case TDateUnits.Year:
                    {
                        DateTime dt = FlxDateTime.FromOADate(d, Workbook.OptionsDates1904);
                        return FlxDateTime.ToOADate(new DateTime(dt.Year, 1, 1), Workbook.OptionsDates1904);
                    }
            }

            return d;
        }

        internal static double AddDateUnit(double d, TDateUnits DateUnits, int Number, bool Dates1904)
        {
            switch (DateUnits)
            {
                case TDateUnits.Day:
                    return FlxDateTime.ToOADate(FlxDateTime.FromOADate(d, Dates1904).AddDays(Number), Dates1904);

                case TDateUnits.Month:
                        return FlxDateTime.ToOADate(FlxDateTime.FromOADate(d, Dates1904).AddMonths(Number), Dates1904);

                case TDateUnits.Year:
                        return FlxDateTime.ToOADate(FlxDateTime.FromOADate(d, Dates1904).AddYears(Number), Dates1904);
            }

            return d;
        }


        internal static int MinDateUnit(TDateUnits DateUnits)
        {
            switch (DateUnits)
            {
                case TDateUnits.Month:
                    return 30;
                case TDateUnits.Year:
                    return 365;
            }
            return 1; // 1 day.
        }
    }

    #endregion

}
