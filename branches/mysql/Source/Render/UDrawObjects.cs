using System;
using System.Text;
using System.IO;
using System.Globalization;

using FlexCel.Core;

#if (MONOTOUCH)
using System.Drawing;
using real = System.Single;
using Color = MonoTouch.UIKit.UIColor;
#else
#if (WPF)
	using RectangleF = System.Windows.Rect;
	using PointF = FlexCel.Core.TPointF;
	using real = System.Double;
	using System.Windows.Media;
	#else
	using real = System.Single;
	using System.Drawing;
	using System.Drawing.Drawing2D;
	using System.Drawing.Imaging;
	using System.Drawing.Text;
	#endif
#endif

namespace FlexCel.Render
{

	/// <summary>
	/// Some methods that might be implemented different depending on where we are rendering the images.
	/// </summary>
	internal interface IDrawObjectMethods
	{
		/// <summary>
		/// Used to know where to place the image in the canvas. 
		/// </summary>
		RectangleF GetImageRectangle(TXlsCellRange PagePrintRange, RectangleF PaintClipRect, TClientAnchor A);

		/// <summary>
		/// This does not make sense when rendering text in images, so the implementation can be empty.
		/// </summary>
		/// <param name="aCol"></param>
		/// <param name="aRow"></param>
		/// <param name="SpawnedCells"></param>
		/// <param name="HAlign"></param>
		/// <param name="TextRect"></param>
		/// <param name="ContainingRect"></param>
		void SpawnCell(int aCol, int aRow, TSpawnedCellList SpawnedCells, THAlign HAlign, ref RectangleF TextRect, ref RectangleF ContainingRect);
	}

    /// <summary>
    /// This is the class that will render any object (images, charts, autoshapes...) inside a given IFlxCanvas.
    /// </summary>
    internal class TDrawObjects
    {
        #region Privates
        private ExcelFile FWorkbook;

        private IFlxGraphics Canvas = null;
		private TFontCache FontCache;
		private real Zoom100;
		private THidePrintObjects FHidePrintObjects;

		private bool ReverseRightToLeftStrings;

		private IDrawObjectMethods SheetMethods;

        
		#endregion

		#region Constants
		internal const real ChartMargin = 5;
		#endregion

		#region Constructor
		internal TDrawObjects(ExcelFile aWorkbook, IFlxGraphics aCanvas, TFontCache aFontCache, real aZoom100, 
			THidePrintObjects aHidePrintObjects, bool aReverseRightToLeftStrings, IDrawObjectMethods aSheetMethods)
		{
			FWorkbook = aWorkbook;
			Canvas = aCanvas;
			FontCache = aFontCache;
			Zoom100 = aZoom100;
			FHidePrintObjects = aHidePrintObjects;
			ReverseRightToLeftStrings = aReverseRightToLeftStrings;
			SheetMethods = aSheetMethods;


		}
		#endregion

        #region Images

        private static bool SwapCoords(ref RectangleF Coords, real Rotation)
        {
            while (Rotation < 0) Rotation += 360;
            Rotation %= 360;
			
            if ((Rotation >= 45 && Rotation < 135) || (Rotation >= 225 && Rotation < 315))
            {
                //Not sure why, but Excel inverts x-y coordinates on this case.
                real ZeroX = Coords.X + Coords.Width / 2f;
                real ZeroY = Coords.Y + Coords.Height / 2f;
                real X = Coords.Width / 2f;
                real Y = Coords.Height / 2f;
				
                Coords.X = ZeroX - Y;
                Coords.Y = ZeroY - X;

                real w = Coords.Width;
                Coords.Width = Coords.Height;
                Coords.Height = w;
				return true;
            }
			return false;
        }

        private bool CalcImageRect(TClientAnchor Anchor, TXlsCellRange PagePrintRange, RectangleF PaintClipRect, real Rotation, out RectangleF ImageRect)
        {
            ImageRect = RectangleF.Empty;
            TClientAnchor A = Anchor;

			if (A.ChartCoords)
			{
				ImageRect = RectangleF.FromLTRB(
					PaintClipRect.Left + ChartMargin + (PaintClipRect.Width - 2 * ChartMargin) * A.Col1 / 4000f,
					PaintClipRect.Top + ChartMargin + (PaintClipRect.Height - 2 * ChartMargin) * A.Row1 / 4000f,
					PaintClipRect.Left + ChartMargin + (PaintClipRect.Width - 2 * ChartMargin) * A.Col2 / 4000f,
					PaintClipRect.Top + ChartMargin + (PaintClipRect.Height - 2 * ChartMargin) * A.Row2 / 4000f);
			}
			else
			{
				if (A.Col1 > PagePrintRange.Right || A.Row1 > PagePrintRange.Bottom) return false;
				if (A.Col2 < PagePrintRange.Left || A.Row2 < PagePrintRange.Top) return false; //Sometimes due to rounding errors we migth ending up allowing a picture where there should be none.

				ImageRect = SheetMethods.GetImageRectangle(PagePrintRange, PaintClipRect, A);
			}

            CalcChildRect(Anchor, ref ImageRect);
			SwapCoords(ref ImageRect, Rotation);
			//the bounds are for the already rotated image. The below line does not apply
			//if (ImageRect.Bottom < PaintClipRect.Top || ImageRect.Right < PaintClipRect.Left) return false;
			return true;
        }

        private static void CalcChildRect(TClientAnchor Anchor, ref RectangleF ImageRect)
        {
            if (Anchor.ChildAnchor != null) //This rect will be always inside the big one
            {
                double Dx1 = Anchor.ChildAnchor.Dx1;
                double Dy1 = Anchor.ChildAnchor.Dy1;
                double Dx2 = Anchor.ChildAnchor.Dx2;
                double Dy2 = Anchor.ChildAnchor.Dy2;

                real w = ImageRect.Width;
                real h = ImageRect.Height;
                ImageRect.Width = (real)(w * (Dx2 - Dx1));
                ImageRect.Height = (real)(h * (Dy2 - Dy1));
                ImageRect.X = (real)(ImageRect.Left + w * Dx1);
                ImageRect.Y = (real)(ImageRect.Top + h * Dy1);
            }
        }


        public void DrawImages(TShapePropertiesList ShapesInPage, TXlsCellRange PagePrintRange, RectangleF PaintClipRect, bool RenderNotPrintable)
        {
			
            if ((FHidePrintObjects & THidePrintObjects.Images)!=0) return;
            
            for (int k = 0; k < ShapesInPage.Count; k++)
            {
                TShapeProperties ShProp = ShapesInPage[k];
                int i = ShProp.zOrder;

                if ((ShProp.Print || RenderNotPrintable) && ShProp.Visible && ShProp.ObjectType != TObjectType.Comment)
                {
                    bool HasObscuredShadow = false;
                    bool dummy = false;
                    try
                    {
                        TShadowInfo ShadowInfo = new TShadowInfo(TShadowStyle.Obscured, 1);
                        DrawObject (i, ShProp, PagePrintRange, PaintClipRect, ShadowInfo, TClippingStyle.Exclude, RectangleF.Empty, false, ref HasObscuredShadow, RenderNotPrintable, false);
                        ShadowInfo.Style = TShadowStyle.Normal;
                        TShadowType ShadowType = TShadowType.Offset;
                        if (ShProp.ShapeOptions != null) ShadowType = (TShadowType)ShProp.ShapeOptions.AsLong(TShapeOption.shadowType, 0);
                        if (ShadowType == TShadowType.EmbossOrEngrave || ShadowType == TShadowType.Double)
                        {
                            ShadowInfo.Pass = 2;
                            DrawObject (i, ShProp, PagePrintRange, PaintClipRect, ShadowInfo, TClippingStyle.None, RectangleF.Empty, false, ref dummy, RenderNotPrintable, false);
                        }
                        ShadowInfo.Pass = 1;
                        DrawObject (i, ShProp, PagePrintRange, PaintClipRect, ShadowInfo, TClippingStyle.None, RectangleF.Empty, false, ref dummy, RenderNotPrintable, false);

                        
                    }
                    finally
                    {
                        if (HasObscuredShadow) Canvas.RestoreState();
                    }

					bool DoHyperlinks = (FHidePrintObjects & THidePrintObjects.Hyperlynks)==0;
					DrawObject (i, ShProp, PagePrintRange, PaintClipRect, new TShadowInfo(TShadowStyle.None, 1), TClippingStyle.None, RectangleF.Empty, false, ref dummy, RenderNotPrintable, DoHyperlinks);
                }
            }
        }

        private void DoAttributes(TXlsCellRange PagePrintRange, RectangleF PaintClipRect, TShapeProperties ShProp, RectangleF ParentCoords, bool HasParentCoords, out RectangleF Coords, out real Rotation, out bool HasCoords, out bool Rotated)
        {
            Rotated = false;
            Coords = new RectangleF();
            Rotation = 0;
            if (ShProp.ShapeOptions != null) 
            {
                Rotation = ShProp.ShapeOptions.As1616(TShapeOption.Rotation, FlxConsts.DefaultRotation); 
                if (ShProp.FlipH ^ ShProp.FlipV) Rotation = - Rotation;
            }
            HasCoords = false;
			if (ShProp.Anchor != null)
			{
				if (HasParentCoords)
				{
					HasCoords = true;
					Coords = ParentCoords;
					CalcChildRect(ShProp.Anchor, ref Coords);
                    SwapCoords(ref Coords, Rotation);
                }
				else
				{
					HasCoords = CalcImageRect(ShProp.Anchor, PagePrintRange, PaintClipRect, Rotation, out Coords);
				}
			}

            if (Rotation != 0 && HasCoords)
            {
                Rotated = true;
            }
        }

        private void RotateXY(RectangleF Coords, real Rotation)
        {
            Canvas.SaveTransform();

            Canvas.Rotate((Coords.Left + Coords.Right) /2f,
                (Coords.Top + Coords.Bottom) /2f,
                -Rotation);
        }

		private static RectangleF RotateCoords(RectangleF Coords, IFlxGraphics Canvas)
		{
			TPointF p1 = Canvas.Transform(new TPointF(Coords.Left, Coords.Top));
			TPointF p2 = Canvas.Transform(new TPointF(Coords.Right, Coords.Top));
			TPointF p3 = Canvas.Transform(new TPointF(Coords.Left, Coords.Bottom));
			TPointF p4 = Canvas.Transform(new TPointF(Coords.Right, Coords.Bottom));

			double xmin = Math.Min(Math.Min(Math.Min(p1.X, p2.X), p3.X), p4.X);
			double xmax = Math.Max(Math.Max(Math.Max(p1.X, p2.X), p3.X), p4.X);

			double ymin = Math.Min(Math.Min(Math.Min(p1.Y, p2.Y), p3.Y), p4.Y);
			double ymax = Math.Max(Math.Max(Math.Max(p1.Y, p2.Y), p3.Y), p4.Y);

			
			return RectangleF.FromLTRB((real)xmin, (real)ymin, (real)xmax, (real)ymax);
		}

        internal static void OffsetCoords(real Zoom100, TShapeProperties ShProp, ref RectangleF Coords, int Pass)
        {
            TShadowType ShadowType = (TShadowType)ShProp.ShapeOptions.AsLong(TShapeOption.shadowType, 0);
            unchecked
            {
                int ShadowX = (int)ShProp.ShapeOptions.AsSignedLong(TShapeOption.shadowOffsetX, 25400);
                int ShadowY = (int)ShProp.ShapeOptions.AsSignedLong(TShapeOption.shadowOffsetY, 25400);

                if (Pass == 2)
                {
                    if (ShadowType == TShadowType.EmbossOrEngrave || ShadowType == TShadowType.Double)
                    {
                        ShadowX = (int)ShProp.ShapeOptions.AsSignedLong(TShapeOption.shadowSecondOffsetX, -ShadowX);
                        ShadowY = (int)ShProp.ShapeOptions.AsSignedLong(TShapeOption.shadowSecondOffsetY, -ShadowY);
                    }
                }


                switch (ShadowType)
                {
                    case TShadowType.Rich:
                        unchecked
                        {
                            real ShadowScaleXToX = ShProp.ShapeOptions.As1616(TShapeOption.shadowScaleXToX, 1);
                            real ShadowScaleYToY = ShProp.ShapeOptions.As1616(TShapeOption.shadowScaleYToY, 1);
                        
                            real dx = ShadowX/ 12700f * Zoom100; 
                            real dy = ShadowY / 12700f * Zoom100;

                            real dxo = ShProp.ShapeOptions.As1616(TShapeOption.shadowOriginX, 0);
                            if (dxo != -0.5) dx += Coords.Width * (1 - ShadowScaleXToX);
                            
                            real dyo = ShProp.ShapeOptions.As1616(TShapeOption.shadowOriginY, 0);
                            if (dyo != -0.5) dy +=  Coords.Height * (1 - ShadowScaleYToY);

                            Coords.X += dx;
                            Coords.Y += dy;
                            Coords.Width *= ShadowScaleXToX;
                            Coords.Height *= ShadowScaleYToY;
                        }
                        break;

                    default:
                        Coords.X += ShadowX / 12700f * Zoom100;
                        Coords.Y += ShadowY / 12700f * Zoom100;
                        break;
                }
            }

			//Coords might be negative here. it is ok, since for example a triangle might be inverted.
        }

		private void DrawObject(int ObjPos, TShapeProperties ShProp, TXlsCellRange PagePrintRange, RectangleF PaintClipRect, TShadowInfo ShadowInfo, 
			TClippingStyle Clipping, RectangleF ParentCoords, bool HasParentCoords, ref bool HasObscuredShadow, bool RenderNotPrintable, bool DoHyperlinks)
		{
			RectangleF Coords = new RectangleF(0,0,0,0);
			RectangleF TextCoords;
			bool HasCoords = false;
			real Rotation = 0;
			bool Rotated = false;
			bool Draw = true;

			//TShapeOption.fshadowObscured, false, 1 is actually TShapeOption.fShadow.
			if (ShadowInfo.Style == TShadowStyle.None || (ShProp.ShapeOptions != null && ShProp.ShapeOptions.AsBool(TShapeOption.fshadowObscured, false, 1)))
			{
				DoAttributes (PagePrintRange, PaintClipRect, ShProp, ParentCoords, HasParentCoords, out Coords, out Rotation, out HasCoords, out Rotated);
			}

			if (HasCoords && ShadowInfo.Style != TShadowStyle.None)
			{
				if (Clipping == TClippingStyle.None) OffsetCoords(Zoom100, ShProp, ref Coords, ShadowInfo.Pass);

				if (ShProp.ShapeOptions != null && ShProp.ShapeOptions.AsBool(TShapeOption.fshadowObscured, false, 0))
				{
					ShadowInfo.Style = TShadowStyle.Obscured;
				}       
				else
					Draw = (Clipping == TClippingStyle.None);
			}

			if (Rotated) RotateXY(Coords, Rotation);

			TextCoords = Coords;
			try
			{
				if (HasCoords && Draw) 
				{
					if (!HasObscuredShadow && Clipping != TClippingStyle.None)
					{
						HasObscuredShadow = true;
						Canvas.SaveState();
					}

					if (ShProp.ShapeType >= TShapeType.TextPlainText && ShProp.ShapeType <= TShapeType.TextCanDown)
					{
						DrawWordArt.DrawPlainText(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
					}

                    if (ShProp.ShapeType != TShapeType.PictureFrame && ShProp.ShapeType != TShapeType.HostControl && ShProp.ShapeGeom != null)
                    {
                        TextCoords = DrawShape2007.DrawCustomShape(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    }
                    else
                    {
                        DrawBiffShape(ObjPos, ShProp, ShadowInfo, Clipping, RenderNotPrintable, ref Coords, ref TextCoords);
                    }

                    if (DoHyperlinks && ShadowInfo.Style == TShadowStyle.None)
					{
						THyperLink hl = ShProp.ShapeOptions.AsHyperLink(TShapeOption.pihlShape, null);
						if (hl != null && hl.Text != null &&hl.Text.Length > 0)
						{
							RectangleF CoordsR = RotateCoords(Coords, Canvas);
							Canvas.AddHyperlink(CoordsR.X, CoordsR.Y, CoordsR.Width, CoordsR.Height, hl.Text);
						}
					}
				}

				if (ShProp.ChildrenCount > 0)
				{
					RectangleF Coords2;
					bool HasCoords2 = false;
					real Rotation2 = 0;
					bool Rotated2 = false;

					try
					{
						if (ShProp.ChildrenCount>1 && ShProp.Children(1).ShapeType == TShapeType.NotPrimitive)
						{
							//This is the shape that governs the others

							DoAttributes (PagePrintRange, PaintClipRect, ShProp.Children(1), ParentCoords, HasParentCoords, out Coords2, out Rotation2, out HasCoords2, out Rotated2);
							if (Rotated2) 
							{
								RotateXY(Coords2, Rotation2);
							}

							ParentCoords = Coords2;
							HasParentCoords = true;
						}

						for (int i = 2; i<= ShProp.ChildrenCount; i++)
						{
							TShapeProperties ChildProp = ShProp.Children(i);
							DrawObject(ObjPos, ChildProp, PagePrintRange, PaintClipRect, ShadowInfo, Clipping, ParentCoords, HasParentCoords, ref HasObscuredShadow, RenderNotPrintable, DoHyperlinks);
						}
					}
					finally
					{
						if (Rotated2)
						{
							Canvas.ResetTransform();
						}
					}
				}
			}
			finally
			{
				if (Rotated)
				{
					Canvas.ResetTransform();
				}
			}

			if (HasCoords && ShadowInfo.Style == TShadowStyle.None) DrawShapeText(ShProp, TextCoords);
		}

        private void DrawBiffShape(int ObjPos, TShapeProperties ShProp, TShadowInfo ShadowInfo, TClippingStyle Clipping, bool RenderNotPrintable, ref RectangleF Coords, ref RectangleF TextCoords)
        {
            #region Shapes
            switch (ShProp.ShapeType)
            {
                #region Basic
                case TShapeType.PictureFrame:
                    DrawShape.DrawImage(Canvas, FWorkbook, ShProp, ObjPos, ShProp.ObjectPath, Coords, ShadowInfo, Zoom100);
                    break;

                case TShapeType.FlowChartProcess:
                case TShapeType.Rectangle:
                case TShapeType.TextBox:
                    DrawShape.DrawRectangle(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;

                case TShapeType.FlowChartAlternateProcess:
                case TShapeType.RoundRectangle:
                    TextCoords = DrawShape.DrawRoundRectangle(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;

                case TShapeType.Ellipse:
                case TShapeType.FlowChartConnector:
                    TextCoords = DrawShape.DrawOval(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;

                case TShapeType.IsocelesTriangle:
                case TShapeType.FlowChartExtract:
                    TextCoords = DrawShape.DrawIsocelesTriangle(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;
                case TShapeType.RightTriangle:
                    TextCoords = DrawShape.DrawRightTriangle(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;

                case TShapeType.Line:
                case TShapeType.StraightConnector1:
                    TextCoords = DrawShape.DrawLine(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;

                #endregion
                #region Basic II
                case TShapeType.Diamond:
                case TShapeType.FlowChartDecision:
                    TextCoords = DrawShape.DrawDiamond(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;
                case TShapeType.Parallelogram:
                    TextCoords = DrawShape.DrawParallelogram(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100, 21600 / 4);
                    break;
                case TShapeType.FlowChartInputOutput:
                    TextCoords = DrawShape.DrawParallelogram(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100, 21600 / 5);
                    break;
                case TShapeType.Trapezoid:
                    TextCoords = DrawShape.DrawTrapezoid(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100, 21600 / 4, false);
                    break;
                case TShapeType.FlowChartManualOperation:
                    TextCoords = DrawShape.DrawTrapezoid(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100, 21600 / 5, false);
                    break;

                case TShapeType.Octagon:
                    TextCoords = DrawShape.DrawOctagon(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;

                case TShapeType.Hexagon:
                    TextCoords = DrawShape.DrawHexagon(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100, 21600 / 4);
                    break;
                case TShapeType.FlowChartPreparation:
                    TextCoords = DrawShape.DrawHexagon(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100, 21600 / 5);
                    break;

                case TShapeType.Pentagon:
                    TextCoords = DrawShape.DrawRegularPentagon(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;

                case TShapeType.Plus:
                    TextCoords = DrawShape.DrawCross(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;
                #endregion
                #region Basic III
                case TShapeType.Can:
                case TShapeType.FlowChartMagneticDisk:
                    TextCoords = DrawShape.DrawCilinder(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;
                case TShapeType.Cube:
                    TextCoords = DrawShape.DrawCube(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;

                case TShapeType.Bevel:
                    TextCoords = DrawShape.DrawBevel(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;

                case TShapeType.FoldedCorner:
                    TextCoords = DrawShape.DrawFoldedSheet(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;

                case TShapeType.SmileyFace:
                    TextCoords = DrawShape.DrawSmiley(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;

                case TShapeType.Donut:
                    TextCoords = DrawShape.DrawDonut(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;

                case TShapeType.NoSmoking:
                    TextCoords = DrawShape.DrawNo(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;
                #endregion
                #region Basic IV
                case TShapeType.Heart:
                    TextCoords = DrawShape.DrawHeart(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;
                case TShapeType.LightningBolt:
                    TextCoords = DrawShape.DrawLightning(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;
                case TShapeType.Sun:
                    TextCoords = DrawShape.DrawSun(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;
                case TShapeType.Moon:
                    TextCoords = DrawShape.DrawMoon(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;
                case TShapeType.Arc:
                    TextCoords = DrawShape.DrawArc(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;

                case TShapeType.BracketPair:
                    TextCoords = DrawShape.DrawRoundRectangle(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;

                case TShapeType.Plaque:
                    TextCoords = DrawShape.DrawPlaque(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;
                #endregion
                #region Callouts
                case TShapeType.WedgeRectCallout:
                    TextCoords = DrawShape.DrawRectCallOut(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100, false);
                    break;

                case TShapeType.WedgeRRectCallout:
                    TextCoords = DrawShape.DrawRectCallOut(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100, true);
                    break;

                #endregion
                #region Block Arrows
                case TShapeType.WedgeEllipseCallout:
                    TextCoords = DrawShape.DrawOval(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;

                case TShapeType.Arrow:
                    TextCoords = DrawShape.DrawHBlockArrow(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100, false);
                    break;

                case TShapeType.LeftArrow:
                    TextCoords = DrawShape.DrawHBlockArrow(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100, true);
                    break;

                case TShapeType.UpArrow:
                    TextCoords = DrawShape.DrawVBlockArrow(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100, true);
                    break;

                case TShapeType.DownArrow:
                    TextCoords = DrawShape.DrawVBlockArrow(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100, false);
                    break;

                case TShapeType.LeftRightArrow:
                    TextCoords = DrawShape.DrawLeftRightBlockArrow(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;

                case TShapeType.UpDownArrow:
                    TextCoords = DrawShape.DrawUpDownBlockArrow(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;

                case TShapeType.QuadArrow:
                case TShapeType.QuadArrowCallout:
                    TextCoords = DrawShape.DrawQuadBlockArrow(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;

                case TShapeType.NotchedRightArrow:
                    TextCoords = DrawShape.DrawNotchedArrow(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;

                case TShapeType.StripedRightArrow:
                    TextCoords = DrawShape.DrawStripedArrow(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;

                case TShapeType.HomePlate:
                    TextCoords = DrawShape.DrawPentagon(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;

                case TShapeType.Chevron:
                    TextCoords = DrawShape.DrawChevron(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;
                #endregion
                #region Stars
                case TShapeType.Seal4:
                    TextCoords = DrawShape.DrawN4Star(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100, 1, 8000);
                    break;
                case TShapeType.Seal8:
                    TextCoords = DrawShape.DrawN4Star(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100, 2, 2800);
                    break;
                case TShapeType.Seal16:
                    TextCoords = DrawShape.DrawN4Star(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100, 4, 2800);
                    break;
                case TShapeType.Seal24:
                    TextCoords = DrawShape.DrawN4Star(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100, 6, 2800);
                    break;
                case TShapeType.Seal32:
                    TextCoords = DrawShape.DrawN4Star(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100, 8, 2800);
                    break;
                case TShapeType.Star:
                    TextCoords = DrawShape.Draw5Star(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;
                case TShapeType.IrregularSeal1:
                    TextCoords = DrawShape.DrawExplosion1(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;
                case TShapeType.IrregularSeal2:
                    TextCoords = DrawShape.DrawExplosion2(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;

                #endregion
                #region FlowChart
                case TShapeType.FlowChartTerminator:
                    TextCoords = DrawShape.DrawTerminator(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;
                case TShapeType.FlowChartManualInput:
                    TextCoords = DrawShape.DrawManualInput(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;
                case TShapeType.FlowChartPunchedCard:
                    TextCoords = DrawShape.DrawCard(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;
                case TShapeType.FlowChartSummingJunction:
                    TextCoords = DrawShape.DrawSummingJunction(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;
                case TShapeType.FlowChartOr:
                    TextCoords = DrawShape.DrawOr(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;
                case TShapeType.FlowChartCollate:
                    TextCoords = DrawShape.DrawCollate(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;
                case TShapeType.FlowChartSort:
                    TextCoords = DrawShape.DrawSort(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;

                case TShapeType.FlowChartMerge:
                    TextCoords = DrawShape.DrawMerge(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;

                case TShapeType.FlowChartOnlineStorage:
                    TextCoords = DrawShape.DrawStoredData(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;

                case TShapeType.FlowChartDelay:
                    TextCoords = DrawShape.DrawDelay(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;

                case TShapeType.FlowChartOffpageConnector:
                    TextCoords = DrawShape.DrawOffPageConnector(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;

                case TShapeType.FlowChartDisplay:
                    TextCoords = DrawShape.DrawDisplay(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;

                case TShapeType.FlowChartPredefinedProcess:
                    TextCoords = DrawShape.DrawPredefinedProcess(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;

                case TShapeType.FlowChartInternalStorage:
                    TextCoords = DrawShape.DrawInternalStorage(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;

                #endregion
                #region Controls
                case TShapeType.HostControl:
                    TextCoords = DrawHostControl(ObjPos, Canvas, FWorkbook, FontCache, ShProp, Coords, ShadowInfo, Clipping, Zoom100, RenderNotPrintable);
                    break;
                #endregion

                default:
                    TextCoords = DrawShape.DrawCustomShape(Canvas, FWorkbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100);
                    break;
            }
            #endregion
        }

		private RectangleF DrawHostControl(int ObjPos, IFlxGraphics aCanvas, ExcelFile Workbook, TFontCache aFontCache, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real aZoom100, bool RenderNotPrintable)
		{
			switch (ShProp.ObjectType)
			{
				case TObjectType.Chart:
                    DrawEmbeddedChart(ObjPos, aCanvas, Workbook, aFontCache, ShProp, Coords, ShadowInfo, Clipping, aZoom100, RenderNotPrintable);
                    return Coords;

                case TObjectType.CheckBox:
                    DrawShape.DrawImage(aCanvas, FWorkbook, ShProp, ObjPos, ShProp.ObjectPath, Coords, ShadowInfo, aZoom100);
                    return DrawCheckbox(ObjPos, aCanvas, Workbook, ShProp, Coords, ShadowInfo, aZoom100);

                case TObjectType.OptionButton:
                    DrawShape.DrawImage(aCanvas, FWorkbook, ShProp, ObjPos, ShProp.ObjectPath, Coords, ShadowInfo, aZoom100);
                    return DrawOptionButton(ObjPos, aCanvas, Workbook, ShProp, Coords, ShadowInfo, aZoom100);

                case TObjectType.GroupBox:
                    return DrawGroupBox(ObjPos, aCanvas, Workbook, ShProp, Coords, ShadowInfo, aZoom100); 

                case TObjectType.Button:
                    DrawShape.DrawImage(aCanvas, FWorkbook, ShProp, ObjPos, ShProp.ObjectPath, Coords, ShadowInfo, aZoom100);
                    return DrawButton(ObjPos, aCanvas, Workbook, ShProp, Coords, ShadowInfo, aZoom100);

                case TObjectType.ComboBox:
                    DrawShape.DrawImage(aCanvas, FWorkbook, ShProp, ObjPos, ShProp.ObjectPath, Coords, ShadowInfo, aZoom100);
                    return DrawComboBox(ObjPos, aCanvas, Workbook, FontCache, ShProp, Coords, ShadowInfo, aZoom100, ReverseRightToLeftStrings, SheetMethods);

                case TObjectType.ListBox:
                    DrawShape.DrawImage(aCanvas, FWorkbook, ShProp, ObjPos, ShProp.ObjectPath, Coords, ShadowInfo, aZoom100);
                    return DrawListBox(ObjPos, aCanvas, Workbook, FontCache, ShProp, Coords, ShadowInfo, aZoom100, ReverseRightToLeftStrings, SheetMethods);

                case TObjectType.Spinner:
                    DrawShape.DrawImage(aCanvas, FWorkbook, ShProp, ObjPos, ShProp.ObjectPath, Coords, ShadowInfo, aZoom100);
                    return DrawSpinner(ObjPos, aCanvas, Workbook, FontCache, ShProp, Coords, ShadowInfo, aZoom100, ReverseRightToLeftStrings, SheetMethods);

                case TObjectType.ScrollBar:
                    DrawShape.DrawImage(aCanvas, FWorkbook, ShProp, ObjPos, ShProp.ObjectPath, Coords, ShadowInfo, aZoom100);
                    return DrawScrollBar(ObjPos, aCanvas, Workbook, FontCache, ShProp, Coords, ShadowInfo, aZoom100, ReverseRightToLeftStrings, SheetMethods);

				default: 
					DrawShape.DrawImage(aCanvas, FWorkbook, ShProp, ObjPos, ShProp.ObjectPath, Coords, ShadowInfo, aZoom100);
					return Coords;

			}

		}

        private static RectangleF DrawCheckbox(int ObjPos, IFlxGraphics aCanvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, float aZoom100)
        {
            float w2 = 4.5f* aZoom100;
            float w = 2 * w2;
            float w4x = w2 / 2.5f;
            float w4y = w2 / 2f;
            float y0 = (Coords.Top + Coords.Bottom) /2f - 1;
            float x0 = Coords.Left + 3 * aZoom100;
            TCheckboxState cbValue = Workbook.GetCheckboxState(ObjPos, ShProp.ObjectPath);
            using (Brush br = GetCheckBoxBrush(cbValue))
            {
                aCanvas.DrawAndFillRectangle(Pens.Black, br, x0, y0 - w2, w , w);
            }
            if (cbValue == TCheckboxState.Checked)
            {
                using (Pen p = new Pen(Color.Black,1.5f))
                {
                    aCanvas.DrawLines(p, new TPointF[] { new TPointF(x0 + w4x, y0), new TPointF(x0 + w2, y0 + w4y), new TPointF(x0 + w - w4x, y0 - w4y) });
                }
            }

            float offs = 4 * aZoom100 + w;
            return new RectangleF(Coords.Left + offs, Coords.Top, Coords.Width - offs, Coords.Height);
        }

        private static Brush GetCheckBoxBrush(TCheckboxState cbValue)
        {
            switch (cbValue)
            {
                case TCheckboxState.Indeterminate:
                    return new HatchBrush(HatchStyle.Percent50, Color.Black, Color.White);

                default:
                    return new SolidBrush(Color.White);
            }
        }

        private static RectangleF DrawOptionButton(int ObjPos, IFlxGraphics aCanvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, float aZoom100)
        {
            float w2 = 4.5f * aZoom100;
            float y0 = (Coords.Top + Coords.Bottom) / 2f - 1;
            float x0 = Coords.Left + 3 * aZoom100 + w2;
            bool rbChecked = Workbook.GetRadioButtonState(ObjPos, ShProp.ObjectPath);

            TPointF[] Circle = TEllipticalArc.GetPoints(x0, y0, w2, w2, 0, 0, 2 * Math.PI);
            aCanvas.DrawAndFillBeziers(Pens.Black, Brushes.White, Circle);

            if (rbChecked)
            {
                float w3 = w2 / 2f;
                TPointF[] Circle2 = TEllipticalArc.GetPoints(x0, y0, w3, w3, 0, 0, 2 * Math.PI);
                aCanvas.DrawAndFillBeziers(Pens.Black, Brushes.Black, Circle2);
            }

            float offs = 4 * aZoom100 + 2 * w2;
            return new RectangleF(Coords.Left + offs, Coords.Top, Coords.Width - offs, Coords.Height);
        }

        private static RectangleF DrawGroupBox(int ObjPos, IFlxGraphics aCanvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, float aZoom100)
        {
            if (ShProp.Text == null || ShProp.Text.Length == 0)
            {
                using (Pen PicturePen = DrawShape.GetPen(ShProp, Workbook, ShadowInfo, false))
                {
                    aCanvas.DrawRectangle(PicturePen, Coords.Left, Coords.Top, Coords.Width, Coords.Height);
                }
                return Coords;
            }

            //If font is different, it will be on the RTF runs.

            string FontName = "Arial"; float FontSize = 10;
            if (ShProp.Text.RTFRunCount > 0 && ShProp.Text.RTFRun(0).FirstChar == 0)
            {
                TFlxFont fnt = Workbook.GetFont(ShProp.Text.RTFRun(0).FontIndex);
                FontName = fnt.Name;
                FontSize = fnt.Size20 / 20f;
            }
            using (Font aFont = new Font(FontName, FontSize))
            {
                float ofs = 5;
                SizeF tz = aCanvas.MeasureString(ShProp.Text.ToString(), aFont);
                RectangleF r = new RectangleF(Coords.Left + ofs * aZoom100, Coords.Top - tz.Height / 2 * aZoom100, tz.Width * aZoom100, tz.Height * aZoom100);

                using (Pen PicturePen = DrawShape.GetPen(ShProp, Workbook, ShadowInfo, false))
                {
                    aCanvas.DrawLine(PicturePen, Coords.Left, Coords.Bottom, Coords.Right, Coords.Bottom);
                    aCanvas.DrawLine(PicturePen, Coords.Left, Coords.Bottom, Coords.Left, Coords.Top);
                    aCanvas.DrawLine(PicturePen, Coords.Right, Coords.Bottom, Coords.Right, Coords.Top);

                    aCanvas.DrawLine(PicturePen, Coords.Left, Coords.Top, r.Left, Coords.Top);
                    aCanvas.DrawLine(PicturePen, r.Right, Coords.Top, Coords.Right, Coords.Top);
                }


                return r;
            }
        }

        private static RectangleF DrawComboBox(int ObjPos, IFlxGraphics aCanvas, ExcelFile Workbook, TFontCache FontCache, 
            TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, float aZoom100,
            bool ReverseRightToLeftStrings, IDrawObjectMethods SheetMethods)
        {
            RectangleF NewCoords;
            using (Pen PicturePen = new Pen(Color.Silver))
            {
                aCanvas.DrawAndFillRectangle(PicturePen, Brushes.White, Coords.Left, Coords.Top, Coords.Width, Coords.Height);
                real w = Coords.Height; if (Coords.Width < w) w = Coords.Width;
                NewCoords = DrawSpin(aCanvas, Coords, PicturePen, w, aZoom100, false);
            }

            int Sel = Workbook.GetObjectSelection(ObjPos, ShProp.ObjectPath);
            if (Sel <= 0) return Coords;

            TCellAddressRange InputRange = Workbook.GetObjectInputRange(ObjPos, ShProp.ObjectPath);
            if (InputRange == null) return Coords;

            int Row1 = InputRange.TopLeft.Row;
            int Row2 = InputRange.BottomRight.Row;

            int r = Row1 + Sel - 1;
            if (r > Row2) return Coords;

            int XF = -1;
            Color aColor = Color.Empty;
            int sheet = string.IsNullOrEmpty(InputRange.TopLeft.Sheet)? Workbook.ActiveSheet: Workbook.GetSheetIndex(InputRange.TopLeft.Sheet);
            object value = Workbook.GetCellValue(sheet, r, InputRange.TopLeft.Col, ref XF);
            TRichString text = TFlxNumberFormat.FormatValue(value, Workbook.GetFormat(XF).Format, ref aColor, Workbook);
            if (text == null) return NewCoords;

            DrawComboLine(aCanvas, Workbook, FontCache, aZoom100, ReverseRightToLeftStrings, SheetMethods, ref NewCoords, text);

            return NewCoords;
        }

        private static void DrawComboLine(IFlxGraphics aCanvas, ExcelFile Workbook, TFontCache FontCache, float aZoom100, bool ReverseRightToLeftStrings, IDrawObjectMethods SheetMethods, ref RectangleF Coords, TRichString text)
        {
            text = new TRichString(FlxConvert.ToString(text)); //remove rich format
            aCanvas.SaveState();
            try
            {
                //If font is different, it will be on the RTF runs.
                Color FontColor = Color.Black;
                TFlxFont DrawFont = new TFlxFont();
                DrawFont.Name = "Tahoma"; DrawFont.Size20 = 160;

                TSubscriptData SubScript = new TSubscriptData(TFlxFontStyles.None);

                THAlign HAlign = THAlign.Left;
                TVAlign VAlign = TVAlign.Center;
                FlexCelRender.DrawText(Workbook, aCanvas, FontCache, aZoom100, ReverseRightToLeftStrings, SheetMethods,
                    -1, -1, ref Coords, ref Coords, null, true, false, 0, false, false, THFlxAlignment.left, TVFlxAlignment.center,
                    0, false, DrawFont, ref FontColor, ref SubScript, ref HAlign, VAlign, 0, text, false, true, 0, null);

            }
            finally
            {
                //Restore clipping area.
                aCanvas.RestoreState();
            }
        }

        private static RectangleF DrawListBox(int ObjPos, IFlxGraphics aCanvas, ExcelFile Workbook, TFontCache FontCache,
            TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, float aZoom100,
            bool ReverseRightToLeftStrings, IDrawObjectMethods SheetMethods)
        {
            using (Pen PicturePen = new Pen(Color.Silver))
            {
                aCanvas.DrawAndFillRectangle(PicturePen, Brushes.White, Coords.Left, Coords.Top, Coords.Width, Coords.Height);
            }

            int Sel = Workbook.GetObjectSelection(ObjPos, ShProp.ObjectPath);

            TCellAddressRange InputRange = Workbook.GetObjectInputRange(ObjPos, ShProp.ObjectPath);
            if (InputRange == null) return Coords;

            real LineHeight = 9.5f;
            int VisibleRows = (int)Math.Truncate(Coords.Height / LineHeight);
            
            int Row1 = InputRange.TopLeft.Row;
            int Row2 = InputRange.BottomRight.Row;

            int Row0 = Sel - VisibleRows;
            if (Row0 < 0) Row0 = 0;

            int rSel = Row1 + Sel - 1;

            RectangleF TextCoords = new RectangleF(Coords.Left + 5 * aZoom100, Coords.Top, Coords.Width - 10 * aZoom100, LineHeight);

            for (int r = Row1 + Row0; r < Math.Min(Row2 + 1,  Row1 + Row0 + VisibleRows); r++)
            {
                if (r == rSel) aCanvas.DrawAndFillRectangle(null, Brushes.LightBlue, TextCoords.X - 5 * aZoom100, TextCoords.Y,
                    TextCoords.Width + 10 * aZoom100, TextCoords.Height);
                int XF = -1;
                Color aColor = Color.Empty;
                int sheet = string.IsNullOrEmpty(InputRange.TopLeft.Sheet) ? Workbook.ActiveSheet : Workbook.GetSheetIndex(InputRange.TopLeft.Sheet);
                object value = Workbook.GetCellValue(sheet, r, InputRange.TopLeft.Col, ref XF);
                TRichString text = TFlxNumberFormat.FormatValue(value, Workbook.GetFormat(XF).Format, ref aColor, Workbook);
                if (text == null) continue;

                DrawComboLine(aCanvas, Workbook, FontCache, aZoom100, ReverseRightToLeftStrings, SheetMethods, ref TextCoords, text);
                TextCoords.Y += LineHeight;
            }
            return Coords;
        }



        private static RectangleF DrawSpinner(int ObjPos, IFlxGraphics aCanvas, ExcelFile Workbook, TFontCache FontCache,
    TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, float aZoom100,
    bool ReverseRightToLeftStrings, IDrawObjectMethods SheetMethods)
        {
            using (Pen PicturePen = new Pen(Color.Silver))
            {
                real w = Coords.Width;
                DrawSpin(aCanvas, new RectangleF(Coords.Left, Coords.Top, w, Coords.Height / 2), PicturePen, w, aZoom100, true);
                DrawSpin(aCanvas, new RectangleF(Coords.Left, Coords.Top + Coords.Height / 2, w, Coords.Height / 2), PicturePen, w, aZoom100, false);
            }
            return Coords;
        }


        private static RectangleF DrawScrollBar(int ObjPos, IFlxGraphics aCanvas, ExcelFile Workbook, TFontCache FontCache,
                TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, float aZoom100,
                bool ReverseRightToLeftStrings, IDrawObjectMethods SheetMethods)
        {
            using (Pen PicturePen = new Pen(Color.Silver))
            {
                aCanvas.DrawAndFillRectangle(PicturePen, Brushes.White, Coords.Left, Coords.Top, Coords.Width, Coords.Height);
                real w = Coords.Width;
                DrawSpin(aCanvas, new RectangleF(Coords.Left, Coords.Top, w, w), PicturePen, w, aZoom100, true);
                DrawSpin(aCanvas, new RectangleF(Coords.Left, Coords.Bottom - w, w, w), PicturePen, w, aZoom100, false);

                TSpinProperties spin = Workbook.GetObjectSpinProperties(ObjPos, ShProp.ObjectPath);
                float min = spin.Min;
                float max = spin.Max;
                float val = Workbook.GetObjectSpinValue(ObjPos, ShProp.ObjectPath);
                float w2 = 2 * aZoom100;
                float y = Coords.Top + w + w2;
                if (max - min > 0 && val >= min && val <= max)
                {
                    y += (Coords.Height - 2 * w - 2 * w2) * (val - min) / (max - min);
                    aCanvas.DrawAndFillRectangle(PicturePen, Brushes.Azure, Coords.Left, y - w2, Coords.Width, 2 * w2);
                }
            }
            return Coords;
        }

        private static RectangleF DrawSpin(IFlxGraphics aCanvas, RectangleF Coords, Pen PicturePen, real w, real Zoom100, bool Up)
        {
            real f = Up ? -8f : 8f;
            aCanvas.DrawAndFillRectangle(PicturePen, Brushes.Gainsboro, Coords.Right - w, Coords.Top, w, Coords.Height);
            aCanvas.DrawAndFillPolygon(Pens.Black, Brushes.Black, new TPointF[] 
            {
                new TPointF(Coords.Right - w * 3 / 4f, (Coords.Top + Coords.Bottom) / 2f - w /f),
                new TPointF(Coords.Right - w / 4f, (Coords.Top + Coords.Bottom) / 2f - w /f),
                new TPointF(Coords.Right - w / 2f, (Coords.Top + Coords.Bottom) / 2f + w /f)
            }
            );

            return new RectangleF(Coords.Left + 5 * Zoom100, Coords.Top, Coords.Width - w - 10 * Zoom100, Coords.Height);
        }
        private static RectangleF DrawButton(int ObjPos, IFlxGraphics aCanvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, float aZoom100)
        {
            return DrawShape.DrawRoundRectangle(aCanvas, TClippingStyle.None, Coords, (real)0.1, Brushes.White, Pens.Black);
       }

        public void DrawEmbeddedChart(int ObjPos, IFlxGraphics aCanvas, ExcelFile Workbook, TFontCache aFontCache, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real aZoom100, bool RenderNotPrintable)
        {
            ExcelChart Chart = Workbook.GetChart(ObjPos, ShProp.ObjectPath);
            if (Chart == null || Chart.SeriesCount == 0) return;

            DrawChart.Draw(Chart, aCanvas, Workbook, aFontCache, ShProp, Coords, ShadowInfo, Clipping, aZoom100);
            TShapePropertiesList ShapesInChart = new TShapePropertiesList();
            for (int i = 1; i <= Chart.ObjectCount; i++)
            {
				TShapeProperties sp = Chart.GetObjectProperties(i, true);
				sp.zOrder = i;
                ShapesInChart.Add(sp);
            }
            if (ShapesInChart.Count > 0)
            {
                DrawImages(ShapesInChart, new TXlsCellRange(1, 1, 1, 1), Coords, RenderNotPrintable);
            }
        }

        #region Text
        internal void DrawShapeText(TShapeProperties ShProp, RectangleF Coords)
        {
            if (ShProp.Text == null || ShProp.Text.Length == 0) return;

			Canvas.SaveState();
			try
			{            
				real LeftMargin = ShProp.ShapeOptions.AsSignedLong(TShapeOption.dxTextLeft, 91440) / 12700f; 
				real RightMargin = ShProp.ShapeOptions.AsSignedLong(TShapeOption.dxTextRight, 91440) / 12700f;
				real TopMargin = ShProp.ShapeOptions.AsSignedLong(TShapeOption.dyTextTop, 91440 / 2) / 12700f;
				real BottomMargin = ShProp.ShapeOptions.AsSignedLong(TShapeOption.dyTextBottom, 91440 / 2) / 12700f;

                if (ShProp.ShapeOptions.AsBool(TShapeOption.fFitTextToShape, false, 3) //Automatic margins
                   )
                {
                    LeftMargin = 91440 / 5 / 12700f;
                    RightMargin = 0;
                    TopMargin =  91440 / 4 / 12700f;
                    BottomMargin = 0;
                }

				RectangleF NewCoords = new RectangleF(Coords.Left + LeftMargin, Coords.Top + TopMargin, Coords.Width - LeftMargin - RightMargin, Coords.Height - TopMargin - BottomMargin);

				bool MultiLine = true;

				THFlxAlignment HJustify = THFlxAlignment.left;
				TVFlxAlignment VJustify = TVFlxAlignment.top;
				THAlign HAlign = THAlign.Left;
				TVAlign VAlign = TVAlign.Top;

				switch ((ShProp.TextFlags & 0x70) >> 4)
				{
					case 2:
						VAlign = TVAlign.Center;
						VJustify = TVFlxAlignment.center;
						break;
					case 3:
						VAlign = TVAlign.Bottom;
						VJustify = TVFlxAlignment.bottom;
						break;
					case 4:
						VAlign = TVAlign.Top;
						VJustify = TVFlxAlignment.justify;
						break;
				}

				switch ((ShProp.TextFlags & 0xE) >> 1)
				{
					case 2:
						HAlign = THAlign.Center;
						HJustify = THFlxAlignment.center;
						break;
					case 3:
						HAlign = THAlign.Right;
						HJustify = THFlxAlignment.right;
						break;
					case 4:
						HAlign = THAlign.Left;
						HJustify = THFlxAlignment.justify;
						break;
				}

				real Alpha = 0;
				bool Vertical = false;

				switch (ShProp.TextRotation)
				{
					case 1: Vertical = true; break;
					case 2: Alpha = 90; break;
					case 3: Alpha = -90; break;
				}

				//If font is different, it will be on the RTF runs.
				Color FontColor = Color.Black;
				TFlxFont DrawFont = new TFlxFont();
				DrawFont.Name = "Arial"; DrawFont.Size20 = 200;
#if (FRAMEWORK30)
                if (ShProp.ShapeThemeFont != null)
                {
                    TThemeFontScheme fs = FWorkbook.GetTheme().Elements.FontScheme;
                    TThemeFont ThemeFont = null;
                    switch (ShProp.ShapeThemeFont.ThemeScheme)
                    {
                        case TFontScheme.None:
                            ThemeFont = null;
                            break;

                        case TFontScheme.Minor:
                            ThemeFont = fs.MinorFont;
                            break;

                        case TFontScheme.Major:
                            ThemeFont = fs.MajorFont;
                            break;
                    }

                    if (ThemeFont != null)
                    {
                        DrawFont.Name = ThemeFont.Latin.Typeface;
                        FontColor = ShProp.ShapeThemeFont.ThemeColor.ToColor(FWorkbook);
                    }
                }

#endif


                TSubscriptData SubScript = new TSubscriptData(TFlxFontStyles.None);

				FlexCelRender.DrawText(FWorkbook, Canvas, FontCache, Zoom100, ReverseRightToLeftStrings, SheetMethods,
					-1, -1, ref NewCoords, ref Coords, null, true, false, 0, MultiLine, false, HJustify, VJustify,
					Alpha, Vertical, DrawFont, ref FontColor, ref SubScript, ref HAlign, VAlign, 0, ShProp.Text, false, true, 0, null); 
            
			}
			finally
			{
				//Restore clipping area.
				Canvas.RestoreState();
			}
        }
        #endregion


		#region RenderObject
        public static Image RenderObject(ExcelFile xls, real Dpi, 
			SmoothingMode aSmoothingMode, bool AntiAliased, InterpolationMode aInterpolationMode,
			int objectIndex, TShapeProperties ShapeProperties, bool ReturnImage, Color BackgroundColor,
			out PointF Origin, out RectangleF ImgCoords, out Size ImgSizePixels)
        {
            Origin = PointF.Empty;
            ImgCoords = RectangleF.Empty;
			ImgSizePixels = Size.Empty;

            if (ShapeProperties.ObjectType == TObjectType.Comment) return null;


			TClientAnchor Anchor = ShapeProperties.NestedAnchor;
			if (Anchor == null) return null;

            TShapePropertiesList ShapeToRender = new TShapePropertiesList();
            ShapeProperties.zOrder = objectIndex;
            ShapeToRender.Add(ShapeProperties);

            real Rotation = 0;
            if (ShapeProperties.ShapeOptions != null)
            {
                Rotation =ShapeProperties.ShapeOptions.As1616(TShapeOption.Rotation, FlxConsts.DefaultRotation);
                if (ShapeProperties.FlipH ^ ShapeProperties.FlipV) Rotation = -Rotation;
            }


            double w = 0;
            double h = 0;
            Anchor.CalcImageCoordsInPoints(ref h, ref w, xls);
            if (w <= 0 || h <= 0) return null;

            ImgCoords = new RectangleF(0, 0, (real)w, (real)h);
            bool Swapped = SwapCoords(ref ImgCoords, Rotation);

            RectangleF FullCoords = AddShadowAndRotation((real)ImgCoords.Height, (real)ImgCoords.Width, ShapeProperties, Rotation);

            int wPix = (int)Math.Ceiling(FullCoords.Width * Dpi / FlexCelRender.DispMul) + 1;
            int hPix = (int)Math.Ceiling(FullCoords.Height * Dpi / FlexCelRender.DispMul) + 1;
			ImgSizePixels = new Size(wPix, hPix);
			ImgCoords = new RectangleF(FullCoords.Left, FullCoords.Top, wPix * FlexCelRender.DispMul / Dpi, hPix * FlexCelRender.DispMul / Dpi); //We need to recalculate the final rectangle dimensions, since they might be bigger than the original one due to rounding errors.

			/*Width and height exchange at 45 degrees. DrawImages will compensate for this, so we need to un-compensate here.
			 * This is why we use w and h instead of the corrected Fullcoords.width, Fullcoords.height, and also correct X and Y*/
			real dx = Swapped ? (real)((h - w) / 2 - FullCoords.X) : -FullCoords.X;
			real dy = Swapped ? (real)(-(h - w) / 2 - FullCoords.Y) : -FullCoords.Y;

			Origin = new PointF(Anchor.Dx1Points(xls) - dx, Anchor.Dy1Points(xls) - dy);

			if (!ReturnImage) return null;

            using (TBitmapCreator bmp = new TBitmapCreator(Dpi, aSmoothingMode, AntiAliased, aInterpolationMode, BackgroundColor, wPix, hPix))
            {
                using (TFontCache FontCache = new TFontCache())
                {
                    GdiPlusGraphics Canvas = new GdiPlusGraphics(bmp.ImgGraphics);
                    Canvas.CreateSFormat();
                    try
                    {
                        TRenderImageMethods RenderMethods = new TRenderImageMethods(new RectangleF(dx, dy, (real)w, (real)h));
                        TDrawObjects DrawObjects = new TDrawObjects(xls, Canvas, FontCache, 1, THidePrintObjects.None, false, RenderMethods);
                        DrawObjects.DrawImages(ShapeToRender, TXlsCellRange.FullRange(), FullCoords, true);
                    }
                    finally
                    {
                        Canvas.DestroySFormat();
                    }
                }

                return bmp.Img;
            }
        }

		private static RectangleF AddShadowAndRotation(real h, real w, TShapeProperties ShProp, real Rotation)
		{
 		    RectangleF Coords = new RectangleF(0, 0 ,w, h);
			RectangleF Result = Coords;
 
            AddShadowOffset(ShProp, Coords, ref Result, Rotation);

			return Result;
		}

        private static void AddShadowOffset(TShapeProperties ShProp, RectangleF Coords, ref RectangleF Result, real Rotation)
        {
			//Shadow is rotated in place. So we cannot add the shadow first and then rotate everything, we need to rotate each shadow on its own.

			AddRotation(ShProp, Result, ref Result, Rotation);
            if (ShProp.ShapeOptions != null && ShProp.ShapeOptions.AsBool(TShapeOption.fshadowObscured, false, 1))
            {
                int PassCount = 1;
                TShadowType ShadowType = TShadowType.Offset;
                if (ShProp.ShapeOptions != null) ShadowType = (TShadowType)ShProp.ShapeOptions.AsLong(TShapeOption.shadowType, 0);
                if (ShadowType == TShadowType.EmbossOrEngrave || ShadowType == TShadowType.Double)
                {
                    PassCount = 2;
                }

                for (int Pass = 0; Pass < PassCount; Pass++)
                {
                    RectangleF ShadowCoords = Coords;
                    TDrawObjects.OffsetCoords(1, ShProp, ref ShadowCoords, Pass);
					GdiPlusGraphics.FixRect(ref ShadowCoords);
					AddRotation(ShProp, ShadowCoords, ref ShadowCoords, Rotation);
                    Result = RectangleF.Union(ShadowCoords, Result);
                }
            }
        }

		private static void RotateX(PointF Center, double SinAlpha, double CosAlpha, ref PointF Pt)
		{
			Pt.X = Center.X + (real)((Pt.X - Center.X) * CosAlpha - (Center.Y - Pt.Y) * SinAlpha);
		}

		private static void RotateY(PointF Center, double SinAlpha, double CosAlpha, ref PointF Pt)
		{
			Pt.Y = Center.Y - (real)((Pt.X - Center.X) * SinAlpha + (Center.Y - Pt.Y) * CosAlpha);
		}

		private static void AddRotation(TShapeProperties ShProp, RectangleF Coords, ref RectangleF Result, real Rotation)
		{
			if (Rotation == 0) return;

			double rot = -Rotation;
			while (rot < 0) rot += 360;
			rot %= 360;
			
			bool Quadrant_1_or_3 = rot < 90 || (rot > 180 && rot  < 270);
			
			PointF ptx1 = new PointF(Result.Right, Result.Bottom);
			PointF ptx2 = new PointF(Result.Left, Result.Top);

			PointF pty1 = new PointF(Result.Right, Result.Top);
			PointF pty2 = new PointF(Result.Left, Result.Bottom);

			if (!Quadrant_1_or_3) 
			{
				PointF tmp = ptx1; ptx1 = pty1; pty1 = tmp;
				tmp = ptx2; ptx2 = pty2; pty2 = tmp;
			}

			PointF Center = new PointF((Coords.Left +  Coords.Right) / 2, (Coords.Top + Coords.Bottom) / 2); //center is on the original image, not the bigger.

			double Alpha = rot * Math.PI / 180f;
			double CosAlpha = Math.Cos(Alpha);
			double SinAlpha = Math.Sin(Alpha);
			RotateX(Center, SinAlpha, CosAlpha, ref ptx1);
			RotateX(Center, SinAlpha, CosAlpha, ref ptx2);
			RotateY(Center, SinAlpha, CosAlpha, ref pty1);
			RotateY(Center, SinAlpha, CosAlpha, ref pty2);

			Result = RectangleF.FromLTRB(Math.Min(ptx1.X, ptx2.X), Math.Min(pty1.Y, pty2.Y), Math.Max(ptx1.X, ptx2.X), Math.Max(pty1.Y, pty2.Y));


		}

		#endregion
        #endregion

    }

	#region Render Separated Images
	internal class TRenderImageMethods: IDrawObjectMethods
	{
		private RectangleF ImageSize;
		public TRenderImageMethods(RectangleF aImageSize)
		{
			ImageSize = aImageSize;
		}


		#region IDrawObjectMethods Members

		public RectangleF GetImageRectangle(TXlsCellRange PagePrintRange, RectangleF PaintClipRect, TClientAnchor A)
		{
			return ImageSize;
		}

		public void SpawnCell(int aCol, int aRow, TSpawnedCellList SpawnedCells, FlexCel.Render.THAlign HAlign, ref RectangleF TextRect, ref RectangleF ContainingRect)
		{
		}

		#endregion
	}


	#endregion

    #region BitmapCreator
    class TBitmapCreator : IDisposable
    {
        private Bitmap bmp;
        Graphics ImageGraphics;

        internal TBitmapCreator(real Dpi, SmoothingMode aSmoothingMode, bool AntiAliased, InterpolationMode aInterpolationMode, Color BackgroundColor, int wPix, int hPix)
        {
            bmp = BitmapConstructor.CreateBitmap(wPix, hPix, PixelFormat.Format32bppArgb);
            bmp.SetResolution(Dpi, Dpi);
            
            ImageGraphics = Graphics.FromImage(bmp);
            {
                if (BackgroundColor != ColorUtil.Empty)
                {
                    using (Brush BackgroundBrush = new SolidBrush(BackgroundColor))
                    {
                        ImageGraphics.FillRectangle(BackgroundBrush, 0, 0, wPix + 1, hPix + 1);
                    }
                }

                ImageGraphics.SmoothingMode = aSmoothingMode;
                ImageGraphics.InterpolationMode = aInterpolationMode;
                if (AntiAliased) ImageGraphics.TextRenderingHint = TextRenderingHint.AntiAlias;
                ImageGraphics.PageUnit = GraphicsUnit.Point;
           }
        }

        internal Bitmap Img { get { return bmp; } }
        internal Graphics ImgGraphics { get { return ImageGraphics; } }

        #region IDisposable Members

        public void Dispose()
        {
            //We will not dispose bmp, since we are using it.
            if (ImageGraphics != null) ImageGraphics.Dispose();
            GC.SuppressFinalize(this);
        }

        #endregion
    }

    #endregion

}
