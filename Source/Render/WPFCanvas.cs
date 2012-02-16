#if (WPF)
using System;
using System.IO;
using FlexCel.Core;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows;
using System.Collections.Generic;

using RectangleF = System.Windows.Rect;
using SizeF = System.Windows.Size;
using PointF = System.Windows.Point;
using real = System.Double;

using System.Windows.Shapes;
using System.Globalization;

namespace FlexCel.Render
{
    internal class WpfGraphics: IFlxGraphics
    {
        private DrawingContext FCanvas;
        private List<int> GraphicsSave;

        internal WpfGraphics(DrawingContext aCanvas)
        {
            FCanvas = aCanvas;
            GraphicsSave = new List<int>();
            GraphicsSave.Add(0);
		}

        #region IFlxGraphics Members

		public void CreateSFormat()
		{
		}

		public void DestroySFormat()
		{
		}


        public RectangleF ClipBounds
        {
            get
            {
                return MainCanvas.Clip.Bounds;
            }
        }

        public void Rotate(real x, real y, real Alpha)
        {
			if (PointOutside.Check(ref x, ref y)) return;

            RotateTransform rt = new RotateTransform();
            rt.Angle = -Alpha;
            rt.CenterX = x;
            rt.CenterY = y;
            if (rt.CanFreeze) rt.Freeze();

            FCanvas.PushTransform(rt);
            GraphicsSave[GraphicsSave.Count - 1]++;
        }

		public TPointF Transform(TPointF p)
		{
			PointOutside.Check(ref p);
            Point p1 = FCanvas.TransformToVisual(MainCanvas).Transform(new Point(p.X, p.Y));
            return new TPointF((real)p1.X, (real)p1.Y);
		}

        public void Scale(real xScale, real yScale)
        {
            ScaleTransform rt = new ScaleTransform();
            rt.ScaleX = xScale;
            rt.ScaleY = yScale;

            if (rt.CanFreeze) rt.Freeze();
            FCanvas.PushTransform(rt);
            GraphicsSave[GraphicsSave.Count - 1]++;
        }

        public void SaveTransform()
        {
            SaveState();
        }

        public void ResetTransform()
        {
            RestoreState();
        }

        public void SaveState()
        {
            GraphicsSave.Add(0);
        }

        public void RestoreState()
        {
            for (int i = 0; i < GraphicsSave[GraphicsSave.Count - 1]; i++)
            {
                FCanvas.Pop();
            }
            GraphicsSave.RemoveAt(GraphicsSave.Count - 1);
        }

        public void DrawString(string Text, Font aFont, Brush aBrush, real x, real y)
        {
			if (PointOutside.Check(ref x, ref y)) return;
			Text = Text.Replace('\u000a','\u0000'); //Replace newlines with empty characters

            Typeface TextTypeface = new Typeface(aFont.Family, aFont.Style, aFont.Weight, FontStretches.Normal);

            FormattedText fmtText = new FormattedText(Text, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, TextTypeface, aFont.SizeInPix, aBrush);
            fmtText.SetTextDecorations(aFont.Decorations);
            FCanvas.DrawText(fmtText, new Point(x, y));
        }

        public void DrawString(string Text, Font aFont, Pen aPen, Brush aBrush, real x, real y)
        {
			if (PointOutside.Check(ref x, ref y)) return;

			Text = Text.Replace('\u000a','\u0000'); //Replace newlines with empty characters
            if (aBrush != null) DrawString(Text, aFont, aBrush, x, y);

            if (aPen != null)
            {
                Typeface TextTypeface = new Typeface(aFont.Family, aFont.Style, aFont.Weight, FontStretches.Normal);

                FormattedText fmtText = new FormattedText(Text, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, TextTypeface, aFont.SizeInPix, Brushes.Black); //Brush doesn't matter here.
                fmtText.SetTextDecorations(aFont.Decorations);
                Geometry TxtGeometry = fmtText.BuildHighlightGeometry(new Point(x, y));
                if (TxtGeometry.CanFreeze) TxtGeometry.Freeze();
                FCanvas.DrawGeometry(null, aPen, TxtGeometry);
            }
        }

        public SizeF MeasureString(string Text, Font aFont, TPointF p)
        {
			PointOutside.Check(ref p);
            Typeface TextTypeface = new Typeface(aFont.Family, aFont.Style, aFont.Weight, FontStretches.Normal);
            FormattedText fmtText = new FormattedText(Text, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, TextTypeface, aFont.SizeInPix, null);
            //fmtText.SetTextDecorations(aFont.Decorations);
            return new SizeF(fmtText.Width, fmtText.Height);
        }


		public SizeF MeasureStringEmptyHasHeight(string Text, Font aFont)
		{
			if (Text == null || Text.Length == 0)
			{
				SizeF s = MeasureString("Mg", aFont);
				s.Width = 0;
				return s;
			}

			return MeasureString(Text, aFont);
		}

        public SizeF MeasureString(string Text, Font aFont)
        {
            return MeasureString(Text, aFont, new TPointF(0,0));
        }

		public real FontDescent(Font aFont)
		{
			FontFamily ff = aFont.Family;
			return aFont.GetHeight(FCanvas) / ff.LineSpacing * ff.CellDescent;	
		}

        public real FontLinespacing(Font aFont)
        {
            FontFamily ff = aFont.FontFamily; //don't dispose
            real LineGap = ff.GetLineSpacing(aFont.Style) - ff.GetCellDescent(aFont.Style) - ff.GetCellAscent(aFont.Style);
            return aFont.GetHeight(FCanvas)/ ff.GetLineSpacing(aFont.Style) * LineGap;	
        }

        private static bool CheckPoints(TPointF[] points)
        {
            bool AllPointsOutside = true;

			for (int i=0; i< points.Length;i++)
			{
				if (!PointOutside.Check(ref points[i])) AllPointsOutside = false;
			}

			if (AllPointsOutside) return false;
            return true;
        }

        public void DrawLines(Pen aPen, TPointF[] points)
        {
            if (aPen == null || points == null || points.Length <=1) return;
			if (!CheckPoints(points)) return;

            StreamGeometry Line = new StreamGeometry();
            using (StreamGeometryContext sc = Line.Open())
            {
                sc.BeginFigure(new Point(points[0].X, points[0].Y), false, false);
                for (int i = 1; i < points.Length; i++)
                {
                    sc.LineTo(new Point(points[i].X, points[i].Y), true, true);
                }
            }

            if (Line.CanFreeze) Line.Freeze();

            FCanvas.DrawGeometry(null, aPen, Line);
        }

		internal static void FixRect(ref real x, ref real y, ref real width, ref real height)
		{
			if (width < 0) {width = -width; x-=width; }
			if (height < 0) {height = -height; y-=height; }
		}

		internal static void FixRect(ref RectangleF r)
		{
			if (r.Width < 0) {r.Width = -r.Width; r.X-=r.Width; }
			if (r.Height < 0) {r.Height = -r.Height; r.Y-=r.Height; }
		}

        public void FillRectangle(Brush b, RectangleF rect)
        {
			FixRect(ref rect);
			if (PointOutside.Check(ref rect)) return;
            if (b == null) return;
			FCanvas.DrawRectangle(b, null, new Rect(rect.Left, rect.Top, rect.Width, rect.Height));
        }

        public void FillRectangle(Brush b, RectangleF rect, TClippingStyle clippingStyle)
        {
			FixRect(ref rect);
			if (PointOutside.Check(ref rect)) return;
			if (b == null) return;

			if (clippingStyle == TClippingStyle.None)
			{
                FCanvas.DrawRectangle(b, null, new Rect(rect.Left, rect.Top, rect.Width, rect.Height));
            }
			else
			{
                RectangleGeometry r = new RectangleGeometry(new Rect(rect.Left, rect.Top, rect.Width, rect.Height));
                if (r.CanFreeze) r.Freeze();

                switch (clippingStyle)
                {
                    case TClippingStyle.Exclude:
                       IntersectClip(Reverse(r));
                        break;
                    case TClippingStyle.Include:
                        IntersectClip(r);
                        break;
                    default:
                        //Already handled.
                        break;
                }
            }
        }

        public void FillRectangle (Brush b, real x1, real y1, real width, real height)
        {
			FixRect(ref x1, ref y1, ref width, ref height);
			if (PointOutside.Check(ref x1, ref y1, ref width, ref height)) return;
			if (b == null) return;
            FCanvas.DrawRectangle(b, null, new Rect(x1, y1, width, height));
        }

        public void DrawRectangle(Pen pen, real x, real y, real width, real height)
        {
			FixRect(ref x, ref y, ref width, ref height);
			if (PointOutside.Check(ref x, ref y, ref width, ref height)) return;
			if (pen == null) return;
            FCanvas.DrawRectangle(null, pen, new Rect(x, y, width, height));
        }

		public void DrawAndFillRectangle(Pen pen, Brush b, real x, real y, real width, real height)
		{
			FixRect(ref x, ref y, ref width, ref height);
			if (PointOutside.Check(ref x, ref y, ref width, ref height)) return;
            if (b == null && pen == null) return;

            FCanvas.DrawRectangle(b, pen, new Rect(x, y, width, height));
		}

        public void DrawAndFillBeziers(Pen pen, Brush brush, TPointF[] points)
        {
            DrawAndFillBeziers(pen, brush, points, TClippingStyle.None);
        }

        public void DrawAndFillBeziers(Pen pen, Brush brush, TPointF[] points, TClippingStyle clippingStyle)
        {
            if (points == null || points.Length <= 1) return;
            if (!CheckPoints(points)) return;

            StreamGeometry Line = new StreamGeometry();
            using (StreamGeometryContext sc = Line.Open())
            {
                sc.BeginFigure(new Point(points[0].X, points[0].Y), brush != null, false);
                for (int i = 1; i < points.Length -2; i+= 3)
                {
                    sc.BezierTo(new Point(points[i].X, points[i].Y), new Point(points[i+1].X, points[i+1].Y), new Point(points[i+2].X, points[i+2].Y), true, true);
                }
            }

            if (Line.CanFreeze) Line.Freeze();

			if (brush != null)
            {
                    switch (clippingStyle)
                    {
                        case TClippingStyle.Exclude:
                            IntersectClip(Reverse(Line));
                            break;
                        case TClippingStyle.Include:
                            IntersectClip(Line);
                            break;
                        default:
                            //Will be handled below.
                            break;
                    }
                }

            if (clippingStyle == TClippingStyle.None) FCanvas.DrawGeometry(brush, pen, Line);
        }

        public void DrawAndFillPolygon(Pen pen, Brush brush, TPointF[] points)
        {
            DrawAndFillPolygon(pen, brush, points, TClippingStyle.None);
        }
         
        public void DrawAndFillPolygon(Pen pen, Brush brush, TPointF[] points, TClippingStyle clippingStyle)
        {
            if (points == null || points.Length <= 1) return;
            if (!CheckPoints(points)) return;

            StreamGeometry Line = new StreamGeometry();
            using (StreamGeometryContext sc = Line.Open())
            {
                sc.BeginFigure(new Point(points[0].X, points[0].Y), brush != null, true);
                for (int i = 1; i < points.Length; i++)
                {
                    sc.LineTo(new Point(points[i].X, points[i].Y), true, true);
                }
            }

            if (Line.CanFreeze) Line.Freeze();

            if (brush != null)
            {
                switch (clippingStyle)
                {
                    case TClippingStyle.Exclude:
                        IntersectClip(Reverse(Line));
                        break;
                    case TClippingStyle.Include:
                        IntersectClip(Line);
                        break;
                    default:
                        //Will be handled below.
                        break;
                }
            }

            if (clippingStyle == TClippingStyle.None) FCanvas.DrawGeometry(brush, pen, Line);

        }

        public void DrawBeziers(Pen pen, TPointF[] points)
        {
            DrawAndFillBeziers(pen, null, points);
        }

        public void DrawImage(Image image, RectangleF destRect, RectangleF srcRect, long transparentColor, int brightness, int contrast, int gamma, Color shadowColor, Stream imgData)
        {
			if (PointOutside.Check(ref srcRect)) return;
			if (PointOutside.Check(ref destRect)) return;
			if (image.Height<=0 || image.Width<=0) return;
            bool ChangedParams = brightness!= FlxConsts.DefaultBrightness || 
								 contrast != FlxConsts.DefaultContrast ||
				                 gamma != FlxConsts.DefaultGamma ||
				                 shadowColor != Colors.Transparent;

            ImageAttributes imgAtt=null;
            try
            {
				if (transparentColor != FlxConsts.NoTransparentColor)
				{
					long cl = transparentColor;
					Color tcl = ColorUtil.FromArgb(0xFF, (byte)(cl & 0xFF), (byte)((cl & 0xFF00) >>8), (byte)((cl & 0xFF0000) >>16));
					imgAtt= new ImageAttributes();
					imgAtt.SetColorKey(tcl, tcl);
				}

				if (gamma != FlxConsts.DefaultGamma)
				{
					if (imgAtt==null) imgAtt= new ImageAttributes();
					imgAtt.SetGamma((real)((UInt32)gamma) / 65536f);
				}

				if (!ChangedParams && srcRect.Top==0 && srcRect.Left==0 && srcRect.Width == image.Width && srcRect.Height== image.Height && imgAtt==null)
				{
					FCanvas.DrawImage(image, destRect);  
				}
				else
				{
					Image FinalImage = image;
					try
					{
						if (image.RawFormat.Equals(ImageFormat.Wmf) || image.RawFormat.Equals(ImageFormat.Emf))
						{
							FinalImage = FlgConsts.RasterizeWMF(image); //metafiles do not like cropping or changing attributes.
						}

						if (ChangedParams) 
						{
							if (shadowColor != ColorUtil.Empty) FlgConsts.MakeImageGray(ref imgAtt, shadowColor);
							else
								FlgConsts.AdjustImage(ref imgAtt, brightness, contrast);
						}

						PointF[] ImageRect = new PointF[]{new PointF(destRect.Left, destRect.Top), new PointF(destRect.Right, destRect.Top), new PointF(destRect.Left, destRect.Bottom)};
						FCanvas.DrawImage(FinalImage, 
							ImageRect, 
							srcRect, GraphicsUnit.Pixel, imgAtt);
					}
					finally
					{
						if (FinalImage != image) FinalImage.Dispose();
					}
				}
            }                    
            finally
            {
                if (imgAtt!=null) imgAtt.Dispose();
            }
        }

        public void DrawLine(Pen aPen, real x1, real y1, real x2, real y2)
        {
			PointOutside.Check(ref x1, ref y1);
			PointOutside.Check(ref x2, ref y2);
			if (aPen == null) return;
            FCanvas.DrawLine(aPen, new Point(x1, y1), new Point(x2, y2));
        }

        private Geometry Reverse(Geometry RegionToReverse)
        {
            CombinedGeometry Reversed = new CombinedGeometry(GeometryCombineMode.Exclude, FPageRect, RegionToReverse);
            Geometry Result = Reversed.GetFlattenedPathGeometry();
            if (Result.CanFreeze) Result.Freeze();
            return Result;
        }

        private void IntersectClip(Geometry NewRegion)
        {
            FCanvas.PushClip(NewRegion);
            GraphicsSave[GraphicsSave.Count - 1]++;
        }

        public void SetClipIntersect(RectangleF rect)
        {
			PointOutside.Check(ref rect);
            IntersectClip(new RectangleGeometry(rect));
        }

        public void SetClipReplace(RectangleF rect)
        {
			PointOutside.Check(ref rect);
			FCanvas.SetClip(rect, System.Drawing.Drawing2D.CombineMode.Replace);
        }

        public void ResetClip()
        {
            FCanvas.ResetClip();
        }

		public void AddHyperlink(real x1, real y1, real width, real height, string Url)
		{
			//Not implemented on Screen preview.
		}

		public void AddComment(real x1, real y1, real width, real height, string comment)
		{
			//Not implemented on Screen preview.
		}


        #endregion

    }

}
#endif
