#if (!WPF)
using System;
using System.IO;
using FlexCel.Core;

using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Collections.Generic;

namespace FlexCel.Render
{
    internal class GdiPlusGraphics: IFlxGraphics
    {
        private Graphics Canvas;
		private StringFormat FSFormat;
        private Stack<GraphicsState> GraphicsSave;

        internal GdiPlusGraphics(Graphics aCanvas)
        {
            Canvas = aCanvas;
            GraphicsSave = new Stack<GraphicsState>();
		}

        #region IFlxGraphics Members
		public void CreateSFormat()
		{
			using (StringFormat sfTemplate = StringFormat.GenericTypographic) //GenericTypographic returns a NEW instance. It has to be disposed.
			{

				FSFormat = (StringFormat) sfTemplate.Clone(); //Even when sfTemplate is a new instance, changing directly on it will change the standard generic typographic :(
				try
				{
                    FSFormat.Alignment = StringAlignment.Near; //this should be set, but just in case someone changed it.
                    FSFormat.Trimming = StringTrimming.None;
					FSFormat.FormatFlags |= StringFormatFlags.NoClip | StringFormatFlags.MeasureTrailingSpaces;
					FSFormat.FormatFlags &= ~StringFormatFlags.LineLimit;
					FSFormat.LineAlignment = StringAlignment.Far;
				}
				catch
				{
					FSFormat.Dispose();
					FSFormat = null;
					throw;
				}
			}
		}

		public void DestroySFormat()
		{
			if (FSFormat != null) FSFormat.Dispose();
			FSFormat = null;
		}


        public RectangleF ClipBounds
        {
            get
            {
                return Canvas.ClipBounds;
            }
        }

        public void Rotate(float x, float y, float Alpha)
        {
			if (PointOutside.Check(ref x, ref y)) return;
            using (Matrix myMatrix = Canvas.Transform)
            {
                myMatrix.RotateAt(-Alpha, new PointF(x,y), MatrixOrder.Prepend);
                Canvas.Transform = myMatrix;
            }
        }

		public TPointF Transform(TPointF p)
		{
			PointOutside.Check(ref p);
			using (Matrix myMatrix = Canvas.Transform)
			{
				float[] DrawingMatrix = myMatrix.Elements;
				return new TPointF(
				(float)(DrawingMatrix[0] * p.X + DrawingMatrix[2] * p.Y + DrawingMatrix[4]),
				(float)(DrawingMatrix[1] * p.X + DrawingMatrix[3] * p.Y + DrawingMatrix[5]));
			}
		}

        public void Scale(float xScale, float yScale)
        {
            using (Matrix myMatrix = Canvas.Transform)
            {
                myMatrix.Scale(xScale, yScale, MatrixOrder.Prepend);
                Canvas.Transform = myMatrix;
            }
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
            GraphicsSave.Push(Canvas.Save());
        }

        public void RestoreState()
        {
            Canvas.Restore(GraphicsSave.Pop());
        }

        public void DrawString(string Text, Font aFont, Brush aBrush, float x, float y)
        {
			if (PointOutside.Check(ref x, ref y)) return;
			Text = Text.Replace('\u000a','\u0000'); //Replace newlines with empty characters
            Canvas.DrawString(Text, aFont, aBrush, x, y, FSFormat);
        }

        public void DrawString(string Text, Font aFont, Pen aPen, Brush aBrush, float x, float y)
        {
			if (PointOutside.Check(ref x, ref y)) return;
			Text = Text.Replace('\u000a','\u0000'); //Replace newlines with empty characters
            if (aBrush != null) Canvas.DrawString(Text, aFont, aBrush, x, y, FSFormat);
            if (aPen != null)
            {
                GraphicsPath TextPath = new GraphicsPath();
                TextPath.AddString(Text, aFont.FontFamily, (int)aFont.Style, aFont.Size, new PointF(x,y), FSFormat);
                Canvas.DrawPath(aPen, TextPath);
            }
        }

        public SizeF MeasureString(string Text, Font aFont, TPointF p)
        {
			PointOutside.Check(ref p);
			return Canvas.MeasureString(Text, aFont, new PointF(p.X, p.Y), FSFormat);
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
            return Canvas.MeasureString(Text, aFont, new PointF(0,0), FSFormat);
        }

		public float FontDescent(Font aFont)
		{
			FontFamily ff = aFont.FontFamily; //don't dispose
			return aFont.GetHeight(Canvas) / ff.GetLineSpacing(aFont.Style) * ff.GetCellDescent(aFont.Style);	
		}

        public float FontLinespacing(Font aFont)
        {
            FontFamily ff = aFont.FontFamily; //don't dispose
            float LineGap = ff.GetLineSpacing(aFont.Style) - ff.GetCellDescent(aFont.Style) - ff.GetCellAscent(aFont.Style);
            return aFont.GetHeight(Canvas)/ ff.GetLineSpacing(aFont.Style) * LineGap;	
        }

        private static PointF[] ToPointF(TPointF[] points)
        {
			bool AllPointsOutside = true;

            PointF[] Result = new PointF[points.Length];
			for (int i=0; i< points.Length;i++)
			{
				if (!PointOutside.Check(ref points[i])) AllPointsOutside = false;
				Result[i] = new PointF(points[i].X, points[i].Y);
			}

			if (AllPointsOutside) return null;
            return Result;
        }

        public void DrawLines(Pen aPen, TPointF[] points)
        {
            if (aPen == null || points == null || points.Length <=1) return;
			PointF[] fpoints = ToPointF(points);
			if (fpoints == null) return;
            Canvas.DrawLines(aPen, fpoints);
        }

		internal static void FixRect(ref float x, ref float y, ref float width, ref float height)
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
			if (b != null) Canvas.FillRectangle(b, rect);
        }

        public void FillRectangle(Brush b, RectangleF rect, TClippingStyle clippingStyle)
        {
			FixRect(ref rect);
			if (PointOutside.Check(ref rect)) return;
			if (b == null) return;

			if (clippingStyle == TClippingStyle.None)
			{
				Canvas.FillRectangle(b, rect);
			}
			else
			{
				using (GraphicsPath gp = new GraphicsPath())
				{
					gp.AddRectangle(rect);
					switch (clippingStyle)
					{
						case TClippingStyle.Exclude:
							Canvas.SetClip(gp, CombineMode.Exclude);
							break;
						case TClippingStyle.Include:
							Canvas.SetClip(gp, CombineMode.Intersect);
							break;
					}
				}
			}
        }

        public void FillRectangle (Brush b, float x1, float y1, float width, float height)
        {
			FixRect(ref x1, ref y1, ref width, ref height);
			if (PointOutside.Check(ref x1, ref y1, ref width, ref height)) return;
			if (b == null) return;
            Canvas.FillRectangle(b, x1, y1, width, height);
        }

        public void DrawRectangle(Pen pen, float x, float y, float width, float height)
        {
			FixRect(ref x, ref y, ref width, ref height);
			if (PointOutside.Check(ref x, ref y, ref width, ref height)) return;
			if (pen == null) return;
            Canvas.DrawRectangle(pen, x, y, width, height);
        }

		public void DrawAndFillRectangle(Pen pen, Brush b, float x, float y, float width, float height)
		{
			FixRect(ref x, ref y, ref width, ref height);
			if (PointOutside.Check(ref x, ref y, ref width, ref height)) return;
			if (b != null)
			{
				Canvas.FillRectangle(b, x, y, width, height);
			}
			if (pen != null)
			{
				Canvas.DrawRectangle(pen, x, y, width, height);
			}
		}

        public void DrawAndFillBeziers(Pen pen, Brush brush, TPointF[] points)
        {
            DrawAndFillBeziers(pen, brush, points, TClippingStyle.None);
        }

        public void DrawAndFillBeziers(Pen pen, Brush brush, TPointF[] points, TClippingStyle clippingStyle)
        {
            if (points.Length < 4 || (points.Length - 4) % 3 != 0)
            {
                int TotalPoints = 4;
                if (points.Length > TotalPoints) TotalPoints = 7 + ((points.Length - 5) / 3) * 3;
                TPointF[] newpoints = new TPointF[TotalPoints];
                Array.Copy(points, newpoints, points.Length);
                for (int i = points.Length; i < TotalPoints; i++)
                {
                    newpoints[i] = points[points.Length - 1];
                }
                points = newpoints;
            }

            PointF[] fpoints = ToPointF(points);
			if (fpoints == null) return;
			if (brush != null)
            {
                using (GraphicsPath gp = new GraphicsPath())
                {
                    gp.AddBeziers(fpoints);
                    switch (clippingStyle)
                    {
                        case TClippingStyle.Exclude:
                            Canvas.SetClip(gp, CombineMode.Exclude);
                            break;
                        case TClippingStyle.Include:
                            Canvas.SetClip(gp, CombineMode.Intersect);
                            break;
                        default:
                            Canvas.FillPath(brush, gp);
                            break;
                    }
                }
            }
            if (pen != null && clippingStyle == TClippingStyle.None) Canvas.DrawBeziers(pen, fpoints);
        }

        public void DrawAndFillPolygon(Pen pen, Brush brush, TPointF[] points)
        {
            DrawAndFillPolygon(pen, brush, points, TClippingStyle.None);
        }
         
        public void DrawAndFillPolygon(Pen pen, Brush brush, TPointF[] points, TClippingStyle clippingStyle)
        {
            PointF[] fpoints = ToPointF(points);
			if (fpoints == null) return;
			if (brush != null)
            {
                if (clippingStyle == TClippingStyle.None)
                {
                    Canvas.FillPolygon(brush, fpoints);
                }
                else
                {
                    using (GraphicsPath gp = new GraphicsPath())
                    {
                        gp.AddLines(fpoints);
                        switch (clippingStyle)
                        {
                            case TClippingStyle.Exclude:
                                Canvas.SetClip(gp, CombineMode.Exclude);
                                break;
                            case TClippingStyle.Include:
                                Canvas.SetClip(gp, CombineMode.Intersect);
                                break;
                        }
                    }
                }
            }
            if (pen != null && clippingStyle == TClippingStyle.None) Canvas.DrawPolygon(pen, fpoints);
        }

        public void DrawBeziers(Pen pen, TPointF[] points)
        {
            if (pen == null) return;
            PointF[] fpoints = ToPointF(points);
			if (fpoints == null) return;
			Canvas.DrawBeziers(pen, fpoints);
        }

        public void DrawImage(Image image, RectangleF destRect, RectangleF srcRect, long transparentColor, int brightness, int contrast, int gamma, Color shadowColor, Stream imgData)
        {
			if (PointOutside.Check(ref srcRect)) return;
			if (PointOutside.Check(ref destRect)) return;
			if (image.Height<=0 || image.Width<=0) return;
            bool ChangedParams = brightness!= FlxConsts.DefaultBrightness || 
								 contrast != FlxConsts.DefaultContrast ||
				                 gamma != FlxConsts.DefaultGamma ||
				                 shadowColor != ColorUtil.Empty;

            ImageAttributes imgAtt=null;
            try
            {
				if (transparentColor != FlxConsts.NoTransparentColor)
				{
					long cl = transparentColor;
					Color tcl = ColorUtil.FromArgb((int)(cl & 0xFF), (int)((cl & 0xFF00) >>8), (int)((cl & 0xFF0000) >>16));
					imgAtt= new ImageAttributes();
					imgAtt.SetColorKey(tcl, tcl);
				}

				if (gamma != FlxConsts.DefaultGamma)
				{
					if (imgAtt==null) imgAtt= new ImageAttributes();
					imgAtt.SetGamma((float)((UInt32)gamma) / 65536f);
				}

				if (!ChangedParams && srcRect.Top==0 && srcRect.Left==0 && srcRect.Width == image.Width && srcRect.Height== image.Height && imgAtt==null)
				{
					bool Retry = false;
					try
					{
						Canvas.DrawImage(image, destRect);  //Optimizes the most common case, but there is also an error when cropping metafiles. At least we make sure here this won't have any error on 99% of the cases.
					}
					catch (System.Runtime.InteropServices.ExternalException ex)
					{
						if (FlexCelTrace.HasListeners) FlexCelTrace.Write(new TRenderMetafileError(ex.Message));

						Retry = true; //metafiles can raise nasty errors here.
					}
					if (Retry)
					{
						using (Image bmp = FlgConsts.RasterizeWMF(image))
						{
							Canvas.DrawImage(bmp, destRect);
						}
					}
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
						Canvas.DrawImage(FinalImage, 
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

        public void DrawLine(Pen aPen, float x1, float y1, float x2, float y2)
        {
			PointOutside.Check(ref x1, ref y1);
			PointOutside.Check(ref x2, ref y2);
			if (aPen == null) return;
            Canvas.DrawLine(aPen, x1, y1, x2, y2);
        }

        public void SetClipIntersect(RectangleF rect)
        {
			PointOutside.Check(ref rect);
			Canvas.SetClip(rect, System.Drawing.Drawing2D.CombineMode.Intersect);
        }

        public void SetClipReplace(RectangleF rect)
        {
			PointOutside.Check(ref rect);
			Canvas.SetClip(rect, System.Drawing.Drawing2D.CombineMode.Replace);
        }

        public void ResetClip()
        {
            Canvas.ResetClip();
        }

		public void AddHyperlink(float x1, float y1, float width, float height, string Url)
		{
			//Not implemented on Screen preview.
		}

		public void AddComment(float x1, float y1, float width, float height, string comment)
		{
			//Not implemented on Screen preview.
		}


        #endregion

    }

}
#endif
