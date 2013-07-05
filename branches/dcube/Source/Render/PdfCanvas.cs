using System;
using System.IO;
using FlexCel.Core;
using FlexCel.Pdf;

#if (WPF)
using RectangleF = System.Windows.Rect;
using SizeF = System.Windows.Size;
using real = System.Double;
using System.Windows.Media;
#else
using real = System.Single;
using System.Drawing;
using System.Drawing.Imaging;
#endif

namespace FlexCel.Render
{
    internal class PdfGraphics: IFlxGraphics
    {
        private PdfWriter Canvas;
		private bool Clipped = false;
		private RectangleF ClipRect = RectangleF.Empty;

        internal PdfGraphics(PdfWriter aCanvas)
        {
            Canvas = aCanvas;
        }

        #region IFlxGraphics Members

		public void CreateSFormat()
		{
		}

		public void DestroySFormat()
		{
		}

        public static RectangleF ConvertToUnits(TPaperDimensions PageBounds)
        {
            real XConv= 1;  //100F/72F;  //100 is Display Units. 72 is a point.
            real YConv= XConv;
            return new RectangleF(0, 0, PageBounds.Width*XConv, PageBounds.Height*YConv);
        }

        public RectangleF ClipBounds
        {
            get
            {
                return ConvertToUnits(Canvas.PageSize);
            }
        }

        public void Rotate(real x, real y, real Alpha)
        {
            Canvas.Rotate(x, y, Alpha);
        }

        public void Scale(real xScale, real yScale)
        {
            Canvas.ScaleBy(xScale, yScale);
        }

		public TPointF Transform(TPointF p)
		{
			return Canvas.Transform(p);
		}

        public void SaveTransform()
        {
            Canvas.SaveState();
        }
        public void ResetTransform()
        {
            Canvas.RestoreState();
        }

        public void DrawString(string Text, Font aFont, Brush aBrush, real x, real y)
        {
            Canvas.DrawString(Text, aFont, aBrush, x, y);
        }

        public void DrawString(string Text, Font aFont, Pen aPen, Brush aBrush, real x, real y)
        {
            Canvas.DrawString(Text, aFont, aPen, aBrush, x, y);
        }

        public SizeF MeasureString(string Text, Font aFont, TPointF p)
        {
            return Canvas.MeasureString(Text, aFont);
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
            return Canvas.MeasureString(Text, aFont);
        }

        public real FontDescent(Font aFont)
        {
            return Canvas.FontDescent(aFont);
        }

        public real FontLinespacing(Font aFont)
        {
            return Canvas.FontLinespacing(aFont);
        }


        public void DrawLines(Pen aPen, TPointF[] points)
        {
            if (aPen == null) return;
            Canvas.DrawLines(aPen, points);
        }

        public void FillRectangle(Brush b, RectangleF rect)
        {
            if (b == null) return;
            Canvas.FillRectangle(b, rect.X, rect.Y, rect.Width, rect.Height);
        }

        public void FillRectangle(Brush b, RectangleF rect, TClippingStyle clippingStyle)
        {
            if (b == null) return;
            switch (clippingStyle)
            {
                case TClippingStyle.None:
                    Canvas.FillRectangle(b, rect.X, rect.Y, rect.Width, rect.Height);
                    break;
                default:
					if (b != null) 
					{
						Canvas.ClipRectangle(rect.X, rect.Y, rect.Width, rect.Height, clippingStyle == TClippingStyle.Exclude);
						SetClipped(true, RectangleF.Empty);

					}
                    break;
            }
        }

        public void FillRectangle (Brush b, real x1, real y1, real width, real height)
        {
            if (b == null) return;
            Canvas.FillRectangle(b, x1, y1, width, height);
        }

        public void DrawRectangle(Pen pen, real x, real y, real width, real height)
        {
            if (pen == null) return;
            Canvas.DrawRectangle(pen, x, y, width, height);
        }

		public void DrawAndFillRectangle(Pen pen, Brush b, real x, real y, real width, real height)
		{
			Canvas.DrawAndFillRectangle(pen, b, x, y, width, height);
		}

        public void DrawAndFillBeziers(Pen pen, Brush brush, TPointF[] points)
        {
            DrawAndFillBeziers(pen, brush, points, TClippingStyle.None);
        }

        public void DrawAndFillBeziers(Pen pen, Brush brush, TPointF[] points, TClippingStyle clippingStyle)
        {
            switch (clippingStyle)
            {
                case TClippingStyle.None:
                    Canvas.DrawAndFillBeziers(pen, brush, points);
                    break;
                default:
					if (brush != null)
					{
						Canvas.ClipBeziers(points, clippingStyle == TClippingStyle.Exclude);
						SetClipped(true, RectangleF.Empty);
					}
					break;
            }
        }

        public void DrawAndFillPolygon(Pen pen, Brush brush, TPointF[] points)
        {
            DrawAndFillPolygon(pen, brush, points, TClippingStyle.None);
        }
         
        public void DrawAndFillPolygon(Pen pen, Brush brush, TPointF[] points, TClippingStyle clippingStyle)
        {
            switch (clippingStyle)
            {
                case TClippingStyle.None:
                    Canvas.DrawAndFillPolygon(pen, brush, points);
                    break;
                default:
					if (brush != null) 
					{
						Canvas.ClipPolygon(points, clippingStyle == TClippingStyle.Exclude);
						SetClipped(true, RectangleF.Empty);
					}
					break;
            }
        }

        public void DrawImage(Image image, RectangleF destRect, RectangleF srcRect, long transparentColor, int brightness, int contrast, int gamma, Color shadowColor, Stream imgData)
        {
			bool ChangedParams = brightness!= FlxConsts.DefaultBrightness || 
				contrast != FlxConsts.DefaultContrast ||
				gamma != FlxConsts.DefaultGamma ||
				shadowColor != ColorUtil.Empty;

			ChangedParams = ChangedParams || (!image.RawFormat.Equals(ImageFormat.Jpeg) && transparentColor != FlxConsts.NoTransparentColor);
            if (!ChangedParams && srcRect.Top==0 && srcRect.Left==0 && srcRect.Width == image.Width && srcRect.Height== image.Height) 
            {
                Canvas.DrawImage(image, destRect, imgData, transparentColor, image.RawFormat.Equals(ImageFormat.Jpeg));
            }     //Optimizes the most common case.
            else
            {
                using (Bitmap bm = BitmapConstructor.CreateBitmap(Convert.ToInt32(srcRect.Width), Convert.ToInt32(srcRect.Height), image.PixelFormat))
                {
                    bm.MakeTransparent();
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
							imgAtt.SetGamma((real)((UInt32)gamma) / 65536f);
						}

						using(Graphics Gr = Graphics.FromImage(bm))
						{
							Rectangle r2=new Rectangle(Convert.ToInt32(srcRect.Left), Convert.ToInt32(srcRect.Top), Convert.ToInt32(srcRect.Width), Convert.ToInt32(srcRect.Height));
							int l=Math.Max(0,-r2.Left);
							int t=Math.Max(0,-r2.Top);
							Rectangle r=new Rectangle( l, t , Math.Min(Convert.ToInt32(image.Width), r2.Width+l), Math.Min(Convert.ToInt32(image.Height), r2.Height+t));

							if (ChangedParams) 
							{
								if (shadowColor != ColorUtil.Empty) FlgConsts.MakeImageGray(ref imgAtt, shadowColor);
								else
									FlgConsts.AdjustImage(ref imgAtt, brightness, contrast);
							}

							Image FinalImage = image;
							try
							{
								if (image.RawFormat.Equals(ImageFormat.Wmf) || image.RawFormat.Equals(ImageFormat.Emf))
								{
									FinalImage = FlgConsts.RasterizeWMF(image); //metafiles do not like cropping or changing attributes.
								}


								Gr.DrawImage(FinalImage, r,
									Math.Max(0, r2.Left), Math.Max(0,r2.Top), Math.Min(Convert.ToInt32(image.Width), r2.Width+l), Math.Min(Convert.ToInt32(image.Height), r2.Height+t), GraphicsUnit.Pixel, imgAtt);
							}
							finally
							{
								if (FinalImage != image) FinalImage.Dispose();
							}


							bool DefaultToJpeg=image.RawFormat.Equals(ImageFormat.Jpeg);
							Canvas.DrawImage(bm, destRect, null, FlxConsts.NoTransparentColor, DefaultToJpeg);
						}
                    }
                    finally
                    {
                        if (imgAtt!=null) imgAtt.Dispose();
                    }

                }
            }
        }

        public void DrawLine(Pen aPen, real x1, real y1, real x2, real y2)
        {
            if (aPen == null) return;
            Canvas.DrawLine(aPen, x1, y1, x2, y2);
        }

		private void SetClipped(bool clipState, RectangleF r)
		{
			Clipped = clipState;
			ClipRect = r;
		}

        public void SetClipIntersect(RectangleF rect)
        {
            Canvas.IntersectClipRegion(rect);
			SetClipped(true, RectangleF.Empty);
		}

        public void SetClipReplace(RectangleF rect)
        {
			if (ClipRect != RectangleF.Empty && ClipRect == rect) return;
			if (Clipped)
			{
				Canvas.RestoreState();  //We always have the original state saved.
				Canvas.SaveState();
			}
            Canvas.IntersectClipRegion(rect);
			SetClipped(true, rect);
		}

        public void ResetClip()
        {
			if (!Clipped) return;
			Canvas.RestoreState();  //We always have the original state saved.
            Canvas.SaveState();
			SetClipped(false, RectangleF.Empty);
		}

        public void SaveState() //always follow with restorestate
        {
            Canvas.SaveState();
        }

        public void RestoreState() 
        {
            Canvas.RestoreState();
        }

		public void AddHyperlink(real x1, real y1, real width, real height, string Url)
		{
            try
            {
                Canvas.Hyperlink(x1, y1, width, height, Url);
            }
            catch (UriFormatException ex)
            {
                if (FlexCelTrace.HasListeners) FlexCelTrace.Write(new TMalformedUrlError(ex.Message, Url));
            }
		}

		private static readonly TPdfCommentProperties StandardPdfCommentProps = 
			new TPdfCommentProperties(TPdfCommentType.Square, TPdfCommentIcon.Note, 0, Color.White, Color.Black); //STATIC**
		public void AddComment(real x1, real y1, real width, real height, string comment)
		{
			Canvas.Comment(x1, y1, width, height, comment, StandardPdfCommentProps);
		}


        #endregion

    }

}
