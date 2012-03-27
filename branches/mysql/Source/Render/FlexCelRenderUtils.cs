using System;
using System.Runtime.CompilerServices;
using FlexCel.Core;

using System.Collections.Generic;

#if (MONOTOUCH)
using System.Drawing;
using real = System.Single;
using Color = MonoTouch.UIKit.UIColor;
using Font = MonoTouch.UIKit.UIFont;
#else
	#if (WPF)
	using SizeF = System.Windows.Size;
	using PointF = System.Windows.Point;
	using RectangleF = System.Windows.Rect;
	using System.Windows.Media;
	using real = System.Double;
	using ColorBlend = System.Windows.Media.GradientStopCollection;
	#else
	using real = System.Single;
	
	using System.Drawing;
	using System.Drawing.Drawing2D;
	using System.Drawing.Imaging;
	#endif
#endif

namespace FlexCel.Render
{
    internal enum TGridDrawState
    {
        Normal     
    }

    internal class TXRichString
    {
        internal TRichString s;
        internal bool Split;
        internal real XExtent;
        internal real YExtent;
        internal TAdaptativeFormats AdaptFormat;

        internal TXRichString(TRichString aString, bool aSplit, real aXExtent, real aYExtent, TAdaptativeFormats aAdaptFormat)
        {
            s = aString;
            Split = aSplit;
            XExtent = aXExtent;
            YExtent = aYExtent;
            AdaptFormat = aAdaptFormat;
        }
    }

#if (FRAMEWORK20)
    internal class TXRichStringList : List<TXRichString>
    {
    }

    internal class TFloatList : List<real>
    {
    }

    internal class TSpawnedCellList : Dictionary<long,object>
    {
        internal TSpawnedCellList()
        {
        }
    }

    internal class TCellMergedCache : Dictionary<long,TXlsCellRange>
    {
        internal TCellMergedCache()
        {
        }
    }

#else

    internal class TXRichStringList: ArrayList
    {
        internal new TXRichString this[int index]
        {
            get {return (TXRichString)base[index];}
            set 
            {
                base[index]=value;
            }
        }
    }

    /// <summary>
    /// Could be made with a template to avoid boxing, but it is not too much used.
    /// </summary>
    internal class TFloatList: ArrayList
    {
        internal new real this[int index]
        {
            get {return (real)base[index];}
            set 
            {
                base[index]=value;
            }
        }
    }

    internal class TSpawnedCellList: Hashtable
    {
        internal TSpawnedCellList()
        {
        }
    }

    internal class TCellMergedCache: Hashtable
    {
        internal TCellMergedCache()
        {
        }

        internal bool TryGetValue(object key, out TXlsCellRange Result)
        {
            Result = (TXlsCellRange)this[key];
            return (Result != null);
        }
        
    }
#endif


    internal enum TVAlign
    {
        Top,
        Center,
        Bottom
    }

    internal enum THAlign
    {
        Left,
        Center,
        Right
    }

    internal sealed class FlgConsts
    {
        private FlgConsts(){}

        private static void FillRect(Bitmap Ac, int x1, int y1, int x2, int y2, Color aColor)
        {
            for (int x=x1; x<x2;x++)
                for (int y=y1; y<y2; y++)
                {
                    Ac.SetPixel(x,y,aColor);
                }
        }

        /// <summary>
        /// Deprecated by the moment. Could be of use in the future.
        /// </summary>
        /*
        internal static Brush CreateBmpPattern( TFlxPatternStyle Pattern, Color ColorFg, Color ColorBg )
        {
            Bitmap Ac=null;
            switch (Pattern)
            {
                case TFlxPatternStyle.None: //No pattern
                    return new SolidBrush(ColorBg);
        
                case TFlxPatternStyle.Solid: //fill pattern
                    return new SolidBrush(ColorFg);
                
                case TFlxPatternStyle.Gray50: //50%
                    Ac= new Bitmap(4,4, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                    for (int y=0; y<4; y++)
                        for (int x=0; x<4; x++) 
                            if ((x + (y%2))%2==0)
                                Ac.SetPixel(x, y ,ColorFg);
                            else
                                Ac.SetPixel(x, y ,ColorBg);
                    return new TextureBrush(Ac);

                case TFlxPatternStyle.Gray75: //75%
                    Ac= new Bitmap(4,4, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                    FillRect(Ac,0,0,4,4, ColorFg );
                    Ac.SetPixel(0,0,ColorBg);
                    Ac.SetPixel(2,1,ColorBg);
                    Ac.SetPixel(0,2,ColorBg);
                    Ac.SetPixel(2,3,ColorBg);
                    return new TextureBrush(Ac);

                case TFlxPatternStyle.Gray25: //25%
                    Ac= new Bitmap(4,4, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                    FillRect(Ac,0,0,4,4,ColorBg);
                    Ac.SetPixel(0,0,ColorFg);
                    Ac.SetPixel(2,1,ColorFg);
                    Ac.SetPixel(0,2,ColorFg);
                    Ac.SetPixel(2,3,ColorFg);
                    return new TextureBrush(Ac);

                case TFlxPatternStyle.Horizontal: //Horz lines
                    Ac= new Bitmap(4,4, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                    FillRect(Ac,0,0,4,2,ColorFg);
                    FillRect(Ac,0,2,4,4,ColorBg);
                    return new TextureBrush(Ac);

                case TFlxPatternStyle.Vertical: //Vert lines
                    Ac= new Bitmap(4,4, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                    FillRect(Ac,0,0,2,4,ColorFg);
                    FillRect(Ac,2,0,4,4,ColorBg);
                    return new TextureBrush(Ac);

                case TFlxPatternStyle.Down: //   \ lines
                    Ac= new Bitmap(4,4, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                    FillRect(Ac,0,0,4,4,ColorBg);
                    Ac.SetPixel(0,0,ColorFg); Ac.SetPixel(1,0,ColorFg);
                    Ac.SetPixel(1,1,ColorFg); Ac.SetPixel(2,1,ColorFg);
                    Ac.SetPixel(2,2,ColorFg); Ac.SetPixel(3,2,ColorFg);
                    Ac.SetPixel(3,3,ColorFg); Ac.SetPixel(0,3,ColorFg);
                    return new TextureBrush(Ac);

                case TFlxPatternStyle.Up: //   / lines
                    Ac= new Bitmap(4,4, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                    FillRect(Ac, 0,0,4,4,ColorBg);
                    Ac.SetPixel(2,0,ColorFg); Ac.SetPixel(3,0,ColorFg);
                    Ac.SetPixel(1,1,ColorFg); Ac.SetPixel(2,1,ColorFg);
                    Ac.SetPixel(0,2,ColorFg); Ac.SetPixel(1,2,ColorFg);
                    Ac.SetPixel(3,3,ColorFg); Ac.SetPixel(0,3,ColorFg);
                    return new TextureBrush(Ac);

                case TFlxPatternStyle.Checker: //  diagonal hatch
                    Ac= new Bitmap(4,4, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                    FillRect(Ac,0,0,4,4,ColorBg);
                    Ac.SetPixel(0,0,ColorFg); Ac.SetPixel(1,0,ColorFg);
                    Ac.SetPixel(0,1,ColorFg); Ac.SetPixel(1,1,ColorFg);
                    Ac.SetPixel(2,2,ColorFg); Ac.SetPixel(3,2,ColorFg);
                    Ac.SetPixel(2,3,ColorFg); Ac.SetPixel(3,3,ColorFg);
                    return new TextureBrush(Ac);

                case TFlxPatternStyle.SemiGray75: //  bold diagonal
                    Ac= new Bitmap(4,4, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                    FillRect(Ac,0,0,4,4,ColorFg);
                    Ac.SetPixel(2,0,ColorBg); Ac.SetPixel(3,0,ColorBg);
                    Ac.SetPixel(0,2,ColorBg); Ac.SetPixel(1,2,ColorBg);
                    return new TextureBrush(Ac);

                case TFlxPatternStyle.LightHorizontal: //  thin horz lines
                    Ac= new Bitmap(4,4, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                    FillRect(Ac,0,0,4,1,ColorFg);
                    FillRect(Ac,0,1,4,4,ColorBg);
                    return new TextureBrush(Ac);
            
                case TFlxPatternStyle.LightVertical: //  thin vert lines
                    Ac= new Bitmap(4,4, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                    FillRect(Ac,0,0,1,4,ColorFg);
                    FillRect(Ac,1,0,4,4,ColorBg);
                    return new TextureBrush(Ac);

                case TFlxPatternStyle.LightDown: //  thin \ lines
                    Ac= new Bitmap(4,4, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                    FillRect(Ac,0,0,4,4,ColorBg);
                    Ac.SetPixel(0,0,ColorFg);
                    Ac.SetPixel(1,1,ColorFg);
                    Ac.SetPixel(2,2,ColorFg);
                    Ac.SetPixel(3,3,ColorFg);
                    return new TextureBrush(Ac);

                case TFlxPatternStyle.LightUp: //  thin / lines
                    Ac= new Bitmap(4,4, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                    FillRect(Ac,0,0,4,4,ColorBg);
                    Ac.SetPixel(3,0,ColorFg);
                    Ac.SetPixel(2,1,ColorFg);
                    Ac.SetPixel(1,2,ColorFg);
                    Ac.SetPixel(0,3,ColorFg);
                    return new TextureBrush(Ac);

                case TFlxPatternStyle.Grid: //  thin horz hatch
                    Ac= new Bitmap(4,4, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                    FillRect(Ac,0,0,4,4,ColorFg);
                    FillRect(Ac,1,1,4,4,ColorBg);
                    return new TextureBrush(Ac);
            
                case TFlxPatternStyle.CrissCross: //  thin diag
                    Ac= new Bitmap(4,4, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                    FillRect(Ac,0,0,4,4,ColorBg);
                    Ac.SetPixel(0,0,ColorFg); Ac.SetPixel(2,0,ColorFg);
                    Ac.SetPixel(1,1,ColorFg);
                    Ac.SetPixel(0,2,ColorFg); Ac.SetPixel(2,2,ColorFg);
                    Ac.SetPixel(3,3,ColorFg);
                    return new TextureBrush(Ac);
            
                case TFlxPatternStyle.Gray16: //  12.5 %
                    Ac= new Bitmap(4,4, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                    FillRect(Ac,0,0,4,4,ColorBg);
                    Ac.SetPixel(0,0,ColorFg);
                    Ac.SetPixel(2,2,ColorFg);
                    return new TextureBrush(Ac);
            
                case TFlxPatternStyle.Gray8: //  6.25 % gray
                    Ac= new Bitmap(8,8, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                    FillRect(Ac,0,0,8,8,ColorBg);
                    Ac.SetPixel(0,0,ColorFg);
                    Ac.SetPixel(4,2,ColorFg);
                    Ac.SetPixel(0,4,ColorFg);
                    Ac.SetPixel(4,6,ColorFg);
                    return new TextureBrush(Ac);
            }//case

            return new SolidBrush(ColorBg);
        }
        */

        internal static Brush CreatePattern(TFlxPatternStyle Pattern, Color ColorFg, Color ColorBg, TExcelGradient GradientFill, RectangleF CellRect, IFlexCelPalette Palette)
        {
            switch (Pattern)
            {
                case TFlxPatternStyle.None: //No pattern
                    return null;
        
                case TFlxPatternStyle.Solid: //fill pattern
                    return new SolidBrush(ColorFg);
                
                case TFlxPatternStyle.Gray50: //50%
                    return new HatchBrush(HatchStyle.Percent50, ColorFg, ColorBg);

                case TFlxPatternStyle.Gray75: //75%
                    return new HatchBrush(HatchStyle.Percent75, ColorFg, ColorBg);

                case TFlxPatternStyle.Gray25: //25%
                    return new HatchBrush(HatchStyle.Percent25, ColorFg, ColorBg);

                case TFlxPatternStyle.Horizontal: //Horz lines
                    return new HatchBrush(HatchStyle.DarkHorizontal, ColorFg, ColorBg);

                case TFlxPatternStyle.Vertical: //Vert lines
                    return new HatchBrush(HatchStyle.DarkVertical, ColorFg, ColorBg);

                case TFlxPatternStyle.Down: //   \ lines
                    return new HatchBrush(HatchStyle.DarkDownwardDiagonal, ColorFg, ColorBg);

                case TFlxPatternStyle.Up: //   / lines
                    return new HatchBrush(HatchStyle.DarkUpwardDiagonal, ColorFg, ColorBg);

                case TFlxPatternStyle.Checker: //  diagonal hatch
                    return new HatchBrush(HatchStyle.SmallCheckerBoard, ColorFg, ColorBg);

                case TFlxPatternStyle.SemiGray75: //  bold diagonal
                    return new HatchBrush(HatchStyle.Percent70, ColorFg, ColorBg);

                case TFlxPatternStyle.LightHorizontal: //  thin horz lines
                    return new HatchBrush(HatchStyle.LightHorizontal, ColorFg, ColorBg);
            
                case TFlxPatternStyle.LightVertical: //  thin vert lines
                    return new HatchBrush(HatchStyle.LightVertical, ColorFg, ColorBg);

                case TFlxPatternStyle.LightDown: //  thin \ lines
                    return new HatchBrush(HatchStyle.LightDownwardDiagonal, ColorFg, ColorBg);

                case TFlxPatternStyle.LightUp: //  thin / lines
                    return new HatchBrush(HatchStyle.LightUpwardDiagonal, ColorFg, ColorBg);

                case TFlxPatternStyle.Grid: //  thin horz hatch
                    return new HatchBrush(HatchStyle.SmallGrid, ColorFg, ColorBg);
            
                case TFlxPatternStyle.CrissCross: //  thin diag
                    return new HatchBrush(HatchStyle.Percent60, ColorFg, ColorBg);
            
                case TFlxPatternStyle.Gray16: //  12.5 %
                    return new HatchBrush(HatchStyle.Percent10, ColorFg, ColorBg);
            
                case TFlxPatternStyle.Gray8: //  6.25 % gray
                    return new HatchBrush(HatchStyle.Percent05, ColorFg, ColorBg);

                case TFlxPatternStyle.Gradient:
                    return GetGradient(ref ColorFg, ref ColorBg, GradientFill, CellRect, Palette);
            }//case

            return new SolidBrush(ColorBg);
        }

        private static Brush GetGradient(ref Color ColorFg, ref Color ColorBg, TExcelGradient GradientFill, RectangleF CellRect, IFlexCelPalette Palette)
        {
            if (GradientFill == null) return new SolidBrush(ColorFg);

            Color Color1 = ColorFg;
            Color Color2 = ColorBg;
            if (GradientFill.Stops != null && GradientFill.Stops.Length > 0)
            {
                Color1 = GradientFill.Stops[0].Color.ToColor(Palette);
                Color2 = GradientFill.Stops[GradientFill.Stops.Length - 1].Color.ToColor(Palette);
            }

            switch (GradientFill.GradientType)
            {
                case TGradientType.Linear:
                    {
                        TExcelLinearGradient lgr = GradientFill as TExcelLinearGradient;
                        LinearGradientBrush Result = new LinearGradientBrush(CellRect, Color1, Color2, (real)(lgr.RotationAngle), true);
                        Result.WrapMode = WrapMode.TileFlipXY;
                        ColorBlend cb = GetInterpolationColors(Color1, Color2, Palette, lgr, false);
                        if (cb != null) Result.InterpolationColors = cb;
                        return Result;
                    }
                case TGradientType.Rectangular:
                    {
                        TExcelRectangularGradient rgr = GradientFill as TExcelRectangularGradient;
                        PathGradientBrush Result = new PathGradientBrush(new PointF[] {
																					  CellRect.Location,
																					  new PointF(CellRect.Right, CellRect.Top),
																					  new PointF(CellRect.Right, CellRect.Bottom),
																					  new PointF(CellRect.Left, CellRect.Bottom)
																				  });
                        Result.WrapMode = WrapMode.TileFlipXY;
                        Result.CenterPoint = new PointF(CellRect.Left + (float)rgr.Left * CellRect.Width, CellRect.Top + (float)rgr.Top * CellRect.Height);
                        Result.CenterColor = Color1;
                        Result.SurroundColors = new Color[] { Color2 };
                        ColorBlend cb = GetInterpolationColors(Color1, Color2, Palette, rgr, true); //must be (Color1, Color2), not (Color2, Color1) even if the blending will be reversed. This is because it will be reversed *after* Color1 and 2 have been added.
                        if (cb != null) Result.InterpolationColors = cb;

                        return Result;
                    }
            }
            return new SolidBrush(ColorFg);
        }

        private static ColorBlend GetInterpolationColors(Color Color1, Color Color2, IFlexCelPalette Palette, TExcelGradient gr, bool InvertColors)
        {
            ColorBlend Result = null;
            if (gr.Stops != null && gr.Stops.Length > 0) //It should be gr.Stops.Length > 2, but we will add them even if not needed, so there aren't exceptions when retrieving them.
            {
                Result = new ColorBlend(gr.Stops.Length);
                for (int i = 0; i < gr.Stops.Length; i++)
                {
                    Result.Colors[i] = gr.Stops[i].Color.ToColor(Palette);
                    Result.Positions[i] = (float)gr.Stops[i].Position;
                }

                FlxGradient.EnsureMinimumAndMaximum(Color1, Color2, ref Result);
                if (InvertColors) FlxGradient.InvertColorBlend(Result);
            }

            return Result;
        }

#if(!COMPACTFRAMEWORK)
        [MethodImpl(MethodImplOptions.NoInlining)]
        internal static void DoAdjustImage(ref ImageAttributes imgAtt, int brightness, int contrast)
        {
            if (imgAtt==null) imgAtt= new ImageAttributes();
            real cf = ((real)contrast)/ FlxConsts.DefaultContrast;   //between 0 and infinitum.
            real bf = (real)brightness/ 0x8000;  //between -1 and +1

            /*To modify contrast, excel substracts half one color, multiplies the rgb components, and adds the half color again.
            The formula for modifying contrast/brightness at the same time sould be:
			
             For Contrast: Multiply by:
                 |1             |       |cf             |       |1             |
                 |   1          |       |   cf          |       |   1          |
                 |      1       |   x   |      cf       |   x   |      1       |
                 |         1    |       |         1     |       |         1    |
                 |-.5 -.5 -.5  1|       |            1  |       |.5 .5 .5     1|
				 
            This will substract the color, modify the contrast, and add the color again.
			
            For Brightness: Multiply by:
                 |1             |
                 |   1          |
                 |      1       |
                 |         1    |
                 |br br br     1|
				 
            The issue is that while this 2 methods work separately, multiplying those matrices
            is not what Excel does when you modify both at the same time.
	 				
            */
            real a = 0.5f * (bf - 1f);
            real b = 0.5f * (bf + 1f);
			
            real af = cf * a + b; 

            ColorMatrix rm = 
                new ColorMatrix(new real[][]
                            {
                                new real[]{cf,    0f,    0f,    0f,    0f},
                                new real[]{0f,    cf,    0f,    0f,    0f},
                                new real[]{0f,    0f,    cf,    0f,    0f},
                                new real[]{0f,    0f,    0f,    1f,    0f},
                                new real[]{af,    af,    af,    0f,    1f}
                            });
            
            imgAtt.SetColorMatrix(rm);
        }

		[MethodImpl(MethodImplOptions.NoInlining)]
		private static void DoMakeImageGray(ref ImageAttributes imgAtt, Color shadowColor)
		{
			if (imgAtt==null) imgAtt = new ImageAttributes();

			ColorMatrix rm = 
				new ColorMatrix(new real[][]
							{
								new real[]{0f,    0f,    0f,    0f,    0f},
								new real[]{0f,    0f,    0f,    0f,    0f},
								new real[]{0f,    0f,    0f,    0f,    0f},
								new real[]{0f,    0f,    0f,    shadowColor.A/255f,    0f},
								new real[]{shadowColor.R/255f,    shadowColor.G/255f,    shadowColor.B/255f,    0,    1f}
							});

			imgAtt.SetColorMatrix(rm);
		}

		[MethodImpl(MethodImplOptions.NoInlining)]
		private static void DoMakeTransparent(ref ImageAttributes imgAtt, real Opacity)
		{
			if (imgAtt==null) imgAtt = new ImageAttributes();

			ColorMatrix rm = 
				new ColorMatrix(new real[][]
							{
								new real[]{1f,    0f,    0f,    0f,    0f},
								new real[]{0f,    1f,    0f,    0f,    0f},
								new real[]{0f,    0f,    1f,    0f,    0f},
								new real[]{0f,    0f,    0f,    Opacity,    0f},
								new real[]{0f,    0f,    0f,    0f,    1f}
							});
  
          
			imgAtt.SetColorMatrix(rm);
		}

		[MethodImpl(MethodImplOptions.NoInlining)]
		private static void DoColorImage(ref ImageAttributes imgAtt, Color bgColor, Color fgColor)
		{
			if (imgAtt==null) imgAtt = new ImageAttributes();

			real r2 = fgColor.R/255f;
			real r1 = bgColor.R/255f;
			real g2 = fgColor.G/255f;
			real g1 = bgColor.G/255f;
			real b2 = fgColor.B/255f;
			real b1 = bgColor.B/255f;
			real a2 = fgColor.A/255f;
			real a1 = bgColor.A/255f;
			ColorMatrix rm = 
				new ColorMatrix(new real[][]
							{
								new real[]{r2-r1,    0f,    0f,    0f,    0f},
								new real[]{0f,    g2-g1,    0f,    0f,    0f},
								new real[]{0f,    0f,    b2-b1,    0f,    0f},
								new real[]{0f,    0f,       0f,    a2-a1, 0f},
								new real[]{r1,    g1,       b1,    a1,    1f}
							});
  
          
			imgAtt.SetColorMatrix(rm);
		}

        /// <summary>
        /// No need for threadstatic.
        /// </summary>
        private static bool MissingFrameworkColorMatrix;

		internal static void MakeImageGray(ref ImageAttributes imgAtt, Color shadowColor)
		{
			if (MissingFrameworkColorMatrix) return;
			try
			{
				DoMakeImageGray(ref imgAtt, shadowColor);
			}
			catch (MissingMethodException)
			{
				MissingFrameworkColorMatrix=true;
			}
		}

		internal static void ColorImage(ref ImageAttributes imgAtt, Color bgColor, Color fgColor)
		{
			if (MissingFrameworkColorMatrix) return;
			try
			{
				DoColorImage(ref imgAtt, bgColor, fgColor);
			}
			catch (MissingMethodException)
			{
				MissingFrameworkColorMatrix=true;
			}
		}

		internal static void MakeTransparent(ref ImageAttributes imgAtt, real Opacity)
		{
			if (MissingFrameworkColorMatrix) return;
			try
			{
				DoMakeTransparent(ref imgAtt, Opacity);
			}
			catch (MissingMethodException)
			{
				MissingFrameworkColorMatrix=true;
			}
		}

		internal static void AdjustImage(ref ImageAttributes imgAtt, int brightness, int contrast)
        {
            if (MissingFrameworkColorMatrix) return;
            try
            {
                DoAdjustImage(ref imgAtt, brightness, contrast);
            }
            catch (MissingMethodException)
            {
                MissingFrameworkColorMatrix=true;
            }
        }
#else
		internal static void AdjustImage(ref ImageAttributes imgAtt, int brightness, int contrast)
		{
		}
#endif

		internal static Image RasterizeWMF(Image Source)
		{
			return new Bitmap(Source);
		}

    }

    internal struct TSubscriptData
    {
        internal real Factor;
        internal real FOffset;

        internal TSubscriptData(TFlxFontStyles aStyle)
        {
            if ((aStyle & TFlxFontStyles.Subscript) != 0)
            {
                Factor = 2f/3f;
                FOffset = 8F;
            }
            else
                if ((aStyle & TFlxFontStyles.Superscript) != 0)
            {
                Factor = 2f/3f;
                FOffset = -1.5F;
            }
            else
            {
                Factor = 1;
                FOffset = 0;
            }

        }

        internal real Offset(IFlxGraphics Canvas, Font MyFont)
        {
            if (FOffset!=0)
            {
                SizeF Sz = Canvas.MeasureString("Mg", MyFont);
                return Sz.Height / FOffset;
            }
            return 0;
        }
    }

    #region Image convert
    internal static class ImgConvert
    {
        /// <summary>
        /// Converts an image to blackandwhite, but just setting the colors to white/black, not by ussing diffussion like floyd steinberg.
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        public static Image ConvertToBiLevel(Image source)
        {
            using (ImageAttributes imgAtt = new ImageAttributes())
            {
                real a = -0.5f;
                real b = 1e5f;
                real af = a * b;
                ColorMatrix cm =
                    new ColorMatrix(new real[][]
                            {
                                new real[]{b,    0f,    0f,    0f,    0f},
                                new real[]{0f,    b,    0f,    0f,    0f},
                                new real[]{0f,    0f,    b,    0f,    0f},
                                new real[]{0f,    0f,    0f,    1f,    0f},
                                new real[]{af,    af,    af,    0f,    1f}
                            });
                
                return ConvertImage(source, cm);
            }
        }

        public static Image ConvertToGrayscale(Image source)
        {
                ColorMatrix cm = new ColorMatrix(new float[][]{   
                                  new float[]{0.3f,0.3f,0.3f,0,0},
                                  new float[]{0.59f,0.59f,0.59f,0,0},
                                  new float[]{0.11f,0.11f,0.11f,0,0},
                                  new float[]{0,0,0,1,0,0},
                                  new float[]{0,0,0,0,1,0},
                                  new float[]{0,0,0,0,0,1}});


                return ConvertImage(source, cm);  
        }

        private static Image ConvertImage(Image source, ColorMatrix cm)
        {
            using (ImageAttributes imgAtt = new ImageAttributes())
            {
                imgAtt.SetColorMatrix(cm);

                Image Result = new Bitmap(source.Width, source.Height, PixelFormat.Format32bppPArgb);
                try
                {
                    using (Graphics gr = Graphics.FromImage(Result))
                    {
                        gr.DrawImage(source, new Rectangle(0, 0, source.Width, source.Height), 0, 0, source.Width, source.Height,
                                GraphicsUnit.Pixel, imgAtt);
                    }
                }
                catch
                {
                    Result.Dispose();
                    throw;
                }
                return Result;
            }
        }
    }
    #endregion

    #region ShapesCache

#if (FRAMEWORK20)
    internal class ZOrderComparer: IComparer<TShapeProperties>
    {
        #region IComparer Members

        public int Compare(TShapeProperties x, TShapeProperties y)
        {
            return x.zOrder.CompareTo(y.zOrder);
        }

        #endregion

    }
#else
    internal class ZOrderComparer: IComparer
    {
        #region IComparer Members

        public int Compare(object x, object y)
        {
            TShapeProperties sx= (TShapeProperties)x;
            TShapeProperties sy= (TShapeProperties)y;
            return sx.zOrder.CompareTo(sy.zOrder);
        }

        #endregion

    }
#endif
 
    /// <summary>
    /// Divides the shapes on pages so we can search faster.
    /// </summary>
    internal class TShapesCache
    {
        private int RowsPerPage;
        private int ColsPerPage;
        private int Top, Left;
        private TShapePropertiesList[,] Pages;

        private static readonly ZOrderComparer ZOrderComparerMethod=new ZOrderComparer(); //STATIC*

        internal TShapesCache(int aRowsPerPage, int aColsPerPage, TXlsCellRange CellRange)
        {
            RowsPerPage = Math.Max(aRowsPerPage, 1);
            ColsPerPage = Math.Max(aColsPerPage, 1);
            Pages = new TShapePropertiesList[(CellRange.RowCount - 1)/aRowsPerPage + 1, (CellRange.ColCount - 1)/ aColsPerPage + 1];
            Top = CellRange.Top;
            Left = CellRange.Left;
        }

        internal void Fill(ExcelFile Workbook)
        {
            for (int i = Workbook.ObjectCount; i >= 1; i--)
            {
                TShapeProperties ShProp = Workbook.GetObjectProperties(i, true);
                if (!ShProp.Print || !ShProp.Visible || ShProp.ObjectType == TObjectType.Comment) continue;
                TClientAnchor Anchor = ShProp.NestedAnchor;
                if (Anchor == null) continue; 

                int MinR = (Anchor.Row1 - Top) / RowsPerPage; if (MinR < 0) MinR = 0;
                int MaxR = (Anchor.Row2 - Top) / RowsPerPage; if (MaxR >= Pages.GetLength(0)) MaxR = Pages.GetLength(0) - 1;
                int MinC = (Anchor.Col1 - Left) / ColsPerPage; if (MinC < 0) MinC = 0;
                int MaxC = (Anchor.Col2 - Left) / ColsPerPage; if (MaxC >= Pages.GetLength(1)) MaxC = Pages.GetLength(1) - 1;

                

                for (int r = MinR; r <= MaxR; r++)
                {
                    for (int c = MinC; c <= MaxC; c++)
                    {
                        ShProp.zOrder = i;
                        if (Pages[r,c] == null) Pages[r,c] = new TShapePropertiesList();
                        Pages[r,c].Add(ShProp);
                    }
                }
            }
        }

        internal TShapePropertiesList GetShapes(int Row1, int Col1, int Row2, int Col2)
        {
            int MinR = (Row1 - Top) / RowsPerPage; if (MinR < 0) MinR = 0;
            int MaxR = (Row2 - Top) / RowsPerPage; if (MaxR >= Pages.GetLength(0)) MaxR = Pages.GetLength(0) - 1;
            int MinC = (Col1 - Left) / ColsPerPage; if (MinC < 0) MinC = 0;
            int MaxC = (Col2 - Left) / ColsPerPage; if (MaxC >= Pages.GetLength(1)) MaxC = Pages.GetLength(1) - 1;

            TShapePropertiesList Result = new TShapePropertiesList();
            for (int r = MinR; r <= MaxR; r++)
            {
                for (int c = MinC; c <= MaxC; c++)
                {
                    if (Pages[r,c] == null) continue;
                    for (int i = Pages[r,c].Count - 1; i >= 0; i--)
                    {
                        TClientAnchor Anchor = Pages[r,c][i].NestedAnchor;
                        if (Anchor.Row1> Row2 || Anchor.Row2 < Row1 || Anchor.Col1 > Col2 || Anchor.Col2 < Col1) continue;

                        int Index = Result.BinarySearch(Pages[r,c][i], ZOrderComparerMethod); 
                        if (Index<0)
                        {
                            Index=~Index;
                            Result.Insert(Index, Pages[r,c][i]);
                        }
                    }
                }
            }
            return Result;
        }

    }
    #endregion

    #region PageFormatCache
    internal class TPageFormatCache
    {
        private TFlxFormat[,] Formats;
        private int FirstRow;
        private int FirstCol;

        internal TPageFormatCache(int aFirstRow, int aFirstCol, int aRowsInPage, int aColsInPage)
        {
            FirstRow = aFirstRow;
            FirstCol = aFirstCol;
            Formats = new TFlxFormat[aRowsInPage, aColsInPage];
        }

        internal bool Includes(int aRow, int aCol)
        {
            return (aRow >= FirstRow && aRow < FirstRow + Formats.GetLength(0) &&
                    aCol >= FirstCol && aCol < FirstCol + Formats.GetLength(1));
        }

        internal bool IsValid(int aRow, int aCol)
        {
            return Includes(aRow, aCol) && Formats[aRow - FirstRow, aCol - FirstCol] != null;
        }

        internal void SetFormat(int aRow, int aCol, TFlxFormat Fmt)
        {
            if (!Includes(aRow, aCol)) return;
            Formats[aRow - FirstRow, aCol - FirstCol] = Fmt;
        }

        internal TFlxFormat GetFormat(int aRow, int aCol)
        {
            return Formats[aRow - FirstRow, aCol - FirstCol];
        }
        
    }

	internal class TXFFormatCache
	{
		#region Privates
		private int LastCachedRow;
		private int LastCachedCol;
		private bool LastCachedRowBiggerThan0;
		private bool LastCachedColBiggerThan0;
		private TPageFormatCache PageFormatCache;
		private Dictionary<int, TFlxFormat> FlxFormatCache;

		#endregion

		#region Constructors
		public TXFFormatCache()
		{
			FlxFormatCache = new Dictionary<int,TFlxFormat>();
			LastCachedRow = -1;
			LastCachedCol = -1;
		}

		#endregion

		private void UpdateRowBiggerThan0(ExcelFile Workbook, int row, int col)
		{
			if (row != LastCachedRow)
			{
				LastCachedRowBiggerThan0 = Workbook.GetRowHeight(row, true)>0;
				LastCachedRow = row;
			}
			if (col != LastCachedCol)
			{
				LastCachedColBiggerThan0 = Workbook.GetColWidth(col, true)>0;
				LastCachedCol = col;
			}
		}

		public TFlxFormat GetCellVisibleFormatDef(ExcelFile Workbook, int row, int col, bool Merged)
		{
			int XF = Workbook.DefaultFormatId;
			UpdateRowBiggerThan0(Workbook, row, col);
            
			bool CacheValid = Merged || (LastCachedRowBiggerThan0 && LastCachedColBiggerThan0);

			if (CacheValid)
			{
				if (PageFormatCache != null && PageFormatCache.IsValid(row, col))
				{
					return PageFormatCache.GetFormat(row, col);
				}

				XF = Workbook.GetCellVisibleFormat(row, col);
			}

			TFlxFormat Result;
            if (!FlxFormatCache.TryGetValue(XF, out Result))
			{
				Result = Workbook.GetFormat(XF);
				FlxFormatCache[XF] = Result;
			}

			TFlxFormat Result2 = Workbook.ConditionallyModifyFormat(Result, row, col);
			if (Result2 != null) 
			{
				if (PageFormatCache != null && CacheValid) PageFormatCache.SetFormat(row, col, Result2);
				return Result2;
			}

			if (PageFormatCache != null && CacheValid) PageFormatCache.SetFormat(row, col, Result);
			return Result;
		}

		public void Clear()
		{
			FlxFormatCache.Clear();
			LastCachedRow = -1;
			LastCachedCol = -1;
		}

		public void CreatePageCache(TXlsCellRange PagePrintRange)
		{
			PageFormatCache = new TPageFormatCache(PagePrintRange.Top, PagePrintRange.Left, PagePrintRange.RowCount, PagePrintRange.ColCount);
		}

		public void DestroyPageCache()
		{
			PageFormatCache = null;
		}

	}

    #endregion

	#region TUserRangeList
	internal class TUsedRangeList
	{
		private List<TXlsCellRange> Ranges;

		internal TUsedRangeList()
		{
			Ranges = new List<TXlsCellRange>();
		}

		internal bool Find(int Row, int Col, out int Index)
		{
			//There should never be too much ranges here, so a linear search should be enough
			for (int i = Ranges.Count - 1; i >= 0; i--)
			{
				TXlsCellRange c = Ranges[i];
				if (c.HasRow(Row) && c.HasCol(Col))
				{
					Index = i;
                    return true;
				}
			}
			Index = -1;
			return false;
		}

		internal bool Find(int Row, int Col)
		{
			int Index;
			return Find(Row, Col, out Index);
		}

		internal void Add(int Row1, int Col1, int Row2, int Col2)
		{
			Ranges.Add(new TXlsCellRange(Row1, Col1, Row2 - 1, Col2 - 1));
		}

		internal void CleanUpUsed(int Row)
		{
			for (int i = Ranges.Count - 1; i >= 0; i--)
			{
				TXlsCellRange c = Ranges[i];
				if (c.Bottom < Row)
				{
					Ranges.RemoveAt(i);
				}
			}
		}

		internal TXlsCellRange this[int Index]
		{
			get
			{
				return Ranges[Index];
			}
		}
	}
	#endregion

	#region TGraphicCanvas
	internal class TGraphicCanvas: IDisposable
	{
		#region Properties
		public readonly TFontCache FontCache;
		public readonly Bitmap bmp;
		public readonly Graphics imgData;
		public readonly IFlxGraphics Canvas;
		#endregion
		
		#region Constructor
		internal TGraphicCanvas()
		{
			FontCache = new TFontCache();
			bmp = BitmapConstructor.CreateBitmap(1,1);
			imgData = Graphics.FromImage(bmp);
			imgData.PageUnit=GraphicsUnit.Point;
			Canvas= new GdiPlusGraphics(imgData);
			Canvas.CreateSFormat();
		}

		#endregion
		#region IDisposable Members

		public void Dispose()
		{
			Canvas.DestroySFormat();
			imgData.Dispose();
			bmp.Dispose();
			FontCache.Dispose();
            GC.SuppressFinalize(this);
        }

		#endregion

	}

	#endregion
}

