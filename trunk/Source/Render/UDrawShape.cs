using System;
using System.IO;

using FlexCel.Core;
using System.Collections.Generic;

#if (MONOTOUCH)
using System.Drawing;
using real = System.Single;
using Color = MonoTouch.UIKit.UIColor;
#else
	#if (WPF)
	using RectangleF = System.Windows.Rect;
	using PointF = System.Windows.Point;
	using real = System.Double;
	using ColorBlend = System.Windows.Media.GradientStopCollection;
	using System.Windows.Media;
	using System.Windows;
	#else
	using real = System.Single;
	using Colors = System.Drawing.Color;
	using DashStyles = System.Drawing.Drawing2D.DashStyle;
	using System.Drawing;
	using System.Drawing.Drawing2D;
	using System.Drawing.Imaging;
	#endif
#endif

namespace FlexCel.Render
{
    internal enum TQuarter 
    {
        None,
        Top,
        Right,
        Bottom, 
        Left
    }

    /// <summary>
    /// Draws different types of autoshapes.
    /// </summary>
	internal static class DrawShape
	{
		#region Images
		internal static void DrawOneImage(IFlxGraphics Canvas, TCropArea c, 
			long TransparentColor, int Brightness, int Contrast, int Gamma, Color ShadowColor,
            bool BiLevel, bool Grayscale,
			RectangleF Coords, MemoryStream ImgData)
		{
			ImgData.Position = 0;

			using (Image Img = Image.FromStream(ImgData))
			{
				try
				{
					bool DoDraw = true;
					RectangleF srcRect = new RectangleF(0, 0, Img.Width, Img.Height);
					if (c.CropFromLeft != 0 || c.CropFromRight != 0 || c.CropFromTop != 0 || c.CropFromBottom != 0)
					{
						real h1 = Img.Height; real h = h1 / 65536F;
						real w1 = Img.Width; real w = w1 / 65536F;

						real srcRectLeft = c.CropFromLeft * w;
						real srcRectTop = c.CropFromTop * h;
						real srcRectRight = w1 - c.CropFromRight * w;
						real srcRectBottom = h1 - c.CropFromBottom * h;

						if (srcRectLeft > w1) DoDraw = false;
						if (srcRectTop > h1) DoDraw = false;
						if (srcRectRight < srcRectLeft) DoDraw = false;
						if (srcRectBottom < srcRectTop) DoDraw = false;

						srcRect = new RectangleF(srcRectLeft, srcRectTop, srcRectRight - srcRectLeft, srcRectBottom - srcRectTop);
					}

                    if (DoDraw && srcRect.Width > 0 && srcRect.Height > 0)
                    {
                        Image Img2 = Img;
                        try
                        {
                            if (Grayscale)
                            {
                                Img2 = ImgConvert.ConvertToGrayscale(Img);
                                ImgData = null;
                            }
                            if (BiLevel)
                            {
                                ImgData = null;
                                Image Img3 = Img2;
                                try
                                {
                                    Img3 = ImgConvert.ConvertToBiLevel(Img2);
                                }
                                finally
                                {
                                    if (Img2 != Img) Img2.Dispose();
                                    Img2 = Img3;
                                }
                            }
                            Canvas.DrawImage(Img2, Coords, srcRect, TransparentColor, Brightness, Contrast, Gamma, ShadowColor, ImgData);
                        }
                        finally
                        {
                            if (Img2 != Img)
                            {
                                Img2.Dispose();
                                Img2 = null;
                            }
                        }
                    }
				}
				catch (System.Runtime.InteropServices.ExternalException ex)
				{
					//Clean the Exception. This is not serious (image has not been drawn) and
					//the framework raises this for small images. (seems like a bug on the framework)
					//The OLE HRESULT of the exception is 0x80004005
					if (FlexCelTrace.HasListeners) FlexCelTrace.Write(new TRenderCorruptImageError(ex.Message));

				}
				catch (FlexCel.Pdf.FlexCelPdfException ex)
				{
					//Clean the Exception. This is not serious (image has not been drawn) but the report still can continue.
					if (FlexCelTrace.HasListeners) FlexCelTrace.Write(new TRenderCorruptImageError(ex.Message));
				}
			}
		}

		private static TCropArea GetCropArea(TShapeOptionList ShOptions)
		{
			TCropArea Result = new TCropArea();
			unchecked
			{
				object o = ShOptions[TShapeOption.cropFromRight]; 
				if (o!=null) Result.CropFromRight = (int)(long)(o);
				o = ShOptions[TShapeOption.cropFromTop]; 
				if (o!=null) Result.CropFromTop = (int)(long)(o);
				o = ShOptions[TShapeOption.cropFromLeft]; 
				if (o!=null) Result.CropFromLeft = (int)(long)(o);
				o = ShOptions[TShapeOption.cropFromBottom]; 
				if (o!=null) Result.CropFromBottom = (int)(long)(o);
			}

			return Result;
		}

		internal static long GetLong(object o, long Default)
		{
			if (o!=null)
				return (long)o;
			return Default;
		}

		internal static int GetInt(object o, int Default)
		{
			unchecked
			{
				if (o!=null)
					return (int)(long)(o);
				return Default;
			}
		}

		internal static uint GetUInt(object o, uint Default)
		{
			unchecked
			{
				if (o!=null)
					return (uint)(long)(o);
				return Default;
			}
		}

		internal static void DrawImage(IFlxGraphics Canvas, ExcelFile FWorkbook, TShapeProperties ShProp, int ObjPos, string ObjectPath, RectangleF Coords, TShadowInfo ShadowInfo, real Zoom100)
		{
			TXlsImgType imageType = TXlsImgType.Unknown;
			using (MemoryStream ImgData = new MemoryStream())
			{
				try
				{
					FWorkbook.GetImage(ObjPos, ObjectPath, ref imageType, ImgData, true);

					//We are going to do this where it matters, this is at the GDIPlusCanvas/PDFCanvas level. So we do not need to rasterize valid metafiles, only those cropped or with changed gamma or transparency
					/*if (imageType == TXlsImgType.Wmf || imageType == TXlsImgType.Emf)  //Rasterize the image. Wmfs give too much issues when dealing with transparency and/or resizing.
					{
						ImgData.Position = 0;
						using (Image Img = Image.FromStream(ImgData))
						{
							ImgData.SetLength(0);				
							using (Bitmap bmp = new Bitmap(Img, Img.Width/WmfScale, Img.Height/WmfScale)) //resolution of wmfs is too high, ot takes too long to convert.
							{
								bmp.Save(ImgData, ImageFormat.Png);  
							}
							imageType=TXlsImgType.Png;
						}
					}*/
					ImgData.Position = 0;

					using (Pen PicturePen = GetPen(ShProp, FWorkbook, ShadowInfo, false))
					{
						using (Brush PictureBrush = GetBrush(Coords, ShProp, FWorkbook, ShadowInfo, true, Zoom100))
						{						
							if (ShadowInfo.Style == TShadowStyle.None)
							{
								if (PictureBrush != null) Canvas.FillRectangle(PictureBrush, Coords);

								if (imageType != TXlsImgType.Unknown)
								{
									DrawOneImage(Canvas, GetCropArea(ShProp.ShapeOptions), 
										GetLong(ShProp.ShapeOptions[TShapeOption.pictureTransparent], FlxConsts.NoTransparentColor), 
										GetInt(ShProp.ShapeOptions[TShapeOption.pictureBrightness], FlxConsts.DefaultBrightness), 
										GetInt(ShProp.ShapeOptions[TShapeOption.pictureContrast], FlxConsts.DefaultContrast), 						
										GetInt(ShProp.ShapeOptions[TShapeOption.pictureGamma], FlxConsts.DefaultGamma), 
										ColorUtil.Empty,
                                        ShProp.ShapeOptions.AsBool(TShapeOption.pictureActive, false, 1),
                                        ShProp.ShapeOptions.AsBool(TShapeOption.pictureActive, false, 2),
										Coords, ImgData);
								}
								if (PicturePen != null) Canvas.DrawRectangle(PicturePen, Coords.Left, Coords.Top, Coords.Width, Coords.Height);
							}
							else			
							{
								if (PicturePen != null) 
								{
									Canvas.FillRectangle(PictureBrush, Coords);
								}
								else
								{
									if (imageType != TXlsImgType.Unknown)
									{
										Color ShadowColor = Colors.Gray;
										SolidBrush sb = PictureBrush as SolidBrush;
										if (sb != null)
										{
											ShadowColor = sb.Color;
										}
										DrawOneImage(Canvas, GetCropArea(ShProp.ShapeOptions), 
											GetLong(ShProp.ShapeOptions[TShapeOption.pictureTransparent], FlxConsts.NoTransparentColor), 
											GetInt(ShProp.ShapeOptions[TShapeOption.pictureBrightness], FlxConsts.DefaultBrightness), 
											GetInt(ShProp.ShapeOptions[TShapeOption.pictureContrast], FlxConsts.DefaultContrast), 						
											GetInt(ShProp.ShapeOptions[TShapeOption.pictureGamma], FlxConsts.DefaultGamma), 						
											ShadowColor,
                                            ShProp.ShapeOptions.AsBool(TShapeOption.pictureActive, false, 1),
                                            ShProp.ShapeOptions.AsBool(TShapeOption.pictureActive, false, 2),
                                            Coords, ImgData);
									}
								}
							}

						}
					}
				}
				catch (FileNotFoundException ex)
				{
					if (FlexCelTrace.HasListeners) FlexCelTrace.Write(new TRenderErrorDrawingImageError(ex.Message));

					//This will be raised if vj# is not installed. it is not a serious thing, we will just not draw the picture.
				}
			}
		}
		#endregion

		#region Common Props
		private static Color ColorFromLong(TShapeProperties ShProp, uint cl, real Opacity, ExcelFile Workbook, bool bkg, out bool IsTransparent, int RecursionLevel)
		{
			Color Result = InternalColorFromLong(ShProp, cl, Workbook, bkg, out IsTransparent, RecursionLevel);
			if (Opacity < 1)
			{
				int A = (int)Math.Round(Opacity * 255f);
				return ColorUtil.FromArgb(A, Result.R, Result.G, Result.B);
			}
			return Result;
		}

		private static Color InternalColorFromLong(TShapeProperties ShProp, uint cl, ExcelFile Workbook, bool bkg, out bool IsTransparent, int RecursionLevel)
		{
			IsTransparent = false;
			int ColorFlags = ((int)((cl & 0xFF000000) >>24));

			if ((ColorFlags & 1) != 0)
				return ColorUtil.FromArgb((int)(cl & 0xFF), (int)((cl & 0xFF00) >>8), (int)((cl & 0xFF0000) >>16));

			if ((ColorFlags & 8) != 0) //Externally Indexed
			{
				int cp = (int)(cl & 0xFFFF) - 7;
				if (cp > 0 && cp <= Workbook.ColorPaletteCount)
				{
					return Workbook.GetColorPalette(cp);
				}
				else if (cp == Workbook.ColorPaletteCount + 1 && bkg) //no fill
				{
					IsTransparent = true;
					return ColorUtil.FromArgb(255, 0, 0, 0);
				}
				else  //auto
					if (bkg)
					return Colors.White; else
					return Colors.Black;
			}

			if ((ColorFlags & 16) != 0) // SysIndex Color
			{
				return GetSysIndexColor(ShProp, Workbook, cl, RecursionLevel);
			}
            
			return ColorUtil.FromArgb((int)(cl & 0xFF), (int)((cl & 0xFF00) >>8), (int)((cl & 0xFF0000) >>16));            
		}

		private static Color GetSystemColor(long cl)
		{
			switch (cl)
			{
#if(WPF)
                case 0: return SystemColors.ControlColor; //COLOR_BTNFACE
                case 1: return SystemColors.WindowTextColor;  // COLOR_WINDOWTEXT
                case 2: return SystemColors.MenuColor; // COLOR_MENU
                case 3: return SystemColors.HighlightColor;           // COLOR_HIGHLIGHT
                case 4: return SystemColors.HighlightTextColor;       // COLOR_HIGHLIGHTTEXT
                case 5: return SystemColors.ActiveCaptionTextColor;        // COLOR_CAPTIONTEXT
                case 6: return SystemColors.ActiveCaptionColor;       // COLOR_ACTIVECAPTION
                case 7: return SystemColors.ControlLightColor;     // COLOR_BTNHIGHLIGHT
                case 8: return SystemColors.ControlDarkColor;        // COLOR_BTNSHADOW
                case 9: return SystemColors.ControlTextColor;          // COLOR_BTNTEXT
                case 10: return SystemColors.GrayTextColor;            // COLOR_GRAYTEXT
                case 11: return SystemColors.InactiveCaptionColor;     // COLOR_INACTIVECAPTION
                case 12: return SystemColors.InactiveCaptionTextColor; // COLOR_INACTIVECAPTIONTEXT
                case 13: return SystemColors.InfoColor;      // COLOR_INFOBK
                case 14: return SystemColors.InfoTextColor;            // COLOR_INFOTEXT
                case 15: return SystemColors.MenuTextColor;            // COLOR_MENUTEXT
                case 16: return SystemColors.ScrollBarColor;           // COLOR_SCROLLBAR
                case 17: return SystemColors.WindowColor;              // COLOR_WINDOW
                case 18: return SystemColors.WindowFrameColor;         // COLOR_WINDOWFRAME
                case 19: return SystemColors.ControlLightLightColor;             // COLOR_3DLIGHT
#else
				case 0: return SystemColors.Control; //COLOR_BTNFACE
				case 1: return SystemColors.WindowText;  // COLOR_WINDOWTEXT
				case 2: return SystemColors.Menu; // COLOR_MENU
				case 3: return SystemColors.Highlight;           // COLOR_HIGHLIGHT
				case 4: return SystemColors.HighlightText;       // COLOR_HIGHLIGHTTEXT
				case 5: return SystemColors.ActiveCaptionText;        // COLOR_CAPTIONTEXT
				case 6: return SystemColors.ActiveCaption;       // COLOR_ACTIVECAPTION
				case 7: return SystemColors.ControlLight;     // COLOR_BTNHIGHLIGHT
				case 8: return SystemColors.ControlDark;        // COLOR_BTNSHADOW
				case 9: return SystemColors.ControlText;          // COLOR_BTNTEXT
				case 10: return SystemColors.GrayText;            // COLOR_GRAYTEXT
				case 11: return SystemColors.InactiveCaption;     // COLOR_INACTIVECAPTION
				case 12: return SystemColors.InactiveCaptionText; // COLOR_INACTIVECAPTIONTEXT
				case 13: return SystemColors.Info;      // COLOR_INFOBK
				case 14: return SystemColors.InfoText;            // COLOR_INFOTEXT
				case 15: return SystemColors.MenuText;            // COLOR_MENUTEXT
				case 16: return SystemColors.ScrollBar;           // COLOR_SCROLLBAR
				case 17: return SystemColors.Window;              // COLOR_WINDOW
				case 18: return SystemColors.WindowFrame;         // COLOR_WINDOWFRAME
				case 19: return SystemColors.ControlLightLight;             // COLOR_3DLIGHT
#endif
				case 20: return ColorUtil.Empty;                 // Count of system colors
			}

			return ColorUtil.Empty;
		}

		private static int ChangeComponent(int R, int p)
		{
			int Result = R + p; 
			if (Result < 0) Result = 0; if (Result > 255) Result = 255;
			return Result;
		}
		private static Color ChangeColor(Color Clr, int p)
		{
			return ColorUtil.FromArgb(Clr.A, ChangeComponent(Clr.R, p), ChangeComponent(Clr.G, p), ChangeComponent(Clr.B, p));
		}

		private static Color ChangeColor(Color Clr, int p1, int p2, int p3)
		{
			return ColorUtil.FromArgb(Clr.A, ChangeComponent(Clr.R, p1), ChangeComponent(Clr.G, p2), ChangeComponent(Clr.B, p3));
		}

		private static int Light(int C, double p)
		{
			int Result = (int)(C * p);
			if (Result < 0) return 0;
			if (Result > 255) return 255;
			return Result;
		}
		private static Color DarkColor(Color Clr, double p)
		{
			return ColorUtil.FromArgb(Clr.A, Light(Clr.R, p), Light(Clr.G, p), Light(Clr.B, p));
		}

		private static Color LightColor(Color Clr, double p)
		{
			return ColorUtil.FromArgb(Clr.A, 255-Light(255 - Clr.R, p), 255-Light(255-Clr.G, p), 255-Light(255-Clr.B, p));
		}

		private static Color GetSysIndexColor(TShapeProperties ShProp, ExcelFile Workbook, long cl, int RecursionLevel)
		{
			if (RecursionLevel > 20) return Colors.Empty;
			if (cl <= 20) return GetSystemColor(cl);

			Color Result = Colors.Empty;

			if ((cl &0x8000) != 0) Result = Colors.Gray;   // Make the color gray (before the above!)
            
			int p = (int)(cl & 0xFF0000) >> 16; // Parameter used as above

			bool IsTransparent;
			switch (cl & 0xFF)
			{
				case 0xF0:    // Use the fillColor property
					Result = ColorFromLong(ShProp, (UInt32)ShProp.ShapeOptions.AsLong(TShapeOption.fillColor, 0x0),
						ShProp.ShapeOptions.As1616(TShapeOption.fillOpacity, 1), Workbook, true, out IsTransparent, RecursionLevel + 1);
					break;

				case 0xF1: // Use the line color only if there is a line
					if (!ShProp.ShapeOptions.AsBool(TShapeOption.fNoLineDrawDash, true, 3))
					{
						Result = ColorFromLong(ShProp, (UInt32)ShProp.ShapeOptions.AsLong(TShapeOption.fillColor, 0x0),
							ShProp.ShapeOptions.As1616(TShapeOption.fillOpacity, 1), Workbook, true, out IsTransparent, RecursionLevel + 1);
					}
					else
					{
						Result = ColorFromLong(ShProp, (UInt32)ShProp.ShapeOptions.AsLong(TShapeOption.lineColor, 0x808080),
							ShProp.ShapeOptions.As1616(TShapeOption.lineOpacity, 1), Workbook, true, out IsTransparent, RecursionLevel + 1);
					}
					break;

				case 0xF2: // Use the lineColor property
					Result = ColorFromLong(ShProp, (UInt32)ShProp.ShapeOptions.AsLong(TShapeOption.lineColor, 0x0),
						ShProp.ShapeOptions.As1616(TShapeOption.lineOpacity, 1), Workbook, true, out IsTransparent, RecursionLevel + 1);
					break;

				case 0xF3:    // Use the shadow color
					Result = ColorFromLong(ShProp, (UInt32)ShProp.ShapeOptions.AsLong(TShapeOption.shadowColor, 0x0),
						ShProp.ShapeOptions.As1616(TShapeOption.shadowOpacity, 1), Workbook, true, out IsTransparent, RecursionLevel + 1);
					break;

				case 0xF4:   // Use this color (only valid as described below)
					break;
				case 0xF5:   // Use the fillBackColor property
					Result = ColorFromLong(ShProp, (UInt32)ShProp.ShapeOptions.AsLong(TShapeOption.fillBackColor, 0x0),
						ShProp.ShapeOptions.As1616(TShapeOption.fillBackOpacity, 1), Workbook, true, out IsTransparent, RecursionLevel + 1);
					break;

				case 0xF6:   // Use the lineBackColor property
					Result = ColorFromLong(ShProp, (UInt32)ShProp.ShapeOptions.AsLong(TShapeOption.lineBackColor, 0x0),
						1, Workbook, true, out IsTransparent, RecursionLevel + 1);
					break;
				case 0xF7:    // Use the fillColor unless no fill and line
					Result = ColorFromLong(ShProp, (UInt32)ShProp.ShapeOptions.AsLong(TShapeOption.fillColor, 0x0),
						ShProp.ShapeOptions.As1616(TShapeOption.fillOpacity, 1), Workbook, true, out IsTransparent, RecursionLevel + 1);
					break;

				case 0xFF: Result = ColorUtil.FromArgb((int)((cl & 0xFF00) >> 8), (int)((cl & 0xFF0000) >>16), (int)((cl & 0xFF000000) >>24)); break;  // Extract the color index
			}

			switch (cl & 0x0F00)   // function to apply
			{
				case 0x0100: Result = DarkColor(Result, p / 255.0); break;   // Darken color by parameter/255
				case 0x0200: Result = LightColor(Result, p / 255.0); break;   // Lighten color by parameter/255
				case 0x0300: Result = ChangeColor(Result, p); break;        // Add grey level RGB(param,param,param)
				case 0x0400: Result = ChangeColor(Result, -p); break;      // Subtract grey level RGB(p,p,p)
				case 0x0500: Result = ChangeColor(ColorUtil.FromArgb(p,p,p), Result.R, Result.G, Result.B); break;   // Subtract from grey level RGB(p,p,p)
					// In the following "black" means maximum component value, white minimum.
					//   The operation is per component, to guarantee white combine with msocolorGray
				case 0x0600: 
					int R =Result.R < p? 0: 255;   // Black if < uParam, else white (>=)
					int G =Result.G < p? 0: 255;   // Black if < uParam, else white (>=)
					int B =Result.B < p? 0: 255;   // Black if < uParam, else white (>=)
					Result = ColorUtil.FromArgb(Result.A, R, G, B); 
					break;
			}

			if ((cl & 0x4000) != 0) Result = ColorUtil.FromArgb(Result.R ^ 128, Result.G ^ 128, Result.B ^ 128);   // Invert by toggling the top bit
			if ((cl & 0x2000) != 0) Result = ColorUtil.FromArgb(255 - Result.R, 255 - Result.G, 255 - Result.B);   // Invert color (at the *end*)

			return Result;
		}

		private static bool DrawLines(TShadowInfo ShadowInfo, bool HasBrush)
		{
			return ShadowInfo.Style == TShadowStyle.None || !HasBrush;
		}

		/// <summary>
		/// Sadly, we cannot change the focus on ColorBlended gradients with SetSigmaShape.
		/// </summary>
		/// <param name="Orig"></param>
		/// <param name="FillFocus"></param>
		/// <returns></returns>
		private static ColorBlend ChangeFocus(ColorBlend Orig, real FillFocus)
		{
			int k = FlxGradient.BlendCount(Orig);
			ColorBlend Result = new ColorBlend(k * 2 - 1);
#if(WPF)
            for (int i = 0; i < k * 2 - 1; i++)
            {
                Result.Add(new GradientStop());
            }
#endif

			for (int i = 0; i < k; i ++)
			{
                FlxGradient.SetColorBlend(Result, i, FlxGradient.BlendColor(Orig, k - 1 - i), (1 - FlxGradient.BlendPosition(Orig, k - i - 1)) * FillFocus);
                FlxGradient.SetColorBlend(Result, k * 2 - 2 - i, FlxGradient.BlendColor(Orig, k - 1 - i), FillFocus + FlxGradient.BlendPosition(Orig, k - i - 1) * (1 - FillFocus));
			}

#if(WPF)
            if (Result.CanFreeze) Result.Freeze();
#endif
            return Result;

		}

		private static ColorBlend GetBlending(TShapeProperties ShProp, ExcelFile Workbook, byte[] Values, real OpacityOrg, real OpacityDest, real FillFocus, bool InvertColors, Color Color1, Color Color2)
		{
			int BlendCount = (Values.Length - 6) / 8;
			ColorBlend Result = new ColorBlend(BlendCount);
#if(WPF)
            for (int i = 0; i < BlendCount; i++)
            {
                Result.Add(new GradientStop());
            }
#endif

			int k = 6;
			for (int i = 0; i < BlendCount; i++)
			{
				bool IsTransparent;
				Color BlendCol = ColorFromLong(ShProp, BitConverter.ToUInt32(Values, k), OpacityDest + (OpacityOrg - OpacityDest) * i /(BlendCount - 1f) , Workbook, true, out IsTransparent,0);
				real BlendPos = 1 - BitConverter.ToUInt16(Values, k + 5)/256f;
                FlxGradient.SetColorBlend(Result, BlendCount - 1 - i, BlendCol, BlendPos);
				k+=8;
			}

            FlxGradient.EnsureMinimumAndMaximum(Color1, Color2, ref Result);

            if (InvertColors) //must be done before changefocus.
            {
                FlxGradient.InvertColorBlend(Result);
            }

			if (FillFocus > 0)
			{
				Result = ChangeFocus(Result, FillFocus);
			}

#if(WPF)
            if (Result.CanFreeze) Result.Freeze();
#endif

			return Result;
		}

		internal static Brush GetGradientBrush(RectangleF Coords, TShapeProperties ShProp, ExcelFile Workbook, TFillType FillType)
		{
			RectangleF Coords2 = Coords;
			if (ShProp.ShapeOptions.AsBool(TShapeOption.fNoFillHitTest, false, 1))
			{
				Coords2 = FlexCelRender.RectangleXY(
					GetInt(ShProp.ShapeOptions[TShapeOption.fillRectLeft], 0),
					GetInt(ShProp.ShapeOptions[TShapeOption.fillRectTop], 0),
					GetInt(ShProp.ShapeOptions[TShapeOption.fillRectRight], 0),
					GetInt(ShProp.ShapeOptions[TShapeOption.fillRectBottom], 0));
			}

			if (Coords2.Width <= 0 || Coords2.Height <= 0) return null;

			bool IsTransparent;

			real FillOpacity = ShProp.ShapeOptions.As1616(TShapeOption.fillOpacity, 1);
			Color Color1 = ColorFromLong(ShProp, GetUInt(ShProp.ShapeOptions[TShapeOption.fillColor], 0xffffffff), FillOpacity, Workbook, true, out IsTransparent, 0);

			real FillBackOpacity = ShProp.ShapeOptions.As1616(TShapeOption.fillBackOpacity, 1);
			Color Color2 = ColorFromLong(ShProp, GetUInt(ShProp.ShapeOptions[TShapeOption.fillBackColor], 0xffffffff), FillBackOpacity, Workbook, true, out IsTransparent, 0);


			//Y coords here go the other way around. so there are alots of "-" symbols.

			real FillAngle = ShProp.ShapeOptions.As1616(TShapeOption.fillAngle, 0);
            while (FillAngle <= -360) FillAngle += 360;
            while (FillAngle >= 360) FillAngle -= 360; //Excel 2007 behaves different here. An angle of 360 = 0. In 2003 it is -360

			//FillFocus does not work with multiple colors!
			real FillFocus = GetInt(ShProp.ShapeOptions[TShapeOption.fillFocus], 0);
            
			bool InvertColors = false;
            if (FillFocus < -100) FillFocus = -100; //Focus on GDI+ must be between 0 and 1
            if (FillFocus > 100) FillFocus = 100;

            if (FillFocus < 0)
            {
                FillFocus += 100;
                InvertColors = true;

            }

            FillFocus = 1 - FillFocus / 100f;

            if (FillAngle < 0)
            {
                FillAngle += 180;
                if (FillFocus > 0 && FillFocus < 1) InvertColors = !InvertColors; //there is a border case here
            }

			if (InvertColors)
			{
                SwapColors(ref FillOpacity, ref Color1, ref FillBackOpacity, ref Color2);
			}

            
			byte[] Blending = ShProp.ShapeOptions[TShapeOption.fillShadeColors] as byte[];

			bool RotateWithShape = true; //This value is located on an undocumented record (ID f122), It is not available on xls2000
            
			switch (FillType)
			{
				case TFillType.ShadeShape:
				{
					PathGradientBrush Result = new PathGradientBrush(new PointF[] {
																					  Coords2.Location,
																					  new PointF(Coords2.Right, Coords2.Top),
																					  new PointF(Coords2.Right, Coords2.Bottom),
																					  new PointF(Coords2.Left, Coords2.Bottom)
																				  });
					Result.CenterColor = Color1;
					Result.SurroundColors = new Color[]{Color2};

					if (Blending != null && Blending.Length > 0) 
					{
						Result.InterpolationColors = GetBlending(ShProp, Workbook, Blending, FillOpacity, FillBackOpacity, FillFocus, InvertColors, Color2, Color1);
					}
					else
					{
						Result.SetSigmaBellShape(1 - FillFocus);
					}
					Result.WrapMode = WrapMode.TileFlipXY;

					return Result;
				}

				case TFillType.ShadeCenter:
				{
					PathGradientBrush Result = new PathGradientBrush(new PointF[] {
																					  Coords2.Location,
																					  new PointF(Coords2.Right, Coords2.Top),
																					  new PointF(Coords2.Right, Coords2.Bottom),
																					  new PointF(Coords2.Left, Coords2.Bottom)
																				  });
					Result.CenterColor = Color1;
					Result.SurroundColors = new Color[]{Color2};
					real ToLeft = ShProp.ShapeOptions.As1616(TShapeOption.fillToLeft, 0);
					real ToTop = ShProp.ShapeOptions.As1616(TShapeOption.fillToTop, 0);
					Result.CenterPoint = new PointF(Coords.Left + (int)ToLeft * Coords.Width, Coords.Top + (int) ToTop * Coords.Height);
					if (Blending != null && Blending.Length > 0) 
					{
						Result.InterpolationColors = GetBlending(ShProp, Workbook, Blending, FillOpacity, FillBackOpacity, FillFocus, InvertColors, Color2, Color1);
					}
					else
					{
						Result.SetSigmaBellShape(1 - FillFocus);
					}
					Result.WrapMode = WrapMode.TileFlipXY;

					return Result;
				}
				default:
				{
					LinearGradientBrush Result = new LinearGradientBrush(Coords2, Color1, Color2, -FillAngle - 90, RotateWithShape);
					if (Blending != null && Blending.Length > 0) 
					{
						Result.InterpolationColors = GetBlending(ShProp, Workbook, Blending, FillOpacity, FillBackOpacity, FillFocus, InvertColors, Color2, Color1);
					}
					else
					{
						Result.SetSigmaBellShape(FillFocus);
					}
					Result.WrapMode = WrapMode.TileFlipXY;
					return Result;
				}
			}

		}

        private static void SwapColors(ref real FillOpacity, ref Color Color1, ref real FillBackOpacity, ref Color Color2)
        {
            Color Tmp = Color1;
            Color1 = Color2;
            Color2 = Tmp;

            real t = FillBackOpacity;
            FillBackOpacity = FillOpacity;
            FillOpacity = t;
        }


		private static bool GetEmbeddedBlip(byte[] Data, Stream Ms)
		{
			TXlsImgType ImageType = TXlsImgType.Unknown;
			if (Data.Length < 4) return false;
			int Instance = BitConverter.ToUInt16(Data, 0) >> 4;
			switch (Instance)
			{
				case 0x216 : ImageType = TXlsImgType.Wmf; break;
				case 0x3D4 : ImageType = TXlsImgType.Emf; break;
					//case 0x542 :	ImageType = TXlsImgType.Pict; break;

				case 0x46A : ImageType = TXlsImgType.Jpeg; break;
				case 0x6E0 : ImageType = TXlsImgType.Png; break;
				case 0x7A8 : ImageType = TXlsImgType.Bmp; break;
			}

			if (ImageType != TXlsImgType.Unknown)
			{
				FlexCel.XlsAdapter.TEscherBSERecord.SaveGraphicToStream(Data, 0, Ms, ImageType);  //A dependency with XlsAdapter. This is not clean, but seems like the best solution. The other solution implies moving a lot of code to core.
				return true;
			}
			return false;

		}

		internal static Image GetTextureImage(TShapeProperties ShProp, ExcelFile Workbook, TFillType FillType, out ImageAttributes Attr, out real w, out real h, real Zoom100)
		{
			w = 0;
			h = 0;
			TXlsImgType ImageType = TXlsImgType.Unknown;
			Image Img = null;
			Attr = null;

			using (MemoryStream Ms = new MemoryStream())
			{
				object Bl = ShProp.ShapeOptions[TShapeOption.fillBlip];
				Byte[] BlBytes = Bl as Byte[];
				if (BlBytes != null)
				{
					if (!GetEmbeddedBlip(BlBytes, Ms)) return null;
				}
				else
				{
					int Blip = GetInt(Bl, 0);
                    if (Blip <= 0) return null;
					Workbook.GetImage(Blip, ShProp.ObjectPath, ref ImageType, Ms);
				}
                if (Ms.Length == 0) return null;
				Ms.Position = 0;
				Img = Image.FromStream(Ms);
			}
			
			w = ShProp.ShapeOptions.AsLong(TShapeOption.fillWidth, 0)/ 12700f * Zoom100;
			h = ShProp.ShapeOptions.AsLong(TShapeOption.fillHeight, 0)/ 12700f * Zoom100;

			if (w == 0 || h == 0)
			{
				w = Img.Width;
				h = Img.Height;
			}

			if (FillType == TFillType.Picture)
			{
				//w = Coords.Width;
				//h = Coords.Height;
			}

			real FillOpacity = ShProp.ShapeOptions.As1616(TShapeOption.fillOpacity, 1);

			if (FillType == TFillType.Pattern)
			{
				bool IsTransparent;
				Color Bg = ColorFromLong(ShProp, GetUInt(ShProp.ShapeOptions[TShapeOption.fillBackColor], 0xffffffff), FillOpacity, Workbook, true, out IsTransparent, 0);
				Color Fg = ColorFromLong(ShProp, GetUInt(ShProp.ShapeOptions[TShapeOption.fillColor], 0xffffffff), FillOpacity, Workbook, true, out IsTransparent, 0);
				FlgConsts.ColorImage(ref Attr, Bg, Fg);
			}
			else if (FillOpacity < 1)
			{
				FlgConsts.MakeTransparent(ref Attr, FillOpacity);
			}

			return Img;

		}					


		internal static Brush GetTextureBrush(TShapeProperties ShProp, ExcelFile Workbook, TFillType FillType, RectangleF Coords, real Zoom100)
		{
			Image Img = null;
			ImageAttributes Attr = null;

			try
			{
				try
				{
					real w;
					real h;
					Img = GetTextureImage(ShProp, Workbook, FillType, out Attr, out w, out h, Zoom100);
                    if (Img == null) return null;

					TextureBrush Result;
					using (Image img1 = new Bitmap(Img)) //Needed to avoid a silly out of memory exception
					{
						if (Attr != null)
						{
							using (Image img2 = BitmapConstructor.CreateBitmap(Img.Width, Img.Height)) //wish it where simpler... but attr on texturebrushes does not seem to work.
							{
								using (Graphics gr = Graphics.FromImage(img2))
								{
									PointF[] ImageRect = new PointF[]{new PointF(0,0), new PointF(img1.Width, 0), new PointF(0, img1.Height)};
									gr.DrawImage(img1, ImageRect, new RectangleF(0,0,img1.Width, img1.Height), GraphicsUnit.Pixel, Attr);
									Result = new TextureBrush(img2, new RectangleF(0,0,w,h));
								}
							}
						}
						else
						{
							Result = new TextureBrush(img1, new RectangleF(0,0,w,h));
						}
						
						Result.TranslateTransform(Coords.Left, Coords.Top);
						if (FillType == TFillType.Picture && img1.Width>0 && img1.Height>0 && Coords.Height > 0 && Coords.Width > 0)
						{
							Result.ScaleTransform(Coords.Width / img1.Width, Coords.Height /img1.Height, MatrixOrder.Prepend) ;
						}
						else
							if (FillType != TFillType.Picture)
						{
							Result.ScaleTransform(Zoom100, Zoom100, MatrixOrder.Prepend) ;
						}
						return Result;
					}
				}
				finally
				{
					if (Img != null) Img.Dispose();
				}
			}
			finally
			{
				if (Attr != null) Attr.Dispose();
			}
		}

        internal static Brush GetSolidBrush(TShapeProperties ShProp, ExcelFile Workbook, TShadowInfo ShadowInfo, TPathFillMode PathFillMode)
        {
            
			real FillOpacity = ShProp.ShapeOptions.As1616(TShapeOption.fillOpacity, 1);
			bool IsTransparent;
			Color aColor = ColorFromLong(ShProp, GetUInt(ShProp.ShapeOptions[TShapeOption.fillColor], 0xffffffff), FillOpacity, Workbook, true, out IsTransparent, 0);
			if (ShadowInfo.Style != TShadowStyle.Obscured && IsTransparent) return null;

			if (ShadowInfo.Style != TShadowStyle.None)
			{
				TShapeOption ShadowColor = ShadowInfo.Pass <= 1? TShapeOption.shadowColor: TShapeOption.shadowHighlight;
				aColor = ColorFromLong(ShProp, (UInt32)ShProp.ShapeOptions.AsLong(ShadowColor, 0x808080),  
					ShProp.ShapeOptions.As1616(TShapeOption.shadowOpacity, 1), Workbook, true, out IsTransparent, 0);
				if (IsTransparent) return null;
			}

            switch (PathFillMode)
            {
                case TPathFillMode.Norm:
                    break;
                case TPathFillMode.None:
                    return null;
                case TPathFillMode.Lighten:
                    aColor = LightColor(aColor, 0.6);
                    break;

                case TPathFillMode.LightenLess:
                    aColor = LightColor(aColor, 0.8);
                    break;

                case TPathFillMode.Darken:
                    aColor = DarkColor(aColor, 0.6);
                    break;

                case TPathFillMode.DarkenLess:
                    aColor = DarkColor(aColor, 0.8);
                    break;
            }

			return new SolidBrush(aColor);
        }

		internal static Brush GetBrush(RectangleF Coords, TShapeProperties ShProp, ExcelFile Workbook, TShadowInfo ShadowInfo, real Zoom100)
		{
			return GetBrush(Coords, ShProp, Workbook, ShadowInfo, false, Zoom100);
		}

        internal static Brush GetBrush(RectangleF Coords, TShapeProperties ShProp, ExcelFile Workbook, TShadowInfo ShadowInfo, bool IsImage, real Zoom100)
        {
            return GetBrush(Coords, ShProp, Workbook, ShadowInfo, IsImage, Zoom100, TPathFillMode.Norm);
        }

        internal static Brush GetBrush(RectangleF Coords, TShapeProperties ShProp, ExcelFile Workbook, TShadowInfo ShadowInfo, 
            bool IsImage, real Zoom100, TPathFillMode PathFillMode)
        {
            if (ShProp.ShapeOptions == null) return null;

			if (ShadowInfo.Style == TShadowStyle.None 
				|| (ShadowInfo.Style == TShadowStyle.Normal && !IsImage))  //Images have shadows even if they are not filled.
			{
				if (!ShProp.ShapeOptions.AsBool(TShapeOption.fNoFillHitTest, !IsImage, 4)) return null;  //ffilled.
			}
			
			if (ShadowInfo.Style == TShadowStyle.None)
            {
                int ShadeType = GetInt(ShProp.ShapeOptions[TShapeOption.fillType], 0);
                switch ((TFillType)ShadeType)
                {
					case TFillType.Pattern:
					case TFillType.Texture:
					case TFillType.Picture:
						return GetTextureBrush(ShProp, Workbook, (TFillType)ShadeType, Coords, Zoom100);

                    case TFillType.Shade:
                    case TFillType.ShadeCenter:
                    case TFillType.ShadeScale:
                    case TFillType.ShadeShape:
                    case TFillType.ShadeTitle:
                        return GetGradientBrush(Coords, ShProp, Workbook, (TFillType)ShadeType);
                }
            }
            //Anything else:
            return GetSolidBrush(ShProp, Workbook, ShadowInfo, PathFillMode);
        }

		internal static Pen GetPen(TShapeProperties ShProp, ExcelFile Workbook, TShadowInfo ShadowInfo)
		{
			return GetPen(ShProp, Workbook, ShadowInfo, true);
		}

        internal static Pen GetPen(TShapeProperties ShProp, ExcelFile Workbook, TShadowInfo ShadowInfo, bool DefaultDrawLine)
        {
            object o = ShProp.ShapeOptions[TShapeOption.lineColor];
            if (o == null || !( o is Int64)) o = 0xff000000;
            
            if (ShadowInfo.Style != TShadowStyle.Obscured && !ShProp.ShapeOptions.AsBool(TShapeOption.fNoLineDrawDash, DefaultDrawLine, 3)) return null; //fline

            object lw = ShProp.ShapeOptions[TShapeOption.lineWidth];
            real LineWidth = TLineStyle.DefaultWidth;
            if (lw != null && lw is Int64)
            {
                LineWidth = Convert.ToSingle(lw);
            }       
         
            bool IsTransparent;
            Color aColor = ColorFromLong(ShProp, Convert.ToUInt32(o),  1, Workbook, false, out IsTransparent, 0);
            if (ShadowInfo.Style != TShadowStyle.Obscured && IsTransparent) return null;
            if (ShadowInfo.Style != TShadowStyle.None)
            {
                TShapeOption ShadowColor = ShadowInfo.Pass <= 1? TShapeOption.shadowColor: TShapeOption.shadowHighlight;
                aColor = ColorFromLong(ShProp, (UInt32)ShProp.ShapeOptions.AsLong(ShadowColor, 0x808080),  
                    ShProp.ShapeOptions.As1616(TShapeOption.shadowOpacity, 1), Workbook, true, out IsTransparent, 0);
                if (IsTransparent) return null;
            }

            Pen Result = new Pen(aColor, LineWidth / 12700f);

            object Dashing = ShProp.ShapeOptions[TShapeOption.lineDashing];
            if (Dashing != null && Dashing is Int64)
            {
                switch ((TLineDashing)Convert.ToInt32(Dashing))
                {
                    case TLineDashing.DashDotDotSys:
                        Result.DashStyle = DashStyles.DashDotDot; break;
                    case TLineDashing.DashDotSys:
                    case TLineDashing.DashDotGEL:
                        Result.DashStyle = DashStyles.DashDot; break;
                    case TLineDashing.DashGEL:
                    case TLineDashing.DashSys:
                        Result.DashStyle = DashStyles.Dash; break;
                    case TLineDashing.DotGEL:
                    case TLineDashing.DotSys:
                        Result.DashStyle = DashStyles.Dot; break;
                    case TLineDashing.LongDashDotDotGEL:
                        Result.DashStyle = DashStyles.DashDotDot; break;
                    case TLineDashing.LongDashDotGEL:
                        Result.DashStyle = DashStyles.DashDot; break;
                    case TLineDashing.LongDashGEL:
                        Result.DashStyle = DashStyles.Dash; break;
                }


            }
            return Result;
            
        }

        #endregion

        #region Utilities
        internal static void Flip(ref RectangleF Source, RectangleF Coords, TShapeProperties ShProp)
        {
            if (ShProp.FlipH)
                FlipH(ref Source, ref Coords);
            if (ShProp.FlipV)
                FlipV(ref Source, ref Coords);
        }

        private static void FlipH(ref RectangleF Source, ref RectangleF Coords)
        {
            Source.X = Coords.Right - (Source.X - Coords.Left) - Source.Width;
        }

        private static void FlipV(ref RectangleF Source, ref RectangleF Coords)
        {
            Source.Y = Coords.Bottom - (Source.Y - Coords.Top) - Source.Height;
        }

        internal static void Flip(ref TPointF[] Source, RectangleF Coords, TShapeProperties ShProp)
        {
            if (ShProp.FlipH)
            {
                FlipH(ref Source, Coords);
            }
            if (ShProp.FlipV)
            {
                FlipV(ref Source, Coords);
            }
        }

        private static void FlipV(ref TPointF[] Source, RectangleF Coords)
        {
            for (int i = 0; i < Source.Length; i++)
            {
                Source[i].Y = Coords.Bottom - (Source[i].Y - Coords.Top);
            }
        }

        private static void FlipH(ref TPointF[] Source, RectangleF Coords)
        {
            for (int i = 0; i < Source.Length; i++)
            {
                Source[i].X = Coords.Right - (Source[i].X - Coords.Left);
            }
        }

        #endregion

        #region Arrows
        private static void DrawOvalArrow(IFlxGraphics Canvas, Pen aPen, real x2, real y2, real h, real w)
        {
            TPointF[] Points = GetOval(x2 - h ,y2 - w, h, w);

            if (aPen != null)
            {
                using (Brush aBrush = new SolidBrush(aPen.Color))
                {
                    Canvas.DrawAndFillBeziers(aPen, aBrush, Points);
                }
            }
        }
        private static void DrawArrow(IFlxGraphics Canvas, Pen aPen, real x1, real y1, real x2, real y2, TArrowStyle Style, real h, real w, TClippingStyle Clipping)
        {
            if (Style == TArrowStyle.None) return;

            real r = (real)Math.Sqrt((x2 - x1) * (x2 - x1) + (y2 - y1) * (y2 - y1));
            real CosAlpha = 0;
            if (x2 != x1) CosAlpha = (x2 - x1) / r;

            real SinAlpha = 0;
            if (y2 != y1) SinAlpha = (y2 - y1) / r;

            if (Style == TArrowStyle.Oval) 
            {
                DrawOvalArrow(Canvas, aPen, x2, y2, h, w);
                return;
            }

            int pcount = 3;
            if (Style == TArrowStyle.Stealth || Style == TArrowStyle.Diamond)
                pcount = 4;

            TPointF[] Points = new TPointF[pcount];
            

            if (Style == TArrowStyle.Diamond)
            {
                Points[0] = new TPointF(x2 + w * SinAlpha, y2 - w * CosAlpha);
                Points[1] = new TPointF(x2 + h * CosAlpha, y2 + h * SinAlpha);
                Points[2] = new TPointF(x2 - w * SinAlpha, y2 + w * CosAlpha);
                Points[3] = new TPointF(x2 - h * CosAlpha, y2 - h * SinAlpha);
            }

            else
            {
                Points[0] = new TPointF(x2 - 2 * h * CosAlpha + w * SinAlpha, y2 - 2 * h * SinAlpha - w * CosAlpha);
                Points[1] = new TPointF(x2, y2);
                Points[2] = new TPointF(x2 - 2 * h * CosAlpha - w * SinAlpha, y2 - 2 * h * SinAlpha + w * CosAlpha);
            }

            if (Style == TArrowStyle.Stealth)
            {
                Points[3] = new TPointF(x2 - h * 2 / 3 * CosAlpha, y2 - h * 2 / 3 * SinAlpha);
            }

            if (Style == TArrowStyle.Open)
            {
                if (Clipping == TClippingStyle.None) Canvas.DrawLines(aPen, Points);
                return;
            }

            if (aPen != null)
            {
                using (Brush aBrush = new SolidBrush(aPen.Color))
                {
                    Canvas.DrawAndFillPolygon(aPen, aBrush, Points);
                }
            }

        }

        #endregion
		
        #region basic shapes
        internal static void DrawRectangle(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                if (br != null) Canvas.FillRectangle(br, Coords, Clipping);
            }

            using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
            {
                if (pe != null && Clipping == TClippingStyle.None) Canvas.DrawRectangle(pe, Coords.X, Coords.Y, Coords.Width, Coords.Height);
            }
            
        }	


        /// <summary>
        /// Draws a quarter circle using splines.
        /// </summary>
        /// <param name="Points">Array to fill.</param>
        /// <param name="i">Position on points where we will fill.</param>
        /// <param name="rx">x radius.</param>
        /// <param name="ry">y radius.</param>
        /// <param name="Quad">0 is Top-Left, from there we move clockwise.</param>
        private static void AddCorner(TPointF[] Points, int i, real rx, real ry, int Quad)
        {
            int sg1 = Quad <= 1? 1 : -1;
            int sg2 = (Quad == 0 || Quad == 3) ? 1 : -1;
            int m = (Quad == 1 || Quad == 3)? 0: 1;
            real rxt = 0.5541f * rx * -sg2;  //fixme: change by 0.551784
            real ryt = 0.5541f * ry * -sg1;

            Points[i+2] = new TPointF(Points[i - 1].X + rx * sg1, Points[i - 1].Y + ry * -sg2);  //endpoint.
            Points[i] = new TPointF(Points[i - 1].X + rxt * (1 - m), Points[i - 1].Y + ryt * m);  //control point 1
            Points[i+1] = new TPointF(Points[i + 2].X + rxt * m, Points[i + 2].Y + ryt * (1 - m));  //control point 2
        }

        private static TPointF[] GetOval(real x, real y, real rx, real ry)
        {
            TPointF[] Result = new TPointF[13];

            Result[0] = new TPointF(x + rx, y);
            AddCorner(Result, 1, rx, ry, 1);
            AddCorner(Result, 4, rx, ry, 2);
            AddCorner(Result, 7, rx, ry, 3);
            AddCorner(Result, 10, rx, ry, 0);
            return Result;
        }

        private static void AddLine(List<TPointF> Points, real x, real y)
        {
            //draw line.
            Points.Add(Points[Points.Count - 1]); //cp1
            Points.Add(new TPointF(x, y));  //cp2
            Points.Add(Points[Points.Count - 1]); //endpoint
        }

        private static void AddLineAndCorner(List<TPointF> Points, real x, real y, real r, int Quad)
        {
            AddLineAndCorner(Points, x, y, r, r, Quad);
        }

        private static void AddLineAndCorner(List<TPointF> Points, real x, real y, real rx ,real ry, int Quad)
        {
            int a1 = Quad == 1? -1: Quad == 3? 1: 0;
            int a2 = Quad == 0? 1: Quad == 2? -1: 0;
            //draw line.
            AddLine(Points, x + a1 * rx, y + a2*ry);

            if (rx > 0 && ry > 0)
            {
                TPointF[] ArrPoints = new TPointF[4];
                ArrPoints[0] = (TPointF)Points[Points.Count - 1];
                AddCorner(ArrPoints, 1, rx, ry, Quad);
                for (int i = 1; i < ArrPoints.Length; i++)
                    Points.Add(ArrPoints[i]);
            }
        }

        private static void AddLineAndCorner(TPointF[] Points, int i, real r, int Quad, real dx, real dy)
        {
            AddLineAndCorner(Points, i, r, r, Quad, dx, dy);
        }
        private static void AddLineAndCorner(TPointF[] Points, int i, real rx, real ry, int Quad, real dx, real dy)
        {
            //draw line.
            Points[i] = Points[i - 1]; //cp1
            Points[i + 1] = new TPointF(Points[i - 1].X + dx, Points[i - 1].Y + dy);  //cp2
            Points[i + 2] = Points[i + 1]; //endpoint

            //draw corner.
            AddCorner(Points, i + 3, rx, ry, Quad);
        }

        internal static RectangleF DrawRoundRectangle(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            real r1 = ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjustValue, 0xd60) / 10800f;
            
            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    return DrawRoundRectangle(Canvas, Clipping, Coords, r1, br, pe);
                }
            }
        }

        public static RectangleF DrawRoundRectangle(IFlxGraphics Canvas, TClippingStyle Clipping, RectangleF Coords, real r1, Brush br, Pen pe)
        {
            real r = Math.Min(Coords.Width, Coords.Height) / 2f * r1;
            TPointF[] Points = new TPointF[25];

            Points[0] = new TPointF(Coords.X + r, Coords.Y);
            AddLineAndCorner(Points, 1, r, 1, Coords.Width - 2 * r, 0);
            AddLineAndCorner(Points, 7, r, 2, 0, Coords.Height - 2 * r);
            AddLineAndCorner(Points, 13, r, 3, -(Coords.Width - 2 * r), 0);
            AddLineAndCorner(Points, 19, r, 0, 0, -(Coords.Height - 2 * r));

            Canvas.DrawAndFillBeziers(pe, br, Points, Clipping);
            
            return new RectangleF(Coords.X + r / 3f, Coords.Y + r / 3f, Coords.Width - 2 * r / 3f, Coords.Height - 2 * r / 3f);
        }	

        internal static void DrawTriangle(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, real r, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            DrawTriangle(Canvas, Workbook, ShProp, Coords, r, ShadowInfo, Clipping, false, Zoom100);
        }

        internal static void DrawTriangle(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, real r, TShadowInfo ShadowInfo, TClippingStyle Clipping, bool VInverted, real Zoom100)
        {
            TPointF[] Points = new TPointF[3];

            real Left = ShProp.FlipH? Coords.Right: Coords.Left;
            real Right = ShProp.FlipH? Coords.Left: Coords.Right;
            real Top = ShProp.FlipV ^ VInverted? Coords.Bottom: Coords.Top;
            real Bottom = ShProp.FlipV  ^ VInverted? Coords.Top: Coords.Bottom;

            real r1 = ShProp.FlipH? -r: r;

            Points[0] = new TPointF(Left + r1, Top);
            Points[1] = new TPointF(Left, Bottom);
            Points[2] = new TPointF(Right, Bottom);


            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillPolygon(pe, br, Points, Clipping);
                }
            }
        }

        internal static RectangleF DrawIsocelesTriangle(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            real r1 = ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjustValue, 10800) / 21600f;           
            real r = Coords.Width * r1;
            DrawTriangle(Canvas, Workbook, ShProp, Coords, r, ShadowInfo, Clipping, Zoom100);
            RectangleF Result = new RectangleF(Coords.Left + (r1 - 0.5f) * Coords.Width / 2f + Coords.Width / 3.65f, Coords.Top + Coords.Height / 2f, Coords.Width * (1 - 2/3.65f), Coords.Height *1f / 3f);
            Flip(ref Result, Coords, ShProp);
            return Result;
        }	

        internal static RectangleF DrawRightTriangle(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            DrawTriangle(Canvas, Workbook, ShProp, Coords, 0, ShadowInfo, Clipping, Zoom100);
            RectangleF Result = new RectangleF(Coords.Left + Coords.Width / 8f, Coords.Top + Coords.Height * 0.58f, Coords.Width * .45f, Coords.Height * 0.325f);
            Flip(ref Result, Coords, ShProp);
            return Result;
        }	

        internal static RectangleF DrawOval(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            real rx = Coords.Width / 2f;
            real ry = Coords.Height / 2f;
            TPointF[] Points = GetOval(Coords.X, Coords.Y, rx, ry);

            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillBeziers(pe, br, Points, Clipping);
                }
            }

            return new RectangleF(Coords.X + rx / 3f, Coords.Y + ry / 3f, Coords.Width - 2*rx/3f, Coords.Height - 2*ry/3f);
        }

        private static real ArrowDimh(long h1)
        {
            return ArrowDimw(h1);
        }

        private static real ArrowDimw(long h1)
        {
            real w = 2.8f;
            if (h1 == 0) w = 1.8f; else if (h1 == 2) w = 4.2f;
            return w;
        }


		internal static void DoArrows(IFlxGraphics Canvas, Pen pe, TShapeProperties ShProp, real Left, real Top, real Right, real Bottom, TClippingStyle Clipping)
		{
			DoArrows(Canvas, pe, ShProp, Left, Top, Right, Bottom, Right, Bottom, Left, Top, Clipping);
		}

		internal static void DoArrows(IFlxGraphics Canvas, Pen pe, TShapeProperties ShProp, TPointF[] Points, TClippingStyle Clipping)
		{
			int n = Points.Length;
			if (n < 2) return;
			int n01 = 1;
			while (Points[n01] == Points[0] && n01 < n) n01++;
			int nn2 = n-1;
			while (Points[nn2] == Points[n-1] && nn2 >0) nn2--;
			DoArrows(Canvas, pe, ShProp, Points[nn2].X, Points[nn2].Y, Points[n-1].X, Points[n-1].Y, Points[n01].X, Points[n01].Y, Points[0].X, Points[0].Y, Clipping);
		}

		internal static void DoArrows(IFlxGraphics Canvas, Pen pe, TShapeProperties ShProp, real Left1, real Top1, real Right1, real Bottom1, real Left2, real Top2, real Right2, real Bottom2, TClippingStyle Clipping)
		{
			TArrowStyle st = (TArrowStyle) ShProp.ShapeOptions.AsLong(TShapeOption.lineStartArrowhead, 0);           
			if (st != TArrowStyle.None)
			{
				long h1 = ShProp.ShapeOptions.AsLong(TShapeOption.lineStartArrowLength, 1);
				real h = ArrowDimh(h1);

				h1 = ShProp.ShapeOptions.AsLong(TShapeOption.lineStartArrowWidth, 1);
				real w = ArrowDimw(h1);
				DrawArrow(Canvas, pe, Left2, Top2, Right2, Bottom2, st, h, w, Clipping); 
			}

			st = (TArrowStyle) ShProp.ShapeOptions.AsLong(TShapeOption.lineEndArrowhead, 0);           
			if (st != TArrowStyle.None)
			{
				long h1 = ShProp.ShapeOptions.AsLong(TShapeOption.lineEndArrowLength, 1);
				real h = ArrowDimh(h1);

				h1 = ShProp.ShapeOptions.AsLong(TShapeOption.lineEndArrowWidth, 1);
				real w = ArrowDimw(h1);
				DrawArrow(Canvas, pe, Left1, Top1, Right1, Bottom1, st, h, w, Clipping); 
			}
		}

        internal static RectangleF DrawLine(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            real Left = ShProp.FlipH? Coords.Right: Coords.Left;
            real Right = ShProp.FlipH? Coords.Left: Coords.Right;
            real Top = ShProp.FlipV? Coords.Bottom: Coords.Top;
            real Bottom = ShProp.FlipV? Coords.Top: Coords.Bottom;
            
            using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
            {
                if (Clipping == TClippingStyle.None) Canvas.DrawLine(pe, Left, Top, Right, Bottom);
				DoArrows(Canvas, pe, ShProp, Left, Top, Right, Bottom, Clipping);
            }

            return Coords;
        }
        #endregion

        #region Basic 2
        internal static RectangleF DrawDiamond(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            TPointF[] Points = new TPointF[4];
            Points[0] = new TPointF(Coords.Left + Coords.Width / 2, Coords.Top);
            Points[1] = new TPointF(Coords.Right, Coords.Top + Coords.Height / 2);
            Points[2] = new TPointF(Coords.Left + Coords.Width / 2, Coords.Bottom);
            Points[3] = new TPointF(Coords.Left, Coords.Top + Coords.Height / 2);

            Flip(ref Points, Coords, ShProp);

            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillPolygon(pe, br, Points, Clipping);
                }
            }

            RectangleF Result = new RectangleF(Coords.Left + Coords.Width / 4, Coords.Top + Coords.Height / 4, Coords.Width /2 , Coords.Height / 2);
            Flip(ref Result, Coords, ShProp);
            return Result;
        }	

        internal static RectangleF DrawParallelogram(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100, long xDefault)
        {
            real rx = ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjustValue, xDefault) * Coords.Width / 21600f;           
            
            TPointF[] Points = new TPointF[4];
            Points[0] = new TPointF(Coords.Left + rx, Coords.Top);
            Points[1] = new TPointF(Coords.Right, Coords.Top);
            Points[2] = new TPointF(Coords.Right - rx, Coords.Bottom);
            Points[3] = new TPointF(Coords.Left, Coords.Bottom);

            Flip(ref Points, Coords, ShProp);

            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillPolygon(pe, br, Points, Clipping);
                }
            }

            RectangleF Result = new RectangleF(Coords.Left + rx, Coords.Top , Coords.Width - 2 * rx , Coords.Height);
            Flip(ref Result, Coords, ShProp);
            return Result;
        }	

        internal static RectangleF DrawTrapezoid(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100, long xDefault, bool inv)
        {
            //This one is different in 2003 / 2007. 2007 inverts the shape upsidedown.
            real rx = ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjustValue, xDefault) * Coords.Width / 21600f;           
            
            TPointF[] Points = new TPointF[4];
            Points[0] = new TPointF(Coords.Left, Coords.Top);
            Points[1] = new TPointF(Coords.Right, Coords.Top);
            Points[2] = new TPointF(Coords.Right - rx, Coords.Bottom);
            Points[3] = new TPointF(Coords.Left  + rx, Coords.Bottom);

            Flip(ref Points, Coords, ShProp);
            if (inv) FlipV(ref Points, Coords);

            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillPolygon(pe, br, Points, Clipping);
                }
            }

            RectangleF Result = new RectangleF(Coords.Left + rx, Coords.Top, Coords.Width - 2 *rx , Coords.Height);
            Flip(ref Result, Coords, ShProp);
            return Result;
        }	

        internal static RectangleF DrawOctagon(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            real MinHW = Math.Min(Coords.Width, Coords.Height);
            real rxy = ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjustValue, 21600 * 2 / 7) * MinHW / 21600f;           
            
            TPointF[] Points = new TPointF[8];
            Points[0] = new TPointF(Coords.Left + rxy, Coords.Top);
            Points[1] = new TPointF(Coords.Right - rxy, Coords.Top);
            Points[2] = new TPointF(Coords.Right, Coords.Top + rxy);
            Points[3] = new TPointF(Coords.Right, Coords.Bottom - rxy);
            Points[4] = new TPointF(Coords.Right - rxy, Coords.Bottom);
            Points[5] = new TPointF(Coords.Left + rxy, Coords.Bottom);
            Points[6] = new TPointF(Coords.Left, Coords.Bottom - rxy);
            Points[7] = new TPointF(Coords.Left, Coords.Top + rxy);

            Flip(ref Points, Coords, ShProp);

            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillPolygon(pe, br, Points, Clipping);
                }
            }

            RectangleF Result = new RectangleF(Coords.Left + rxy / 2, Coords.Top + rxy / 2, Coords.Width - rxy , Coords.Height - rxy);
            Flip(ref Result, Coords, ShProp);
            return Result;
        }	

        internal static RectangleF DrawHexagon(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100, long xDefault)
        {
            real rxy = ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjustValue, xDefault) * Coords.Width / 21600f;           
            
            TPointF[] Points = new TPointF[6];
            Points[0] = new TPointF(Coords.Left + rxy, Coords.Top);
            Points[1] = new TPointF(Coords.Right - rxy, Coords.Top);
            Points[2] = new TPointF(Coords.Right, Coords.Top + Coords.Height / 2);
            Points[3] = new TPointF(Coords.Right - rxy, Coords.Bottom);
            Points[4] = new TPointF(Coords.Left + rxy, Coords.Bottom);
            Points[5] = new TPointF(Coords.Left, Coords.Top + Coords.Height / 2);

            Flip(ref Points, Coords, ShProp);

            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillPolygon(pe, br, Points, Clipping);
                }
            }

            RectangleF Result = new RectangleF(Coords.Left + rxy / 2, Coords.Top + rxy / 2, Coords.Width - rxy , Coords.Height - rxy);
            Flip(ref Result, Coords, ShProp);
            return Result;
        }
	
        internal static RectangleF DrawRegularPentagon(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            real ry = Coords.Height * 3 / 8f;
            real rx = Coords.Width * 6 / 31f;
            
            TPointF[] Points = new TPointF[5];
            Points[0] = new TPointF(Coords.Left + Coords.Width / 2, Coords.Top);
            Points[1] = new TPointF(Coords.Right, Coords.Top + ry);
            Points[2] = new TPointF(Coords.Right - rx, Coords.Bottom);
            Points[3] = new TPointF(Coords.Left + rx, Coords.Bottom);
            Points[4] = new TPointF(Coords.Left, Coords.Top + ry);

            Flip(ref Points, Coords, ShProp);

            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillPolygon(pe, br, Points, Clipping);
                }
            }

            RectangleF Result = new RectangleF(Coords.Left + rx, Coords.Top + ry *2 / 3, Coords.Width - 2 *rx , Coords.Height - ry * 2/3);
            Flip(ref Result, Coords, ShProp);
            return Result;
        }	

        internal static RectangleF DrawCross(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            real MinHW = Math.Min(Coords.Width, Coords.Height);
            real rxy = ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjustValue, 21600 * 1 / 4) * MinHW / 21600f;           
            
            TPointF[] Points = new TPointF[12];
            Points[0] = new TPointF(Coords.Left + rxy, Coords.Top);
            Points[1] = new TPointF(Coords.Right - rxy, Coords.Top);
            Points[2] = new TPointF(Coords.Right - rxy, Coords.Top + rxy);
            Points[3] = new TPointF(Coords.Right, Coords.Top + rxy);
            Points[4] = new TPointF(Coords.Right, Coords.Bottom - rxy);
            Points[5] = new TPointF(Coords.Right - rxy, Coords.Bottom - rxy);
            Points[6] = new TPointF(Coords.Right - rxy, Coords.Bottom);
            Points[7] = new TPointF(Coords.Left + rxy, Coords.Bottom);
            Points[8] = new TPointF(Coords.Left + rxy, Coords.Bottom - rxy);
            Points[9] = new TPointF(Coords.Left, Coords.Bottom - rxy);
            Points[10] = new TPointF(Coords.Left, Coords.Top + rxy);
            Points[11] = new TPointF(Coords.Left + rxy, Coords.Top + rxy);

            Flip(ref Points, Coords, ShProp);

            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillPolygon(pe, br, Points, Clipping);
                }
            }

            RectangleF Result = new RectangleF(Coords.Left + rxy, Coords.Top + rxy, Coords.Width - 2*rxy , Coords.Height - 2*rxy);
            Flip(ref Result, Coords, ShProp);
            return Result;
        }	


        #endregion

        #region Basic 3
        internal static RectangleF DrawCilinder(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            real ry = ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjustValue, 21600 / 4) * Coords.Height / 21600f;
            
            TPointF[] Points = new TPointF[19];
            Points[0] = new TPointF(Coords.X, Coords.Bottom - ry / 2);
            AddLineAndCorner(Points, 1, Coords.Width / 2, ry / 2,  0, 0, -(Coords.Height - ry));
            AddCorner(Points, 7, Coords.Width / 2, ry/2, 1);
            AddLineAndCorner(Points, 10, Coords.Width / 2, ry / 2,  2, 0, Coords.Height - ry);
            AddCorner(Points, 16, Coords.Width / 2, ry/2, 3);

            Flip(ref Points, Coords, ShProp);

            TPointF[] Points2 = new TPointF[7];
            Points2[0] = new TPointF(Coords.Right, Coords.Top + ry / 2);
            AddCorner(Points2, 1, Coords.Width / 2, ry/2, 2);
            AddCorner(Points2, 4, Coords.Width / 2, ry/2, 3);

            Flip(ref Points2, Coords, ShProp);

            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillBeziers(pe, br, Points, Clipping);
                    if (DrawLines(ShadowInfo, br != null)) Canvas.DrawAndFillBeziers(pe, null, Points2, Clipping);
                }
            }

            RectangleF Result = new RectangleF(Coords.Left, Coords.Top + ry, Coords.Width , Coords.Height - ry * 3/ 2);
            Flip(ref Result, Coords, ShProp);
            return Result;
        }	
        
        internal static int LightUp(int c, int light)
        {
            int Result = c;
            switch (light)
            {
                case 2:
                    Result = (int) Math.Round(101.49 + c * 0.6047);
                    break;

                case 1:
                    Result = (int) Math.Ceiling(c + 50 - c / 5.0);
                    break;

                case -1:
                    Result = (int) Math.Round( 0.1113 + c * 0.8031);
                    break;

                case -2:
                    Result = (int) Math.Round( -0.1184 + c * 0.6);
                    break;

            }

        
            if (Result > 255) return 255;
            if (Result < 0) return 0;
            return Result;
        }

        private static Brush GetBrFinal(Brush br, ref Brush br2, int light)
        {
            return GetBrFinal(br, ref br2, light, TShadowStyle.None);
        }
        private static Brush GetBrFinal(Brush br, ref Brush br2, int light, TShadowStyle ShadowStyle)
        {
            if (ShadowStyle != TShadowStyle.None) return br;
            SolidBrush brs = br as SolidBrush;
            if (brs != null)
            {
                Color c = brs.Color;
                br2 = new SolidBrush(ColorUtil.FromArgb(c.A, LightUp(c.R, light), LightUp(c.G, light), LightUp(c.B, light)));
                return br2;
            }
            else
            {
                return br;
            }
        }

        internal static RectangleF DrawCube(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            real MinHW = Math.Min(Coords.Width, Coords.Height);
            real rxy = ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjustValue, 21600 / 4) * MinHW / 21600f;
            
            TPointF[] Points1 = new TPointF[4];
            Points1[0] = new TPointF(Coords.Left, Coords.Top + rxy);
            Points1[1] = new TPointF(Coords.Right - rxy, Coords.Top + rxy);
            Points1[2] = new TPointF(Coords.Right - rxy, Coords.Bottom);
            Points1[3] = new TPointF(Coords.Left, Coords.Bottom);
            Flip(ref Points1, Coords, ShProp);

            TPointF[] Points2 = new TPointF[4];
            Points2[0] = new TPointF(Coords.Left, Coords.Top + rxy);
            Points2[1] = new TPointF(Coords.Right - rxy, Coords.Top + rxy);
            Points2[2] = new TPointF(Coords.Right , Coords.Top);
            Points2[3] = new TPointF(Coords.Left + rxy, Coords.Top);
            Flip(ref Points2, Coords, ShProp);

            TPointF[] Points3 = new TPointF[4];
            Points3[0] = new TPointF(Coords.Right - rxy, Coords.Top + rxy);
            Points3[1] = new TPointF(Coords.Right , Coords.Top);
            Points3[2] = new TPointF(Coords.Right, Coords.Bottom - rxy);
            Points3[3] = new TPointF(Coords.Right - rxy, Coords.Bottom);
            Flip(ref Points3, Coords, ShProp);


            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillPolygon(pe, br, Points1, Clipping);

                    Brush br2 = null;
                    try
                    {
                        Brush Br2Final = GetBrFinal(br, ref br2, 1, ShadowInfo.Style); //We will not dispose brxFinal
                        Canvas.DrawAndFillPolygon(pe, Br2Final, Points2, Clipping);
                    }
                    finally
                    {
                        if (br2 != null) br2.Dispose();
                        br2 = null;
                    }

                    Brush br3 = null;
                    try
                    {
                        Brush Br3Final = GetBrFinal(br, ref br3, -1, ShadowInfo.Style); //We will not dispose brxFinal
                        Canvas.DrawAndFillPolygon(pe, Br3Final, Points3, Clipping);
                    }
                    finally
                    {
                        if (br3 != null) br3.Dispose();
                        br3 = null;
                    }
                }
            }

            RectangleF Result = new RectangleF(Coords.Left, Coords.Top + rxy, Coords.Width - rxy , Coords.Height - rxy);
            Flip(ref Result, Coords, ShProp);
            return Result;
        }	


        internal static RectangleF DrawBevel(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
                real MinHW = Math.Min(Coords.Width, Coords.Height);
                real rxy = ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjustValue, 21600 / 8) * MinHW / 21600f;           

            if (ShadowInfo.Style == TShadowStyle.None) 
            {
                TPointF[] Points1 = new TPointF[4];
                Points1[0] = new TPointF(Coords.Left + rxy, Coords.Top + rxy);
                Points1[1] = new TPointF(Coords.Right - rxy, Coords.Top + rxy);
                Points1[2] = new TPointF(Coords.Right - rxy, Coords.Bottom - rxy);
                Points1[3] = new TPointF(Coords.Left + rxy, Coords.Bottom - rxy);
                Flip(ref Points1, Coords, ShProp);

                TPointF[] Points2 = new TPointF[4];
                Points2[0] = new TPointF(Coords.Left, Coords.Top);
                Points2[1] = new TPointF(Coords.Right, Coords.Top);
                Points2[2] = new TPointF(Coords.Right - rxy, Coords.Top + rxy);
                Points2[3] = new TPointF(Coords.Left + rxy, Coords.Top + rxy);
                Flip(ref Points2, Coords, ShProp);

                TPointF[] Points3 = new TPointF[4];
                Points3[0] = new TPointF(Coords.Right, Coords.Top);
                Points3[1] = new TPointF(Coords.Right , Coords.Bottom);
                Points3[2] = new TPointF(Coords.Right - rxy, Coords.Bottom - rxy);
                Points3[3] = new TPointF(Coords.Right - rxy, Coords.Top + rxy);
                Flip(ref Points3, Coords, ShProp);

                TPointF[] Points4 = new TPointF[4];
                Points4[0] = new TPointF(Coords.Right, Coords.Bottom);
                Points4[1] = new TPointF(Coords.Left , Coords.Bottom);
                Points4[2] = new TPointF(Coords.Left + rxy, Coords.Bottom - rxy);
                Points4[3] = new TPointF(Coords.Right - rxy, Coords.Bottom - rxy);
                Flip(ref Points4, Coords, ShProp);

                TPointF[] Points5 = new TPointF[4];
                Points5[0] = new TPointF(Coords.Left, Coords.Bottom);
                Points5[1] = new TPointF(Coords.Left , Coords.Top);
                Points5[2] = new TPointF(Coords.Left + rxy, Coords.Top + rxy);
                Points5[3] = new TPointF(Coords.Left + rxy, Coords.Bottom - rxy);
                Flip(ref Points5, Coords, ShProp);

                using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
                {
                    using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                    {
                        Canvas.DrawAndFillPolygon(pe, br, Points1, Clipping);

                        Brush br2 = null;
                        try
                        {
                            Brush Br2Final = GetBrFinal(br, ref br2, 1, ShadowInfo.Style); //We will not dispose brxFinal
                            Canvas.DrawAndFillPolygon(pe, Br2Final, Points2, Clipping);
                        }
                        finally
                        {
                            if (br2 != null) br2.Dispose();
                            br2 = null;
                        }

                        Brush br3 = null;
                        try
                        {
                            Brush Br3Final = GetBrFinal(br, ref br3, -2, ShadowInfo.Style); //We will not dispose brxFinal
                            Canvas.DrawAndFillPolygon(pe, Br3Final, Points3, Clipping);
                        }
                        finally
                        {
                            if (br3 != null) br3.Dispose();
                            br3 = null;
                        }

                        Brush br4 = null;
                        try
                        {
                            Brush Br4Final = GetBrFinal(br, ref br4, -1); //We will not dispose brxFinal
                            Canvas.DrawAndFillPolygon(pe, Br4Final, Points4, Clipping);
                        }
                        finally
                        {
                            if (br4 != null) br4.Dispose();
                            br4 = null;
                        }

                        Brush br5 = null;
                        try
                        {
                            Brush Br5Final = GetBrFinal(br, ref br5, 2); //We will not dispose brxFinal
                            Canvas.DrawAndFillPolygon(pe, Br5Final, Points5, Clipping);
                        }
                        finally
                        {
                            if (br5 != null) br5.Dispose();
                            br5 = null;
                        }

                    }
                }
            }
            else
            {
                TPointF[] Points1 = new TPointF[4];
                Points1[0] = new TPointF(Coords.Left, Coords.Top);
                Points1[1] = new TPointF(Coords.Right, Coords.Top);
                Points1[2] = new TPointF(Coords.Right, Coords.Bottom);
                Points1[3] = new TPointF(Coords.Left, Coords.Bottom);
                using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
                {
                    using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                    {
                        Canvas.DrawAndFillPolygon(pe, br, Points1, Clipping);
                    }
                }
            }

            RectangleF Result = new RectangleF(Coords.Left + rxy, Coords.Top + rxy, Coords.Width - 2 *rxy , Coords.Height - 2 *rxy);
            Flip(ref Result, Coords, ShProp);
            return Result;
        }	


        internal static RectangleF DrawFoldedSheet(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            real rxy0 = 1 - ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjustValue, 21600 * 7 / 8) / 21600f;           
            real rx = rxy0 * Coords.Width;
            real ry = rxy0 * Coords.Height;

            TPointF[] Points1 = new TPointF[5];
            Points1[0] = new TPointF(Coords.Left, Coords.Top);
            Points1[1] = new TPointF(Coords.Right, Coords.Top);
            Points1[2] = new TPointF(Coords.Right, Coords.Bottom - ry);
            Points1[3] = new TPointF(Coords.Right - rx, Coords.Bottom);
            Points1[4] = new TPointF(Coords.Left, Coords.Bottom);
            Flip(ref Points1, Coords, ShProp);

            TPointF[] Points2 = new TPointF[7];
            Points2[0] = new TPointF(Coords.Right, Coords.Bottom - ry);
            Points2[1] = new TPointF(Coords.Right - rx * 2/4f, Coords.Bottom - ry * 3/4f); //control
            Points2[2] = new TPointF(Coords.Right - rx * 3/4f, Coords.Bottom - ry * 7/8f); //control

            Points2[3] = new TPointF(Coords.Right - rx * 3 /4f, Coords.Bottom - ry);
            Points2[4] = Points2[3];  //control

            Points2[5] = new TPointF(Coords.Right - rx, Coords.Bottom); //control
            Points2[6] = Points2[5];

            Flip(ref Points2, Coords, ShProp);

            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillPolygon(pe, br, Points1, Clipping);

                    if (DrawLines(ShadowInfo, br != null)) 
                    {
                        Brush br2 = null;
                        try
                        {
                            Brush Br2Final = GetBrFinal(br, ref br2, -1); //We will not dispose brxFinal
                            Canvas.DrawAndFillBeziers(pe, Br2Final, Points2, Clipping);
                        }
                        finally
                        {
                            if (br2 != null) br2.Dispose();
                            br2 = null;
                        }
                    }
                }
            }

            RectangleF Result = new RectangleF(Coords.Left, Coords.Top, Coords.Width, Coords.Height - ry);
            Flip(ref Result, Coords, ShProp);
            return Result;
        }	


        internal static RectangleF DrawSmiley(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            real ryMouth = (ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjustValue, 17400) - 16500) / 10000f * Coords.Height;           

            real rx = Coords.Width / 2f;
            real ry = Coords.Height / 2f;
            TPointF[] Points1 = GetOval(Coords.X, Coords.Y, rx, ry);

            real rx2 = rx * 2 / 17f;
            real ry2 = ry * 2 / 17f;
            TPointF[] Points2 = GetOval(Coords.X + Coords.Width * (9f / 31f), Coords.Y + Coords.Height * (5f /17f), rx2, ry2);

            TPointF[] Points3 = new TPointF[Points2.Length];
            for (int i = 0; i < Points2.Length; i++)
            {
                Points3[i] = new TPointF(Coords.Left + Coords.Right - Points2[i].X, Points2[i].Y);
            }

            Flip(ref Points2, Coords, ShProp); //Flip after loadng points3
            Flip(ref Points3, Coords, ShProp);

            real xHang = Coords.Width / 4f;
            real StartMouth = ryMouth < 0? Coords.Left + xHang: Coords.Right - xHang;
            TPointF[] Points4 = new TPointF[7];
            Points4[0] = new TPointF(StartMouth, Coords.Top + Coords.Height * 3f / 4f);
            int Quad1 = ryMouth < 0? 0: 2;
            int Quad2 = ryMouth < 0? 1: 3;
            AddCorner(Points4, 1, Coords.Width / 2f - xHang, Math.Abs(ryMouth) , Quad1); 
            AddCorner(Points4, 4, Coords.Width / 2f - xHang, Math.Abs(ryMouth), Quad2);
            Points4[1] = Points4[0];
            Points4[Points4.Length - 2] = Points4[Points4.Length - 1];


            Flip(ref Points4, Coords, ShProp);

            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillBeziers(pe, br, Points1, Clipping);
                    if (DrawLines(ShadowInfo, br != null)) 
                    {
                        Canvas.DrawAndFillBeziers(pe, null, Points4, Clipping);

                        Brush br2 = null;
                        try
                        {
                            Brush Br2Final = GetBrFinal(br, ref br2, -1); //We will not dispose brxFinal
                            Canvas.DrawAndFillBeziers(pe, Br2Final, Points2, Clipping);
                            Canvas.DrawAndFillBeziers(pe, Br2Final, Points3, Clipping);
                        }
                        finally
                        {
                            if (br2 != null) br2.Dispose();
                            br2 = null;
                        }
                    }

                }
            }

            return new RectangleF(Coords.X + rx / 3f, Coords.Y + ry / 3f, Coords.Width - 2*rx/3f, Coords.Height - 2*ry/3f);
        }
    

        internal static RectangleF DrawDonut(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            real rx = Coords.Width / 2;
            real ry = Coords.Height / 2;
            real r1 = 0.5f - ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjustValue, 21600 / 4) / 21600f;
            real rx2 = r1 * Coords.Width;           
            real ry2 = r1 * Coords.Height;           

            TPointF[] Points1 = GetOval(Coords.X, Coords.Y, rx, ry);
            TPointF[] Points2 = GetOval(Coords.X + rx - rx2, Coords.Y + ry - ry2, rx2, ry2);

            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    switch (Clipping)
                    {
                        case TClippingStyle.Exclude:
                            //There is no way on the current API to exclude the donut. Excel does not do it either, anyway.
                            //Canvas.DrawAndFillBeziers(pe, br, Points2, TClippingStyle.Include);
                            //Canvas.DrawAndFillBeziers(pe, br, Points1, Clipping);
                            break;
                        case TClippingStyle.Include:
                            Canvas.DrawAndFillBeziers(pe, br, Points1, Clipping);
                            Canvas.DrawAndFillBeziers(pe, br, Points2, TClippingStyle.Exclude);
                            break;

                        default:
                            Canvas.SaveState();
                            try
                            {
                                Canvas.DrawAndFillBeziers(pe, br, Points2, TClippingStyle.Exclude);
                                Canvas.DrawAndFillBeziers(pe, br, Points1, Clipping);
                            }
                            finally
                            {
                                Canvas.RestoreState();
                            }
                            Canvas.DrawAndFillBeziers(pe, null, Points2, Clipping);
                            break;
                    }
                }
            }

            return new RectangleF(Coords.X + rx / 3f, Coords.Y + ry / 3f, Coords.Width - 2*rx/3f, Coords.Height - 2*ry/3f);
        }
    

        internal static RectangleF DrawNo(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            real rx = Coords.Width / 2;
            real ry = Coords.Height / 2;
            real r1 = ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjustValue, 21600 / 9) / 21600f * 2;
            real rx2 = (1 -r1) * Coords.Width / 2;           
            real ry2 = (1 -r1) * Coords.Height / 2;           

            TPointF[] Points1 = GetOval(Coords.X, Coords.Y, rx, ry);
            TPointF[] Points2 = GetOval(Coords.X + rx - rx2, Coords.Y + ry - ry2, rx2, ry2);
            TPointF[] Points3 = new TPointF[4];
            real rxr = (rx - rx2)*2/3;
            real ryr = (ry - ry2)*2/3;
            Points3[0] = new TPointF(Coords.Left + rxr, Coords.Top);
            Points3[1] = new TPointF(Coords.Right, Coords.Bottom - ryr);
            Points3[2] = new TPointF(Coords.Right - rxr, Coords.Bottom);
            Points3[3] = new TPointF(Coords.Left, Coords.Top + ryr);
            Flip(ref Points3, Coords, ShProp);



            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    switch (Clipping)
                    {
                        case TClippingStyle.Exclude:
                            //There is no way on the current API to exclude the donut. Excel does not do it either, anyway.
                            //Canvas.DrawAndFillBeziers(pe, br, Points2, TClippingStyle.Include);
                            //Canvas.DrawAndFillBeziers(pe, br, Points1, Clipping);
                            break;
                        case TClippingStyle.Include:
                            Canvas.DrawAndFillBeziers(pe, br, Points1, Clipping);
                            Canvas.DrawAndFillBeziers(pe, br, Points2, TClippingStyle.Exclude);
                            break;

                        default:
                            Canvas.SaveState();
                            try
                            {
                                Canvas.DrawAndFillBeziers(pe, br, Points2, TClippingStyle.Exclude);
                                Canvas.DrawAndFillBeziers(pe, br, Points1, Clipping);
                            }
                            finally
                            {
                                Canvas.RestoreState();
                            }
                            Canvas.SaveState();
                            try
                            {
                                Canvas.DrawAndFillPolygon(pe, Brushes.Black, Points3, TClippingStyle.Exclude);
                                Canvas.DrawAndFillBeziers(pe, null, Points2, Clipping);
                            }
                            finally
                            {
                                Canvas.RestoreState();
                            }

                            Canvas.SaveState();
                            try
                            {
                                Canvas.DrawAndFillBeziers(pe, Brushes.Black, Points2, TClippingStyle.Include);
                                Canvas.DrawAndFillPolygon(pe, br, Points3, Clipping);
                            }
                            finally
                            {
                                Canvas.RestoreState();
                            }
                            break;
                    }
                }
            }

            return new RectangleF(Coords.X + rx / 3f, Coords.Y + ry / 3f, Coords.Width - 2*rx/3f, Coords.Height - 2*ry/3f);
        }
    

        #endregion

        #region Basic 4
        internal static RectangleF DrawHeart(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            TPointF[] Points = new TPointF[25];
            Points[0] = new TPointF(Coords.Left + Coords.Width / 2, Coords.Bottom);
            Points[1] = Points[0];
            Points[2] = new TPointF(Coords.Left + Coords.Width / 40, Coords.Top + Coords.Height * 17 / 48);
            Points[3] = Points[2];
            
            Points[4] = new TPointF(Coords.Left, Coords.Top + Coords.Height * 13 / 48);
            Points[5] = new TPointF(Coords.Left + Coords.Width * 0 / 80, Coords.Top + Coords.Height * 8/ 48);
            Points[6] = new TPointF(Coords.Left + Coords.Width * 7 / 80, Coords.Top + Coords.Height * 3/ 48);

            Points[7] = new TPointF(Coords.Left + Coords.Width *14 / 80, Coords.Top + 1 / 48);
            Points[8] = new TPointF(Coords.Left + Coords.Width / 4, Coords.Top);
            Points[9] = new TPointF(Coords.Left + Coords.Width / 4, Coords.Top);

            Points[10] = new TPointF(Coords.Left + Coords.Width /4, Coords.Top);
            Points[11] = new TPointF(Coords.Left + Coords.Width * 64 / 160, Coords.Top);
            Points[12] = new TPointF(Coords.Left + Coords.Width / 2, Coords.Top + Coords.Height *5 / 48);

            for (int i= (Points.Length + 1) / 2; i < Points.Length; i++)
                Points[i] = new TPointF(Coords.Left + Coords.Right - Points[Points.Length - 1 - i].X, Points[Points.Length - 1 - i].Y);


            Flip(ref Points, Coords, ShProp);



            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillBeziers(pe, br, Points, Clipping);
                }
            }

            RectangleF Result = new RectangleF(Coords.X + Coords.Width / 4 , Coords.Y + Coords.Height / 8, Coords.Width /2f, Coords.Height /2);
            Flip(ref Result, Coords, ShProp);
            return Result;
        }
    
        internal static RectangleF DrawLightning(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            TPointF[] Points = new TPointF[11];
            Points[0] = new TPointF(Coords.Left, Coords.Top + Coords.Height * 8 / 44);
            Points[1] = new TPointF(Coords.Left + Coords.Width * 22.8f / 58, Coords.Top);
            Points[2] = new TPointF(Coords.Left + Coords.Width * 34.5f/ 58, Coords.Top + Coords.Height * 12.2f / 44);
            Points[3] = new TPointF(Coords.Left + Coords.Width * 29.8f/ 58, Coords.Top + Coords.Height * 13.9f / 44);
            
            Points[4] = new TPointF(Coords.Left + Coords.Width * 44.5f/ 58, Coords.Top + Coords.Height * 24.4f / 44);
            Points[5] = new TPointF(Coords.Left + Coords.Width * 39.6f / 58, Coords.Top + Coords.Height * 26.1f/ 44);
            Points[6] = new TPointF(Coords.Right, Coords.Bottom);

            Points[7] = new TPointF(Coords.Left + Coords.Width * 26.9f/ 58, Coords.Top + Coords.Height * 30.2f / 44);
            Points[8] = new TPointF(Coords.Left + Coords.Width * 32.6f/ 58, Coords.Top + Coords.Height * 28.2f / 44);
            Points[9] = new TPointF(Coords.Left + Coords.Width * 13.5f/ 58, Coords.Top + Coords.Height * 19.6f / 44);

            Points[10] = new TPointF(Coords.Left + Coords.Width * 20.3f/ 58, Coords.Top + Coords.Height * 17 / 44);


            Flip(ref Points, Coords, ShProp);



            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillPolygon(pe, br, Points, Clipping);
                }
            }

            RectangleF Result = new RectangleF(Coords.X + Coords.Width * 13 / 32f , Coords.Y + Coords.Height / 2.96f, Coords.Width * 0.245f, Coords.Height / 3.2f);
            Flip(ref Result, Coords, ShProp);
            return Result;
        }
    

        private static void Rotate(TPointF[] Source, ref TPointF[] Dest, real x, real y, real Alpha, real xToy, real yTox)
        {
            real SinAlpha = (real)Math.Sin(Alpha);
            real CosAlpha = (real)Math.Cos(Alpha);
            for (int i = 0; i < Source.Length; i++)
            {
                Dest[i].X = x + (Source[i].X - x) * CosAlpha - (Source[i].Y - y) * SinAlpha * yTox; 
                Dest[i].Y = y + (Source[i].X - x) * SinAlpha * xToy + (Source[i].Y - y) * CosAlpha; 
            }
        }
        
        internal static RectangleF DrawSun(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            real r0 = (ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjustValue, 21600 / 4)) / 21600f;           

            real rx = (0.5f-r0) * Coords.Width;
            real ry = (0.5f-r0) * Coords.Height;
            real rx2 = r0 * Coords.Width;
            real ry2 = r0 * Coords.Height;

            TPointF[] Points1 = GetOval(Coords.X + rx2, Coords.Y + ry2, rx, ry);

            TPointF[] Ray = new TPointF[3];
            Ray[0] = new TPointF(Coords.Left + Coords.Width / 2, Coords.Top);
            Ray[1] = new TPointF(Coords.Left + Coords.Width / 2 - rx / 4f, Coords.Top + ry2 * 4f / 5f);
            Ray[2] = new TPointF(Coords.Left + Coords.Width / 2 + rx / 4f, Coords.Top + ry2 * 4f / 5f);


            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillBeziers(pe, br, Points1, Clipping);
                    Canvas.DrawAndFillPolygon(pe, br, Ray, Clipping);
                    TPointF[] Ray2 = new TPointF[3];
                    real xToy = Coords.Width > 0? Coords.Height / Coords.Width: 0;
                    real yTox = Coords.Height > 0? Coords.Width / Coords.Height: 0;
                    for (int i = 1; i < 8; i++)
                    {
                        //Rotate uses a reference to avoid creating objects on the heap.
                        real Angle = (real) (i * Math.PI / 4f);
                        Rotate(Ray, ref Ray2, Coords.Left + Coords.Width / 2f, Coords.Top + Coords.Height / 2f, Angle, xToy, yTox);
                        Canvas.DrawAndFillPolygon(pe, br, Ray2, Clipping);
                    }
                }
            }

            return new RectangleF(Coords.X + rx2 + rx / 3f, Coords.Y + ry2 + ry / 3f, Coords.Width - 2*rx2 - 2*rx/3f, Coords.Height - 2*ry2 - 2*ry/3f);
        }
    

        internal static RectangleF DrawMoon(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            real r0 = (ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjustValue, 21600 / 2)) / 21600f;           

            real rx2 = (1f-r0) * Coords.Width;
            real rx = Coords.Width;
            real ry = Coords.Height / 2;

            TPointF[] Points = new TPointF[13];
            Points[0] = new TPointF(Coords.Left + rx, Coords.Bottom);
            AddCorner(Points, 1, rx, ry, 3);
            AddCorner(Points, 4, rx, ry, 0);
            AddCorner(Points, 7, rx2, ry, 1);
            AddCorner(Points, 10, rx2, ry, 2);

            real x0 = Points[6].X;
            for (int i = 7; i < Points.Length; i++) //Invert the draw direction.
            {
                Points[i].X = 2* x0 - Points[i].X;
            }

            Flip(ref Points, Coords, ShProp);

            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillBeziers(pe, br, Points, Clipping);
                }
            }

            RectangleF Result = new RectangleF(Coords.X + rx / 6f, Coords.Y + Coords.Height / 4, Coords.Width  - rx2 - rx/6f, Coords.Height / 2);
            Flip(ref Result, Coords, ShProp);
            return Result;
        }
    

        internal static RectangleF DrawArc(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            real Left = ShProp.FlipH? Coords.Right: Coords.Left;
            real Right = ShProp.FlipH? Coords.Left: Coords.Right;
            real Top = ShProp.FlipV? Coords.Bottom: Coords.Top;
            real Bottom = ShProp.FlipV? Coords.Top: Coords.Bottom;
            
            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    int ExtraPt = br != null? 3: 0;
                    TPointF[] Points = new TPointF[4 + ExtraPt];
                    int Quad = 0;
                    if (Left <= Right)
                    {
                        if (Bottom > Top) Quad = 1;
                        else Quad = 2;
                    }
                    else
                    {
                        if (Bottom > Top) Quad = 0;
                        else Quad = 3;
                    }

                    if (Quad == 1 || Quad == 3)                     
                        Points[0] = new TPointF(Left, Top);
                    else
                        Points[0] = new TPointF(Right, Bottom);

                    AddCorner(Points, 1, Coords.Width, Coords.Height, Quad);
                    if (ExtraPt > 0) 
                    {
                        Points[4] = new TPointF(Left, Bottom);
                        Points[5] = new TPointF(Left, Bottom);
                        Points[6] = new TPointF(Left, Bottom);
                    }
                    Canvas.DrawAndFillBeziers(pe, br, Points, Clipping);

                    TArrowStyle st = (TArrowStyle) ShProp.ShapeOptions.AsLong(TShapeOption.lineStartArrowhead, 0);           
                    if (st != TArrowStyle.None)
                    {
                        long h1 = ShProp.ShapeOptions.AsLong(TShapeOption.lineStartArrowLength, 1);
                        real h = ArrowDimh(h1);

                        h1 = ShProp.ShapeOptions.AsLong(TShapeOption.lineStartArrowWidth, 1);
                        real w = ArrowDimw(h1);
                        DrawArrow(Canvas, pe, Right, Top, Left, Top, st, h, w, Clipping); 
                    }

                    st = (TArrowStyle) ShProp.ShapeOptions.AsLong(TShapeOption.lineEndArrowhead, 0);           
                    if (st != TArrowStyle.None)
                    {
                        long h1 = ShProp.ShapeOptions.AsLong(TShapeOption.lineEndArrowLength, 1);
                        real h = ArrowDimh(h1);

                        h1 = ShProp.ShapeOptions.AsLong(TShapeOption.lineEndArrowWidth, 1);
                        real w = ArrowDimw(h1);
                        DrawArrow(Canvas, pe, Right, Top, Right, Bottom, st, h, w, Clipping); 
                    }
                }
            }

            return Coords;
        }

        internal static RectangleF DrawPlaque(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            real r1 = ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjustValue, 0xd60) / 10800f;
            
            real r = Math.Min(Coords.Width, Coords.Height) / 2f * r1;
            TPointF[] Points = new TPointF[25];

            Points[0] = new TPointF(Coords.Right - r, Coords.Y);
            AddLineAndCorner(Points, 1, r, 2, -(Coords.Width - 2*r), 0);
            AddLineAndCorner(Points, 7, r, 1, 0, Coords.Height - 2*r);
            AddLineAndCorner(Points, 13, r, 0, (Coords.Width - 2*r), 0);
            AddLineAndCorner(Points, 19, r, 3, 0, -(Coords.Height - 2*r));


            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillBeziers(pe, br, Points, Clipping);
                }
            }

            return new RectangleF(Coords.X + r * 2/ 3f, Coords.Y + r *2/ 3f, Coords.Width - 2*r*2/3f, Coords.Height - 2*r*2/3f);
        }	


        #endregion

        #region Callouts
        internal static void AddHCallOut(List<TPointF> Points, RectangleF Coords, real x, real y, bool LeftToRight)
        {
            real x1 = x > Coords.Width / 2? Coords.X + Coords.Width * 7f / 12f: Coords.X + Coords.Width / 6f;
            real x2 = x > Coords.Width / 2? Coords.X + Coords.Width * 5f / 6f: Coords.X + Coords.Width * 5f / 12f;
            if (!LeftToRight)
            {
                real tmp = x1; x1 = x2; x2 = tmp;
            }
            
            real Y0 = LeftToRight? Coords.Y: Coords.Bottom;
            AddLine(Points, x1, Y0);
            AddLine(Points,Coords.X + x, Coords.Y + y);
            AddLine(Points, x2 , Y0);
        }

        internal static void AddVCallOut(List<TPointF> Points, RectangleF Coords, real x, real y, bool TopToBottom)
        {
            real y1 = y > Coords.Height / 2? Coords.Y + Coords.Height * 7f / 12f: Coords.Y + Coords.Height / 6f;
            real y2 = y > Coords.Height / 2? Coords.Y + Coords.Height * 5f / 6f: Coords.Y + Coords.Height * 5f / 12f;
            if (!TopToBottom)
            {
                real tmp = y1; y1 = y2; y2 = tmp;
            }
            
            real X0 = TopToBottom? Coords.Right: Coords.X;
            AddLine(Points, X0, y1);
            AddLine(Points, Coords.X + x, Coords.Y + y);
            AddLine(Points, X0, y2);
        }

        internal static TQuarter CalcQuarter(RectangleF Coords, real x, real y)
        {
            if (Coords.Height == 0)
            {
                if (y < 0) return TQuarter.Top;
                return TQuarter.Bottom;
            }

            real wh = Coords.Width / Coords.Height;

            if (x >=0 && x <= Coords.Width && y >=0 && y <= Coords.Height) return TQuarter.None;
            x -= Coords.Width / 2;
            y -= Coords.Height / 2;
            if ( x >= -y * wh)
            {
                if (x >= y * wh)
                    return TQuarter.Right;
                return TQuarter.Bottom;
            }

            if (x >= y * wh)
                return TQuarter.Top;
            return TQuarter.Left;
        }

        internal static RectangleF DrawRectCallOut(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100, bool Round)
        {
            long xDef = (long)(21600f / 16f);
            long yDef = (long)(21600f * 1.2f);
            real x; real y;
            unchecked
            {
                x = (int)(ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjustValue, xDef)) * Coords.Width / 21600f;
                y = (int)(ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjust2Value, yDef)) * Coords.Height / 21600f;
            }

            if (ShProp.FlipH) x = Coords.Width-x;
            if (ShProp.FlipV) y = Coords.Height-y;

            real r = 0;
            if (Round)
            {
                real r1 = 0xd60 / 10800f;
                r = Math.Min(Coords.Width, Coords.Height) / 2f * r1;
            }
            
            TQuarter Quarter = CalcQuarter(Coords, x, y);

            List<TPointF> Points = new List<TPointF>();
            Points.Add(new TPointF(Coords.X + r, Coords.Y));
            if (Quarter == TQuarter.Top) AddHCallOut(Points, Coords, x, y, true);
            AddLineAndCorner(Points, Coords.Right, Coords.Y, r, 1);
            if (Quarter == TQuarter.Right) AddVCallOut(Points, Coords, x, y, true);
            AddLineAndCorner(Points, Coords.Right, Coords.Bottom, r, 2);
            if (Quarter == TQuarter.Bottom) AddHCallOut(Points, Coords, x, y, false);
            AddLineAndCorner(Points, Coords.Left, Coords.Bottom, r, 3);
            if (Quarter == TQuarter.Left) AddVCallOut(Points, Coords, x, y, false);
            AddLineAndCorner(Points, Coords.X, Coords.Y, r, 0);

            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    TPointF[] PointArr = Points.ToArray();
                    Canvas.DrawAndFillBeziers(pe, br, PointArr, Clipping);
                }
            }

            return new RectangleF(Coords.X + r / 3f, Coords.Y + r / 3f, Coords.Width - 2*r/3f, Coords.Height - 2*r/3f);
        }	

        #endregion

        #region Block Arrows
        internal static RectangleF DrawHBlockArrow(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100, bool ToLeft)
        {
            long xDef = ToLeft? 21600 - 16200: 16200;
            real rx = ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjustValue, xDef) * Coords.Width / 21600f;
            real ry = ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjust2Value, 10800 / 2) * Coords.Height / 21600f;  

            real rx1 = 0;
            if (ToLeft)
            {
                rx1 = 1;
                rx = Coords.Width - rx;
            }

            TPointF[] Points = new TPointF[7];
            DoHBlockArrow(ref Points, 0, 1, Coords, rx, ry, ToLeft ^ ShProp.FlipH);

            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillPolygon(pe, br, Points, Clipping);
                }
            }

            real dx = 0;
            real ArrowWidth = Coords.Width - rx;
            if (Coords.Height > 0)
            {
                dx = ry * (ArrowWidth) /(Coords.Height / 2);
            }

            RectangleF Result = new RectangleF(Coords.Left + rx1 * (ArrowWidth - dx), Coords.Top + ry, rx + dx, Coords.Height - ry * 2);
            Flip(ref Result, Coords, ShProp);
            return Result;
        }	

        internal static void DoHBlockArrow(ref TPointF[] Points, int p0, int d, RectangleF Coords, real rx, real ry, bool DrawToLeft)
        {
            real Left = DrawToLeft? Coords.Right: Coords.Left;
            real Right = DrawToLeft? Coords.Left: Coords.Right;
            real Top = Coords.Top;
            real Bottom = Coords.Bottom;
            if (DrawToLeft) 
            {
                rx = -rx;
            }

            if (p0 + 0 * d < Points.Length) Points[p0 + 0*d] = new TPointF(Left, Top + ry);
            Points[p0 + 1*d] = new TPointF(Left + rx, Top + ry);
            Points[p0 + 2*d] = new TPointF(Left + rx, Top);
            Points[p0 + 3*d] = new TPointF(Right, Top + Coords.Height /2f);
            Points[p0 + 4*d] = new TPointF(Left + rx, Bottom);
            Points[p0 + 5*d] = new TPointF(Left + rx, Bottom - ry);
            Points[p0 + 6*d] = new TPointF(Left, Bottom - ry);
        }


        internal static RectangleF DrawVBlockArrow(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100, bool ToTop)
        {
            long yDef = ToTop? 21600 - 16200: 16200;
            real rx = ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjust2Value, 10800 / 2) * Coords.Width / 21600f;
            real ry = ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjustValue, yDef) * Coords.Height / 21600f;  

            real ry1 = 0;
            if (ToTop)
            {
                ry1 = 1;
                ry = Coords.Height - ry;
            }

            TPointF[] Points = new TPointF[7];
            DoVBlockArrow(ref Points, 0, 1, Coords, rx, ry, ToTop ^ ShProp.FlipV);

            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillPolygon(pe, br, Points, Clipping);
                }
            }

            real dy = 0;
            real ArrowHeight = Coords.Height - ry;
            if (Coords.Width > 0)
            {
                dy = rx * (ArrowHeight) /(Coords.Width / 2);
            }

            RectangleF Result = new RectangleF(Coords.Left + rx, Coords.Top + ry1 * (ArrowHeight - dy), Coords.Width - rx * 2, ry + dy);
            Flip(ref Result, Coords, ShProp);
            return Result;
        }	

        internal static void DoVBlockArrow(ref TPointF[] Points, int p0, int d, RectangleF Coords, real rx, real ry, bool DrawToTop)
        {
            real Left = Coords.Left;
            real Right = Coords.Right;
            real Top = DrawToTop? Coords.Bottom: Coords.Top;
            real Bottom = DrawToTop? Coords.Top: Coords.Bottom;
            if (DrawToTop) 
            {
                ry = -ry;
            }

            if (p0 + 0 * d < Points.Length) Points[p0 + 0 * d] = new TPointF(Left + rx, Top);
            Points[p0 + 1*d] = new TPointF(Left + rx, Top + ry);
            Points[p0 + 2*d] = new TPointF(Left, Top + ry);
            Points[p0 + 3*d] = new TPointF(Left + Coords.Width / 2f, Bottom);
            Points[p0 + 4*d] = new TPointF(Right, Top + ry);
            Points[p0 + 5*d] = new TPointF(Right - rx, Top + ry);
            Points[p0 + 6*d] = new TPointF(Right - rx, Top);
        }


        internal static RectangleF DrawLeftRightBlockArrow(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            long xDef = 21600 - 16200;
            real rx = ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjustValue, xDef) * Coords.Width / 21600f;
            real ry = ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjust2Value, 10800 / 2) * Coords.Height / 21600f;  

            real ArrowWidth = rx;
            rx = Coords.Width / 2 - rx;

            TPointF[] Points = new TPointF[12];
            RectangleF Coords1 = new RectangleF(Coords.X + Coords.Width / 2f, Coords.Y, Coords.Width / 2f, Coords.Height);
            DoHBlockArrow(ref Points, 0, 1, Coords1, rx, ry, false);
            Coords1.X = Coords.X;
            DoHBlockArrow(ref Points, 12, -1, Coords1, rx, ry, true);

            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillPolygon(pe, br, Points, Clipping);
                }
            }

            real dx = 0;
            if (Coords.Height > 0)
            {
                dx = ry * (ArrowWidth) /(Coords.Height / 2);
            }

            RectangleF Result = new RectangleF(Coords.Left + (ArrowWidth - dx), Coords.Top + ry, 2* rx + 2* dx, Coords.Height - ry * 2);
            return Result;
        }	


        internal static RectangleF DrawUpDownBlockArrow(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            long yDef =21600 - 16200;
            real rx = ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjustValue, 10800 / 2) * Coords.Width / 21600f;
            real ry = ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjust2Value, yDef) * Coords.Height / 21600f;  

            real ArrowHeight = ry;
            ry = Coords.Height / 2 - ry;

            TPointF[] Points = new TPointF[12];
            RectangleF Coords1 = new RectangleF(Coords.X, Coords.Y + Coords.Height / 2f, Coords.Width, Coords.Height / 2f);
            DoVBlockArrow(ref Points, 0, 1, Coords1, rx, ry, false);
            Coords1.Y = Coords.Y;
            DoVBlockArrow(ref Points, 12, -1, Coords1, rx, ry, true);

            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillPolygon(pe, br, Points, Clipping);
                }
            }

            real dy = 0;
            if (Coords.Width > 0)
            {
                dy = rx * (ArrowHeight) /(Coords.Width / 2);
            }

            RectangleF Result = new RectangleF(Coords.Left + rx, Coords.Top + (ArrowHeight - dy), Coords.Width - rx * 2, 2* ry + 2* dy);
            return Result;
        }	


        internal static RectangleF DrawQuadBlockArrow(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            long rxExt0 = ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjustValue, 21600 * 3 / 10); 
            real rxExtV = rxExt0 * Coords.Width / 21600f; 
            real rxExtH = rxExt0 * Coords.Height / 21600f;

            real rx0 = ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjust2Value, 21600 * 2 / 5);
            real rxV =  rx0 * Coords.Width / 21600f;
            real rxH =  rx0 * Coords.Height / 21600f;

            real ArrowHeight0 = ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjust3Value, 21600 * 1 / 5);
            real ArrowHeightV =  ArrowHeight0 * Coords.Height / 21600f;
            real ArrowWidthH =  ArrowHeight0 * Coords.Width / 21600f;

            real TailWidthV = Coords.Width - 2 * rxV;
            real TailHeightH = Coords.Height - 2 * rxH;

            real TailHeightV = Coords.Height / 2 - ArrowHeightV;
            real TailWidthH = Coords.Width / 2 - ArrowWidthH;
            
            TPointF[] Points = new TPointF[24];
            RectangleF Coords1 = FlexCelRender.RectangleXY(Coords.X + rxExtV, Coords.Y, Coords.Right - rxExtV, (Coords.Top + Coords.Bottom - TailHeightH) / 2f);
            DoVBlockArrow(ref Points, 0, 1, Coords1, rxV - rxExtV, TailHeightV - TailHeightH / 2, true);
            
            RectangleF Coords2 = FlexCelRender.RectangleXY((Coords.Left + Coords.Right + TailWidthV ) /2f, Coords.Y + rxExtH, Coords.Right, Coords.Bottom - rxExtH);
            DoHBlockArrow(ref Points, 6, 1, Coords2, TailWidthH - TailWidthV / 2, rxH - rxExtH, false);

            Coords1.Y += (Coords.Height + TailHeightH) / 2;
            DoVBlockArrow(ref Points, 18, -1, Coords1, rxV - rxExtV, TailHeightV - TailHeightH / 2, false);

            Coords2.X -= (Coords.Width + TailWidthV) / 2;
            DoHBlockArrow(ref Points, 24, -1, Coords2, TailWidthH - TailWidthV / 2, rxH - rxExtH, true);

            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillPolygon(pe, br, Points, Clipping);
                }
            }

            real dx = 0;
            if (Coords.Height > 0)
            {
                dx = (rxExtV - rxExtV) * (ArrowWidthH) /(Coords.Height / 2);
            }

            RectangleF Result = new RectangleF(Coords.Left + (ArrowWidthH - dx), Coords.Top + rxH, 2* TailWidthH + 2* dx, Coords.Height - rxH * 2);
            return Result;
        }	

        internal static RectangleF DrawNotchedArrow(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            long xDef = 16200;
            real rx = ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjustValue, xDef) * Coords.Width / 21600f;
            real ry = ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjust2Value, 10800 / 2) * Coords.Height / 21600f;  

            TPointF[] Points = new TPointF[8];
            DoHBlockArrow(ref Points, 0, 1, Coords, rx, ry, ShProp.FlipH);

            real ArrowWidth = Coords.Width - rx;
            real Left = ShProp.FlipH? Coords.Right: Coords.Left;
            real TailHeight = Coords.Height - 2 * ry;
            real x0 = TailHeight * ArrowWidth;
            if (Coords.Height > 0) x0 /= (Coords.Height) ; else x0 = 0;
            real x1 = ShProp.FlipH? -x0: x0;
            Points[7] = new TPointF(Left + x1, (Coords.Top + Coords.Bottom )/2);

            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillPolygon(pe, br, Points, Clipping);
                }
            }

            real dx = 0;
            if (Coords.Height > 0)
            {
                dx = ry * (ArrowWidth) /(Coords.Height / 2);
            }

            RectangleF Result = new RectangleF(Coords.Left + x0, Coords.Top + ry, rx + dx - x0, Coords.Height - ry * 2);
            Flip(ref Result, Coords, ShProp);
            return Result;
        }	


        internal static RectangleF DrawStripedArrow(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            long xDef = 16200;
            real rx = ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjustValue, xDef) * Coords.Width / 21600f;
            real ry = ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjust2Value, 10800 / 2) * Coords.Height / 21600f;  

            TPointF[] Points = new TPointF[7];
            real StartX = Coords.Width * 5 / 32;
            RectangleF Coords2 = new RectangleF(Coords.Left + StartX, Coords.Top, Coords.Width - StartX, Coords.Height);
            DoHBlockArrow(ref Points, 0, 1, Coords2, rx - StartX, ry, false);

            RectangleF Rect1 = new RectangleF(Coords.Left, Coords.Top + ry, Coords.Width / 32, Coords.Height - 2 * ry);
            RectangleF Rect2 = new RectangleF(Coords.Left + Coords.Width / 16, Coords.Top + ry, Coords.Width / 16, Coords.Height - 2*ry);

            Flip(ref Points, Coords, ShProp);
            Flip(ref Rect1, Coords, ShProp);
            Flip(ref Rect2, Coords, ShProp);

            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillPolygon(pe, br, Points, Clipping);
                    Canvas.FillRectangle(br, Rect1);
                    Canvas.DrawRectangle(pe, Rect1.Left, Rect1.Top, Rect1.Width, Rect1.Height);
                    Canvas.FillRectangle(br, Rect2);
                    Canvas.DrawRectangle(pe, Rect2.Left, Rect2.Top, Rect2.Width, Rect2.Height);

                }
            }

            real dx = 0;
            if (Coords.Height > 0)
            {
                real ArrowWidth = Coords.Width - rx;
                dx = ry * (ArrowWidth) /(Coords.Height / 2);
            }

            RectangleF Result = new RectangleF(Coords.Left + StartX , Coords.Top + ry, rx + dx - StartX, Coords.Height - ry * 2);
            Flip(ref Result, Coords, ShProp);
            return Result;
        }	

        
        internal static RectangleF DrawPentagon(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            long xDef = 21600 * 3 / 4;
            real rx = ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjustValue, xDef) * Coords.Width / 21600f;           

            TPointF[] Points = new TPointF[5];
            Points[0] = new TPointF(Coords.Left, Coords.Top);
            Points[1] = new TPointF(Coords.Left + rx, Coords.Top);
            Points[2] = new TPointF(Coords.Right, Coords.Top + Coords.Height / 2);
            Points[3] = new TPointF(Coords.Left + rx, Coords.Bottom);
            Points[4] = new TPointF(Coords.Left, Coords.Bottom);

            Flip(ref Points, Coords, ShProp);

            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillPolygon(pe, br, Points, Clipping);
                }
            }

            real ArrowWidth = Coords.Width - rx;

            RectangleF Result = new RectangleF(Coords.Left, Coords.Top, Coords.Width - ArrowWidth / 2 , Coords.Height);
            Flip(ref Result, Coords, ShProp);
            return Result;
        }	


        internal static RectangleF DrawChevron(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            long xDef = 21600 * 3 / 4;
            real rx = ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjustValue, xDef) * Coords.Width / 21600f;           

            TPointF[] Points = new TPointF[6];
            Points[0] = new TPointF(Coords.Left, Coords.Top);
            Points[1] = new TPointF(Coords.Left + rx, Coords.Top);
            Points[2] = new TPointF(Coords.Right, Coords.Top + Coords.Height / 2);
            Points[3] = new TPointF(Coords.Left + rx, Coords.Bottom);
            Points[4] = new TPointF(Coords.Left, Coords.Bottom);
            Points[5] = new TPointF(Coords.Right - rx, Coords.Top + Coords.Height / 2);

            Flip(ref Points, Coords, ShProp);

            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillPolygon(pe, br, Points, Clipping);
                }
            }

            real ArrowWidth = Coords.Width - rx;

            RectangleF Result = new RectangleF(Coords.Left, Coords.Top, Coords.Width - ArrowWidth / 2 , Coords.Height);
            Flip(ref Result, Coords, ShProp);
            return Result;
        }	

        #endregion

        #region Stars
        internal static RectangleF DrawN4Star(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100, int n, int rDef)
        {
            real r0 = (ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjustValue, rDef)) / 21600f;           

            real x0 = Coords.Left + Coords.Width / 2;
            real y0 = Coords.Top + Coords.Height / 2;
            real rx = Coords.Width / 2;
            real ry = Coords.Height / 2;

            real rx1 = (0.5f-r0) * Coords.Width;
            real ry1 = (0.5f-r0) * Coords.Height;

            TPointF[] Ray = new TPointF[n * 8];
            for (int i = 0; i < n * 4; i++)
            {
                real Alpha = (real)(Math.PI / 2 / n * i);
                Ray[i*2] = new TPointF((real)(x0 + rx * Math.Cos(Alpha)), (real)(y0 + ry * Math.Sin(Alpha)));
                
                Alpha += (real)(Math.PI / 4 / n);
                Ray[i*2 + 1] = new TPointF((real)(x0 + rx1 * Math.Cos(Alpha)), (real)(y0 + ry1 * Math.Sin(Alpha)));
            }


            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillPolygon(pe, br, Ray, Clipping);
                }
            }


            real rt =  1/6f +  2*r0 / 3;
            return new RectangleF(Coords.X + Coords.Width * rt, Coords.Y + Coords.Height * rt, Coords.Width * (1 - 2* rt), Coords.Height * (1 - 2 * rt));
        }
    
        internal static RectangleF Draw5Star(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {

            TPointF[] Point = new TPointF[10];
            real rx = Coords.Width;
            real ry = Coords.Height;
            real x0 = Coords.Left + rx / 2;

            Point[0] = new TPointF(x0, Coords.Top);
            Point[1] = new TPointF(Coords.Right - rx * 5 / 13f, Coords.Top + ry * 3 / 8f);
            Point[2] = new TPointF(Coords.Right, Coords.Top + ry * 3 / 8f);
            Point[3] = new TPointF(Coords.Left + rx * 9/13f, Coords.Top + ry * 5/8f);
            Point[4] = new TPointF(Coords.Left + rx * 25/31f, Coords.Bottom);
            Point[5] = new TPointF(x0, Coords.Top + ry * 14/17f);
            Point[6] = new TPointF(Coords.Left + rx * 6/31f, Coords.Bottom);
            Point[7] = new TPointF(Coords.Left + rx * 4/13f, Coords.Top + ry * 5/8f);
            Point[8] = new TPointF(Coords.Left, Coords.Top + ry * 3 / 8f);
            Point[9] = new TPointF(Coords.Left + rx * 5 / 13f, Coords.Top + ry * 3 / 8f);

            Flip(ref Point, Coords, ShProp);

            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillPolygon(pe, br, Point, Clipping);
                }
            }


            real rxt =  4 / 13f;
            real ryt =  3 / 8f;
            return new RectangleF(Coords.X + Coords.Width * rxt, Coords.Y + Coords.Height * ryt, Coords.Width * (1 - 2* rxt), Coords.Height * 0.4f);
        }
    

        internal static RectangleF DrawExplosion1(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {

            TPointF[] Point = new TPointF[24];
            real rx = Coords.Width / 44f;
            real ry = Coords.Height / 38f;
            real x0 = Coords.Left;
            real y0 = Coords.Top;

            Point[0] = new TPointF(x0 + 0.9f * rx, y0 + 4f * ry);
            Point[1] = new TPointF(x0 + 15f * rx, y0 + 11f * ry);
            Point[2] = new TPointF(x0 + 17f * rx, y0 + 4f * ry);
            Point[3] = new TPointF(x0 + 22f * rx, y0 + 10f * ry);
            Point[4] = new TPointF(x0 + 29.5f * rx, y0 + 0f * ry);
            Point[5] = new TPointF(x0 + 28.9f * rx, y0 + 9.4f * ry);
            Point[6] = new TPointF(x0 + 37.5f * rx, y0 + 7.8f * ry);
            Point[7] = new TPointF(x0 + 34f * rx, y0 + 12.9f * ry);
            Point[8] = new TPointF(x0 + 43f * rx, y0 + 14.2f * ry);
            Point[9] = new TPointF(x0 + 36f * rx, y0 + 18.5f * ry);
            Point[10] = new TPointF(x0 + 44f * rx, y0 + 23.2f * ry);
            Point[11] = new TPointF(x0 + 34.4f * rx, y0 + 22.6f * ry);
            Point[12] = new TPointF(x0 + 37f * rx, y0 + 31.8f * ry);
            Point[13] = new TPointF(x0 + 28.5f * rx, y0 + 25.4f * ry);
            Point[14] = new TPointF(x0 + 27f * rx, y0 + 34.5f * ry);
            Point[15] = new TPointF(x0 + 21.4f * rx, y0 + 26.2f * ry);
            Point[16] = new TPointF(x0 + 17.2f * rx, y0 + 38f * ry);
            Point[17] = new TPointF(x0 + 15.8f * rx, y0 + 27.5f * ry);
            Point[18] = new TPointF(x0 + 9.6f * rx, y0 + 31f * ry);
            Point[19] = new TPointF(x0 +  11.5f * rx, y0 + 24.5f * ry);
            Point[20] = new TPointF(x0 + 0.2f * rx, y0 + 25.5f * ry);
            Point[21] = new TPointF(x0 + 7.5f * rx, y0 + 20.5f * ry);
            Point[22] = new TPointF(x0 + 0f * rx, y0 + 15f * ry);
            Point[23] = new TPointF(x0 + 9.4f * rx, y0 + 13.4f * ry);

            Flip(ref Point, Coords, ShProp);

            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillPolygon(pe, br, Point, Clipping);
                }
            }


            return FlexCelRender.RectangleXY(Point[23].X, Point[1].Y, Point[7].X, Point[19].Y);
        }
    
        internal static RectangleF DrawExplosion2(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {

            TPointF[] Point = new TPointF[28];
            real rx = Coords.Width / 44f;
            real ry = Coords.Height / 38f;
            real x0 = Coords.Left;
            real y0 = Coords.Top;

            Point[0] = new TPointF(x0,  y0 + 22.5f * ry);
            Point[1] = new TPointF(x0 + 8f * rx, y0 + 20.2f * ry);
            Point[2] = new TPointF(x0 + 2.4f * rx, y0 + 14.4f * ry);
            Point[3] = new TPointF(x0 + 11f * rx, y0 + 13.8f * ry);
            Point[4] = new TPointF(x0 + 9.1f * rx, y0 + 6.2f * ry);
            Point[5] = new TPointF(x0 + 17.3f * rx, y0 + 11.2f * ry);
            Point[6] = new TPointF(x0 + 19.8f * rx, y0 + 3.2f * ry);
            Point[7] = new TPointF(x0 + 23.4f * rx, y0 + 7.5f * ry);
            Point[8] = new TPointF(x0 + 30f * rx, y0 + 0f * ry);
            Point[9] = new TPointF(x0 + 29.7f * rx, y0 + 10f * ry);
            Point[10] = new TPointF(x0 + 36.6f * rx, y0 + 5.5f * ry);
            Point[11] = new TPointF(x0 + 33.5f * rx, y0 + 11.4f * ry);
            Point[12] = new TPointF(x0 + 44f * rx, y0 + 11.6f * ry);
            Point[13] = new TPointF(x0 + 34.5f * rx, y0 + 16.4f * ry);
            Point[14] = new TPointF(x0 + 37.2f * rx, y0 + 19.8f * ry);
            Point[15] = new TPointF(x0 + 33.5f * rx, y0 + 21.5f * ry);
            Point[16] = new TPointF(x0 + 38.5f * rx, y0 + 27.5f * ry);
            Point[17] = new TPointF(x0 + 29.9f * rx, y0 + 25f * ry);
            Point[18] = new TPointF(x0 + 30.5f * rx, y0 + 30.5f * ry);
            Point[19] = new TPointF(x0 + 24.9f * rx, y0 + 28f * ry);
            Point[20] = new TPointF(x0 + 23.6f * rx, y0 + 33f * ry);
            Point[21] = new TPointF(x0 + 20f * rx, y0 + 30.5f * ry);
            Point[22] = new TPointF(x0 + 17.6f * rx, y0 + 34.6f * ry);
            Point[23] = new TPointF(x0 + 15.2f * rx, y0 + 31.8f * ry);
            Point[24] = new TPointF(x0 + 10f * rx, y0 + 38f * ry);
            Point[25] = new TPointF(x0 + 9.8f * rx, y0 + 32f * ry);
            Point[26] = new TPointF(x0 + 1.5f * rx, y0 + 31.2f * ry);
            Point[27] = new TPointF(x0 + 6.6f * rx, y0 + 27f * ry);

            Flip(ref Point, Coords, ShProp);

            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillPolygon(pe, br, Point, Clipping);
                }
            }


            return FlexCelRender.RectangleXY(Point[3].X, Point[5].Y, Point[17].X, Point[19].Y);
        }
    

        #endregion

        #region FlowCharts
        internal static RectangleF DrawTerminator(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            real rx = Coords.Width / 8f;
            real ry = Coords.Height/ 2f;
            TPointF[] Points = new TPointF[25];

            Points[0] = new TPointF(Coords.X + rx, Coords.Y);
            AddLineAndCorner(Points, 1, rx, ry, 1, Coords.Width - 2*rx, 0);
            AddLineAndCorner(Points, 7, rx, ry, 2, 0, Coords.Height - 2*ry);
            AddLineAndCorner(Points, 13, rx, ry, 3, - (Coords.Width - 2*rx), 0);
            AddLineAndCorner(Points, 19, rx, ry, 0, 0, -(Coords.Height - 2*ry));


            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillBeziers(pe, br, Points, Clipping);
                }
            }

            return new RectangleF(Coords.X + rx / 3f, Coords.Y + ry / 3f, Coords.Width - 2*rx/3f, Coords.Height - 2*ry/3f);
        }	

        internal static RectangleF DrawManualInput(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            TPointF[] Points = new TPointF[4];

            Points[0] = new TPointF(Coords.X, Coords.Y + Coords.Height / 5f);
            Points[1] = new TPointF(Coords.Right, Coords.Y);
            Points[2] = new TPointF(Coords.Right, Coords.Bottom);
            Points[3] = new TPointF(Coords.Left, Coords.Bottom);

            Flip(ref Points, Coords, ShProp);

            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillPolygon(pe, br, Points, Clipping);
                }
            }

            RectangleF Result = new RectangleF(Coords.X, Coords.Y + Coords.Height / 5f, Coords.Width, Coords.Height*4/5f );
            Flip(ref Result, Coords, ShProp);
            return Result;
        }	

        internal static RectangleF DrawCard(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            TPointF[] Points = new TPointF[5];

            Points[0] = new TPointF(Coords.X, Coords.Y + Coords.Height / 5f);
            Points[1] = new TPointF(Coords.Left + Coords.Width / 5f, Coords.Y);
            Points[2] = new TPointF(Coords.Right, Coords.Top);
            Points[3] = new TPointF(Coords.Right, Coords.Bottom);
            Points[4] = new TPointF(Coords.Left, Coords.Bottom);

            Flip(ref Points, Coords, ShProp);

            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillPolygon(pe, br, Points, Clipping);
                }
            }

            RectangleF Result = new RectangleF(Coords.X, Coords.Y + Coords.Height / 5f, Coords.Width, Coords.Height*4/5f );
            Flip(ref Result, Coords, ShProp);
            return Result;
        }	

        internal static RectangleF DrawSummingJunction(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            real rx = Coords.Width / 2f;
            real ry = Coords.Height / 2f;
            TPointF[] Points1 = GetOval(Coords.X, Coords.Y, rx, ry);

            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillBeziers(pe, br, Points1, Clipping);
                    if (DrawLines(ShadowInfo, br != null) && pe != null) 
                    {
                        real dx = (real)( rx * Math.Cos(Math.PI / 4) );
                        dx = rx - dx;
                        real dy = (real)( ry * Math.Sin(Math.PI / 4) );
                        dy = ry - dy;
                        Canvas.DrawLine(pe, Coords.Left + dx, Coords.Top + dy, Coords.Right - dx, Coords.Bottom - dy); 
                        Canvas.DrawLine(pe, Coords.Right - dx, Coords.Top + dy, Coords.Left + dx, Coords.Bottom - dy); 
                    }
                }
            }

            return new RectangleF(Coords.X + rx / 3f, Coords.Y + ry / 3f, Coords.Width - 2*rx/3f, Coords.Height - 2*ry/3f);
        }

        internal static RectangleF DrawOr(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            real rx = Coords.Width / 2f;
            real ry = Coords.Height / 2f;
            TPointF[] Points1 = GetOval(Coords.X, Coords.Y, rx, ry);

            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillBeziers(pe, br, Points1, Clipping);
                    if (DrawLines(ShadowInfo, br != null) && pe != null) 
                    {
                        Canvas.DrawLine(pe, Coords.Left + rx, Coords.Top, Coords.Left + rx, Coords.Bottom); 
                        Canvas.DrawLine(pe, Coords.Left, Coords.Top + ry, Coords.Right, Coords.Top + ry); 
                    }
                }
            }

            return new RectangleF(Coords.X + rx / 3f, Coords.Y + ry / 3f, Coords.Width - 2*rx/3f, Coords.Height - 2*ry/3f);
        }


        internal static RectangleF DrawCollate(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {

            TPointF[] Points = new TPointF[4];
            Points[0] = new TPointF(Coords.Left, Coords.Top);
            Points[1] = new TPointF(Coords.Right, Coords.Bottom);
            Points[2] = new TPointF(Coords.Left, Coords.Bottom);
            Points[3] = new TPointF(Coords.Right, Coords.Top);

            Flip(ref Points, Coords, ShProp);

            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillPolygon(pe, br, Points, Clipping);
                }
            }

            RectangleF Result = new RectangleF(Coords.Left + Coords.Width / 4f, Coords.Top + Coords.Height / 4f, Coords.Width /2f, Coords.Height/2f);
            Flip(ref Result, Coords, ShProp);
            return Result;
        }	

        internal static RectangleF DrawSort(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {

            TPointF[] Points = new TPointF[4];
            real rx = Coords.Width / 2;
            real ry = Coords.Height / 2;
            Points[0] = new TPointF(Coords.Left + rx, Coords.Top);
            Points[1] = new TPointF(Coords.Right, Coords.Top+ry);
            Points[2] = new TPointF(Coords.Left + rx, Coords.Bottom);
            Points[3] = new TPointF(Coords.Left, Coords.Top + ry);

            Flip(ref Points, Coords, ShProp);

            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillPolygon(pe, br, Points, Clipping);
                    if (DrawLines(ShadowInfo, br != null) && pe != null) 
                    {
                        Canvas.DrawLine(pe, Coords.Left, Coords.Top + ry, Coords.Right, Coords.Top + ry); 
                    }
                }
            }

            RectangleF Result = new RectangleF(Coords.Left + Coords.Width / 4f, Coords.Top + Coords.Height / 4f, Coords.Width /2f, Coords.Height/2f);
            Flip(ref Result, Coords, ShProp);
            return Result;
        }	

        internal static RectangleF DrawMerge(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            real r = Coords.Width * 0.5f;
            DrawTriangle(Canvas, Workbook, ShProp, Coords, r, ShadowInfo, Clipping, true, Zoom100);
            RectangleF Result = new RectangleF(Coords.Left + Coords.Width / 3.65f, Coords.Top, Coords.Width * (1 - 2/3.65f), Coords.Height *1f / 3f);
            Flip(ref Result, Coords, ShProp);
            return Result;
        }	

        internal static RectangleF DrawStoredData(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            real rx = Coords.Width / 6f;
            real ry = Coords.Height/ 2f;
            TPointF[] Points = new TPointF[25];

            Points[0] = new TPointF(Coords.X + rx, Coords.Y);
            AddLineAndCorner(Points, 1, rx, ry, 1, Coords.Width - rx, 0);
            AddLineAndCorner(Points, 7, rx, ry, 2, 0, Coords.Height - 2*ry);
            AddLineAndCorner(Points, 13, rx, ry, 3, - (Coords.Width - rx), 0);
            AddLineAndCorner(Points, 19, rx, ry, 0, 0, -(Coords.Height - 2*ry));

            real x0 = Points[3].X;
            for (int i = 4; i < 14; i++) //Invert the draw direction.
            {
                Points[i].X = 2* x0 - Points[i].X;
            }


            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillBeziers(pe, br, Points, Clipping);
                }
            }

            return new RectangleF(Coords.X + rx / 3f, Coords.Y + ry / 3f, Coords.Width - 4*rx/3f, Coords.Height - 2*ry/3f);
        }	

        internal static RectangleF DrawDelay(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            real rx = Coords.Width / 2f;
            real ry = Coords.Height/ 2f;
            TPointF[] Points = new TPointF[19];

            Points[0] = new TPointF(Coords.X, Coords.Y);
            AddLineAndCorner(Points, 1, rx, ry, 1, Coords.Width - rx, 0);
            AddLineAndCorner(Points, 7, rx, ry, 2, 0, Coords.Height - 2*ry);
            Points[13] = Points[12];
            Points[14] = new TPointF(Coords.Left, Coords.Bottom);
            Points[15] = Points[14];

            Points[16] = Points[15];
            Points[17] = Points[0];
            Points[18] = Points[0];

            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillBeziers(pe, br, Points, Clipping);
                }
            }

            return new RectangleF(Coords.X, Coords.Y + ry / 3f, Coords.Width - rx/3f, Coords.Height - 2*ry/3f);
        }	


        internal static RectangleF DrawDisplay(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            real rx = Coords.Width / 6f;
            real ry = Coords.Height/ 2f;
            real rx2 = Coords.Width / 6;
            
            TPointF[] Points = new TPointF[22];

            Points[0] = new TPointF(Coords.X + rx2, Coords.Y);
            AddLineAndCorner(Points, 1, rx, ry, 1, Coords.Width - rx - rx2, 0);
            AddLineAndCorner(Points, 7, rx, ry, 2, 0, Coords.Height - 2*ry);
            Points[13] = Points[12];
            Points[14] = new TPointF(Coords.Left + rx2, Coords.Bottom);
            Points[15] = Points[14];

            Points[16] = Points[15];
            Points[17] = new TPointF(Coords.Left, Coords.Top + Coords.Height/ 2f);
            Points[18] = Points[17];
            
            Points[19] = Points[17];
            Points[20] = Points[0];
            Points[21] = Points[0];

            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillBeziers(pe, br, Points, Clipping);
                }
            }

            return new RectangleF(Coords.X + rx2, Coords.Y, Coords.Width - rx - rx2, Coords.Height);
        }	

        internal static RectangleF DrawOffPageConnector(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            TPointF[] Points = new TPointF[5];

            Points[0] = new TPointF(Coords.X, Coords.Y);
            Points[1] = new TPointF(Coords.Right, Coords.Y);
            Points[2] = new TPointF(Coords.Right, Coords.Bottom  - Coords.Height / 5f);
            Points[3] = new TPointF(Coords.Left + Coords.Width/2, Coords.Bottom);
            Points[4] = new TPointF(Coords.Left, Coords.Bottom - Coords.Height / 5f);

            Flip(ref Points, Coords, ShProp);

            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Canvas.DrawAndFillPolygon(pe, br, Points, Clipping);
                }
            }

            RectangleF Result = new RectangleF(Coords.X, Coords.Y, Coords.Width, Coords.Height*4/5f );
            Flip(ref Result, Coords, ShProp);
            return Result;
        }	

        internal static RectangleF DrawPredefinedProcess(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            bool HasBrush = false;
            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                if (br != null) 
                {
                    Canvas.FillRectangle(br, Coords, Clipping);
                    HasBrush = true;
                }
            }

            using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
            {
                if (pe != null && Clipping == TClippingStyle.None && DrawLines(ShadowInfo, HasBrush))
                {
                    Canvas.DrawRectangle(pe, Coords.X, Coords.Y, Coords.Width, Coords.Height);
                    Canvas.DrawLine(pe, Coords.X + Coords.Width * 5f /41f, Coords.Top, Coords.Left +  Coords.Width *  5f/41f, Coords.Bottom);
                    Canvas.DrawLine(pe, Coords.Right - Coords.Width * 5f /41f, Coords.Top, Coords.Right -  Coords.Width * 5f/41f, Coords.Bottom);
                }
            }

            RectangleF Result = new RectangleF(Coords.X+ Coords.Width * 5f /41f, Coords.Y, Coords.Width * (1 - 10f/41f), Coords.Height);
            Flip(ref Result, Coords, ShProp);
            return Result;
            
        }	


        internal static RectangleF DrawInternalStorage(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            bool HasBrush = false;
            using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
            {
                if (br != null) 
                {
                    Canvas.FillRectangle(br, Coords, Clipping);
                    HasBrush = true;
                }
            }

            using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
            {
                if (pe != null && Clipping == TClippingStyle.None && DrawLines(ShadowInfo, HasBrush))
                {
                    Canvas.DrawRectangle(pe, Coords.X, Coords.Y, Coords.Width, Coords.Height);
                    Canvas.DrawLine(pe, Coords.X + Coords.Width / 5f, Coords.Top, Coords.Left +  Coords.Width /5f, Coords.Bottom);
                    Canvas.DrawLine(pe, Coords.Left, Coords.Top + Coords.Height / 5f, Coords.Right, Coords.Top + Coords.Height / 5f);
                }
            }

            RectangleF Result = new RectangleF(Coords.X+ Coords.Width / 5f, Coords.Y + Coords.Height / 5f, Coords.Width * 4f/5f, Coords.Height*4f/5f);
            Flip(ref Result, Coords, ShProp);
            return Result;
            
        }	
        
        #endregion

		#region Custom Shapes
		private static TPathInfo[] GetBezierPoints(TPointF[] Points, TShapeProperties ShProp)
		{
			byte[] SegmentInfo = ShProp.ShapeOptions[TShapeOption.pSegmentInfo] as byte[];
			if (SegmentInfo == null) return null;

			int SiPos = 6;
			int i = 0;

            List<TPathInfo> Result = new List<TPathInfo>();

            while (SiPos < SegmentInfo.Length)
            {
                TPathInfo PathInfo = GetOnePath(Points, SegmentInfo, ref SiPos, ref i);
                if (PathInfo.Points != null && PathInfo.Points.Length > 1)
                {
                    Result.Add(PathInfo);
                }
                SiPos += 2;
            }

            return Result.ToArray();
		}

        private static TPathInfo GetOnePath(TPointF[] Points, byte[] SegmentInfo, ref int SiPos, ref int i)
        {
            TPathInfo Result = new TPathInfo(false, true, true);
            List<TPointF> ResultPoints = new List<TPointF>();
            bool AtEnd = false;
            int StartOfPath = i;
            while (SiPos < SegmentInfo.Length && !AtEnd)
            {
                byte Info0 = SegmentInfo[SiPos];
                byte Info1 = SegmentInfo[SiPos + 1];
                TSegmentType SegmentType = (TSegmentType)(Info1 >> 5);
                int SegmentCount = ((Info1 & 0x1F) << 8) | Info0;

                if (SegmentType == TSegmentType.Escape)
                {
                    ProcessEscape(Points, ResultPoints, Info0, Info1, ref i, StartOfPath, ref Result);
                }
                else
                {
                    if (!ProcessNormalSegments(Points, ref i, ResultPoints, SegmentType, SegmentCount, ref Result, out AtEnd)) return new TPathInfo(false, false, false);
                }

                if (!AtEnd) SiPos += 2;
            }
            
            Result.Points = ResultPoints.ToArray();
            return Result;
        }

        private static bool ProcessNormalSegments(TPointF[] Points, ref int i, List<TPointF> ResultPoints, TSegmentType SegmentType, int SegmentCount, ref TPathInfo PathInfo, out bool AtEnd)
        {
            AtEnd = false;
            if (SegmentCount == 0) SegmentCount = 1;

            for (int seg = 0; seg < SegmentCount; seg++)
            {
                switch (SegmentType)
                {
                    case TSegmentType.LineTo:
                        if (i < 1 || i >= Points.Length) return false;
                        if (ResultPoints.Count < 1) return false;
                        ResultPoints.Add(ResultPoints[ResultPoints.Count - 1]);
                        ResultPoints.Add(Points[i]);
                        ResultPoints.Add(Points[i]);
                        i++;
                        break;

                    case TSegmentType.CurveTo:
                        if (i < 1 || i + 2 >= Points.Length) return false;
                        ResultPoints.Add(Points[i]);
                        ResultPoints.Add(Points[i + 1]);
                        ResultPoints.Add(Points[i + 2]);
                        i += 3;
                        break;

                    case TSegmentType.MoveTo:
                        if (i >= Points.Length) return false;
                        ResultPoints.Add(Points[i]);
                        i++;
                        break;

                    case TSegmentType.Close:
                        PathInfo.Close = true;
                        break;

                    case TSegmentType.EndPath:
                        AtEnd = true;
                        break;

                    case TSegmentType.Escape: //Not needed for rendering.
                        break;

                    default:
                        return false;

                }
            }
            return true;
        }

        private static bool ProcessEscape(TPointF[] Points, List<TPointF> ResultPoints, byte Info0, byte Info1, ref int i, int StartOfPath, ref TPathInfo PathInfo)
        {
            TSegmentTypeEscaped SegmentTypeEsc = (TSegmentTypeEscaped)(Info1 & 0x1F);
            int VertexCount = Info0;

            switch (SegmentTypeEsc)
            {
                case TSegmentTypeEscaped.msopathEscapeExtension:
                    i += VertexCount;
                    break;

                case TSegmentTypeEscaped.msopathEscapeAngleEllipseTo:
                    i += VertexCount;
                    break;

                case TSegmentTypeEscaped.msopathEscapeAngleEllipse:
                    i += VertexCount;
                    break;

                case TSegmentTypeEscaped.msopathEscapeArcTo:
                    for (int k = 0; k < VertexCount / 4; k++)
                    {
                        if (i < 0 || i + 3 >= Points.Length) return false;
                        AddArc(Points, ResultPoints, i, false);
                        i += 4;
                    }
                    break;
                
                case TSegmentTypeEscaped.msopathEscapeArc:
                    for (int k = 0; k < VertexCount / 4; k++)
                    {
                        if (i < 0 || i + 3 >= Points.Length) return false;
                        AddArc(Points, ResultPoints, i, false);
                        i += 4;
                    }
                    break;

                case TSegmentTypeEscaped.msopathEscapeClockwiseArcTo:
                    for (int k = 0; k < VertexCount / 4; k++)
                    {
                        if (i < 0 || i + 3 >= Points.Length) return false;
                        AddArc(Points, ResultPoints, i, true);
                        i += 4;
                    }
                    break;
                
                case TSegmentTypeEscaped.msopathEscapeClockwiseArc:
                    for (int k = 0; k < VertexCount / 4; k++)
                    {
                        if (i < 0 || i + 3 >= Points.Length) return false;
                        AddArc(Points, ResultPoints, i, true);
                        i += 4;
                    }
                    break;
                
                case TSegmentTypeEscaped.msopathEscapeEllipticalQuadrantX:
                    i += VertexCount;
                    break;
                
                case TSegmentTypeEscaped.msopathEscapeEllipticalQuadrantY:
                    i += VertexCount;
                    break;
                
                case TSegmentTypeEscaped.msopathEscapeQuadraticBezier:
                    for (int k = 0; k < VertexCount; k++) ResultPoints.Add(Points[k]);
                    i += VertexCount;
                    break;

                case TSegmentTypeEscaped.msopathEscapeNoFill:
                    PathInfo.FillPath = false;
                    break;

                case TSegmentTypeEscaped.msopathEscapeNoLine:
                    PathInfo.DrawPath = false;
                    break;
                
                default:
                    i += VertexCount;
                    break;
            }

            return true;
        }

        private static void AddArc(TPointF[] Points, List<TPointF> Result, int i, bool ClockWise)
        {
            double cx = (Points[i].X + Points[i + 1].X) / 2;
            double cy = (Points[i].Y + Points[i + 1].Y) / 2;
            double a = Math.Abs(Points[i].X - cx);
            double b = Math.Abs(Points[i].Y - cy);
            double l1 = Math.Atan2(Points[i + 2].Y - cy, Points[i + 2].X - cx);
            double l2 = Math.Atan2(Points[i + 3].Y - cy, Points[i + 3].X - cx);

            double ls1 = ClockWise ? l1 : l2;
            double ls2 = ClockWise ? l2 : l1;
            TPointF[] EPoints = TEllipticalArc.GetPoints(cx, cy, a, b, 0, ls1, ls2);

            if (!ClockWise) Array.Reverse(EPoints);
            if (Result.Count > 0)
            {
                Result.Add(Result[Result.Count - 1]);
                Result.Add(Result[Result.Count - 1]);
            }
            else
            {
                Result.Add(EPoints[0]);
                Result.Add(EPoints[0]);
            }
            Result.AddRange(EPoints);
        }

		internal static RectangleF DrawCustomShape(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
		{
            real geox1 = ShProp.ShapeOptions.AsSignedLong(TShapeOption.geoLeft, 0);
            real geox2 = ShProp.ShapeOptions.AsSignedLong(TShapeOption.geoRight, 21600);
            real geoy1 = ShProp.ShapeOptions.AsSignedLong(TShapeOption.geoTop, 0);
            real geoy2 = ShProp.ShapeOptions.AsSignedLong(TShapeOption.geoBottom, 21600);
			real geox = geox2 - geox1;
			real geoy = geoy2 - geoy1;

			byte[] Vertices = ShProp.ShapeOptions[TShapeOption.pVertices] as byte[];

			if (Vertices == null || geox == 0 || geoy == 0 || Vertices.Length <= 6) return Coords;

            int bSize = BitConverter.ToUInt16(Vertices, 4);

            byte[] Guide = (byte[])ShProp.ShapeOptions[TShapeOption.pGuides];

            TPointF[] Points = new TPointF[BitConverter.ToInt16(Vertices, 0)];
			for (int i = 0; i < Points.Length; i++)
			{
                Point p = GetMsoArray(Vertices, i, bSize); 
				Points[i] = new TPointF(Coords.Left + Coords.Width  * GetGuide(p.X, Guide, 0, ShProp, Coords, geox, geoy)/geox,
                    Coords.Top + Coords.Height * GetGuide(p.Y, Guide, 0, ShProp, Coords, geox, geoy) / geoy);
			}

			Flip(ref Points, Coords, ShProp);

			using (Brush br = GetBrush(Coords, ShProp, Workbook, ShadowInfo, Zoom100))
			{
				using (Pen pe = GetPen(ShProp, Workbook, ShadowInfo))
				{
					switch ((TShapePathBiff8)ShProp.ShapeOptions.AsLong(TShapeOption.shapePath, 4))
					{
						case TShapePathBiff8.Lines:
							if (Clipping == TClippingStyle.None) Canvas.DrawLines(pe, Points);
							if (br != null)Canvas.DrawAndFillPolygon(null, br, Points, Clipping);
							if (pe != null)
							{
								DoArrows(Canvas, pe, ShProp, Points, Clipping);
							}
							break;
						case TShapePathBiff8.LinesClosed:
							Canvas.DrawAndFillPolygon(pe, br, Points, Clipping);
							break;
						case TShapePathBiff8.Curves:
							Canvas.DrawAndFillBeziers(pe, null, Points, Clipping);
							if (br != null)Canvas.DrawAndFillBeziers(null, br, Points, Clipping);
							if (pe != null)
							{
								DoArrows(Canvas, pe, ShProp, Points, Clipping);
							}
							break;
						case TShapePathBiff8.CurvesClosed:
							Canvas.DrawAndFillBeziers(pe, br, Points, Clipping);
							break;
						case TShapePathBiff8.Complex:
							TPathInfo[] Paths = GetBezierPoints(Points, ShProp);
                            if (Paths == null) return Coords;

                            for (int i = 0; i < Paths.Length; i++)
                            {
                                //if (i!=3) continue;
                                TPathInfo PathInfo = Paths[i];
                                Pen pe1 = PathInfo.DrawPath? pe: null;
                                Brush br1 = PathInfo.FillPath ? br : null;

                                if (PathInfo.Points == null) return Coords;
                                Canvas.DrawAndFillBeziers(pe1, br1, PathInfo.Points, Clipping);
                                if (pe1 != null && !PathInfo.Close)
                                {
                                    DoArrows(Canvas, pe, ShProp, Points, Clipping);
                                }
                            }
							break;
					}
				}
			}
			return Coords;
		}

        private static float GetGuide(int x, byte[] Guide, int Level, TShapeProperties ShProp, RectangleF Coords, real geox, real geoy)
        {
            if (Level > 128) return 0; //avoid infinite recursion. There can't be more than 128 SG entries.
            uint ux;
            unchecked
            {
                ux = (UInt32)x;
            }

            if (ux < 0x80000000 || ux > 0x8000007F) return x;

            ux -= 0x80000000;

            if (Guide == null) return 0;

            int start = (int)(6 + ux * 8);
            int first = BitConverter.ToUInt16(Guide, start);
            TSgFormula sgf = (TSgFormula) (first & 0x1FFF);
            bool fCalculatedParam1 = ((first >> 13) & 1) != 0;
            bool fCalculatedParam2 = ((first >> 14) & 1) != 0;
            bool fCalculatedParam3 = ((first >> 15) & 1) != 0;

            float fParam1 = CalcParam(BitConverter.ToUInt16(Guide, start + 2), fCalculatedParam1, Guide, Level, ShProp, Coords, geox, geoy);
            float fParam2 = CalcParam(BitConverter.ToUInt16(Guide, start + 4), fCalculatedParam2, Guide, Level, ShProp, Coords, geox, geoy);
            float fParam3 = CalcParam(BitConverter.ToUInt16(Guide, start + 6), fCalculatedParam3, Guide, Level, ShProp, Coords, geox, geoy);

            switch (sgf)
            {
                case TSgFormula.sgfSum:
                    return fParam1 + fParam2 - fParam3;

                case TSgFormula.sgfProduct:
                    if (fParam3 == 0) return 0;
                    return (float)((double)fParam1 * (double)fParam2 / (double)fParam3);

                case TSgFormula.sgfMid:
                    return (float)((double)fParam1 + (double)fParam2 / 2.0);

                case TSgFormula.sgfAbsolute:
                    return Math.Abs(fParam1);

                case TSgFormula.sgfMin:
                    return Math.Min(fParam1, fParam2);

                case TSgFormula.sgfMax:
                    return Math.Max(fParam1, fParam2);

                case TSgFormula.sgfIf:
                        return fParam1 > 0 ? fParam2 : fParam3;

                case TSgFormula.sgfMod:
                    return (float)Math.Sqrt(Math.Pow(fParam1, 2) + Math.Pow(fParam2, 2) + Math.Pow(fParam3, 2));

                case TSgFormula.sgfATan2:
                    return (float)Math.Atan2(Angle(fParam2), Angle(fParam1));

                case TSgFormula.sgfSin:
                    return (float) (fParam1 * Math.Sin(Angle(fParam2)));

                case TSgFormula.sgfCos:
                    return (float)(fParam1 * Math.Cos(Angle(fParam2)));

                case TSgFormula.sgfCosATan2:
                    return (float)  (fParam1 * Math.Cos(Math.Atan2(Angle(fParam3), Angle(fParam2))));

                case TSgFormula.sgfSinATan2:
                    return (float)(fParam1 * Math.Sin(Math.Atan2(Angle(fParam3), Angle(fParam2))));

                case TSgFormula.sgfSqrt:
                    return (float)Math.Sqrt(fParam1);

                case TSgFormula.sgfSumAngle:
                    return (float)(fParam1 + fParam2 * 65536.0 - fParam3 * 65536.0);

                case TSgFormula.sgfEllipse:
                    return (float)(fParam3 * Math.Sqrt(1 - Math.Pow(((double)fParam1 / (double)fParam2), 2)));
                
                case TSgFormula.sgfTan:
                    return (float)(fParam1 * Math.Tan(Angle(fParam2)));
            }

            return 0;

        }

        private static double Angle(float fParam)
        {
            return TShapeOptionList.Get1616((long)fParam) * Math.PI / 180;
        }

        private static float CalcParam(ushort p, bool fCalculatedParam, byte[] Guide, int Level, TShapeProperties ShProp, RectangleF Coords, real geox, real geoy)
        {
            if (!fCalculatedParam)
            {
                return p;
            }

            float a = 0;
            switch (p)
            {
                // X coordinate of the center of the geometry space of this shape. 
                case 0x0140: return geox / 2;

                // Y coordinate of the center of the geometry space of this shape. 
                case 0x0141: return geoy / 2;

                // Width of the geometry space of this shape. 
                case 0x0142: return geox;

                // Height of the geometry space of this shape. 
                case 0x0143: return geoy;

                // The value of this shape's adjustValue property. 
                case 0x0147: return ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjustValue, 0x0);

                // The value of this shape's adjust2Value property. 
                case 0x0148: return ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjust2Value, 0x0);

                // The value of this shape's adjust3Value property. 
                case 0x0149: return ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjust3Value, 0x0);

                // The value of this shape's adjust4Value property. 
                case 0x014A: return ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjust4Value, 0x0);

                // The value of this shape's adjust5Value property. 
                case 0x014B: return ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjust5Value, 0x0);

                // The value of this shape's adjust6Value property. 
                case 0x014C: return ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjust6Value, 0x0);

                // The value of this shape's adjust7Value property. 
                case 0x014D: return ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjust7Value, 0x0);

                // The value of this shape's adjust8Value property. 
                case 0x014E: return ShProp.ShapeOptions.AsSignedLong(TShapeOption.adjust8Value, 0x0);

                // The value of this shape's xLimo property. 
                case 0x0153: return ShProp.ShapeOptions.AsSignedLong(TShapeOption.xLimo, 0x80000000);

                // The value of this shape's yLimo property. 
                case 0x0154: return ShProp.ShapeOptions.AsSignedLong(TShapeOption.yLimo, 0x80000000);

                // The value of the fLine bit from this shape's Line Style Boolean Properties. 
                case 0x01FC: return a;

                // The width of a line in this shape in pixels. 
                case 0x04F7: return a;

                // The width of this shape in pixels.             
                case 0x04F8: return (real)( Coords.Width / FlxConsts.PixToPoints);
            }

            if (p >= 0x0400 && p <= 0x047F)
            {
                UInt32 lp;
                unchecked
                {
                    lp = (UInt32) (0x80000000 + (p - 0x400));

                    return GetGuide((Int32)lp, Guide, Level + 1, ShProp, Coords, geox, geoy);
                }
            }

            return 0;
        }

        private static Point GetMsoArray(byte[] Vertices, int i, int bSize)
        {
            unchecked
            {
                switch (bSize)
                {
                    case 2: return new Point((sbyte)Vertices[6 + i * 2], (sbyte)Vertices[6 + i * 2 + 1]);
                    case 0xFFF0: return new Point(BitConverter.ToUInt16(Vertices, 6 + i * 4), BitConverter.ToUInt16(Vertices, 6 + i * 4 + 2)); //this is truncated 8 bytes, so values are unsigned
                    case 4: return new Point(BitConverter.ToInt16(Vertices, 6 + i * 4), BitConverter.ToInt16(Vertices, 6 + i * 4 + 2));
                    case 8: return new Point(BitConverter.ToInt32(Vertices, 6 + i * 8), BitConverter.ToInt32(Vertices, 6 + i * 8 + 4));
                }
            }

            return new Point(0, 0);
        }

		#endregion
    }

    #region Supporting enums and structs
    internal enum TShadowType
    {
        Offset,    // N pixel offset shadow
        Double,    // Use second offset too
        Rich,      // Rich perspective shadow (cast relative to shape)
        Shape,     // Rich perspective shadow (cast in shape space)
        Drawing,   // Perspective shadow cast in drawing space
        EmbossOrEngrave
    }
            
	internal enum TShadowStyle
	{
		None,    
		Normal,    
		Obscured
	}

	internal enum TShapePathBiff8
	{
		Lines = 0,    
		LinesClosed = 1,    
		Curves = 2,
		CurvesClosed = 3,
		Complex = 4
	}

	internal enum TSegmentType
	{
		LineTo, // Draw a straight line (one point) 
		CurveTo, // Draw a cubic Bezier curve (three points) 
		MoveTo, // Move to a new point (one point) 
		Close, // Close a sub-path (no points) 
		EndPath, // End a path (no points) 
		Escape, // Escape code 
		ClientEscape, // Escape code interpreted by the client msopathInvalid // Invalid - should never be found
	}

    internal enum TSegmentTypeEscaped
    {
        /// <summary>
        /// Adds additional POINT values to the escape code that follows msopathEscapeExtension. 
        /// </summary>
        msopathEscapeExtension = 0x00000000,

        /// <summary>
        /// The first POINT specifies the center of the ellipse. 
        /// The second POINT specifies the starting radius in x and the ending radius in y. 
        /// The third POINT specifies the starting angle in the x value and the ending angle in the y value. 
        /// Angles are in degrees. The number of ellipse segments drawn is equal to the number of segments divided by three. 
        /// </summary>
        msopathEscapeAngleEllipseTo = 0x00000001,

        /// <summary>
        /// The first POINT specifies the center of the ellipse. 
        /// The second POINT specifies the starting radius in x and the ending radius in y. 
        /// The third POINT specifies the starting angle in the x value and the ending angle in the y value. 
        /// Angles are in degrees. The number of ellipse segments drawn is equal to the number of segments divided by three. 
        /// The first POINT of the ellipse becomes the first POINT of a new path. 
        /// </summary>
        msopathEscapeAngleEllipse = 0x00000002,

        /// <summary>
        /// The first two POINT values specify the bounding rectangle of the ellipse. 
        /// The second two POINT values specify the radial vectors for the ellipse. 
        /// The radial vectors are cast from the center of the bounding rectangle. 
        /// The path will start at the POINT where the first radial vector intersects the bounding rectangle to the POINT where the second radial vector intersects the bounding rectangle. 
        /// The drawing direction is always counterclockwise. 
        /// If the path has already been started, a line is drawn from the last POINT to the starting POINT of the arc; 
        /// otherwise a new path is started. The number of arc segments drawn is equal to the number of segments divided by four. 
        /// </summary>
        msopathEscapeArcTo = 0x00000003,

        /// <summary>
        /// The first two POINT values specify the bounding rectangle of the ellipse. 
        /// The second two POINT values specify the radial vectors for the ellipse. 
        /// The radial vectors are cast from the center of the bounding rectangle. 
        /// The path will start at the POINT where the first radial vector intersects the bounding rectangle to the POINT where the second radial vector intersects the bounding rectangle. The drawing direction is always counterclockwise. 
        /// The number of arc segments drawn is equal to the number of segments divided by four. 
        /// </summary>
        msopathEscapeArc = 0x00000004,

        /// <summary>
        /// The first two POINT values specify the bounding rectangle of the ellipse. 
        /// The second two POINT values specify the radial vectors for the ellipse. 
        /// The radial vectors are cast from the center of the bounding rectangle. 
        /// The path will start at the POINT where the first radial vector intersects the bounding rectangle to the POINT where the second radial vector intersects the bounding rectangle. 
        /// The drawing direction is always clockwise. If the path has already been started, a line is drawn from the last POINT to the starting POINT of the arc; otherwise a new path is started. 
        /// The number of arc segments drawn is equal to the number of segments divided by four. 
        /// </summary>
        msopathEscapeClockwiseArcTo = 0x00000005,

        /// <summary>
        /// The first two POINT values specify the bounding rectangle of the ellipse. 
        /// The second two POINT values specify the radial vectors for the ellipse. 
        /// The radial vectors are cast from the center of the bounding rectangle. 
        /// The path will start at the POINT where the first radial vector intersects the bounding rectangle to the POINT where the second radial vector intersects the bounding rectangle. 
        /// The drawing direction is always clockwise. 
        /// The number of arc segments drawn is equal to the number of segments divided by four. 
        /// This escape code always starts a new path.
        /// </summary>
        msopathEscapeClockwiseArc = 0x00000006,

        /// <summary>
        ///         Add an ellipse to the path from the current POINT to the next POINT starting. 
        ///         The ellipse is drawn as a quadrant that starts as a tangent to the X axis. 
        ///         Multiple elliptical quadrants are joined by a straight line. 
        ///         The number of elliptical quadrants drawn is equal to the number of segments. 
        /// </summary>
        msopathEscapeEllipticalQuadrantX = 0x00000007,

        /// <summary>
        /// Add an ellipse to the path from the current POINT to the next POINT starting. 
        /// The ellipse is drawn as a quadrant that starts as a tangent to the Y axis. 
        /// Multiple elliptical quadrants are joined by a straight line. 
        /// The number of elliptical quadrants drawn is equal to the number of segments. 
        /// </summary>
        msopathEscapeEllipticalQuadrantY = 0x00000008,

        /// <summary>
        /// Each POINT defines a control point for a quadratic Bezier curve. 
        /// The number of control POINT values is defined by the segments parameter. 
        /// </summary>
        msopathEscapeQuadraticBezier = 0x00000009,

        /// <summary>
        /// Path should not be filled
        /// </summary>
        msopathEscapeNoFill = 0x0000000A,

        /// <summary>
        /// Path should not have line
        /// </summary>
        msopathEscapeNoLine = 0x0000000B
    }

    internal struct TPathInfo
    {
        internal bool Close;
        internal bool DrawPath;
        internal bool FillPath;
        internal TPointF[] Points;

        internal TPathInfo(bool aClose, bool aDrawPath, bool aFillPath)
        {
            Close = aClose;
            DrawPath = aDrawPath;
            FillPath = aFillPath;
            Points = null;
        }   
    }
        
    internal class TShadowInfo
    {
        internal TShadowStyle Style;
        internal int Pass;

        internal TShadowInfo(TShadowStyle aStyle, int aPass)
        {
            Style = aStyle;
            Pass = aPass;
        }
    }
    #endregion

    internal static class DrawShape2007
    {
        internal static RectangleF DrawCustomShape(IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100)
        {
            TShapeGeom ShapeDef = ShProp.ShapeGeom;

            TDrawingRelativeRect Bounds = new TDrawingRelativeRect(
                Coords.Left * TDrawingCoordinate.PointsToEmu, Coords.Top * TDrawingCoordinate.PointsToEmu,
                Coords.Right * TDrawingCoordinate.PointsToEmu, Coords.Bottom * TDrawingCoordinate.PointsToEmu);

            Guid CacheID = Guid.NewGuid();

            foreach (TShapePath ShapePath in ShapeDef.PathList)
            {
                DrawPath(Canvas, Workbook, ShProp, Coords, ShadowInfo, Clipping, Zoom100, ShapeDef, ShapePath, Bounds, CacheID);
            }

            if (ShapeDef.TextRect == null) return Coords;
            RectangleF Result = RectangleF.FromLTRB(
                (real)ShapeDef.TextRect.Left.XInPoints(Bounds, CacheID, 0),
                (real)ShapeDef.TextRect.Top.YInPoints(Bounds, CacheID, 0),
                (real)ShapeDef.TextRect.Right.XInPoints(Bounds, CacheID, 0),
                (real)ShapeDef.TextRect.Bottom.YInPoints(Bounds, CacheID, 0));

            DrawShape.Flip(ref Result, Coords, ShProp);
            return Result;

        }

        private static void DrawPath(
            IFlxGraphics Canvas, ExcelFile Workbook, TShapeProperties ShProp, 
            RectangleF Coords, TShadowInfo ShadowInfo, TClippingStyle Clipping, real Zoom100, 
            TShapeGeom ShapeDef, TShapePath ShapePath, TDrawingRelativeRect Bounds, Guid CacheID)
        {
            List<List<TPointF>> Points = new List<List<TPointF>>();

            bool ClosePath = CalcPoints(ShapePath, Points, Bounds, CacheID);

            using (Brush br = DrawShape.GetBrush(Coords, ShProp, Workbook, ShadowInfo, false, Zoom100, ShapePath.PathFill))
            {
                using (Pen pe = DrawShape.GetPen(ShProp, Workbook, ShadowInfo))
                {
                    Pen pe1 = ShapePath.PathStroke ? pe : null;
                    Brush br1 = ShapePath.PathFill != TPathFillMode.None ? br : null;

                    foreach (List<TPointF> CurrPoints in Points)
                    {
                        TPointF[] PointArray = CurrPoints.ToArray();
                        DrawShape.Flip(ref PointArray, Coords, ShProp);

                        Canvas.DrawAndFillBeziers(pe1, br1, PointArray, Clipping);
                        if (pe1 != null && !ClosePath)
                        {
                            DrawShape.DoArrows(Canvas, pe, ShProp, PointArray, Clipping);
                        }
                    }
                }
            }
        }

        private static bool CalcPoints(TShapePath ShapePath, List<List<TPointF>> Points, TDrawingRelativeRect Bounds, Guid CacheID)
        {
            bool ClosePath = false;
            List<TPointF> CurrPoints = null;
            foreach (TShapeAction Action in ShapePath.Actions)
            {
                switch (Action.ActionType)
                {
                    case TShapeActionType.Close:
                        AddDuplicate(CurrPoints);
                        CurrPoints.Add(CurrPoints[0]);
                        AddDuplicate(CurrPoints);
                        ClosePath = true;
                        break;

                    case TShapeActionType.MoveTo:
                        Points.Add(new List<TPointF>());
                        CurrPoints = Points[Points.Count - 1];
                        TShapeActionMoveTo mt = (TShapeActionMoveTo)Action;
                        CurrPoints.Add(new TPointF((real)mt.Target.x.XInPoints(Bounds, CacheID, ShapePath.Width), (real)mt.Target.y.YInPoints(Bounds, CacheID, ShapePath.Height)));
                        break;

                    case TShapeActionType.LineTo:
                        TShapeActionLineTo lt = (TShapeActionLineTo)Action;
                        AddDuplicate(CurrPoints);
                        CurrPoints.Add(new TPointF((real)lt.Target.x.XInPoints(Bounds, CacheID, ShapePath.Width), (real)lt.Target.y.YInPoints(Bounds, CacheID, ShapePath.Height)));
                        AddDuplicate(CurrPoints);
                        break;

                    case TShapeActionType.ArcTo:
                        TShapeActionArcTo art = (TShapeActionArcTo)Action;
                        TPointF StartP = CurrPoints[CurrPoints.Count - 1];

                        double rx = art.WidthRadius.AbsValueInPoints(Bounds, CacheID, ShapePath.Width, Bounds.Width);
                        double ry = art.HeightRadius.AbsValueInPoints(Bounds, CacheID, ShapePath.Height, Bounds.Height);
                        double ls1 = GetRadians(art.StartAngle.Value(0, Bounds, CacheID));
                        double Swing = GetRadians(art.SwingAngle.Value(0, Bounds, CacheID));
                        double ls2 = ls1 + Swing;
                        bool CounterClockWise = Swing < 0;

                        double StartAngle = CounterClockWise ? ls2 : ls1; 
                        double EndAngle = CounterClockWise? ls1 : ls2;

                        double r =rx*ry / Math.Sqrt (ry * ry * Math.Cos(ls1) * Math.Cos(ls1) + rx*rx* Math.Sin(ls1)*Math.Sin(ls1));

                        TPointF[] EPoints = TEllipticalArc.GetPoints(StartP.X - r * Math.Cos(ls1), StartP.Y - r * Math.Sin(ls1),
                            rx,
                            ry,
                            0, StartAngle, EndAngle);

                        if (CounterClockWise) { Array.Reverse(EPoints);}
                       
                        for (int i = 1; i < EPoints.Length; i++)
                        {
                            CurrPoints.Add(EPoints[i]);                            
                        }
                        break;

                    case TShapeActionType.CubicBezierTo:
                        TShapeActionCubicBezierTo cbt = (TShapeActionCubicBezierTo)Action;
                        foreach (TShapePoint pt in cbt.Target)
                        {
                            CurrPoints.Add(new TPointF((real)pt.x.XInPoints(Bounds, CacheID, ShapePath.Width), (real)pt.y.YInPoints(Bounds, CacheID, ShapePath.Height)));
                        }
                        break;

                    case TShapeActionType.QuadBezierTo:
                        TShapeActionQuadBezierTo qbt = (TShapeActionQuadBezierTo)Action;
                        TPointF[] pquad = new TPointF[3];
                        pquad[0] = CurrPoints[CurrPoints.Count - 1];
                        for (int i = 0; i < 2; i++)
                        {
                            pquad[i + 1] = new TPointF((real)qbt.Target[i].x.XInPoints(Bounds, CacheID, ShapePath.Width), (real)qbt.Target[i].y.YInPoints(Bounds, CacheID, ShapePath.Height));
                        }

                        CurrPoints.Add(SumPoint(MultPoint(pquad[0], 1d / 3d), MultPoint(pquad[1], 2d / 3d)));
                        CurrPoints.Add(SumPoint(MultPoint(pquad[2], 1d / 3d), MultPoint(pquad[1], 2d / 3d)));
                        CurrPoints.Add(pquad[2]);

                        break;
                }
            }
            return ClosePath;
        }

        private static double GetRadians(double p)
        {
            return p / 60000.0 * Math.PI / 180d;
        }

        private static TPointF SumPoint(TPointF point1, TPointF point2)
        {
            return new TPointF(point1.X + point2.X, point1.Y + point2.Y);
        }

        private static TPointF MultPoint(TPointF point, double p)
        {
            return new TPointF((real)(point.X * p),(real)(point.Y * p));
        }

        private static void AddDuplicate(List<TPointF> Points)
        {
            Points.Add(Points[Points.Count - 1]);
        }
    }
}
