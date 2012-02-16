using System;
using System.Text;
using System.Globalization;

using System.Collections.Generic;

#if (MONOTOUCH)
	using real = System.Single;
	using System.Drawing;
    using Color = MonoTouch.UIKit.UIColor;
    using Image = MonoTouch.UIKit.UIImage;
    using MonoTouch.Foundation;
#else
	#if (WPF)
	using RectangleF = System.Windows.Rect;
	using SizeF = System.Windows.Size;
	using real = System.Double;
	using ColorBlend = System.Windows.Media.GradientStopCollection;
	using System.Windows.Media;
	using System.Windows;
	#else
	using real = System.Single;
	using System.Drawing;
	using System.Drawing.Drawing2D;
	using System.Drawing.Imaging;
	#endif
#endif

using System.IO;

namespace FlexCel.Core
{
    #region Img Utils
    /// <summary>
    /// Utilities for manipulating images.
    /// </summary>
    public static class ImageUtils
    {
        private static bool CompareHeader(byte[] data, byte[] header, int dataStartPos)
        {
            if (data == null || data.Length - dataStartPos <= header.Length) return false;
            for (int i = 0; i < header.Length; i++)
                if (data[i + dataStartPos] != header[i]) return false;
            return true;
        }

        /// <summary>
        /// Access stores images encapsulated on an OLE container. This function will load an OLE image and
        /// try to return the raw image data.
        /// </summary>
        /// <remarks>See Ms kb Q175261</remarks>
        /// <param name="data">Image in OLE format.</param>
        /// <returns>Image on raw format.</returns>
        public static byte[] StripOLEHeader(byte[] data)
        {
            byte[] OLEHeader = { 0x15, 0x1C };
            if (data == null || data.Length < 4 || !CompareHeader(data, OLEHeader, 0)) return data;

            int OleOfs = BitConverter.ToUInt16(data, 2) - 1; //see Q175261
            int AccessOfs = -1; //not documented, we will have to search.
            int MaxSearch = Math.Min(512, data.Length - OleOfs);
            for (int i = 0; i < MaxSearch; i++)
            {
                if (GetImageType(data, OleOfs + i) != TXlsImgType.Unknown)
                {
                    AccessOfs = i;
                    break;
                }
            }
            if (AccessOfs < 0) return data;
            byte[] Result = new byte[data.Length - OleOfs - AccessOfs];
            Array.Copy(data, data.Length - Result.Length, Result, 0, Result.Length);
            return Result;
        }

        /// <summary>
        /// Returns the image type for a byte array.
        /// </summary>
        /// <param name="data">Array with the image.</param>
        /// <param name="position">Position on data where the image begins.</param>
        /// <returns>Image type</returns>
        public static TXlsImgType GetImageType(byte[] data, int position)
        {
            byte[] EmfHeader = { 0x20, 0x45, 0x4D, 0x46 };
            if (CompareHeader(data, EmfHeader, position + 40)) return TXlsImgType.Emf;

            byte[] PngHeader = { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
            if (CompareHeader(data, PngHeader, position)) return TXlsImgType.Png;

            byte[] JpegHeader1 = { 0xFF, 0xD8, 0xFF, 0xE0 };
            byte[] JpegHeader2 = { 0x4A, 0x46, 0x49, 0x46, 0x00 };
            if (CompareHeader(data, JpegHeader1, position) && CompareHeader(data, JpegHeader2, position + 6)) return TXlsImgType.Jpeg;

            byte[] WmfHeader = { 0xD7, 0xCD, 0xC6, 0x9A };
            if (CompareHeader(data, WmfHeader, position)) return TXlsImgType.Wmf;

            byte[] BmpHeader = { 0x42, 0x4D };
            if (CompareHeader(data, BmpHeader, position)) return TXlsImgType.Bmp;

            byte[] Tiff1Header = { 0x4D, 0x4D, 0x00, 0x2A }; //byte order in second word is reversed
            if (CompareHeader(data, Tiff1Header, position)) return TXlsImgType.Tiff;

            byte[] Tiff2Header = { 0x49, 0x49, 0x2A, 0x00 };
            if (CompareHeader(data, Tiff2Header, position)) return TXlsImgType.Tiff;

            byte[] GifHeader = { 0x47, 0x49, 0x46 };
            if (CompareHeader(data, GifHeader, position)) return TXlsImgType.Gif;

            return TXlsImgType.Unknown;
        }

        /// <summary>
        /// Returns the image type for a byte array.
        /// </summary>
        /// <param name="data">Array with the image.</param>
        /// <returns>Image type</returns>
        public static TXlsImgType GetImageType(byte[] data)
        {
            return GetImageType(data, 0);
        }

        internal static void CheckImgValid(ref byte[] data, ref TXlsImgType imgType, bool AllowBmp)
        {
            TXlsImgType bmp = AllowBmp ? TXlsImgType.Unknown : TXlsImgType.Bmp;
            if (imgType == TXlsImgType.Unknown || imgType == bmp || imgType == TXlsImgType.Gif || imgType == TXlsImgType.Tiff)  //We will try to convert bmps to png. We can't on CF
            {
                data = TCompactFramework.ImgConvert(data, ref imgType);
            }
            if (imgType == TXlsImgType.Unknown || imgType == TXlsImgType.Gif || imgType == TXlsImgType.Tiff)
                FlxMessages.ThrowException(FlxErr.ErrInvalidImage);
        }

    }

    #endregion

     //Classes here mimic system.Drawing classes.

    /// <summary>
    /// Represents a point on X,Y coordinates. We used instead of System.Drawing.Point because we do not want references to System.Drawing 
    /// (because .NET 3.0 applications do not need GDI+)
    /// </summary>
    public struct TPoint
    {
        #region Privates
        private int FX;
        private int FY;
        #endregion

        #region Constructors
        /// <summary>
        /// Creates a new point.
        /// </summary>
        /// <param name="aX">Coordinates for the X value.</param>
        /// <param name="aY">Coordinates for the Y value.</param>
        public TPoint(int aX, int aY)
        {
            FX = aX;
            FY = aY;
        }
        #endregion

        #region Publics
        
        /// <summary>
        /// X coord.
        /// </summary>
        public int X { get { return FX; } set { FX = value; } }

        /// <summary>
        /// Y coord.
        /// </summary>
        public int Y { get { return FY; } set { FY = value; } }

        /// <summary>
        /// True if both objects are equal.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            if (!(obj is TPoint)) return false;
            TPoint o2 = (TPoint)obj;
            return (o2.X == X && o2.Y == Y);
        }

        /// <summary></summary>
        public static bool operator==(TPoint b1, TPoint b2)
        {
            return b1.Equals(b2);
        }

        /// <summary></summary>
        public static bool operator!=(TPoint b1, TPoint b2)
        {
            return !(b1 == b2);
        }

        /// <summary>
        /// Hash code for the point.
        /// </summary>
        /// <returns>hashcode.</returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHash(X, Y);
            
        }


        #endregion
    }

#if (!COMPACTFRAMEWORK && !MONOTOUCH)
    internal sealed class FlxGradient
    {
        private FlxGradient() { }

        public static Color BlendColor(ColorBlend BlendColors, int Position)
        {			
#if (WPF)
            return BlendColors[Position].Color;
#else
            return BlendColors.Colors[Position];
#endif
        }

        public static real BlendPosition(ColorBlend BlendColors, int Position)
        {
#if (WPF)
            return BlendColors[Position].Offset;
#else
            return BlendColors.Positions[Position];
#endif
        }

        public static int BlendCount(ColorBlend BlendColors)
        {
#if (WPF)
            return BlendColors.Count;
#else
                return BlendColors.Colors.Length;
#endif
        }

        public static void SetColorBlend(ColorBlend BlendColors, int i, Color ColorInPos, real Position)
        {
#if (WPF)
            BlendColors[i].Color = ColorInPos;
            BlendColors[i].Offset = Position;
#else
            BlendColors.Colors[i] = ColorInPos;
            BlendColors.Positions[i] = Position;
#endif
        }

        internal static void EnsureMinimumAndMaximum(Color Color1, Color Color2, ref ColorBlend BlendedColors)
        {
            int Blends = BlendCount(BlendedColors);
            bool NeedsZero = BlendPosition(BlendedColors, 0) != 0;
            bool NeedsOne = BlendPosition(BlendedColors, Blends - 1) != 1;

            if (NeedsZero || NeedsOne)
            {
                if (NeedsZero) Blends++;
                if (NeedsOne) Blends++;
                ColorBlend Result = new ColorBlend(Blends);
                int k1 = 0;
                if (NeedsZero) { SetColorBlend(Result, 0, Color1, 0); k1++; }
                for (int i = 0; i < BlendCount(BlendedColors); i++)
                {
                    SetColorBlend(Result, k1 + i, BlendColor(BlendedColors, i), BlendPosition(BlendedColors, i));
                }
                if (NeedsOne) {SetColorBlend(Result, Blends - 1, Color2, 1); }
                BlendedColors = Result;

            }

        }

        internal static void InvertColorBlend(ColorBlend Result)
        {
            for (int i = 0; i < (FlxGradient.BlendCount(Result) + 1) / 2; i++)
            {
                int n = FlxGradient.BlendCount(Result) - 1 - i;

                Color Tmp = FlxGradient.BlendColor(Result, n);
                real TmpPos = FlxGradient.BlendPosition(Result, n);

                FlxGradient.SetColorBlend(Result, n, FlxGradient.BlendColor(Result, i), 1 - FlxGradient.BlendPosition(Result, i));
                FlxGradient.SetColorBlend(Result, i, Tmp, 1 - TmpPos);
            }
        }


    }
#endif

#if (WPF)
    internal static class ColorExtender
    {
        internal static int ToArgb(this Color clr)
        {
            unchecked
            {
                uint Result = (((uint)clr.A) << 24) | (((uint)clr.R) << 16) | (((uint)clr.G) << 8) | (((uint)clr.B) << 0);
                return (int)Result;
            }
        }

        internal static byte A(this Color clr)
        {
            return clr.A;
        }

        internal static byte R(this Color clr)
        {
            return clr.R;
        }

        internal static byte G(this Color clr)
        {
            return clr.G;
        }

        internal static byte B(this Color clr)
        {
            return clr.B;
        }

        internal static float Rd(this Color clr)
        {
            return clr.R / 255f;
        }

        internal static float Gd(this Color clr)
        {
            return clr.G / 255f;
        }

        internal static float Bd(this Color clr)
        {
            return clr.B / 255f;
        }

        internal static bool IsSystemColor(this Color clr)
        {
            return false;
        }

        internal static bool IsNamedColor(this Color clr)
        {
            return false;
        }

        internal static bool IsEmpty(this Color clr)
        {
            return clr == ColorUtil.Empty;
        }
    }
#endif

#if (!WPF && !MONOTOUCH && FRAMEWORK30 && !COMPACTFRAMEWORK)
    internal static class ColorExtender
    {
        internal static byte A(this Color clr)
        {
            return clr.A;
        }

        internal static byte R(this Color clr)
        {
            return clr.R;
        }

        internal static byte G(this Color clr)
        {
            return clr.G;
        }
       
        internal static byte B(this Color clr)
        {
            return clr.B;
        }

        internal static float Rd(this Color clr)
        {
            return clr.R / 255f;
        }

        internal static float Gd(this Color clr)
        {
            return clr.G / 255f;
        }

        internal static float Bd(this Color clr)
        {
            return clr.B / 255f;
        }

        internal static bool IsSystemColor(this Color clr)
        {
            return clr.IsSystemColor;
        }

        internal static bool IsNamedColor(this Color clr)
        {
            return clr.IsNamedColor;
        }

        internal static bool IsEmpty(this Color clr)
        {
            return clr.IsEmpty;
        }

    }
#endif
	
#if(!MONOTOUCH)
    internal static class ImageExtender
    {
#if (!WPF && FRAMEWORK30 && !COMPACTFRAMEWORK)
        internal static int Height(this Image Img)
        {
            return Img.Height;
        }

        internal static int Width(this Image Img)
        {
            return Img.Width;
        }
#endif
#if (!COMPACTFRAMEWORK)
        internal static Image FromStream(Stream St)
        {
            return Image.FromStream(St);
        }
#endif
    }
#endif
	
#if (MONOTOUCH)
	internal enum ImageFormat
	{
		Png,
		Jpg
	}
	
	internal static class ImageExtender
	{
		internal static int Height(this Image Img)
		{
			return (int)Img.Size.Height;	
		}

		internal static int Width(this Image Img)
		{
			return (int)Img.Size.Width;	
		}
		
		internal static Image FromStream(Stream St)
		{
			using (NSData data = NSData.FromStream(St))
			{
				return Image.LoadFromData(data);
			}
		}
		
		internal static void Save(this Image Img, Stream st, ImageFormat ImgFormat)
		{
			using (NSData data = ImgFormat == ImageFormat.Jpg? Img.AsJPEG(): Img.AsPNG())
			{
				using (Stream OutStream = data.AsStream())
				{
					Sh.CopyStream(st, OutStream);
				}
			}
		}
	}
	
	internal static class ColorExtender
	{
		internal static int ToArgb(this Color clr)
        {
            unchecked
            {
                uint Result = (((uint)clr.A()) << 24) | (((uint)clr.R()) << 16) | (((uint)clr.G()) << 8) | (((uint)clr.B()) << 0);
                return (int)Result;
            }
        }
		
		internal static float Rd(this Color clr)
		{
			float r, g, b, a;
			clr.GetRGBA(out r, out g, out b, out a);
			if (r < 0) r = 0; if (r > 1) r = 1;
			return r;
		}

		internal static float Gd(this Color clr)
		{
			float r, g, b, a;
			clr.GetRGBA(out r, out g, out b, out a);
			if (g < 0) g = 0; if (g > 1) g = 1;
			return g;
		}
		
		internal static float Bd(this Color clr)
		{
			float r, g, b, a;
			clr.GetRGBA(out r, out g, out b, out a);
			if (b < 0) b = 0; if (b > 1) b = 1;
			return b;
		}
		
		internal static float Ad(this Color clr)
		{
			float r, g, b, a;
			clr.GetRGBA(out r, out g, out b, out a);
			if (a < 0) a = 0; if (a > 1) a = 1;
			return a;
		}

		internal static byte R(this Color clr)
		{
			return (byte)(clr.Rd() * 255);
		}

		internal static byte G(this Color clr)
		{
			return (byte)(clr.Gd() * 255);
		}
		
		internal static byte B(this Color clr)
		{
			return (byte)(clr.Bd() * 255);
		}

		internal static byte A(this Color clr)
		{
			return (byte)(clr.Ad() * 255);
		}
		
		internal static bool IsEmpty(this Color clr)
		{
			return object.Equals(clr, ColorUtil.Empty);
		}
		
		internal static bool IsSystemColor(this Color clr)
		{
			return false;
		}
		
		internal static bool IsNamedColor(this Color clr)
		{
			return false;
		}
		
		
	}
	
#endif

    internal sealed class ColorUtil
    {
        private ColorUtil() { }

#if (!COMPACTFRAMEWORK)
        internal static Color FromArgb(int a, int r, int g, int b)
        {
#if(MONOTOUCH)
			return Color.FromRGBA(r / 255f, g / 255f, b / 255f, a / 255f);
#else
#if(WPF)
            return Color.FromArgb((byte)a, (byte)r, (byte)g, (byte)b);
#else
            return Color.FromArgb(a, r, g, b);
#endif
#endif
        }
#endif

        internal static Color FromArgb(int r, int g, int b)
        {
#if(MONOTOUCH)
			return Color.FromRGBA(r / 255f, g / 255f, b / 255f, 1);
#else
#if(WPF)
            return Color.FromArgb(255, (byte)r, (byte)g, (byte)b);
#else
            return Color.FromArgb(r, g, b);
#endif
#endif
        }

		internal static Color FromArgb(int argb)
		{
#if(MONOTOUCH)
            unchecked
            {
            	uint uargb = (uint)argb;
            	byte a = (byte)(uargb >> 24);
				byte r = (byte)((uargb >> 16) & 0xFF);
				byte g = (byte)((uargb >> 8) & 0xFF);
				byte b = (byte)(uargb & 0xFF);
				return Color.FromRGBA(r / 255f, g / 255f, b / 255f, a / 255f);
			}
#else
#if(WPF)
            unchecked
            {
                uint uargb = (uint)argb;
                return Color.FromArgb((byte)(uargb >> 24), (byte)((uargb >> 16) & 0xFF), (byte)((uargb >> 8) & 0xFF), (byte)(uargb & 0xFF));
            }

#else
			return Color.FromArgb(argb);
#endif
#endif
		}

#if (!COMPACTFRAMEWORK)
		internal static Color FromArgb(int Alpha, Color Clr)
		{
#if(MONOTOUCH)
			float r,g,b,a;
			Clr.GetRGBA(out r, out g, out b, out a);
			return Color.FromRGBA(r, g, b, Alpha / 255f);
#else
#if(WPF)
            unchecked
            {
                uint uargb = (uint)argb;
                return Color.FromArgb((byte)(uargb >> 24), (byte)((uargb >> 16) & 0xFF), (byte)((uargb >> 8) & 0xFF), (byte)(uargb & 0xFF));
            }

#else
			return Color.FromArgb(Alpha, Clr);
#endif
#endif
		}
#endif

        internal static Color Empty
        {
            get
            {
#if(MONOTOUCH)
				return Colors.Transparent;
#else
#if(WPF)
                return Colors.Transparent;
#else
                return Color.Empty;
#endif
#endif
            }
        }

        internal static int BgrToRgb(long aColor)
        {
            return (int)((aColor & 0x0000FF00) | ((aColor & 0x00FF0000) >> 16) | ((aColor & 0x000000FF) << 16));
        }

        internal static TSystemColor GetSystemColor(long c)
        {
            switch (c)
            {
                case 0: return TSystemColor.ScrollBar;
                case 1: return TSystemColor.Background;
                case 2: return TSystemColor.ActiveCaption;
                case 3: return TSystemColor.InactiveCaption;
                case 4: return TSystemColor.Menu;
                case 5: return TSystemColor.Window;
                case 6: return TSystemColor.WindowFrame;
                case 7: return TSystemColor.MenuText;
                case 8: return TSystemColor.WindowText;
                case 9: return TSystemColor.CaptionText;
                case 10: return TSystemColor.ActiveBorder;
                case 11: return TSystemColor.InactiveBorder;
                case 12: return TSystemColor.AppWorkspace;
                case 13: return TSystemColor.Highlight;
                case 14: return TSystemColor.HighlightText;
                case 15: return TSystemColor.BtnFace;
                case 16: return TSystemColor.BtnShadow;
                case 17: return TSystemColor.GrayText;
                case 18: return TSystemColor.BtnText;
                case 19: return TSystemColor.InactiveCaptionText;
                case 20: return TSystemColor.BtnHighlight;
                case 21: return TSystemColor.DkShadow3d;
                case 22: return TSystemColor.Light3d;
                case 23: return TSystemColor.InfoText;
                case 24: return TSystemColor.InfoBk;

                case 26: return TSystemColor.HotLight;
                case 27: return TSystemColor.GradientActiveCaption;
                case 28: return TSystemColor.GradientInactiveCaption;

                case 29: return TSystemColor.MenuHighlight;
                case 30: return TSystemColor.MenuBar;

            }

            return TSystemColor.None;
        }

        internal static int GetSysColor(TSystemColor sysc)
        {
            switch (sysc)
            {
                case TSystemColor.ScrollBar: return 0;
                case TSystemColor.Background: return 1;
                case TSystemColor.ActiveCaption: return 2;
                case TSystemColor.InactiveCaption: return 3;
                case TSystemColor.Menu: return 4;
                case TSystemColor.Window: return 5;
                case TSystemColor.WindowFrame: return 6;
                case TSystemColor.MenuText: return 7;
                case TSystemColor.WindowText: return 8;
                case TSystemColor.CaptionText: return 9;
                case TSystemColor.ActiveBorder: return 10;
                case TSystemColor.InactiveBorder: return 11;
                case TSystemColor.AppWorkspace: return 12;
                case TSystemColor.Highlight: return 13;
                case TSystemColor.HighlightText: return 14;
                case TSystemColor.BtnFace: return 15;
                case TSystemColor.BtnShadow: return 16;
                case TSystemColor.GrayText: return 17;
                case TSystemColor.BtnText: return 18;
                case TSystemColor.InactiveCaptionText: return 19;
                case TSystemColor.BtnHighlight: return 20;
                case TSystemColor.DkShadow3d: return 21;
                case TSystemColor.Light3d: return 22;
                case TSystemColor.InfoText: return 23;
                case TSystemColor.InfoBk: return 24;

                case TSystemColor.HotLight: return 26;
                case TSystemColor.GradientActiveCaption: return 27;
                case TSystemColor.GradientInactiveCaption: return 28;

                case TSystemColor.MenuHighlight: return 29;
                case TSystemColor.MenuBar: return 30;
            }

            return -1;
        }

        internal static string GetSystemColorName(TSystemColor SysCol)
        {
            switch (SysCol)
            {
                case TSystemColor.ActiveBorder: return "activeBorder";
                case TSystemColor.ActiveCaption: return "activeCaption";
                case TSystemColor.AppWorkspace: return "appWorkspace";
                case TSystemColor.Background: return "background";
                case TSystemColor.BtnFace: return "buttonFace";
                case TSystemColor.BtnHighlight: return "buttonHighlight";
                case TSystemColor.BtnShadow: return "buttonShadow";
                case TSystemColor.BtnText: return "buttonText";
                case TSystemColor.CaptionText: return "captionText";
                case TSystemColor.GrayText: return "grayText";
                case TSystemColor.Highlight: return "highlight";
                case TSystemColor.HighlightText: return "highlightText";
                case TSystemColor.InactiveBorder: return "inactiveBorder";
                case TSystemColor.InactiveCaption: return "inactiveCaption";
                case TSystemColor.InactiveCaptionText: return "inactiveCaptionText";
                case TSystemColor.InfoBk: return "infoBackground";
                case TSystemColor.InfoText: return "infoText";
                case TSystemColor.Menu: return "menu";
                case TSystemColor.MenuText: return "menuText";
                case TSystemColor.ScrollBar: return "scrollbar";
                case TSystemColor.DkShadow3d: return "threeDDarkShadow";
                case TSystemColor.Light3d: return "threeDFace";
                case TSystemColor.HotLight: return "threeDHighlight";
                //case TSystemColor.: return "ThreeDLightShadow"; Can't find an equivalent.
                //case TSystemColor.DkShadow3d: return "ThreeDShadow";
                case TSystemColor.Window: return "window";
                case TSystemColor.WindowFrame: return "windowFrame";
                case TSystemColor.WindowText: return "windowText";
 
            }

            return null;
        }
    }


#if (WPF)

    public enum SmoothingMode
    {
        AntiAlias
    }

    public enum InterpolationMode
    {
        HighQualityBicubic
    }

    internal enum HatchStyle
    {
        Percent50,
        Percent75,
        Percent25,
        DarkHorizontal,
        DarkVertical,
        DarkUpwardDiagonal,
        DarkDownwardDiagonal,
        SmallCheckerBoard,
        Percent70,
        LightHorizontal, //  thin horz lines
        LightVertical, //  thin vert lines
        LightUpwardDiagonal,
        LightDownwardDiagonal,
        SmallGrid,
        Percent60,
        Percent10,
        Percent05
    }

    internal class HatchBrush : Brush
    {
    }


    /// <summary>
    /// Represents a Font in WPF, including font size, family and style.
    /// </summary>
    public struct Font
    {
        double FSize;
        FontFamily FFamily;
        FontStyle FStyle;
        FontWeight FWeight;
        TextDecorationCollection FDecorations;

        /// <summary>
        /// Creates a new font.
        /// </summary>
        /// <param name="aFamily">Familiy for the font.</param>
        /// <param name="aStyle">Style for the font.</param>
        /// <param name="aSize">Size in points.</param>
        /// <param name="aWeight">Weight of the font.</param>
        /// <param name="aDecorations">Decorations for the font.</param>
        public Font(FontFamily aFamily, double aSize, FontStyle aStyle, FontWeight aWeight, TextDecorationCollection aDecorations)
        {
            FFamily = aFamily;
            FSize = aSize;
            FStyle = aStyle;
            FWeight = aWeight;
            FDecorations = aDecorations;
        }

        /// <summary>
        /// Creates a new font.
        /// </summary>
        /// <param name="aName">Name of the font.</param>
        /// <param name="aStyle">Style for the font.</param>
        /// <param name="aSize">Size in points.</param>
        /// <param name="aWeight">Weight of the font.</param>
        /// <param name="aDecorations">Decorations for the font.</param>
        public Font(string aName, double aSize, FontStyle aStyle, FontWeight aWeight, TextDecorationCollection aDecorations)
        {
            FFamily = new FontFamily(aName);
            FSize = aSize;
            FStyle = aStyle;
            FWeight = aWeight;
            FDecorations = aDecorations;
        }

        /// <summary>
        /// Size of the font in points.
        /// </summary>
        public double SizeInPoints { get { return FSize; } set { FSize = value; } }

        /// <summary>
        /// Size of the font in pixels.
        /// </summary>
        public double SizeInPix { get { return FSize / FlxConsts.PixToPoints; } set { FSize = value * FlxConsts.PixToPoints; } }

        /// <summary>
        /// Family of the font.
        /// </summary>
        public FontFamily Family { get { return FFamily; } set { FFamily = value; } }

        /// <summary>
        /// Style of the font.
        /// </summary>
        public FontStyle Style { get { return FStyle; } set { FStyle = value; } }

        /// <summary>
        /// Weight of the font.
        /// </summary>
        public FontWeight Weight { get { return FWeight; } set { FWeight = value; } }

        /// <summary>
        /// Decorations for the font.
        /// </summary>
        public TextDecorationCollection Decorations { get { return FDecorations; } set { FDecorations = value; } }


        /// <summary>
        /// Freezes all freezable items in the font
        /// </summary>
        internal void Freeze()
        {
#if (!SILVERLIGHT)
            if (Decorations.CanFreeze) Decorations.Freeze();
#endif
        }
    }

#endif
#if (!COMPACTFRAMEWORK && !MONOTOUCH && !SILVERLIGHT)

    /// <summary>
    /// An utility class to create bitmaps and wrap the error in a FlexCelException. Internal use.
    /// </summary>
    public sealed class BitmapConstructor
    {
        private BitmapConstructor() { }

        /// <summary>
        /// Returns a new bitmap.
        /// </summary>
        /// <param name="height">Height of the bitmap in pixels.</param>
        /// <param name="width">Width of the bitmap in pixels.</param>
        /// <returns></returns>
        public static Bitmap CreateBitmap(int height, int width) 
        {
            return CreateBitmap(height, width, PixelFormat.Format32bppPArgb);
        }

        /// <summary>
        /// Creates a bitmap.
        /// </summary>
        /// <param name="height">Height of the bitmap in pixels.</param>
        /// <param name="width">Width of the bitmap in pixels.</param>
        /// <param name="aPixelFormat">Pixel Format.</param>
        /// <returns></returns>
        public static Bitmap CreateBitmap(int height, int width, PixelFormat aPixelFormat)
        {
            try
            {
                return new Bitmap(height, width, aPixelFormat);
            }
            catch (ArgumentException)
            {
                FlxMessages.ThrowException(FlxErr.ErrCreatingImage, height, width, TCompactFramework.EnumGetName(typeof(PixelFormat), aPixelFormat));
            }
            return null;
        }
    }
#endif
}
