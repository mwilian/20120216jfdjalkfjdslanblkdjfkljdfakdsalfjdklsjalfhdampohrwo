using System.Text;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

#if (MONOTOUCH)
	using real = System.Single;
	using System.Drawing;
    using Color = MonoTouch.UIKit.UIColor;
#else
	#if(WPF)
	using real = System.Double;
	using System.Windows;
	using System.Windows.Media;
	#else
	using real = System.Single;
	using System.Drawing;
	using Colors = System.Drawing.Color;
	#endif
#endif

namespace FlexCel.Core
{
    #region HSL
    /// <summary>
    /// Implements a simple representation of a color in Hue/Saturation/Lum colorspace.
    /// </summary>
    public struct THSLColor: IComparable
    {
        #region Variables
        private double FHue;
        private double FSat;
        private double FLum;

        private byte FR;
        private byte FG;
        private byte FB;
        #endregion

        #region Properties
        /// <summary>
        /// Color hue. (between 0 and 360)
        /// </summary>
        public double Hue { get { return FHue; } }

        /// <summary>
        /// Color Saturation. (between 0 and 1)
        /// </summary>
        public double Sat { get { return FSat; } }

        /// <summary>
        /// Color brigthtness. (between 0 and 1)
        /// </summary>
        public double Lum { get { return FLum; } }




        /// <summary>
        /// Red component in the RGB space. 
        /// </summary>
        public byte R { get { return FR; } }

        /// <summary>
        /// Green component in the RGB space. 
        /// </summary>
        public byte G { get { return FG; } }

        /// <summary>
        /// Blue component in the RGB space. 
        /// </summary>
        public byte B { get { return FB; } }
        #endregion

        #region Constructors
        /// <summary>
        /// Creates a new instance from a system color.
        /// </summary>
        /// <param name="rGBColor"></param>
        public THSLColor(Color rGBColor)
        {
#if (COMPACTFRAMEWORK)
            CalcHSL(rGBColor.R, rGBColor.G, rGBColor.B, out FHue, out FSat, out FLum);
#else
            FHue = rGBColor.GetHue();
            FSat = rGBColor.GetSaturation();
            FLum = rGBColor.GetBrightness();
#endif

#if (COMPACTFRAMEWORK || (!FRAMEWORK30 && !MONOTOUCH))
            FR = rGBColor.R;
            FG = rGBColor.G;
            FB = rGBColor.B;
#else
            FR = rGBColor.R();
            FG = rGBColor.G();
            FB = rGBColor.B();
#endif
        }

        /// <summary>
        /// Creates a Color from the hue, saturation and luminescence.
        /// </summary>
        /// <param name="aHue">Hue for the color</param>
        /// <param name="aSat">Saturation for the color.</param>
        /// <param name="aLum">Luminescence for the color.</param>
        public THSLColor(double aHue, double aSat, double aLum)
        {
            CheckBoundsHue(ref aHue);
            FHue = aHue;
            CheckBounds01(ref aSat);
            FSat = aSat;
            CheckBounds01(ref aLum);
            FLum = aLum;

            CalcRGB(FHue, FSat, FLum, out FR, out FG, out FB);
        }

        /// <summary>
        /// Assigns a system color to this instance.
        /// </summary>
        /// <param name="aColor"></param>
        /// <returns></returns>
        public static implicit operator THSLColor(Color aColor)
        {
            return new THSLColor(aColor);
        }

        /// <summary>
        /// Assigns this instance to a system color.
        /// </summary>
        /// <param name="aColor"></param>
        /// <returns></returns>
        public static implicit operator Color(THSLColor aColor)
        {
            return ColorUtil.FromArgb(aColor.R, aColor.G, aColor.B);
        }

        /// <summary>
        /// Returns a system color from this instance. This method is only needed in Visual Basic, in C# you can just assign the HslColor to the Color:
        /// <code>
        /// Color myColor = hslColor; 
        /// </code>
        /// Is the same as:
        /// <code>
        /// Color myColor = hslColor.ToColor(); 
        /// </code>
        /// </summary>
        public Color ToColor()
        {
            return ColorUtil.FromArgb(R, G, B);            
        }

        #endregion

        #region Utilities
        private static void CheckBounds01(ref double value)
        {
            if (value < 0) value = 0;
            if (value > 1) value = 1;
        }

        private static void CheckBoundsHue(ref double value)
        {
            while (value < 0) value += 360;
            value = value % 360;
        }

        private static void CheckBoundsTint(ref double value)
        {
            if (value < -1) value = -1;
            if (value > 1) value = 1;
        }
        #endregion

        #region HSL->RBG
        private static void CalcRGB(double Hue, double Sat, double Lum, out byte FR, out byte FG, out byte FB)
        {
            double q = Lum < 0.5 ?
                Lum * (1 + Sat) :
                (Lum + Sat) - (Lum * Sat);

            double p = 2 * Lum - q;
            double hk = Hue / 360.0;

            double tr = hk + 1.0 / 3.0;
            double tg = hk;
            double tb = hk - 1.0 / 3.0;

            FR = ColorToByte(CalcComponent(p, q, tr));
            FG = ColorToByte(CalcComponent(p, q, tg));
            FB = ColorToByte(CalcComponent(p, q, tb));

        }

        private static double CalcComponent(double p, double q, double tc)
        {
            if (tc < 0) tc += 1;
            if (tc > 1) tc -= 1;

            if (tc < 1.0 / 6.0)
            {
                return p + ((q - p) * 6.0 * tc);
            }

            if (tc < 1.0 / 2.0)
            {
                return q;
            }

            if (tc < 2.0 / 3.0)
            {
                return p + ((q - p) * 6.0 * (2.0 / 3.0 - tc));
            }

            return p;
        }

        private static byte ColorToByte(double c)
        {
            if (c > 1) c = 1;
            if (c < 0) c = 0;
            return (byte)Math.Round(c * 255);
        }
        #endregion

        #region RGB->HSL
        #endregion

        #region Public methods
        /// <summary>
        /// This method returns the brightness that results from applying tint to brightness.
        /// </summary>
        /// <param name="brightness">Brightness of the color (between 0 and 1)</param>
        /// <param name="tint">Tint of the color (between -1 and 1)</param>
        public static double ApplyTint(double brightness, double tint)
        {
            CheckBounds01(ref brightness);
            CheckBoundsTint(ref tint);

            if (tint < 0) return brightness * (tint + 1);
            return brightness + (1 - brightness) * tint;
        }

        /// <summary>
        /// Returns the tint needed to go from originalBrightness to newBrightness.
        /// A tint of 0 means no change (OriginalBrightness == NewBrightness), a tint of -1 means NewBrightness = 0,
        /// and a tint of 1 means NewBrightness = 1. So this method just does a simple interpolation to find out the needed tint.
        /// <para></para>This method is an inverse of <see cref="ApplyTint(double,double)"/>
        /// </summary>
        /// <param name="originalBrightness">Brightness of the original color. (between 0 and 1)</param>
        /// <param name="newBrightness">Brightness of the new color that we want to produce by applying tint to originalBrightness/></param>
        /// <returns>The tint we need to apply to go from OriginalBrightness to NewBrigtness.</returns>
        public static double GetTint(double originalBrightness, double newBrightness)
        {
            CheckBounds01(ref originalBrightness);
            CheckBounds01(ref newBrightness);

            if (newBrightness == originalBrightness) return 0;

            if (newBrightness > originalBrightness) //Tint must be > 0
            {
                if (Math.Abs(1 - originalBrightness) < double.Epsilon * 2) return 0;
                return (newBrightness - originalBrightness) / (1 - originalBrightness);  //denominator cannot be 0. OriginalBrightness <NewBrightness <=1  -> OriginalBrighness < 1. There could be some numeric stability issues here, though.
            }

            //Tint must be < 0
            if (Math.Abs(originalBrightness) < double.Epsilon * 2) return 0;
            return newBrightness / originalBrightness - 1;  //denominator cannot be 0. OriginalBrightness >NewBrightness >=0  -> OriginalBrighness > 0. There could be some numeric stability issues here, though.

        }

        /// <summary>
        /// Returns the distance between 2 colors. Not that this is not the euclidean distance, but a distance calculated to improve Hue matching.
        /// When converting cell colors, we try to preserve hues, so even a very pale red cell will be converted to bright red and not white or a very pale blue.
        /// This make it different from standard color matching as is done when adjusting images to a color palette, and where hue is not as important as here.
        /// </summary>
        /// <param name="hue1"></param>
        /// <param name="sat1"></param>
        /// <param name="hue2"></param>
        /// <param name="sat2"></param>
        /// <returns></returns>
        public static double DistanceSquared(double hue1, double sat1, double hue2, double sat2)
        {
            //Real distance would be:
            //return Math.Pow(SinAlpha * sat2, 2) + Math.Pow(sat1 - sat2 * CosAlpha, 2); 
            //or simplified:
            //return Math.Pow(sat1, 2) + Math.Pow(sat2, 2) - 2 * sat1 * sat2 * CosAlpha;
            //but this doesn't make sense in a HSL space. Hue variations look much more different to the eye than saturation variations.
            //Except when sat is very small. In fact, at sat = 0, hue variations don't matter.
            //So we will use the not simplified expression, weigthed so diffs in the hue plane have much less impact than in the perpendicular plane.

            double alpha = Math.Abs(hue1 - hue2) * Math.PI / 180;
            double CosAlpha = Math.Cos(alpha);
            double SinAlpha = Math.Sin(alpha);

            sat1 = Math.Pow(sat1, 0.1); //saturation isn't linear, it affects more near 0 sat.
            sat2 = Math.Pow(sat2, 0.1);
            const double a1 = 100;
            const double a2 = 32;

            return a1 * Math.Pow(SinAlpha * sat2, 2) + a2 * Math.Pow(sat1 - sat2 * CosAlpha, 2); 

        }

        internal static double DistanceSquared(double hue1, double sat1, double lum1, double hue2, double sat2, double lum2)
        {
            //lum grows vertically, hue is the angle in the cilynder, sat is the distance from the center of the cilinder.
            //So looking at the plane vertically, distance is the distance in (sat,hue)2 + (lum2-lum1)2

            return DistanceSquared(hue1, sat1, hue2, sat2) + Math.Pow(lum2 - lum1, 2);
        }

        #endregion

        #region Conversion
        /// <summary>
        /// Returns true if both colors are the same.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            if (!(obj is THSLColor)) return false;
            THSLColor o = (THSLColor)obj;

            return FHue == o.FHue && FSat == o.FSat && FLum == o.FLum;
        }

        /// <summary>
        /// Returns a hashcode for the color.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHash(FHue.GetHashCode(), FSat.GetHashCode(), FLum.GetHashCode());
        }

        /// <summary>
        /// Returns true if both colors are equal.
        /// </summary>
        /// <param name="o1">First color to compare.</param>
        /// <param name="o2">Second color to compare.</param>
        /// <returns></returns>
        public static bool operator ==(THSLColor o1, THSLColor o2)
        {
            return o1.Equals(o2);
        }

        /// <summary>
        /// Returns true if both colors do not have the same value.
        /// </summary>
        /// <param name="o1">First color to compare.</param>
        /// <param name="o2">Second color to compare.</param>
        /// <returns></returns>
        public static bool operator !=(THSLColor o1, THSLColor o2)
        {
            return !(o1.Equals(o2));
        }

        /// <summary>
        /// Returns true if o1 is bigger than o2.
        /// </summary>
        /// <param name="o1"></param>
        /// <param name="o2"></param>
        /// <returns></returns>
        public static bool operator >(THSLColor o1, THSLColor o2)
        {
            return o1.CompareTo(o2) > 0;
        }

        /// <summary>
        /// Returns true if o1 is less than o2.
        /// </summary>
        /// <param name="o1"></param>
        /// <param name="o2"></param>
        /// <returns></returns>
        public static bool operator <(THSLColor o1, THSLColor o2)
        {
            return o1.CompareTo(o2) < 0;
        }

        #endregion

        #region IComparable Members

        /// <summary>
        /// Returns -1 if obj is more than color, 0 if both colors are the same, and 1 if obj is less than color.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public int CompareTo(object obj)
        {
            if (!(obj is THSLColor)) return -1;
            THSLColor o2 = (THSLColor)obj;
            int c;
            c = FHue.CompareTo(o2.FHue); if (c != 0) return c;
            c = FSat.CompareTo(o2.FSat); if (c != 0) return c;
            c = FLum.CompareTo(o2.FLum); if (c != 0) return c;

            return 0;
        }

        #endregion
    }
    #endregion

    #region ScRGB
    /// <summary>
    /// Implements a simple representation of a color in scRGB colorspace. Components are doubles going from 0 to 1.
    /// </summary>
    public struct TScRGBColor : IComparable
    {
        #region Variables
        private double FscR;
        private double FscG;
        private double FscB;
        #endregion

        #region Properties
        /// <summary>
        /// ScRed component. (between 0 and 1)
        /// </summary>
        public double ScR { get { return FscR; } }

        /// <summary>
        /// ScGreen component. (between 0 and 1)
        /// </summary>
        public double ScG { get { return FscG; } }

        /// <summary>
        /// ScBlue component. (between 0 and 1)
        /// </summary>
        public double ScB { get { return FscB; } }

        /// <summary>
        /// Red component in the RGB space. (0-255)
        /// </summary>
        public byte R { get { return To255(SRGBtoRGB(FscR)); } }

        /// <summary>
        /// Green component in the RGB space. (0-255)
        /// </summary>
        public byte G { get { return To255(SRGBtoRGB(FscG)); } }

        /// <summary>
        /// Blue component in the RGB space. (0-255)
        /// </summary>
        public byte B { get { return To255(SRGBtoRGB(FscB)); } }

        internal static byte To255(double p)
        {
            if (p < 0) p = 0;
            if (p > 1) p = 1;
            return (byte)Math.Round(p * 255);
        }

        #endregion

        #region Constructors
        /// <summary>
        /// Creates a new instance from a system color.
        /// </summary>
        /// <param name="RGBColor"></param>
        public TScRGBColor(Color RGBColor)
        {
#if (COMPACTFRAMEWORK || (!FRAMEWORK30 && !MONOTOUCH))
            FscR = RGBtoSRGB(RGBColor.R / 255f);
            FscG = RGBtoSRGB(RGBColor.G / 255f);
            FscB = RGBtoSRGB(RGBColor.B / 255f);
#else
            FscR = RGBtoSRGB(RGBColor.Rd());
            FscG = RGBtoSRGB(RGBColor.Gd());
            FscB = RGBtoSRGB(RGBColor.Bd());
#endif
        }

        /// <summary>
        /// Creates a Color from the sc components.
        /// <param name="scRed">Red component. (0-1)</param>
        /// <param name="scGreen">Green component. (0-1)</param>
        /// <param name="scBlue">Blue component. (0-1)</param>
        /// </summary>
        public TScRGBColor(double scRed, double scGreen, double scBlue)
        {
            FscR = scRed;
            FscG = scGreen;
            FscB = scBlue;
        }

        /// <summary>
        /// Assigns a system color to this instance.
        /// </summary>
        /// <param name="aColor"></param>
        /// <returns></returns>
        public static implicit operator TScRGBColor(Color aColor)
        {
            return new TScRGBColor(aColor);
        }

        /// <summary>
        /// Assigns this instance to a system color.
        /// </summary>
        /// <param name="aColor"></param>
        /// <returns></returns>
        public static implicit operator Color(TScRGBColor aColor)
        {
            return ColorUtil.FromArgb(aColor.R, aColor.G, aColor.B);
        }

        #endregion

        #region Utilities
        private static void CheckBounds01(ref double value)
        {
            if (value < 0) value = 0;
            if (value > 1) value = 1;
        }

        /// <summary>
        /// Converts a RGB value to sRGB using a gamma of 2.2
        /// </summary>
        /// <param name="component">R, G or B component of the color.</param>
        /// <returns></returns>
        public static double RGBtoSRGB(double component)
        {
            CheckBounds01(ref component);
            if (component > 0.04045) return Math.Pow((component + 0.055) / 1.055, 2.4); else return component / 12.92;
        }

        /// <summary>
        /// Converts a sRGB value to RGB using a gamma of 2.2
        /// </summary>
        /// <param name="component">sR, sG or sB component of the color.</param>
        /// <returns></returns>
        public static double SRGBtoRGB(double component)
        {
            CheckBounds01(ref component);
            if (component <= 0.0031308)
            {
                return component * 12.92;
            }

            return 1.055 * (Math.Pow(component, (1 / 2.4))) - 0.055;
        }

        #endregion

        #region Conversion
        /// <summary>
        /// Returns true if both colors are the same.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            if (!(obj is TScRGBColor)) return false;
            TScRGBColor o = (TScRGBColor)obj;

            return FscR == o.FscR && FscG == o.FscG && FscB == o.FscB;
        }

        /// <summary>
        /// Returns a hashcode for the color.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHash(FscR.GetHashCode(), FscG.GetHashCode(), FscB.GetHashCode());
        }

        /// <summary>
        /// Returns true if both colors are equal.
        /// </summary>
        /// <param name="o1">First color to compare.</param>
        /// <param name="o2">Second color to compare.</param>
        /// <returns></returns>
        public static bool operator ==(TScRGBColor o1, TScRGBColor o2)
        {
            return o1.Equals(o2);
        }

        /// <summary>
        /// Returns true if both colors do not have the same value.
        /// </summary>
        /// <param name="o1">First color to compare.</param>
        /// <param name="o2">Second color to compare.</param>
        /// <returns></returns>
        public static bool operator !=(TScRGBColor o1, TScRGBColor o2)
        {
            return !(o1.Equals(o2));
        }

        /// <summary>
        /// Returns true if o1 is bigger than o2.
        /// </summary>
        /// <param name="o1"></param>
        /// <param name="o2"></param>
        /// <returns></returns>
        public static bool operator >(TScRGBColor o1, TScRGBColor o2)
        {
            return o1.CompareTo(o2) > 0;
        }

        /// <summary>
        /// Returns true if o1 is less than o2.
        /// </summary>
        /// <param name="o1"></param>
        /// <param name="o2"></param>
        /// <returns></returns>
        public static bool operator <(TScRGBColor o1, TScRGBColor o2)
        {
            return o1.CompareTo(o2) < 0;
        }

        #endregion

        #region IComparable Members

        /// <summary>
        /// Returns -1 if obj is more than color, 0 if both colors are the same, and 1 if obj is less than color.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public int CompareTo(object obj)
        {
            if (!(obj is TScRGBColor)) return -1;
            TScRGBColor o2 = (TScRGBColor)obj;
            int c;
            c = FscR.CompareTo(o2.FscR); if (c != 0) return c;
            c = FscG.CompareTo(o2.FscG); if (c != 0) return c;
            c = FscB.CompareTo(o2.FscB); if (c != 0) return c;

            return 0;
        }

        #endregion
    }
    #endregion

    #region CIE-L*a*b*
    /// <summary>
    /// Implements a simple representation of a color in CIE-L*a*b* colorspace. This colorspace is mostly used for finding distances between colors.
    /// </summary>
    public struct TLabColor : IComparable
    {
        #region Variables
        private double FL0;
        private double Fa0;
        private double Fb0;

        private byte FR;
        private byte FG;
        private byte FB;
        #endregion

        #region Properties
        /// <summary>
        /// L* (Lightness) (Between 0 and 100)
        /// </summary>
        public double L0 { get { return FL0; } }

        /// <summary>
        /// a*
        /// </summary>
        public double a0 { get { return Fa0; } }

        /// <summary>
        /// b*
        /// </summary>
        public double b0 { get { return Fb0; } }




        /// <summary>
        /// Red component in the RGB space. 
        /// </summary>
        public byte R { get { return FR; } }

        /// <summary>
        /// Green component in the RGB space. 
        /// </summary>
        public byte G { get { return FG; } }

        /// <summary>
        /// Blue component in the RGB space. 
        /// </summary>
        public byte B { get { return FB; } }
        #endregion

        #region Constructors
        /// <summary>
        /// Creates a new instance from a system color.
        /// </summary>
        /// <param name="rGBColor"></param>
        public TLabColor(Color rGBColor)
        {
#if (COMPACTFRAMEWORK || (!FRAMEWORK30 && !MONOTOUCH))
            CalcLab(rGBColor.R, rGBColor.G, rGBColor.B, out FL0, out Fa0, out Fb0);

            FR = rGBColor.R;
            FG = rGBColor.G;
            FB = rGBColor.B;
#else
   CalcLab(rGBColor.R(), rGBColor.G(), rGBColor.B(), out FL0, out Fa0, out Fb0);

            FR = rGBColor.R();
            FG = rGBColor.G();
            FB = rGBColor.B();
#endif
        }

        /// <summary>
        /// Assigns a system color to this instance.
        /// </summary>
        /// <param name="aColor"></param>
        /// <returns></returns>
        public static implicit operator TLabColor(Color aColor)
        {
            return new TLabColor(aColor);
        }

        /// <summary>
        /// Assigns this instance to a system color.
        /// </summary>
        /// <param name="aColor"></param>
        /// <returns></returns>
        public static implicit operator Color(TLabColor aColor)
        {
            return ColorUtil.FromArgb(aColor.R, aColor.G, aColor.B);
        }

        /// <summary>
        /// Returns a system color from this instance. This method is only needed in Visual Basic, in C# you can just assign the LabColor to the Color:
        /// <code>
        /// Color myColor = labColor; 
        /// </code>
        /// Is the same as:
        /// <code>
        /// Color myColor = labColor.ToColor(); 
        /// </code>
        /// </summary>
        public Color ToColor()
        {
            return ColorUtil.FromArgb(R, G, B);
        }

        #endregion

        #region RGB->Lab
        private static void CalcXYZ(byte FR, byte FG, byte FB, out double X, out double Y, out double Z)
        {
            double r = TScRGBColor.RGBtoSRGB(FR / 255.0);
            double g = TScRGBColor.RGBtoSRGB(FG / 255.0);
            double b = TScRGBColor.RGBtoSRGB(FB / 255.0);

            X = r * 0.412453 + g * 0.357580 + b * 0.180423;
            Y = r * 0.212671 + g * 0.715160 + b * 0.072169;
            Z = r * 0.019334 + g * 0.119193 + b * 0.950227;
        }

        private static double flab(double xyz)
        {
            if (xyz > 0.008856) return Math.Pow(xyz, (1.0/3.0)); else return 7.787*xyz + 16.0/116.0; 
        }

        private static void CalcLab(byte FR, byte FG, byte FB, out double L, out double a, out double b)
        {
            const double Xn = 0.9505; //D65 White point
            const double Yn = 1.0;
            const double Zn = 1.0890;

            double X, Y, Z;
            CalcXYZ(FR, FG, FB, out X, out Y, out Z);

            double fX = flab(X/Xn);
            double fY = flab(Y/Yn);
            double fZ = flab(Z/Zn);

            L = (116.0 * fY) - 16;
            a = 500.0 * (fX - fY);
            b = 200.0 * (fY - fZ);


        }
        #endregion

        #region Public methods
        /// <summary>
        /// Returns the euclidean distance squared (DeltaE CIE 1976 squared) between this color and other color.
        /// </summary>
        /// <param name="Color2"></param>
        /// <returns></returns>
        public double DistanceSquared(TLabColor Color2)
        {
            double a = (a0 * Color2.a0 <= 0) ? 100000 : 1; //avoid opposite colors appearing
            double b = (b0 * Color2.b0 <= 0) ? 100000 : 1;

            return  Math.Pow(Color2.L0 - L0, 2) + a * Math.Pow(Color2.a0 - a0, 2) + b * Math.Pow(Color2.b0 - b0, 2) ;
        }


        /// <summary>
        /// Returns the CMC color distance between this color and color2 (distance returned is squared, so you need to get the sqrt if you want the real CMC value). Note that CMC is not symetric (Color1.CMC(Color2) != Color2.CMC(Color1), so this color is the one used as reference.
        /// </summary>
        /// <param name="Color2">Color that will be compared against this reference.</param>
        /// <param name="l">L parameter for CMC calculation. For acceptability (CMC2:1) this is normally 2, and for perceptibility (CMC1:1) this should be 1.</param>
        /// <param name="c">C parameter for CMC calculation. This is normally 1.</param>
        /// <returns></returns>
        public double CMCSquared(TLabColor Color2, int l, int c)
        {
            double C1Squared = Math.Pow(a0, 2) + Math.Pow(b0, 2);

            double C1 = Math.Sqrt(C1Squared);
            double C2 = Math.Sqrt(Math.Pow(Color2.a0, 2) + Math.Pow(Color2.b0, 2));
            double DeltaCSquared = (C1 - C2) * (C1 - C2);

            double DeltaHSquared = Math.Pow(a0 - Color2.a0, 2) + Math.Pow(b0 - Color2.b0, 2) - DeltaCSquared;
            double DeltaL = L0 - Color2.L0;

            double H1 = Math.Atan2(b0, a0) * 180 / Math.PI;
            if (H1 < 0) H1 += 360; if (H1 >= 360) H1 -= 360;

            double C1SquaredSquared = C1Squared * C1Squared;
            double F = Math.Sqrt(C1SquaredSquared / (C1SquaredSquared + 1900));

            double T = H1 >= 164 && H1 <= 345 ?
                0.56 + 0.2 * Math.Abs(Math.Cos((H1 + 168) * Math.PI / 180)) :
                0.36 + 0.4 * Math.Abs(Math.Cos((H1 + 35) * Math.PI / 180));

            double Sl = L0 < 16 ? 0.511 : 0.040975 * L0 / (1 + 0.01765 * L0);
            double Sc = 0.0638 * C1 / (1 + 0.0131 * C1) + 0.638;
            double Sh = Sc * (F * T + 1 - F);

            return Math.Pow(DeltaL / (l * Sl), 2) + DeltaCSquared / Math.Pow(c * Sc, 2) + DeltaHSquared / Math.Pow(Sh, 2);
        }
        #endregion

        #region Conversion
        /// <summary>
        /// Returns true if both colors are the same.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            if (!(obj is TLabColor)) return false;
            TLabColor o = (TLabColor)obj;

            return FL0 == o.FL0 && Fa0 == o.Fa0 && Fb0 == o.Fb0;
        }

        /// <summary>
        /// Returns a hashcode for the color.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHash(FL0.GetHashCode(), Fa0.GetHashCode(), Fb0.GetHashCode());
        }

        /// <summary>
        /// Returns true if both colors are equal.
        /// </summary>
        /// <param name="o1">First color to compare.</param>
        /// <param name="o2">Second color to compare.</param>
        /// <returns></returns>
        public static bool operator ==(TLabColor o1, TLabColor o2)
        {
            return o1.Equals(o2);
        }

        /// <summary>
        /// Returns true if both colors do not have the same value.
        /// </summary>
        /// <param name="o1">First color to compare.</param>
        /// <param name="o2">Second color to compare.</param>
        /// <returns></returns>
        public static bool operator !=(TLabColor o1, TLabColor o2)
        {
            return !(o1.Equals(o2));
        }

        /// <summary>
        /// Returns true if o1 is bigger than o2.
        /// </summary>
        /// <param name="o1"></param>
        /// <param name="o2"></param>
        /// <returns></returns>
        public static bool operator >(TLabColor o1, TLabColor o2)
        {
            return o1.CompareTo(o2) > 0;
        }

        /// <summary>
        /// Returns true if o1 is less than o2.
        /// </summary>
        /// <param name="o1"></param>
        /// <param name="o2"></param>
        /// <returns></returns>
        public static bool operator <(TLabColor o1, TLabColor o2)
        {
            return o1.CompareTo(o2) < 0;
        }


        #endregion

        #region IComparable Members

        /// <summary>
        /// Returns -1 if obj is more than color, 0 if both colors are the same, and 1 if obj is less than color.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public int CompareTo(object obj)
        {
            if (!(obj is TLabColor)) return -1;
            TLabColor o2 = (TLabColor)obj;
            int c;
            c = FL0.CompareTo(o2.FL0); if (c != 0) return c;
            c = Fa0.CompareTo(o2.Fa0); if (c != 0) return c;
            c = Fb0.CompareTo(o2.Fb0); if (c != 0) return c;

            return 0;
        }

        #endregion
    }
    #endregion
}
