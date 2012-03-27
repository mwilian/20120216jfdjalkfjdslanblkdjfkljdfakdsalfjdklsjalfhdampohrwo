using System;
using System.Collections.Generic;
using System.Text;
#if (WPF)
using System.Windows.Media;
#else
using System.Drawing;
#endif

namespace FlexCel.Core
{

    enum TDrawingUnit
    {
        emu,
        cm,
        mm,
        inches,
        pt,
        pc,
        pi
    }

    /// <summary>
    /// A coordinate in a drawing.
    /// </summary>
    public struct TDrawingCoordinate: IComparable
    {
        private long FEmu;

        /// <summary>
        /// Returns how many EMUs in 1 cm.
        /// </summary>
        public const double CmToEmu = 360000.0;

        /// <summary>
        /// Returns how many EMUs in 1 inch.
        /// </summary>
        public const double InchesToEmu = 914400.0;

        /// <summary>
        /// Returns how many EMUs in 1 point.
        /// </summary>
        public const double PointsToEmu = 12700.0;

        /// <summary>
        /// Returns how many EMUs in 1 pc.
        /// </summary>
        public const double PcToEmu = 12.0 * PointsToEmu;

        /// <summary>
        /// Returns how many EMUs in 1 pi.
        /// </summary>
        public const double PiToEmu = 12.0 * PointsToEmu; //Yes, it is the SAME as pc (!)

        TDrawingCoordinate(double aMeasure, TDrawingUnit DrawingUnit)
        {
            switch (DrawingUnit)
            {
                case TDrawingUnit.emu:
                    FEmu = (long) Math.Round(aMeasure);
                    return;
                
                case TDrawingUnit.cm:
                    FEmu = (long)Math.Round(aMeasure * CmToEmu);
                    return;

                case TDrawingUnit.mm:
                    FEmu = (long)Math.Round(aMeasure * CmToEmu / 10.0);
                    return;
                
                case TDrawingUnit.inches:
                    FEmu = (long)Math.Round(aMeasure * InchesToEmu);
                    return;
                
                case TDrawingUnit.pt:
                    FEmu = (long)Math.Round(aMeasure * PointsToEmu);
                    return;
                
                case TDrawingUnit.pc:
                    FEmu = (long)Math.Round(aMeasure * PcToEmu);
                    return;
                
                case TDrawingUnit.pi:
                    FEmu = (long)Math.Round(aMeasure * PiToEmu);
                    return;
                default:
                    break;
            }
            FlxMessages.ThrowException(FlxErr.ErrInternal);
            FEmu = 0;
        }

        /// <summary>
        /// Creates a coordinate in Emus. To use other units, use the "From..." methods of this struct.
        /// </summary>
        /// <param name="aEmu">Value of the coordinate.</param>
        public TDrawingCoordinate(long aEmu)
        {
            FEmu = aEmu;
        }

        #region From
        /// <summary>
        /// Creates a drawing coordinate from a measunement in centimeters.
        /// </summary>
        /// <param name="p">Value in cm.</param>
        /// <returns>The corresponding DrawingCoordinate.</returns>
        public static TDrawingCoordinate FromCm(double p)
        {
            return new TDrawingCoordinate(p, TDrawingUnit.cm);
        }

        /// <summary>
        /// Creates a drawing coordinate from a measunement in milimeters.
        /// </summary>
        /// <param name="p">Value in mm.</param>
        /// <returns>The corresponding DrawingCoordinate.</returns>
        public static TDrawingCoordinate FromMm(double p)
        {
            return new TDrawingCoordinate(p, TDrawingUnit.mm);
        }

        /// <summary>
        /// Creates a drawing coordinate from a measunement in inches.
        /// </summary>
        /// <param name="p">Value in inches.</param>
        /// <returns>The corresponding DrawingCoordinate.</returns>
        public static TDrawingCoordinate FromInches(double p)
        {
            return new TDrawingCoordinate(p, TDrawingUnit.inches);
        }

        /// <summary>
        /// Creates a drawing coordinate from a measunement in points. (1/72 of an inch)
        /// </summary>
        /// <param name="p">Value in points (1/72 of an inch).</param>
        /// <returns>The corresponding DrawingCoordinate.</returns>
        public static TDrawingCoordinate FromPoints(double p)
        {
            return new TDrawingCoordinate(p, TDrawingUnit.pt);
        }

        internal static TDrawingCoordinate FromPixels(double p)
        {
            return new TDrawingCoordinate(p / FlxConsts.PixToPoints, TDrawingUnit.pt);
        }


        /// <summary>
        /// Creates a drawing coordinate from a measunement in Pi Excel units.
        /// </summary>
        /// <param name="p">Value in pi Excel units.</param>
        /// <returns>The corresponding DrawingCoordinate.</returns>
        internal static TDrawingCoordinate FromPi(double p)
        {
            return new TDrawingCoordinate(p, TDrawingUnit.pi);
        }

        internal static TDrawingCoordinate FromPc(double p)
        {
            return new TDrawingCoordinate(p, TDrawingUnit.pc);
        }
        #endregion

        /// <summary>
        /// Value of the coordinate in EMUs (English Metric Units)
        /// </summary>
        public long Emu { get { return FEmu; } }

        /// <summary>
        /// Value of the coordinate in cm 
        /// </summary>
        public double Cm { get { return FEmu / CmToEmu; } }

        /// <summary>
        /// Value of the coordinate in inches
        /// </summary>
        public double Inches { get { return FEmu / InchesToEmu; } }

        /// <summary>
        /// Value of the coordinate in points
        /// </summary>
        public double Points { get { return FEmu / PointsToEmu; } }

        /// <summary>
        /// Value of the coordinate in pixels
        /// </summary>
        public double Pixels { get { return FEmu / PointsToEmu * FlxConsts.PixToPoints; } }

        #region Compare
        /// <summary>
        /// Returns true if this instance has the same data as the object obj.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            if (!(obj is TDrawingCoordinate)) return false;
            return ((TDrawingCoordinate)obj).FEmu == FEmu;
        }

        /// <summary>
        /// Returns the hashcode for this object
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return FEmu.GetHashCode();
        }
        #endregion

        /// <summary></summary>
        public static bool operator ==(TDrawingCoordinate b1, TDrawingCoordinate b2)
        {
            return b1.FEmu == b2.FEmu;
        }

        /// <summary></summary>
        public static bool operator !=(TDrawingCoordinate b1, TDrawingCoordinate b2)
        {
            return !(b1 == b2);
        }



        #region IComparable Members

        /// <summary>
        /// 
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public int CompareTo(object obj)
        {
            if (!(obj is TDrawingCoordinate)) return -1;
            TDrawingCoordinate s2 = (TDrawingCoordinate)obj;
            int r;
            r = FEmu.CompareTo(s2.FEmu); if (r != 0) return r;
            return 0;
        }

        #endregion
    }

    /// <summary>
    /// A point with x and y coordinates.
    /// </summary>
    public struct TDrawingPoint
    {
        private TDrawingCoordinate Fx;
        private TDrawingCoordinate Fy;

        /// <summary>
        /// Creates a new Drawing point.
        /// </summary>
        /// <param name="aX">X Coordinate.</param>
        /// <param name="aY">Y Coordinate.</param>
        public TDrawingPoint(TDrawingCoordinate aX, TDrawingCoordinate aY)
        {
            Fx = aX;
            Fy = aY;
        }


        /// <summary>
        /// X coordinate.
        /// </summary>
        public TDrawingCoordinate X { get { return Fx; } set { Fx = value; } }

        /// <summary>
        /// Y coordinate.
        /// </summary>
        public TDrawingCoordinate Y { get { return Fy; } set { Fy = value; } }


        #region Compare
        /// <summary>
        /// Returns true if this instance has the same data as the object obj.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            if (!(obj is TDrawingPoint)) return false;
            TDrawingPoint o2 = (TDrawingPoint)obj;
            return o2.Fx == Fx && o2.Fy == Fy;
        }

        /// <summary>
        /// Returns the hashcode for this object
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHash(Fx.GetHashCode(), Fy.GetHashCode());
        }
        #endregion

        /// <summary></summary>
        public static bool operator ==(TDrawingPoint b1, TDrawingPoint b2)
        {
            return b1.Fx == b2.Fx && b1.Fy == b2.Fy;
        }

        /// <summary></summary>
        public static bool operator !=(TDrawingPoint b1, TDrawingPoint b2)
        {
            return !(b1 == b2);
        }
    }

    /// <summary>
    /// A rectangle with coordinates used in a drawing.
    /// </summary>
    public struct TDrawingRelativeRect: IComparable
    {
        private double FLeft;
        private double FTop;
        private double FRight;
        private double FBottom;

        /// <summary>
        /// Creates a new instance.
        /// </summary>
        /// <param name="aLeft">Percentage of the left coordinate. Might be negative.</param>
        /// <param name="aTop">Percentage of the top coordinate. Might be negative.</param>
        /// <param name="aRight">Percentage of the right coordinate. Might be negative.</param>
        /// <param name="aBottom">Percentage of the bottom coordinate. Might be negative.</param>
        public TDrawingRelativeRect(double aLeft, double aTop, double aRight, double aBottom)
        {
            FLeft = aLeft;
            FTop = aTop;
            FRight = aRight;
            FBottom = aBottom;
        }

        /// <summary>
        /// Percentage of the left coordinate. Might be negative.
        /// </summary>
        public double Left { get { return FLeft; } }

        /// <summary>
        /// Percentage of the top coordinate. Might be negative.
        /// </summary>
        public double Top { get { return FTop; } }

        /// <summary>
        /// Percentage of the right coordinate. Might be negative.
        /// </summary>
        public double Right { get { return FRight; } }

        /// <summary>
        /// Percentage of the bottom coordinate. Might be negative.
        /// </summary>
        public double Bottom { get { return FBottom; } }

        /// <summary>
        /// Bottom - Top
        /// </summary>
        public double Height { get { return FBottom - FTop; } }

        /// <summary>
        /// Right - Left
        /// </summary>
        public double Width { get { return FRight - FLeft; } }


        /// <summary>
        /// Returns true if both classes contain the same rectangle.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            if (!(obj is TDrawingRelativeRect)) return false;
            TDrawingRelativeRect o2 = (TDrawingRelativeRect)obj;
            if (FTop != o2.FTop) return false;
            if (FLeft != o2.FLeft) return false;
            if (FBottom != o2.FBottom) return false;
            if (FRight != o2.FRight) return false;
            return true;
        }

        /// <summary>
        /// Returns the hashcode for this object.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHash(FTop.GetHashCode(), FLeft.GetHashCode(), FBottom.GetHashCode(), FRight.GetHashCode());
        }

        /// <summary></summary>
        public static bool operator ==(TDrawingRelativeRect b1, TDrawingRelativeRect b2)
        {
            return b1.Equals(b2);
        }

        /// <summary></summary>
        public static bool operator !=(TDrawingRelativeRect b1, TDrawingRelativeRect b2)
        {
            return !(b1 == b2);
        }


        /// <summary></summary>
        public int CompareTo(object obj)
        {
            if (!(obj is TDrawingRelativeRect)) return -1;
            TDrawingRelativeRect s2 = (TDrawingRelativeRect)obj;
            int r;
            r = FLeft.CompareTo(s2.FLeft); if (r != 0) return r;
            r = FTop.CompareTo(s2.FTop); if (r != 0) return r;
            r = FRight.CompareTo(s2.FRight); if (r != 0) return r;
            r = FBottom.CompareTo(s2.FBottom); if (r != 0) return r;
            return 0;
        }
    }

    /// <summary>
    /// How to position two rectangles relative to each other.
    /// </summary>
    public enum TDrawingRectAlign
    {
        /// <summary>
        /// Align at the top left.
        /// </summary>
        TopLeft,

        /// <summary>
        /// Align at the top.
        /// </summary>
        Top,

        /// <summary>
        /// Align at the top right.
        /// </summary>
        TopRight,

        /// <summary>
        /// Align at the left.
        /// </summary>
        Left,

        /// <summary>
        /// Align at the center.
        /// </summary>
        Center,

        /// <summary>
        /// Align at the right.
        /// </summary>
        Right,

        /// <summary>
        /// Align at the bottom left.
        /// </summary>
        BottomLeft,

        /// <summary>
        /// Align at the bottom.
        /// </summary>
        Bottom,

        /// <summary>
        /// Align at the bottom right.
        /// </summary>
        BottomRight
    }

    /// <summary>
    /// How an image will be flipped when filling a pattern.
    /// </summary>
    public enum TFlipMode
    {
        /// <summary>
        /// Image will not be flipped. 
        /// </summary>
        None,

        /// <summary>
        /// Tiles are flipped horizontally.
        /// </summary>
        X,

        /// <summary>
        /// Tiles are flipped vertically.
        /// </summary>
        Y,

        /// <summary>
        /// Tiles are flipped horizontally and vertically.
        /// </summary>
        XY
    }


    /// <summary>
    /// Represents one of the points in a Gradient definition for a drawing (autoshapes, charts, etc). Note that Excel cells
    /// use a different Gradient definition: <see cref="TGradientStop"/>
    /// </summary>
    public struct TDrawingGradientStop : IComparable
    {
        private double FPosition;
        private TDrawingColor FColor;

        /// <summary>
        /// This value must be between 0 and 1, and represents the position in the gradient where the <see cref="Color"/> in this structure is pure.
        /// </summary>
        public double Position { get { return FPosition; } set { if (value < 0 || value > 1) FlxMessages.ThrowException(FlxErr.ErrInvalidValue, "Position", value, 0, 1); FPosition = value; } }

        /// <summary>
        /// Color for this definition.
        /// </summary>
        public TDrawingColor Color { get { return FColor; } set { FColor = value; } }

        /// <summary>
        /// Creates a new Gradient stop.
        /// </summary>
        /// <param name="aPosition">Position for the stop.</param>
        /// <param name="aColor">Color for the stop.</param>
        public TDrawingGradientStop(double aPosition, TDrawingColor aColor)
        {
            FPosition = 0;
            FColor = aColor;// to compile

            Position = aPosition; //to set the real values
            Color = aColor;
        }

        #region IComparable Members

        /// <summary>
        /// Compares 2 instances of this struct.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public int CompareTo(object obj)
        {
            if (!(obj is TDrawingGradientStop)) return -1;
            TDrawingGradientStop o2 = (TDrawingGradientStop)obj;

            int Result = Position.CompareTo(o2.Position);
            if (Result != 0) return Result;
            Result = Color.CompareTo(o2.Color);
            if (Result != 0) return Result;

            return 0;
        }

        /// <summary>
        /// Returns if this struct has the same values as other one.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            return CompareTo(obj) == 0;
        }

        /// <summary>
        /// Returns the hashcode for this struct.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHash(Position.GetHashCode(), Color.GetHashCode());
        }

        /// <summary>
        /// Returns true if both gradient stops are equal.
        /// </summary>
        /// <param name="o1">First stop to compare.</param>
        /// <param name="o2">Second stop to compare.</param>
        /// <returns></returns>
        public static bool operator ==(TDrawingGradientStop o1, TDrawingGradientStop o2)
        {
            return o1.Equals(o2);
        }

        /// <summary>
        /// Returns true if both gradient stops are different.
        /// </summary>
        /// <param name="o1">First stop to compare.</param>
        /// <param name="o2">Second stop to compare.</param>
        /// <returns></returns>
        public static bool operator !=(TDrawingGradientStop o1, TDrawingGradientStop o2)
        {
            return !o1.Equals(o2);
        }

        /// <summary>
        /// Returns true if o1 is bigger than o2.
        /// </summary>
        /// <param name="o1"></param>
        /// <param name="o2"></param>
        /// <returns></returns>
        public static bool operator >(TDrawingGradientStop o1, TDrawingGradientStop o2)
        {
            return o1.CompareTo(o2) > 0;
        }

        /// <summary>
        /// Returns true if o1 is less than o2.
        /// </summary>
        /// <param name="o1"></param>
        /// <param name="o2"></param>
        /// <returns></returns>
        public static bool operator <(TDrawingGradientStop o1, TDrawingGradientStop o2)
        {
            return o1.CompareTo(o2) < 0;
        }


        #endregion
    }


    /// <summary>
    /// Different types of patterns for filling an Excel 2007 object. This is different from the patterns to fill a cell.
    /// </summary>
    public enum TDrawingPattern
    {
        /// <summary>
        /// pct5
        /// </summary>
        pct5,

        /// <summary>
        /// pct10
        /// </summary>
        pct10,

        /// <summary>
        /// pct20
        /// </summary>
        pct20,

        /// <summary>
        /// pct25
        /// </summary>
        pct25,

        /// <summary>
        /// pct30
        /// </summary>
        pct30,

        /// <summary>
        /// pct40
        /// </summary>
        pct40,

        /// <summary>
        /// pct50
        /// </summary>
        pct50,

        /// <summary>
        /// pct60
        /// </summary>
        pct60,

        /// <summary>
        /// pct70
        /// </summary>
        pct70,

        /// <summary>
        /// pct75
        /// </summary>
        pct75,

        /// <summary>
        /// pct80
        /// </summary>
        pct80,

        /// <summary>
        /// pct90
        /// </summary>
        pct90,

        /// <summary>
        /// horz
        /// </summary>
        horz,

        /// <summary>
        /// vert
        /// </summary>
        vert,

        /// <summary>
        /// ltHorz
        /// </summary>
        ltHorz,

        /// <summary>
        /// ltVert
        /// </summary>
        ltVert,

        /// <summary>
        /// dkHorz
        /// </summary>
        dkHorz,

        /// <summary>
        /// dkVert
        /// </summary>
        dkVert,

        /// <summary>
        /// narHorz
        /// </summary>
        narHorz,

        /// <summary>
        /// narVert
        /// </summary>
        narVert,

        /// <summary>
        /// dashHorz
        /// </summary>
        dashHorz,

        /// <summary>
        /// dashVert
        /// </summary>
        dashVert,

        /// <summary>
        /// cross
        /// </summary>
        cross,

        /// <summary>
        /// dnDiag
        /// </summary>
        dnDiag,

        /// <summary>
        /// upDiag
        /// </summary>
        upDiag,

        /// <summary>
        /// ltDnDiag
        /// </summary>
        ltDnDiag,

        /// <summary>
        /// ltUpDiag
        /// </summary>
        ltUpDiag,

        /// <summary>
        /// dkDnDiag
        /// </summary>
        dkDnDiag,

        /// <summary>
        /// dkUpDiag
        /// </summary>
        dkUpDiag,

        /// <summary>
        /// wdDnDiag
        /// </summary>
        wdDnDiag,

        /// <summary>
        /// wdUpDiag
        /// </summary>
        wdUpDiag,

        /// <summary>
        /// dashDnDiag
        /// </summary>
        dashDnDiag,

        /// <summary>
        /// dashUpDiag
        /// </summary>
        dashUpDiag,

        /// <summary>
        /// diagCross
        /// </summary>
        diagCross,

        /// <summary>
        /// smCheck
        /// </summary>
        smCheck,

        /// <summary>
        /// lgCheck
        /// </summary>
        lgCheck,

        /// <summary>
        /// smGrid
        /// </summary>
        smGrid,

        /// <summary>
        /// lgGrid
        /// </summary>
        lgGrid,

        /// <summary>
        /// dotGrid
        /// </summary>
        dotGrid,

        /// <summary>
        /// smConfetti
        /// </summary>
        smConfetti,

        /// <summary>
        /// lgConfetti
        /// </summary>
        lgConfetti,

        /// <summary>
        /// horzBrick
        /// </summary>
        horzBrick,

        /// <summary>
        /// diagBrick
        /// </summary>
        diagBrick,

        /// <summary>
        /// solidDmnd
        /// </summary>
        solidDmnd,

        /// <summary>
        /// openDmnd
        /// </summary>
        openDmnd,

        /// <summary>
        /// dotDmnd
        /// </summary>
        dotDmnd,

        /// <summary>
        /// plaid
        /// </summary>
        plaid,

        /// <summary>
        /// sphere
        /// </summary>
        sphere,

        /// <summary>
        /// weave
        /// </summary>
        weave,

        /// <summary>
        /// divot
        /// </summary>
        divot,

        /// <summary>
        /// shingle
        /// </summary>
        shingle,

        /// <summary>
        /// wave
        /// </summary>
        wave,

        /// <summary>
        /// trellis
        /// </summary>
        trellis,

        /// <summary>
        /// zigZag
        /// </summary>
        zigZag
    }

    internal struct TDrawingPresetGeom
    {
        internal static Dictionary<TShapeType, string> FromShapeType = CreateShapeTypeToString();
        internal static Dictionary<string, TShapeType> FromString = CreateStringToShapeType();

        private static Dictionary<string, TShapeType> CreateStringToShapeType()
        {
            Dictionary<string, TShapeType> Result = new Dictionary<string, TShapeType>();
            foreach (TShapeType st in TCompactFramework.EnumGetValues(typeof(TShapeType)))
            {
                if (st == TShapeType.NotPrimitive) continue;
                if (FromShapeType.ContainsKey(st)) Result.Add(FromShapeType[st], st);
            }
            return Result;
        }

        private static Dictionary<TShapeType, string> CreateShapeTypeToString()
        {
            Dictionary<TShapeType, string> Result = new Dictionary<TShapeType, string>();
            Result.Add(TShapeType.Line, "line");
            Result.Add(TShapeType.LineInv, "lineInv");
            Result.Add(TShapeType.IsocelesTriangle, "triangle");
            Result.Add(TShapeType.RightTriangle, "rtTriangle");
            Result.Add(TShapeType.Rectangle, "rect");
            Result.Add(TShapeType.Diamond, "diamond");
            Result.Add(TShapeType.Parallelogram, "parallelogram");
            Result.Add(TShapeType.Trapezoid2007, "trapezoid"); //Trapezoid is different in Excel 2007. It is inverted upside down.
            Result.Add(TShapeType.NonIsoscelesTrapezoid, "nonIsoscelesTrapezoid");
            Result.Add(TShapeType.Pentagon, "pentagon");
            Result.Add(TShapeType.Hexagon, "hexagon");
            Result.Add(TShapeType.Heptagon, "heptagon");
            Result.Add(TShapeType.Octagon, "octagon");
            Result.Add(TShapeType.Decagon, "decagon");
            Result.Add(TShapeType.Dodecagon, "dodecagon");
            Result.Add(TShapeType.Seal4, "star4");
            Result.Add(TShapeType.Star, "star5");
            Result.Add(TShapeType.Star6, "star6");
            Result.Add(TShapeType.Star7, "star7");
            Result.Add(TShapeType.Seal8, "star8");
            Result.Add(TShapeType.Star10, "star10");
            Result.Add(TShapeType.Star12, "star12");
            Result.Add(TShapeType.Seal16, "star16");
            Result.Add(TShapeType.Seal24, "star24");
            Result.Add(TShapeType.Seal32, "star32");
            Result.Add(TShapeType.RoundRectangle, "roundRect");
            Result.Add(TShapeType.Round1Rect, "round1Rect");
            Result.Add(TShapeType.Round2SameRect, "round2SameRect");
            Result.Add(TShapeType.Round2DiagRect, "round2DiagRect");
            Result.Add(TShapeType.SnipRoundRect, "snipRoundRect");
            Result.Add(TShapeType.Snip1Rect, "snip1Rect");
            Result.Add(TShapeType.Snip2SameRect, "snip2SameRect");
            Result.Add(TShapeType.Snip2DiagRect, "snip2DiagRect");
            Result.Add(TShapeType.Plaque, "plaque");
            Result.Add(TShapeType.Ellipse, "ellipse");
            Result.Add(TShapeType.Teardrop, "teardrop");
            Result.Add(TShapeType.HomePlate, "homePlate");
            Result.Add(TShapeType.Chevron, "chevron");
            Result.Add(TShapeType.PieWedge, "pieWedge");
            Result.Add(TShapeType.Pie, "pie");
            Result.Add(TShapeType.BlockArc, "blockArc");
            Result.Add(TShapeType.Donut, "donut");
            Result.Add(TShapeType.NoSmoking, "noSmoking");
            Result.Add(TShapeType.RightArrow, "rightArrow");
            Result.Add(TShapeType.LeftArrow, "leftArrow");
            Result.Add(TShapeType.UpArrow, "upArrow");
            Result.Add(TShapeType.DownArrow, "downArrow");
            Result.Add(TShapeType.StripedRightArrow, "stripedRightArrow");
            Result.Add(TShapeType.NotchedRightArrow, "notchedRightArrow");
            Result.Add(TShapeType.BentUpArrow, "bentUpArrow");
            Result.Add(TShapeType.LeftRightArrow, "leftRightArrow");
            Result.Add(TShapeType.UpDownArrow, "upDownArrow");
            Result.Add(TShapeType.LeftUpArrow, "leftUpArrow");
            Result.Add(TShapeType.LeftRightUpArrow, "leftRightUpArrow");
            Result.Add(TShapeType.QuadArrow, "quadArrow");
            Result.Add(TShapeType.LeftArrowCallout, "leftArrowCallout");
            Result.Add(TShapeType.RightArrowCallout, "rightArrowCallout");
            Result.Add(TShapeType.UpArrowCallout, "upArrowCallout");
            Result.Add(TShapeType.DownArrowCallout, "downArrowCallout");
            Result.Add(TShapeType.LeftRightArrowCallout, "leftRightArrowCallout");
            Result.Add(TShapeType.UpDownArrowCallout, "upDownArrowCallout");
            Result.Add(TShapeType.QuadArrowCallout, "quadArrowCallout");
            Result.Add(TShapeType.BentArrow, "bentArrow");
            Result.Add(TShapeType.UturnArrow, "uturnArrow");
            Result.Add(TShapeType.CircularArrow, "circularArrow");
            Result.Add(TShapeType.LeftCircularArrow, "leftCircularArrow");
            Result.Add(TShapeType.LeftRightCircularArrow, "leftRightCircularArrow");
            Result.Add(TShapeType.CurvedRightArrow, "curvedRightArrow");
            Result.Add(TShapeType.CurvedLeftArrow, "curvedLeftArrow");
            Result.Add(TShapeType.CurvedUpArrow, "curvedUpArrow");
            Result.Add(TShapeType.CurvedDownArrow, "curvedDownArrow");
            Result.Add(TShapeType.SwooshArrow, "swooshArrow");
            Result.Add(TShapeType.Cube, "cube");
            Result.Add(TShapeType.Can, "can");
            Result.Add(TShapeType.LightningBolt, "lightningBolt");
            Result.Add(TShapeType.Heart, "heart");
            Result.Add(TShapeType.Sun, "sun");
            Result.Add(TShapeType.Moon, "moon");
            Result.Add(TShapeType.SmileyFace, "smileyFace");
            Result.Add(TShapeType.IrregularSeal1, "irregularSeal1");
            Result.Add(TShapeType.IrregularSeal2, "irregularSeal2");
            Result.Add(TShapeType.FoldedCorner, "foldedCorner");
            Result.Add(TShapeType.Bevel, "bevel");
            Result.Add(TShapeType.Frame, "frame");
            Result.Add(TShapeType.HalfFrame, "halfFrame");
            Result.Add(TShapeType.Corner, "corner");
            Result.Add(TShapeType.DiagStripe, "diagStripe");
            Result.Add(TShapeType.Chord, "chord");
            Result.Add(TShapeType.Arc, "arc");
            Result.Add(TShapeType.LeftBracket, "leftBracket");
            Result.Add(TShapeType.RightBracket, "rightBracket");
            Result.Add(TShapeType.LeftBrace, "leftBrace");
            Result.Add(TShapeType.RightBrace, "rightBrace");
            Result.Add(TShapeType.BracketPair, "bracketPair");
            Result.Add(TShapeType.BracePair, "bracePair");
            Result.Add(TShapeType.StraightConnector1, "straightConnector1");
            Result.Add(TShapeType.BentConnector2, "bentConnector2");
            Result.Add(TShapeType.BentConnector3, "bentConnector3");
            Result.Add(TShapeType.BentConnector4, "bentConnector4");
            Result.Add(TShapeType.BentConnector5, "bentConnector5");
            Result.Add(TShapeType.CurvedConnector2, "curvedConnector2");
            Result.Add(TShapeType.CurvedConnector3, "curvedConnector3");
            Result.Add(TShapeType.CurvedConnector4, "curvedConnector4");
            Result.Add(TShapeType.CurvedConnector5, "curvedConnector5");
            Result.Add(TShapeType.Callout1, "callout1");
            Result.Add(TShapeType.Callout2, "callout2");
            Result.Add(TShapeType.Callout3, "callout3");
            Result.Add(TShapeType.AccentCallout1, "accentCallout1");
            Result.Add(TShapeType.AccentCallout2, "accentCallout2");
            Result.Add(TShapeType.AccentCallout3, "accentCallout3");
            Result.Add(TShapeType.BorderCallout1, "borderCallout1");
            Result.Add(TShapeType.BorderCallout2, "borderCallout2");
            Result.Add(TShapeType.BorderCallout3, "borderCallout3");
            Result.Add(TShapeType.AccentBorderCallout1, "accentBorderCallout1");
            Result.Add(TShapeType.AccentBorderCallout2, "accentBorderCallout2");
            Result.Add(TShapeType.AccentBorderCallout3, "accentBorderCallout3");
            Result.Add(TShapeType.WedgeRectCallout, "wedgeRectCallout");
            Result.Add(TShapeType.WedgeRRectCallout, "wedgeRoundRectCallout");
            Result.Add(TShapeType.WedgeEllipseCallout, "wedgeEllipseCallout");
            Result.Add(TShapeType.CloudCallout, "cloudCallout");
            Result.Add(TShapeType.Cloud, "cloud");
            Result.Add(TShapeType.Ribbon, "ribbon");
            Result.Add(TShapeType.Ribbon2, "ribbon2");
            Result.Add(TShapeType.EllipseRibbon, "ellipseRibbon");
            Result.Add(TShapeType.EllipseRibbon2, "ellipseRibbon2");
            Result.Add(TShapeType.LeftRightRibbon, "leftRightRibbon");
            Result.Add(TShapeType.VerticalScroll, "verticalScroll");
            Result.Add(TShapeType.HorizontalScroll, "horizontalScroll");
            Result.Add(TShapeType.Wave, "wave");
            Result.Add(TShapeType.DoubleWave, "doubleWave");
            Result.Add(TShapeType.Plus, "plus");
            Result.Add(TShapeType.FlowChartProcess, "flowChartProcess");
            Result.Add(TShapeType.FlowChartDecision, "flowChartDecision");
            Result.Add(TShapeType.FlowChartInputOutput, "flowChartInputOutput");
            Result.Add(TShapeType.FlowChartPredefinedProcess, "flowChartPredefinedProcess");
            Result.Add(TShapeType.FlowChartInternalStorage, "flowChartInternalStorage");
            Result.Add(TShapeType.FlowChartDocument, "flowChartDocument");
            Result.Add(TShapeType.FlowChartMultidocument, "flowChartMultidocument");
            Result.Add(TShapeType.FlowChartTerminator, "flowChartTerminator");
            Result.Add(TShapeType.FlowChartPreparation, "flowChartPreparation");
            Result.Add(TShapeType.FlowChartManualInput, "flowChartManualInput");
            Result.Add(TShapeType.FlowChartManualOperation, "flowChartManualOperation");
            Result.Add(TShapeType.FlowChartConnector, "flowChartConnector");
            Result.Add(TShapeType.FlowChartPunchedCard, "flowChartPunchedCard");
            Result.Add(TShapeType.FlowChartPunchedTape, "flowChartPunchedTape");
            Result.Add(TShapeType.FlowChartSummingJunction, "flowChartSummingJunction");
            Result.Add(TShapeType.FlowChartOr, "flowChartOr");
            Result.Add(TShapeType.FlowChartCollate, "flowChartCollate");
            Result.Add(TShapeType.FlowChartSort, "flowChartSort");
            Result.Add(TShapeType.FlowChartExtract, "flowChartExtract");
            Result.Add(TShapeType.FlowChartMerge, "flowChartMerge");
            Result.Add(TShapeType.FlowChartOfflineStorage, "flowChartOfflineStorage");
            Result.Add(TShapeType.FlowChartOnlineStorage, "flowChartOnlineStorage");
            Result.Add(TShapeType.FlowChartMagneticTape, "flowChartMagneticTape");
            Result.Add(TShapeType.FlowChartMagneticDisk, "flowChartMagneticDisk");
            Result.Add(TShapeType.FlowChartMagneticDrum, "flowChartMagneticDrum");
            Result.Add(TShapeType.FlowChartDisplay, "flowChartDisplay");
            Result.Add(TShapeType.FlowChartDelay, "flowChartDelay");
            Result.Add(TShapeType.FlowChartAlternateProcess, "flowChartAlternateProcess");
            Result.Add(TShapeType.FlowChartOffpageConnector, "flowChartOffpageConnector");
            Result.Add(TShapeType.ActionButtonBlank, "actionButtonBlank");
            Result.Add(TShapeType.ActionButtonHome, "actionButtonHome");
            Result.Add(TShapeType.ActionButtonHelp, "actionButtonHelp");
            Result.Add(TShapeType.ActionButtonInformation, "actionButtonInformation");
            Result.Add(TShapeType.ActionButtonForwardNext, "actionButtonForwardNext");
            Result.Add(TShapeType.ActionButtonBackPrevious, "actionButtonBackPrevious");
            Result.Add(TShapeType.ActionButtonEnd, "actionButtonEnd");
            Result.Add(TShapeType.ActionButtonBeginning, "actionButtonBeginning");
            Result.Add(TShapeType.ActionButtonReturn, "actionButtonReturn");
            Result.Add(TShapeType.ActionButtonDocument, "actionButtonDocument");
            Result.Add(TShapeType.ActionButtonSound, "actionButtonSound");
            Result.Add(TShapeType.ActionButtonMovie, "actionButtonMovie");
            Result.Add(TShapeType.Gear6, "gear6");
            Result.Add(TShapeType.Gear9, "gear9");
            Result.Add(TShapeType.Funnel, "funnel");
            Result.Add(TShapeType.MathPlus, "mathPlus");
            Result.Add(TShapeType.MathMinus, "mathMinus");
            Result.Add(TShapeType.MathMultiply, "mathMultiply");
            Result.Add(TShapeType.MathDivide, "mathDivide");
            Result.Add(TShapeType.MathEqual, "mathEqual");
            Result.Add(TShapeType.MathNotEqual, "mathNotEqual");
            Result.Add(TShapeType.CornerTabs, "cornerTabs");
            Result.Add(TShapeType.SquareTabs, "squareTabs");
            Result.Add(TShapeType.PlaqueTabs, "plaqueTabs");
            Result.Add(TShapeType.ChartX, "chartX");
            Result.Add(TShapeType.ChartStar, "chartStar");
            Result.Add(TShapeType.ChartPlus, "chartPlus");
            return Result;
        }
    }

}
